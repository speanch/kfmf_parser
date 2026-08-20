"""
Парсер каталога фасадов kfmf-skat.ru
Коллекция фрезеровок -> раздел (Premium, Classic, ...) -> модель -> Толщина фасада
Результат: Excel-файл.
"""
import re
import sys
import time
import urllib.parse
import os

import requests
from bs4 import BeautifulSoup

BASE = "https://kfmf-skat.ru"
COLLECTION_URL = BASE + "/catalog/mebelnye-fasady/kollektsiya-frezerovok/"
PREFIX = "/catalog/mebelnye-fasady/kollektsiya-frezerovok/"

UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
)
HEADERS = {
    "User-Agent": UA,
    "Accept": (
        "text/html,application/xhtml+xml,application/xml;q=0.9,"
        "image/avif,image/webp,*/*;q=0.8"
    ),
    "Accept-Language": "ru-RU,ru;q=0.9,en;q=0.8",
    "Accept-Encoding": "gzip, deflate, br",
    "Sec-CH-UA": '"Not_A Brand";v="8", "Chromium";v="120", "Google Chrome";v="120"',
    "Upgrade-Insecure-Requests": "1",
    "Connection": "keep-alive",
}

SLEEP = 0.4
session = requests.Session()
session.headers.update(HEADERS)

# Разделы, которые парсить не нужно
SKIP_SECTIONS = {"LITTLE PARADISE"}


def log(msg=""):
    print(msg, flush=True)


def restart_session():
    """Пересоздаём сессию и разогреваем заново (при пустых ответах)."""
    global session
    session = requests.Session()
    session.headers.update(HEADERS)
    fetch(BASE + "/")
    time.sleep(0.5)


def fetch(url, retries=4):
    """GET, возвращает html-строку или None."""
    for i in range(retries):
        try:
            r = session.get(url, timeout=30)
            r.raise_for_status()
            return r.text
        except Exception as e:
            log(f"  [fetch] ошибка {url}: {e} (попытка {i + 1}/{retries})")
            time.sleep(2 * (i + 1))
    return None


def warmup():
    """
    Разогрев сессии: сначала заходим на главную, потом на коллекцию.
    Без этого сайт отдаёт страницы-списки разделов пустыми.
    """
    log("[1/4] Разогрев сессии...")
    for url in [BASE + "/"]:
        log(f"  GET {url}")
        html = fetch(url)
        if not html:
            log("  Ошибка: главная не получена")
            sys.exit(1)
    log(f"  GET {COLLECTION_URL}")
    html = fetch(COLLECTION_URL)
    if not html:
        log("  Ошибка: коллекция не получена")
        sys.exit(1)
    log(f"  Получено байт: {len(html)}")
    return html


def discover_sections(collection_html):
    """Ищем разделы коллекции в html коллекции."""
    log("[2/4] Поиск разделов коллекции...")
    soup = BeautifulSoup(collection_html, "html.parser")
    sections = {}
    seen = set()
    for a in soup.find_all("a", href=True):
        href = a["href"]
        key = href.rstrip("/")
        if not key.startswith(PREFIX):
            continue
        if key == COLLECTION_URL.rstrip("/"):
            continue
        if key in seen or key.endswith(".html"):
            continue
        seen.add(key)
        name = a.get_text(strip=True) or key.rsplit("/", 1)[-1]
        sections[href] = name
    log(f"  Найдено разделов: {len(sections)}")
    for name in sections.values():
        log(f"    - {name}")
    return sections


def collect_section(section_path):
    """
    Обходит раздел (включая вложенные подкаталоги, напр. 3D -> 3d-fasady)
    и возвращает dict {url_модели: имя_модели}. Имя берём из карточки в списке.
    """
    root = section_path.rstrip("/")
    queue = [root]
    seen_dirs = set()
    queued = set()
    links = {}  # url -> name
    empty_strikes = 0

    while queue:
        d = queue.pop(0)
        if d in seen_dirs:
            continue
        seen_dirs.add(d)

        page = 1
        while True:
            dir_url = f"{BASE + d}/" if page == 1 else f"{BASE + d}/?PAGEN_1={page}"
            html = fetch(dir_url)
            if html is None:
                log(f"  [разд {root.split('/')[-1]}] {d} стр.{page}: не получено, стоп")
                break
            soup = BeautifulSoup(html, "html.parser")

            n_before = len(links)
            # модели на странице
            for a in soup.select("a.product-item[href]"):
                href = a["href"]
                if not (href.startswith(PREFIX) and href.endswith(".html")):
                    continue
                full = urllib.parse.urljoin(BASE, href)
                nm = a.select_one(".product-item__name")
                links.setdefault(full, nm.get_text(strip=True) if nm else None)

            # вложенные подкаталоги: ссылки, заканчивающиеся на '/', строго внутри раздела
            sub_before = len(queued)
            root_prefix = root + "/"
            for a in soup.find_all("a", href=True):
                h = a["href"]
                if not (h.startswith(root_prefix) and h.endswith("/")):
                    continue
                sub = h.rstrip("/")
                if sub not in queued and sub != root:
                    queue.append(sub)
                    queued.add(sub)
            new_sub = len(queued) - sub_before

            new = len(links) - n_before
            pagen = [int(m) for m in re.findall(r"PAGEN_1=(\d+)", html)]
            last = max(pagen) if pagen else page
            dname = d.split("/")[-1] or d.split("/")[-2]
            log(
                f"  [разд {root.split('/')[-1]}] {dname} стр.{page}: "
                f"новых моделей {new}, подкат. {new_sub}, всего {len(links)}, "
                f"страниц {last}"
            )

            # защита от пустой выдачи: нет ни моделей, ни подкаталогов, ни пагинации
            if not pagen and new == 0 and new_sub == 0:
                empty_strikes += 1
                if empty_strikes >= 3:
                    log(f"  пустой ответ {d} x3, пропуск")
                    break
                log(f"  пустой ответ {d} ({empty_strikes}/3), перезапуск сессии...")
                restart_session()
                continue  # повторяем ту же страницу

            if page >= last:
                break
            page += 1
            time.sleep(SLEEP)

    return links


def parse_model(url, fallback_name=None):
    """Тянем модель: имя, толщина фасада, габариты."""
    html = fetch(url)
    if html is None:
        return None
    soup = BeautifulSoup(html, "html.parser")

    name = fallback_name
    if not name:
        h2 = soup.find("h2", class_="section-heading")
        cand = h2.get_text(strip=True) if h2 else ""
        if cand and cand not in ("ФАСАДЫ ИЗ ЭТОЙ КОЛЛЕКЦИИ",):
            name = cand
        else:
            name = url.rsplit("/", 1)[-1].replace(".html", "")

    thickness = []
    height, width = [], []
    for item in soup.select("div.product-card__data-item"):
        b = item.find("b")
        if not b:
            continue
        label = b.get_text(strip=True)
        full = item.get_text(" ", strip=True)
        value = full[len(label):].strip()
        if label == "Толщина фасада" and value:
            thickness.append(value)
        elif label in ("Высота", "Ширина") and value:
            (height if label == "Высота" else width).append(value)

    return {
        "model": name,
        "thickness": "; ".join(dict.fromkeys(thickness)),
        "height": "; ".join(dict.fromkeys(height)),
        "width": "; ".join(dict.fromkeys(width)),
        "url": url,
    }


def save_model_images(url, section_name, model_name, images_root):
    """
    Качает главные фото модели (там, где толщина указана только на картинке)
    в папку images_root/<Раздел>/<Модель>__NN.<ext>.
    Возвращает список сохранённых путей.
    """
    html = fetch(url)
    if html is None:
        return []
    soup = BeautifulSoup(html, "html.parser")

    srcs = []
    for img in soup.select(".product-card__image img"):
        s = img.get("src")
        if s:
            srcs.append(urllib.parse.urljoin(BASE, s))
    # fallback: любой слайд из главного слайдера
    if not srcs:
        for img in soup.select(".product-card__main-slider img"):
            s = img.get("src")
            if s:
                srcs.append(urllib.parse.urljoin(BASE, s))

    seen, uniq = set(), []
    for s in srcs:
        if s not in seen:
            seen.add(s)
            uniq.append(s)
    if not uniq:
        log(f"    [img] {model_name}: картинок в карточке не найдено")
        return []

    safe = "".join(c for c in model_name if c not in '\\/:*?"<>|').strip() or "model"
    folder = os.path.join(images_root, section_name)
    os.makedirs(folder, exist_ok=True)

    saved = []
    for i, s in enumerate(uniq, 1):
        ext = os.path.splitext(urllib.parse.urlparse(s).path)[1] or ".jpg"
        fp = os.path.join(folder, f"{safe}__{i:02d}{ext}")
        if os.path.isfile(fp):
            saved.append(fp)
            continue
        try:
            r = session.get(s, timeout=30)
            r.raise_for_status()
            with open(fp, "wb") as f:
                f.write(r.content)
            saved.append(fp)
        except Exception as e:
            log(f"    [img] ошибка {s}: {e}")
        time.sleep(0.2)
    return saved


def main():
    collection_html = warmup()

    sections = discover_sections(collection_html)
    if not sections:
        log("Ошибка: разделы не найдены")
        sys.exit(1)

    log("[3/4] Сбор моделей, толщины и картинок по разделам...")
    rows = []
    EXCEL = "/home/sereg/personal-projects/kfmf_parser/kfmf_facady_tolshina.xlsx"
    IMAGES_ROOT = os.path.splitext(EXCEL)[0] + "_images"  # рядом с Excel
    for section_path, section_name in sections.items():
        if section_name in SKIP_SECTIONS:
            log(f"\n== Раздел: {section_name} (пропущен, не нужен) ==")
            continue
        log(f"\n== Раздел: {section_name} ==")
        name_map = collect_section(section_path)
        for model_url, model_name in name_map.items():
            rec = parse_model(model_url, fallback_name=model_name)
            if not rec:
                continue
            pics = []
            if not rec["thickness"]:
                pics = save_model_images(
                    model_url, section_name, rec["model"], IMAGES_ROOT
                )
                log(
                    f"    {rec['model']}: толщина только на картинке -> сохранено "
                    f"{len(pics)} фото"
                )
            else:
                log(f"    {rec['model']}: {rec['thickness']}")
            rows.append(
                {
                    "Коллекция": "Коллекция фрезеровок",
                    "Раздел": section_name,
                    "Модель": rec["model"],
                    "Толщина фасада": rec["thickness"],
                    "Высота (мм)": rec["height"],
                    "Ширина (мм)": rec["width"],
                    "Файлы картинок": "; ".join(pics),
                    "Ссылка": rec["url"],
                }
            )
            time.sleep(SLEEP)

    log(f"\n[4/4] Строк собрано: {len(rows)}")
    import pandas as pd

    pd.DataFrame(rows).to_excel(EXCEL, index=False)
    log(f"Готово. Файл: {EXCEL}")
    log(f"Картинки: {IMAGES_ROOT}")


if __name__ == "__main__":
    main()

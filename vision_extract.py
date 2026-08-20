"""
Вижн-проход: достаёт толщину фасада с картинок (GRAFIKA/3D), где она не указана
текстом. Толщина на таких картинках нарисована таблицей характеристик. Фото
отправляются в Qwen3.7 Flash (vision) через OpenRouter; значение 'Толщина … N мм'
записывается в Excel.

Устойчивость:
- прогресс хранится в vision_progress.csv (по пути картинки), Excel пишется один
  раз в конце — прерывание не портит файл, повтор докачивает из CSV;
- reasoning отключён (иначе уходит весь бюджет токенов, content пустой);
- на 429 длинные паузы и повторы; на каждую картинку до RETRIES_PER_IMAGE попыток.
"""
import base64
import csv
import io
import json
import os
import re
import sys
import time
import urllib.request
import urllib.error

import pandas as pd
from PIL import Image

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL = os.path.join(BASE_DIR, "kfmf_facady_tolshina.xlsx")
PROGRESS_CSV = os.path.join(BASE_DIR, "vision_progress.csv")
MODEL = "qwen/qwen3.7-flash"
API_URL = "https://openrouter.ai/api/v1/chat/completions"

PROMPT = (
    "Внимательно посмотри на изображение. Это образец мебельного фасада с "
    "таблицей характеристик. Выпиши ТОЧНО весь текст и цифры, которые видишь: "
    "подписи таблицы и значения (толщина, мм, сечение кромки, глубина "
    "фрезеровки, и т.п.), по одному на строку. Если текста нет вообще - ответь "
    "одним словом НЕТТЕКСТА."
)
SLEEP = 2.0
MAX_ATTEMPTS = 6
RETRIES_PER_IMAGE = 3


def api_key():
    key = os.getenv("OPENROUTER_API_KEY")
    if key:
        return key
    env_path = os.path.join(BASE_DIR, "..", "t_bot", ".env")
    if os.path.isfile(env_path):
        with open(env_path, encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if line.startswith("OPENROUTER_API_KEY="):
                    return line.split("=", 1)[1].strip()
    print("Не найден OPENROUTER_API_KEY", file=sys.stderr)
    sys.exit(1)


def mime_for(path):
    return {
        ".png": "image/png",
        ".jpg": "image/jpeg",
        ".jpeg": "image/jpeg",
        ".webp": "image/webp",
        ".gif": "image/gif",
    }.get(os.path.splitext(path)[1].lower(), "application/octet-stream")


def prepare_bytes(image_path):
    """
    Читает изображение. Если есть прозрачность (текст белый на прозрачном фоне) -
    кладёт на чёрный фон, чтобы текст стал виден. Возвращает PNG-байты.
    """
    im = Image.open(image_path).convert("RGBA")
    if im.mode == "RGBA" and im.getextrema()[3][0] < 255:
        bg = Image.new("RGBA", im.size, (18, 18, 18, 255))
        bg.alpha_composite(im)
        im = bg
    buf = io.BytesIO()
    im.convert("RGB").save(buf, "PNG")
    return buf.getvalue()


def ask_vision(key, image_path):
    b64 = base64.b64encode(prepare_bytes(image_path)).decode()
    payload = {
        "model": MODEL,
        "messages": [
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": PROMPT},
                    {
                        "type": "image_url",
                        "image_url": {
                            "url": f"data:{mime_for(image_path)};base64,{b64}"
                        },
                    },
                ],
            }
        ],
        "max_tokens": 2048,
        "reasoning": {"enabled": False},
    }
    req = urllib.request.Request(
        API_URL,
        data=json.dumps(payload).encode(),
        headers={
            "Authorization": f"Bearer {key}",
            "Content-Type": "application/json",
        },
    )
    with urllib.request.urlopen(req, timeout=180) as r:
        resp = json.load(r)
    return (resp["choices"][0]["message"].get("content") or "").strip()


def extract_thickness(transcription):
    """Ищем 'Толщина ... N мм'. Без слова 'Толщина' - толщину не берём
    (на картинке может быть только 'Глубина фрезеровки' и т.п.)."""
    low = transcription.lower()
    idx = low.rfind("толщина")
    if idx == -1:
        return ""
    mm = list(re.finditer(r"от\s*\d+(?:[.,]\d+)?\s*мм|\d+(?:[.,]\d+)?\s*мм", low))
    for m in mm:
        if m.start() >= idx:
            return m.group(0).strip()
    return mm[-1].group(0).strip() if mm else ""


def load_progress():
    cache = {}
    if os.path.isfile(PROGRESS_CSV):
        with open(PROGRESS_CSV, newline="", encoding="utf-8") as f:
            for row in csv.reader(f):
                # храним только найденные значения; пустые не кэшируем,
                # чтобы их можно было переспрашивать
                if len(row) >= 2 and row[1].strip():
                    cache.setdefault(row[0], row[1])
    return cache


def save_progress(cache):
    with open(PROGRESS_CSV, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        for path, val in cache.items():
            w.writerow([path, val])


def main():
    limit = int(os.getenv("LIMIT", "0") or 0)
    only_section = os.getenv("SECTION", "") or ""
    key = api_key()
    cache = load_progress()

    df = pd.read_excel(EXCEL)
    th_col = "Толщина фасада"
    pics_col = "Файлы картинок"
    df[th_col] = df[th_col].fillna("").astype(str)
    df[pics_col] = df[pics_col].fillna("").astype(str)

    todo = []
    for idx, row in df.iterrows():
        if row.get("Толщина по фото") == "да":
            continue
        if row[th_col].strip():
            continue
        if only_section and str(row["Раздел"]).strip() != only_section:
            continue
        paths = [p.strip() for p in row[pics_col].split(";") if p.strip()]
        if paths:
            todo.append((idx, paths))
    if limit:
        todo = todo[:limit]

    total = len(todo)
    print(f"Обработка {total} моделей через {MODEL}", flush=True)
    found = 0
    done = 0
    last_results = []
    for n, (idx, paths) in enumerate(todo, 1):
        model_name = df.at[idx, "Модель"]
        result = ""
        for path in paths:
            if not os.path.isfile(path):
                continue
            if cache.get(path):  # уже найдено ранее
                result = cache[path]
                break
            for read_attempt in range(RETRIES_PER_IMAGE):
                text = ""
                for attempt in range(1, MAX_ATTEMPTS + 1):
                    try:
                        text = ask_vision(key, path)
                        break
                    except urllib.error.HTTPError as e:
                        if e.code == 429:
                            wait = min(15 * attempt, 90)
                            print(
                                f"  [{n}/{total}] 429, пауза {wait}с ({attempt}/{MAX_ATTEMPTS})...",
                                file=sys.stderr,
                            )
                            time.sleep(wait)
                        else:
                            msg = e.read().decode()[:200] if hasattr(e, "read") else ""
                            print(f"  [{n}/{total}] HTTP {e.code}: {msg}", file=sys.stderr)
                            time.sleep(10)
                    except Exception as e:
                        print(f"  [{n}/{total}] {model_name}: {e}", file=sys.stderr)
                        time.sleep(10)
                result = extract_thickness(text)
                if result:
                    break
            cache[path] = result
            if result:  # пишем в кэш только найденные
                save_progress(cache)
            if result:
                break
        if result:
            df.at[idx, th_col] = result
            df.at[idx, "Толщина по фото"] = "да"
            found += 1
            last_results.append((model_name, result))
            print(f"  [{n}/{total}] {model_name}: {result}", flush=True)
        else:
            print(f"  [{n}/{total}] {model_name}: толщина не найдена", flush=True)
        done += 1
        time.sleep(SLEEP)

        # промежуточное логирование каждые 5 картинок
        if done % 5 == 0:
            tail = "; ".join(f"{m}->{v}" for m, v in last_results[-5:])
            print(
                f"  --- промежуточно: обработано {done}/{total}, "
                f"найдено {found}: {tail}",
                flush=True,
            )

    # один финальный вызов записи в Excel
    df.to_excel(EXCEL, index=False)
    print(f"\nГотово. Заполнено: {found} из {total}")
    print(f"Excel обновлён: {EXCEL}")


if __name__ == "__main__":
    main()
# Парсер каталога фасадов KFMF

Собирает каталог фасадов с kfmf-skat.ru: коллекция → раздел → модель → толщина фасада. Результат — Excel-файл.

## Файлы

- `kfmf_parser.py` — парсинг каталога фасадов
- `vision_extract.py` — извлечение толщины фасада с картинок (GRAFIKA/3D), где она нарисована таблицей характеристик, через vision-модель (Qwen3.7 Flash через OpenRouter)

## Запуск

```
pip install requests beautifulsoup4 pandas pillow openpyxl
python kfmf_parser.py
python vision_extract.py
```

Для `vision_extract.py` задайте ключ OpenRouter:

```
export OPENROUTER_API_KEY=...
```

## Примечание о конфиденциальности

Изображения фасадов и промежуточные CSV/Excel не хранятся в репозитории.

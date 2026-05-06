# gost standardizer

локальный mcp-инструмент для приведения word-документов к аккуратному гост-подобному виду.

## что это умеет

- анализировать `.docx` и `.docm`
- предлагать подходящий пресет оформления
- собирать копию документа с единым шрифтом, полями, интервалами и абзацными отступами
- работать как mcp-сервер через tools

## что внутри

- `plugins/gost-standardizer/.codex-plugin/plugin.json`
- `plugins/gost-standardizer/.mcp.json`
- `scripts/gost_standardizer.py`
- `scripts/mcp_server.py`

## текущая база правил

пресеты собраны с опорой на действующие источники росстандарта, актуальные в 2026 году:

- [гост р 2.105-2019](https://protect.gost.ru/gost/details/0dd3cdfa-c15a-47b3-9e3b-4a90a43005c8)
- [гост 7.32-2017](https://protect.gost.ru/gost/details/7d280e43-7036-4a69-8e6e-d15867028343)
- [гост р 7.0.97-2016](https://protect.gost.ru/document1.aspx?control=31&id=205885)

это не заменяет требования конкретного колледжа, кафедры или организации. это базовый слой стандартизации.

## требования

- python 3.12
- `python-docx`

установка зависимостей:

```bash
python -m pip install -r requirements.txt
```

## запуск mcp

сервер запускается через python-скрипт:

```bash
python scripts/mcp_server.py
```

`.mcp.json` уже указывает на этот вход.

## инструменты

### `list_presets`

показывает доступные пресеты и их назначение.

пример:

```json
[
  {
    "key": "report",
    "title": "GOST report",
    "description": "Balanced preset for reports, coursework, explanatory notes, and formal documents."
  }
]
```

### `inspect_document`

читает документ и возвращает оценку состояния.

пример запроса:

```json
{
  "path": "C:\\work\\input.docx",
  "sample_size": 5
}
```

пример ответа:

```json
{
  "preset_guess": "technical",
  "statistics": {
    "paragraphs": 19,
    "tables": 0,
    "sections": 2
  },
  "issues": [],
  "deviations": [
    "Section margin left_mm is 27.87 mm, target is 30 mm"
  ]
}
```

### `standardize_document`

создает новую стандартизированную копию.

пример запроса:

```json
{
  "path": "C:\\work\\input.docx",
  "preset": "report",
  "overwrite": true
}
```

пример ответа:

```json
{
  "source_path": "C:\\work\\input.docx",
  "output_path": "C:\\work\\input_gost.docx",
  "preset": {
    "key": "report"
  },
  "paragraph_actions": {
    "title": 1,
    "heading": 3,
    "body": 12
  }
}
```

## форматирование

по умолчанию проект использует:

- `times new roman`
- размер текста 14 pt
- межстрочный интервал 1.5
- левое поле 30 мм для отчетных документов
- правое поле 10 мм
- верхнее и нижнее поле 20 мм

есть также пресеты:

- `report`
- `office`
- `technical`
- `legacy-college`

## ограничения

- поддерживаются только `.docx` и `.docm`
- `.doc` пока не конвертируется
- сложные колонтитулы, таблицы и секции лучше проверять вручную после стандартизации

## примеры использования

### проверить документ

```bash
python scripts/mcp_server.py
```

после старта в mcp-клиенте вызвать `inspect_document`.

### привести к гост-стилю

```json
{
  "path": "C:\\Users\\eegor\\Documents\\report.docx",
  "preset": "technical",
  "overwrite": false
}
```

результат будет сохранен рядом с исходником в файл вида `report_gost.docx`.

### `refresh_meganorm_cache`

обновляет локальный кэш html-источника meganorm.

пример ответа:

```json
{
  "kind": "refresh",
  "cache": {
    "state": "refreshed",
    "fetched_at": "2026-05-07T10:30:00+00:00"
  },
  "pages_refreshed": 3
}
```

### `search_meganorm_catalog`

ищет категории и документы по названию, с опорой на кэш и живой источник.

пример ответа:

```json
{
  "kind": "search",
  "cache": {
    "state": "hit"
  },
  "categories": [
    {
      "title": "ГОСТ (Государственный стандарт)",
      "origin": "cache"
    }
  ],
  "documents": [
    {
      "title": "ГОСТ 7.32-2017 ...",
      "origin": "source"
    }
  ]
}
```

### `get_meganorm_topics`

возвращает текущие темы и ссылки по каталогу или по выбранной категории.

пример ответа:

```json
{
  "kind": "current-topics",
  "scope": "catalog",
  "topics": [
    {
      "title": "ГОСТ (Государственный стандарт)",
      "origin": "cache"
    },
    {
      "title": "ГОСТ Р (Государственный стандарт)",
      "origin": "source"
    }
  ]
}
```

### `find_current_gost`

ищет только актуальные `ГОСТ` и `ГОСТ Р` записи, с более точным совпадением по номеру.

пример запроса:

```json
{
  "query": "ГОСТ 7.32-2017",
  "max_pages": 10,
  "limit": 5
}
```

пример ответа:

```json
{
  "kind": "current-gost-search",
  "documents": [
    {
      "title": "\"ГОСТ 7.32-2017 ...\"",
      "origin": "source",
      "match": {
        "exact": true,
        "number_hint": true,
        "partial": 3
      }
    }
  ]
}
```

## meganorm

источник берется отсюда:

- [meganorm актуализированная база](https://meganorm.ru/mega_doc/norm/norm.html)

в выводе всегда указывай происхождение записи:

- `origin=cache` значит запись взята из локального кэша
- `origin=source` значит запись получена живьем из нормативного html

это касается и категорий, и документов, и тем.

при очистке html-источника из вывода вырезаются служебные блоки вроде `script`, `style`, `noscript`, `footer`, `nav` и `aside`, а также шумные wrapper-элементы, если они встречаются в dom.

## разработчики

- `@glutinosa`

## лицензия

mit. см. [license](LICENSE).

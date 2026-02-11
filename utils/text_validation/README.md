# Text Diff Tool

Пакет для сравнения текстов и создания отчетов об ошибках с подсветкой в Excel.

## Установка

Скопируйте папку `text_diff_tool` в ваш проект или добавьте путь к ней в `PYTHONPATH`.

## Использование

### Базовый пример

```python
from text_diff_tool import (
    load_classification_data,
    load_reference_data,
    find_errors,
    create_error_report,
    save_workbook
)

# Загрузка данных
orig_names, klass_codes = load_classification_data(
    "классификация.xls",
    name_column=1,  # Столбец с наименованием
    code_columns=0  # Столбец(ы) с кодом (можно указать список: [0, 1, 2] для склеивания)
)

reference_names, reference_codes = load_reference_data(
    "справочник.xlsx",
    name_column=0,  # Столбец с наименованием
    code_columns=10  # Столбец(ы) с кодом (можно указать список: [10, 11, 12] для склеивания)
)

# Пример: если код состоит из нескольких столбцов (например, столбцы 10, 11, 12)
# reference_names, reference_codes = load_reference_data(
#     "справочник.xlsx",
#     name_column=0,
#     code_columns=[10, 11, 12]  # Код будет склеен из этих столбцов
# )

# Поиск ошибок
errors = find_errors(
    orig_names.tolist(),
    reference_names.tolist(),
    max_distance=25
)

# Создание отчета
wb = create_error_report(
    errors,
    klass_codes.tolist(),
    reference_codes.tolist()
)

# Сохранение
save_workbook(wb, "отчет.xlsx")
```

### Использование отдельных модулей

```python
# Сравнение двух строк
from text_diff_tool import find_differences

diff_idx, corrections = find_differences("текст1", "текст2")

# Создание подсвеченного текста
from text_diff_tool import create_highlighted_text

highlighted = create_highlighted_text("текст", diff_idx, corrections)

# Поиск лучшего совпадения
from text_diff_tool import find_best_match

best_idx, best_ref, dist = find_best_match(
    "исходный текст",
    ["вариант1", "вариант2", "вариант3"],
    max_distance=25
)
```

## Структура пакета

- `text_comparator.py` - сравнение строк и поиск различий
- `text_highlighter.py` - создание подсвеченного текста для Excel
- `data_loader.py` - загрузка данных из Excel файлов
- `error_finder.py` - поиск ошибок в данных
- `excel_exporter.py` - экспорт результатов в Excel

## Зависимости

- pandas
- openpyxl
- python-Levenshtein
- difflib (встроенный модуль)


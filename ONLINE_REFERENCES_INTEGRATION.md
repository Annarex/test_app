# Руководство по интеграции онлайн справочников

## Оглавление
1. [Архитектура системы](#архитектура-системы)
2. [Процесс загрузки справочников](#процесс-загрузки-справочников)
3. [Структура базы данных](#структура-базы-данных)
4. [Добавление нового справочника](#добавление-нового-справочника)
5. [Примеры: КВР и РЗПР](#примеры-квр-и-рзпр)

---

## Архитектура системы

### Компоненты системы

```
┌─────────────────────────────────────────────────────────────────┐
│                    API Электронного бюджета                      │
│              http://budget.gov.ru/epbs/registry/                 │
└────────────────────────────┬────────────────────────────────────┘
                             │
                             ▼
┌─────────────────────────────────────────────────────────────────┐
│         BudgetReferencesService (Загрузка данных)               │
│  • Пагинация (pageSize=1000)                                    │
│  • Фильтрация по региону                                        │
│  • Обработка NPA                                                │
└────────────────────────────┬────────────────────────────────────┘
                             │
                             ▼
┌─────────────────────────────────────────────────────────────────┐
│            SQLite База данных (budget_forms.db)                 │
│  • Базовые таблицы (ФУ и МО отдельно)                          │
│  • VIEW для объединения (ФУ + МО)                               │
│  • Таблица отслеживания обновлений                              │
└────────────────────────────┬────────────────────────────────────┘
                             │
                             ▼
┌─────────────────────────────────────────────────────────────────┐
│         GUI: Управление справочниками                           │
│  • BudgetReferencesUpdateDialog - обновление                   │
│  • ReferencesManagementDialog - просмотр                        │
└─────────────────────────────────────────────────────────────────┘
```

### Файловая структура

```
services/
└── budget_references_service.py     # Сервис загрузки данных

models/
└── database.py                       # Структура БД и таблиц

views/
├── budget_references_update_dialog.py   # Диалог обновления
└── references_management_dialog.py      # Диалог просмотра
```

---

## Процесс загрузки справочников

### 1. Запрос к API

**Формат URL:**
```
http://budget.gov.ru/epbs/registry/{REGISTRY_ID}/data
```

**Параметры запроса:**
- `pageSize` - размер страницы (по умолчанию 1000)
- `offset` - смещение для пагинации
- `filter{FieldName}` - фильтры (например, `filterppocode=21______`)

**Пример:**
```
http://budget.gov.ru/epbs/registry/7710568760-BUDGETCLASKVR/data?pageSize=1000&offset=0
```

### 2. Структура ответа API

```json
{
  "recordCount": 1500,
  "data": [
    {
      "code": "244",
      "name": "Прочая закупка товаров, работ и услуг",
      "startdate": "2024-01-01",
      "enddate": "2024-12-31",
      "level": "3",
      "stagename": "Исполнение",
      "budgetname": "Бюджет субъекта РФ",
      "pponame": "",
      "ppocode": "",
      "year": "2024",
      "npa": [
        {
          
        }
      ]
    }
  ]
}
```

### 3. Алгоритм загрузки

```python
1. Получить общее количество записей (recordCount)
2. Проверить текущее количество в БД
3. Если количество отличается:
   a. Очистить таблицу
   b. Загрузить все данные с пагинацией
   c. Обработать NPA (если есть)
   d. Вставить записи в БД
   e. Обновить дату последнего обновления
4. Если количество совпадает - пропустить загрузку
```

### 4. Обработка NPA (Нормативно-правовых актов)

Справочники могут содержать связанные НПА:

```python
# 1. Вставить/получить NPA из таблицы npa
npa_id = get_or_insert_npa(cursor, npa_item)

# 2. Связать с основной таблицей
cursor.execute(
    f"UPDATE {table_name} SET npa_id = ? WHERE id = ?",
    (','.join(npa_ids), record_id)
)
```

---

## Структура базы данных

### Базовая таблица справочника

#### Пример: Таблица ОКТМО
```sql
CREATE TABLE IF NOT EXISTS oktmo (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    code TEXT,
    name TEXT,
    centrename TEXT,
    munname TEXT,
    regioncode TEXT,
    status TEXT,
    created_at TEXT DEFAULT (strftime('%d.%m.%Y %H:%M:%S', 'now', 'localtime'))
)
```

#### Пример: Таблица с NPA (КВР)
```sql
CREATE TABLE IF NOT EXISTS budgetclaskvr (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    code TEXT,
    name TEXT,
    startdate TEXT,
    enddate TEXT,
    level TEXT,
    stagename TEXT,
    budgetname TEXT,
    pponame TEXT,
    ppocode TEXT,
    year TEXT,
    npa_id INTEGER,
    created_at TEXT DEFAULT (strftime('%d.%m.%Y %H:%M:%S', 'now', 'localtime')),
    FOREIGN KEY (npa_id) REFERENCES npa(id)
)
```

### Индексы

Создаются для ускорения поиска:
```sql
CREATE INDEX IF NOT EXISTS idx_kvr_ppocode ON budgetclaskvr(ppocode);
CREATE INDEX IF NOT EXISTS idx_kvr_code ON budgetclaskvr(code);
CREATE INDEX IF NOT EXISTS idx_kvr_dates ON budgetclaskvr(startdate, enddate);
CREATE INDEX IF NOT EXISTS idx_kvr_npa_id ON budgetclaskvr(npa_id);
```

### Таблица отслеживания обновлений

```sql
CREATE TABLE IF NOT EXISTS budget_references_updates (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    table_name TEXT NOT NULL,
    records_count INTEGER NOT NULL,
    created_at TEXT DEFAULT (strftime('%d.%m.%Y %H:%M:%S', 'now', 'localtime')),
    update_at TEXT DEFAULT (strftime('%d.%m.%Y %H:%M:%S', 'now', 'localtime'))
)
```

### VIEW для объединения (опционально)

Для справочников с версиями ФУ и МО:
```sql
CREATE VIEW IF NOT EXISTS v_budgetclaskvr_merged AS
SELECT 
    id, code, name, startdate, enddate, level,
    stagename, budgetname, pponame, ppocode, year,
    npa_id, created_at
FROM (
    SELECT * FROM budgetclaskvr
    UNION ALL
    SELECT * FROM budgetclaskvrmo
)
ORDER BY ppocode, code, startdate, enddate, year
```

---

## Добавление нового справочника

### Шаг 1: Определить URL API

**Формат:** `http://budget.gov.ru/epbs/registry/{REGISTRY_ID}/data`

Примеры:
- КВР: `7710568760-BUDGETCLASKVR`
- РЗПР: `7710568760-BUDGETCLASRZPR`

### Шаг 2: Добавить URL в `budget_references_service.py`

```python
# services/budget_references_service.py

# Добавить константу URL
URL_BUDGETCLASKVR = "http://budget.gov.ru/epbs/registry/7710568760-BUDGETCLASKVR/data"
URL_BUDGETCLASRZPR = "http://budget.gov.ru/epbs/registry/7710568760-BUDGETCLASRZPR/data"

# Добавить в словарь URL_TO_TABLE
URL_TO_TABLE = {
    # ... существующие ...
    URL_BUDGETCLASKVR: 'budgetclaskvr',
    URL_BUDGETCLASRZPR: 'budgetclasrzpr',
}

# Добавить фильтры по умолчанию (если нужны)
def get_default_filters(default_region: str = "21") -> Dict[str, Optional[Dict]]:
    return {
        # ... существующие ...
        'budgetclaskvr': {"ppocode": f"{default_region}______"},
        'budgetclasrzpr': {"ppocode": f"{default_region}______"},
    }
```

### Шаг 3: Создать таблицу в `database.py`

```python
# models/database.py

def _init_budget_references_tables(self, cursor: sqlite3.Cursor) -> None:
    # ... существующий код ...
    
    # Таблица: КВР
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS budgetclaskvr (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            code TEXT,
            name TEXT,
            startdate TEXT,
            enddate TEXT,
            level TEXT,
            stagename TEXT,
            budgetname TEXT,
            pponame TEXT,
            ppocode TEXT,
            year TEXT,
            npa_id INTEGER,
            created_at TEXT DEFAULT (strftime('%d.%m.%Y %H:%M:%S', 'now', 'localtime')),
            FOREIGN KEY (npa_id) REFERENCES npa(id)
        )
    ''')
    
    # Индексы
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_kvr_ppocode ON budgetclaskvr(ppocode)')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_kvr_code ON budgetclaskvr(code)')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_kvr_dates ON budgetclaskvr(startdate, enddate)')
    cursor.execute('CREATE INDEX IF NOT EXISTS idx_kvr_npa_id ON budgetclaskvr(npa_id)')
```

### Шаг 4: Добавить в `budget_references_update_dialog.py`

```python
# views/budget_references_update_dialog.py

# Добавить в REFERENCE_NAMES
REFERENCE_NAMES = {
    # ... существующие ...
    'budgetclaskvr': 'Коды видов расходов (КВР)',
    'budgetclasrzpr': 'Коды разделов и подразделов (РЗПР)',
}
```
# Правила сортировки для различных таблиц справочников
SORT_RULES = {
    ...
    'budgetclaskvr': ['ppocode', 'code', 'startdate', 'enddate', 'year'],
    'budgetclasrzpr': ['ppocode', 'code', 'startdate', 'enddate', 'year'],
}
### Шаг 5: Добавить в `references_management_dialog.py`

```python
# views/references_management_dialog.py

REFERENCE_TYPES = {
    # ... существующие онлайн справочники ...
    'Коды видов расходов (КВР)': {
        'table': 'budgetclaskvr',
        'load_method': None,
        'columns': None,
        'is_online': True,
        'has_npa': True
    },
    'Коды разделов и подразделов (РЗПР)': {
        'table': 'budgetclasrzpr',
        'load_method': None,
        'columns': None,
        'is_online': True,
        'has_npa': True
    },
}
```

### Шаг 6: Экспортировать константы (опционально)

```python
# services/__init__.py

from .budget_references_service import (
    # ... существующие ...
    URL_BUDGETCLASKVR,
    URL_BUDGETCLASRZPR,
)

__all__ = [
    # ... существующие ...
    'URL_BUDGETCLASKVR',
    'URL_BUDGETCLASRZPR',
]
```

---

## Примеры: КВР и РЗПР

### КВР (Коды видов расходов)

**URL:** `http://budget.gov.ru/epbs/registry/7710568760-BUDGETCLASKVR/data`

**Структура данных:**
```json
{
  "code": "244",
  "name": "Прочая закупка товаров, работ и услуг",
  "startdate": "2024-01-01",
  "enddate": "2024-12-31",
  "level": "3",
  "stagename": "Исполнение",
  "budgetname": "Бюджет субъекта РФ",
  "pponame": "Чувашская Республика",
  "ppocode": "21000000000",
  "year": "2024",
  "npa": [...]
}
```

**Плейсхолдеры для полей:**
- `code` → "Код"
- `name` → "Наименование"
- `startdate` → "Дата начала действия"
- `enddate` → "Дата окончания"
- `level` → "Уровень иерархии"
- `stagename` → "Наименование этапа"
- `budgetname` → "Бюджет"
- `pponame` → "Наименование ППО"
- `ppocode` → "Код ППО"
- `year` → "Год"
- `npa` → "НПА"

### РЗПР (Коды разделов и подразделов)

**URL:** `https://budget.gov.ru/epbs/registry/7710568760-BUDGETCLASRZPR/data`

**Структура данных:** (аналогична КВР)
```json
{
  "code": "0104",
  "name": "Функционирование Правительства Российской Федерации",
  "startdate": "2024-01-01",
  "enddate": "2024-12-31",
  "level": "2",
  "stagename": "Исполнение",
  "budgetname": "Федеральный бюджет",
  "pponame": "Российская Федерация",
  "ppocode": "00000000000",
  "year": "2024",
  "npa": [...]
}
```

**Плейсхолдеры:** (те же, что и для КВР)

---

## Особенности реализации

### 1. Пагинация

Загрузка данных происходит постранично (по 1000 записей):
```python
PAGE_SIZE = 1000
offset = 0

while True:
    rows = fetch_chunk(url, offset, PAGE_SIZE, filters)
    if not rows:
        break
    all_rows.extend(rows)
    offset += PAGE_SIZE
```

### 2. Фильтрация по региону

Для справочников МО применяется фильтр по коду региона:
```python
filters = {"ppocode": "21______"}  # Чувашия (21)
```

### 3. Обработка дублирования

- Для обычных таблиц: `INSERT OR IGNORE`
- Для таблиц с direct_id (costs): `INSERT OR REPLACE`

### 4. Обновление метаданных

После успешной загрузки обновляется таблица `budget_references_updates`:
```python
db_manager.update_budget_reference_date(table_name, inserted_count)
```

### 5. Отмена загрузки

Поддерживается отмена загрузки через callback:
```python
def cancel_check():
    return self._is_cancelled.get(table_name, False)

records = fetch_all_rows(url, filters, cancel_check=cancel_check)
```

---

## Тестирование

### Проверка URL

```python
import requests

url = "http://budget.gov.ru/epbs/registry/7710568760-BUDGETCLASKVR/data"
params = {"pageSize": 10, "offset": 0}

response = requests.get(url, params=params)
data = response.json()

print(f"Всего записей: {data.get('recordCount')}")
print(f"Первая запись: {data['data'][0]}")
```

### Проверка таблицы

```sql
-- Количество записей
SELECT COUNT(*) FROM budgetclaskvr;

-- Записи с NPA
SELECT COUNT(*) FROM budgetclaskvr WHERE npa_id IS NOT NULL;

-- Последнее обновление
SELECT * FROM budget_references_updates 
WHERE table_name = 'budgetclaskvr' 
ORDER BY update_at DESC LIMIT 1;
```

---

## Дополнительные ресурсы

### Полезные VIEW

```sql
-- Статистика по таблице
SELECT 
    table_name,
    COUNT(*) as total_records,
    COUNT(npa_id) as records_with_npa
FROM budgetclaskvr
GROUP BY table_name;

-- Актуальные записи
SELECT * FROM budgetclaskvr
WHERE (enddate IS NULL OR enddate = '' OR enddate >= date('now'))
  AND startdate <= date('now');
```

### Логирование

Все операции логируются через модуль `logger`:
```python
from logger import logger

logger.info(f"[{table_name}] Загрузка начата")
logger.warning(f"[{table_name}] Количество не совпадает")
logger.error(f"[{table_name}] Ошибка: {e}")
```

---

## Контрольный список

При добавлении нового справочника проверьте:

- [ ] URL добавлен в `budget_references_service.py`
- [ ] Таблица создана в `database.py`
- [ ] Индексы созданы для ключевых полей
- [ ] Название добавлено в `REFERENCE_NAMES`
- [ ] Справочник добавлен в `REFERENCE_TYPES`
- [ ] Фильтры настроены (если нужны)
- [ ] VIEW создано (если нужно)
- [ ] Проверена загрузка данных
- [ ] Проверен вывод в GUI

---

**Документация актуальна для версии:** 2026-02-18

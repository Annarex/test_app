# Анализ логики фильтрации в диалоговом окне справочников

## Обзор

В диалоговом окне управления справочниками (`ReferencesManagementDialog`) реализована комплексная система фильтрации данных, включающая:
1. **Фильтрацию по дате** - для получения актуальных записей на указанную дату
2. **Текстовый поиск** - поиск по всем текстовым колонкам
3. **Пагинацию** - разбиение больших объемов данных на страницы

---

## 1. Компоненты фильтрации

### 1.1. Фильтрация по дате

**Элементы UI:**
- `date_filter_checkbox` (QCheckBox) - чекбокс включения/выключения фильтрации
- `date_filter` (QDateEdit) - поле выбора даты

**Состояние:**
- По умолчанию фильтрация **отключена** (`filter_date = None`)
- При включении чекбокса поле даты становится активным
- Дата сохраняется в формате `'YYYY-MM-DD'` в переменной `self.filter_date`

**Обработчики:**
- `on_date_filter_toggled()` - при включении/выключении чекбокса
- `on_date_filter_changed()` - при изменении даты
- `_update_filter_date()` - обновляет `self.filter_date` из UI

**Логика фильтрации:**
Фильтрует записи по условию:
```sql
(startdate IS NULL OR date(startdate) <= date(?))
AND (enddate IS NULL OR enddate = '' OR date(enddate) >= date(?))
```

То есть запись считается актуальной, если:
- `startdate` не установлена ИЛИ меньше или равна выбранной дате
- `enddate` не установлена/пустая ИЛИ больше или равна выбранной дате

### 1.2. Текстовый поиск

**Элемент UI:**
- `search_input` (QLineEdit) - поле ввода текста для поиска

**Обработчик:**
- `on_search_text_changed()` - вызывается при изменении текста

**Логика поиска:**
- Ищет по всем текстовым колонкам (исключая служебные: `id`, `guid`, `created_at`, `loaddate`, `startdate`, `enddate`)
- Использует SQL `LIKE` с паттерном `%текст%` (регистронезависимый поиск)
- Для каждой колонки создается условие `column LIKE ?`, все условия объединяются через `OR`

**SQL условие:**
```sql
(column1 LIKE ? OR column2 LIKE ? OR ...)
```

### 1.3. Пагинация

**Параметры:**
- `current_page` - текущая страница (начинается с 1)
- `page_size` - размер страницы (по умолчанию 1000 записей)
- `total_records` - общее количество записей (с учетом фильтров)
- `total_pages` - общее количество страниц

**Элементы UI:**
- Кнопки навигации: первая, предыдущая, следующая, последняя страница
- Поле ввода номера страницы
- Поле изменения размера страницы

---

## 2. Алгоритм загрузки данных

### 2.1. Определение типа справочника

Метод `load_current_reference()` определяет тип справочника:

1. **Справочники конфигурации** (`is_config=True`)
   - Загружаются без фильтрации и пагинации
   - Используют специальные методы `_load_years()`, `_load_municipalities()`, etc.

2. **VIEW (объединенные представления)** (`is_view=True`)
   - Используют специальную логику фильтрации через `get_filtered_view()`

3. **Обычные таблицы** (включая онлайн справочники)
   - Используют стандартную SQL фильтрацию

### 2.2. Проверка наличия полей дат

```python
has_date_fields = 'startdate' in existing_columns and 'enddate' in existing_columns
use_date_filter = has_date_fields and self.filter_date is not None
```

Фильтрация по дате применяется только если:
- Таблица содержит поля `startdate` и `enddate`
- Фильтрация включена пользователем (`filter_date is not None`)

### 2.3. Два пути обработки данных

#### Путь 1: VIEW с фильтрацией по дате

**Условие:** `is_view and use_date_filter`

**Алгоритм:**
1. Загружаются ВСЕ данные через `get_filtered_view()` с фильтрацией по дате
2. Применяется текстовый поиск в pandas DataFrame
3. Вычисляется общее количество записей после фильтрации
4. Применяется пагинация через `iloc[offset:offset+limit]`

**Особенности:**
- ⚠️ **ПРОБЛЕМА:** Загружаются ВСЕ записи в память, затем фильтруются
- Поиск выполняется в pandas, а не в SQL
- Неэффективно для больших объемов данных

**Код:**
```python
if is_view and use_date_filter:
    df_filtered = get_filtered_view(conn, table_name, self.filter_date)
    
    # Применяем поиск к отфильтрованным данным
    if self.search_text:
        search_mask = pd.Series([False] * len(df_filtered))
        for col in df_filtered.columns:
            if col not in ['id', 'guid', 'created_at', 'loaddate', 'startdate', 'enddate']:
                search_mask |= df_filtered[col].astype(str).str.contains(
                    self.search_text, case=False, na=False
                )
        df_filtered = df_filtered[search_mask]
    
    # Применяем пагинацию
    self.total_records = len(df_filtered)
    offset = (self.current_page - 1) * self.page_size
    df = df_filtered.iloc[offset:offset + limit].copy()
```

#### Путь 2: Обычные таблицы (SQL фильтрация)

**Условие:** Не VIEW или фильтрация по дате отключена

**Алгоритм:**
1. Формируется SQL запрос с WHERE условиями для:
   - Фильтрации по дате (если включена)
   - Текстового поиска (если указан)
2. Выполняется COUNT запрос для подсчета общего количества записей
3. Выполняется SELECT запрос с LIMIT/OFFSET для пагинации

**Особенности:**
- ✅ Фильтрация выполняется на уровне SQL
- ✅ Загружаются только нужные записи
- ✅ Эффективно для больших объемов данных

**Код:**
```python
# Подсчет общего количества записей
count_query = f'SELECT COUNT(*) FROM {table_name}'
search_where, search_params = build_search_where(existing_columns, add_date_filter=use_date_filter)
if search_where:
    count_query += f" {search_where}"
cursor.execute(count_query, search_params)
self.total_records = cursor.fetchone()[0]

# Загрузка данных с пагинацией
offset = (self.current_page - 1) * self.page_size
query = f'SELECT ... FROM {table_name}'
if search_where:
    query += f" {search_where}"
query += f" LIMIT {limit} OFFSET {offset}"
df = self._execute_query(conn, query, search_params)
```

---

## 3. Функция build_search_where()

**Назначение:** Формирует WHERE условие для SQL запроса

**Параметры:**
- `base_columns` - список колонок для поиска
- `add_date_filter` - добавить фильтр по дате
- `table_prefix` - префикс таблицы (для JOIN запросов)

**Логика:**

1. **Фильтрация по дате** (если включена):
   ```python
   if add_date_filter and self.filter_date:
       conditions.append(f"({prefix}startdate IS NULL OR date({prefix}startdate) <= date(?))")
       conditions.append(f"({prefix}enddate IS NULL OR {prefix}enddate = '' OR date({prefix}enddate) >= date(?))")
       params.extend([self.filter_date, self.filter_date])
   ```

2. **Текстовый поиск** (если указан):
   ```python
   if self.search_text:
       text_columns = [col for col in base_columns 
                      if col not in ['id', 'guid', 'created_at', 'loaddate', 'startdate', 'enddate']]
       if text_columns:
           search_pattern = f"%{self.search_text}%"
           search_conditions = " OR ".join([f"{prefix}{col} LIKE ?" for col in text_columns])
           conditions.append(f"({search_conditions})")
           params.extend([search_pattern] * len(text_columns))
   ```

3. **Объединение условий:**
   - Все условия объединяются через `AND`
   - Возвращается строка `"WHERE ..."` и список параметров

---

## 4. Особые случаи

### 4.1. Справочники с JOIN к таблице NPA

**Условие:** `has_npa=True` и есть колонка `npa_id`

**Особенности:**
- Используется `LEFT JOIN` с таблицей `npa`
- В WHERE условиях используется префикс таблицы (`table_name.column`) для избежания неоднозначности
- Колонки из `npa` добавляются с префиксом `npa_`

**Код:**
```python
if has_npa and has_npa_column:
    count_query = f'SELECT COUNT(*) FROM {table_name} LEFT JOIN npa ON {table_name}.npa_id = npa.id'
    search_where, search_params = build_search_where(existing_columns, add_date_filter=use_date_filter, table_prefix=table_name)
```

### 4.2. Справочник "Коды доходов"

**Особенность:** Использует VIEW `v_budgetclastypeinc_merged` с `display_columns` и фильтром по дате (`concatenated_code AS код`, `name AS наименование`, `level AS уровень`).

**Код:** общая ветка для таблиц с `display_columns` и `search_columns` (в т.ч. `v_budgetclastypeinc_merged`), поиск по `search_columns`, после `get_filtered_view` — маппинг колонок на отображаемые имена.

---

## 5. Проблемы и недостатки

### 5.1. ⚠️ Неэффективная обработка VIEW с фильтрацией по дате

**Проблема:**
- Загружаются ВСЕ записи из VIEW в память
- Поиск выполняется в pandas, а не в SQL
- Пагинация применяется после загрузки всех данных

**Последствия:**
- Высокое потребление памяти
- Медленная работа при больших объемах данных
- Неэффективное использование ресурсов БД

**Рекомендация:**
- Переписать логику для применения фильтров и пагинации на уровне SQL
- Использовать подзапросы или CTE для фильтрации VIEW

### 5.2. ⚠️ Дублирование логики фильтрации

**Проблема:**
- Логика фильтрации по дате дублируется:
  - В `get_filtered_view()` (для VIEW)
  - В `build_search_where()` (для обычных таблиц)

**Последствия:**
- Сложность поддержки
- Возможность расхождений в логике

**Рекомендация:**
- Унифицировать логику фильтрации
- Использовать единый подход для всех типов справочников

### 5.3. ⚠️ Отсутствие индексов для поиска

**Проблема:**
- Текстовый поиск использует `LIKE '%текст%'`
- Такой поиск не может использовать индексы
- Медленно работает на больших таблицах

**Рекомендация:**
- Рассмотреть использование полнотекстового поиска (FTS)
- Или добавить индексы на часто используемые колонки

### 5.4. ⚠️ Нет кэширования результатов

**Проблема:**
- При каждом изменении фильтров выполняется новый запрос к БД
- Нет кэширования результатов поиска

**Рекомендация:**
- Добавить кэширование результатов для часто используемых запросов
- Использовать debounce для текстового поиска

---

## 6. Рекомендации по улучшению

### 6.1. Унификация логики фильтрации

**Предложение:**
Создать единую функцию для фильтрации всех типов справочников:

```python
def build_filtered_query(
    table_name: str,
    columns: list,
    filter_date: Optional[str] = None,
    search_text: str = "",
    is_view: bool = False,
    has_npa: bool = False,
    limit: int = None,
    offset: int = None
) -> tuple[str, list]:
    """Строит SQL запрос с фильтрацией для любого типа справочника"""
    # Единая логика для всех случаев
```

### 6.2. Оптимизация VIEW фильтрации

**Предложение:**
Применять фильтры и пагинацию на уровне SQL:

```python
# Вместо загрузки всех данных:
df_filtered = get_filtered_view(conn, table_name, self.filter_date)
df_filtered = df_filtered[search_mask]
df = df_filtered.iloc[offset:offset+limit]

# Использовать подзапрос:
query = f"""
    SELECT * FROM (
        SELECT * FROM (
            {get_filtered_view_query(table_name, filter_date)}
        ) WHERE {search_conditions}
    ) LIMIT {limit} OFFSET {offset}
"""
```

### 6.3. Добавление debounce для поиска

**Предложение:**
Использовать QTimer для задержки выполнения поиска:

```python
from PyQt5.QtCore import QTimer

def __init__(self):
    self.search_timer = QTimer()
    self.search_timer.setSingleShot(True)
    self.search_timer.timeout.connect(self._execute_search)

def on_search_text_changed(self, text: str):
    self.search_text = text.strip()
    self.search_timer.stop()
    self.search_timer.start(500)  # Задержка 500 мс
```

### 6.4. Сохранение состояния фильтров

**Предложение:**
Сохранять выбранные фильтры в конфигурацию:

```python
def save_filter_state(self):
    config = {
        'filter_date_enabled': self.date_filter_checkbox.isChecked(),
        'filter_date': self.filter_date,
        'page_size': self.page_size
    }
    self.db_manager.save_config('references_filters', config)

def load_filter_state(self):
    config = self.db_manager.load_config('references_filters', {})
    # Восстанавливаем состояние
```

---

## 7. Выводы

### Текущее состояние:
- ✅ Фильтрация по дате работает корректно для обычных таблиц
- ✅ Текстовый поиск реализован и функционирует
- ✅ Пагинация работает правильно
- ⚠️ VIEW фильтрация неэффективна (загружает все данные)
- ⚠️ Дублирование логики фильтрации
- ⚠️ Нет оптимизации для больших объемов данных

### Приоритеты улучшений:
1. **Высокий:** Оптимизация VIEW фильтрации (применение фильтров в SQL)
2. **Средний:** Унификация логики фильтрации
3. **Низкий:** Добавление debounce для поиска
4. **Низкий:** Сохранение состояния фильтров

---

## 8. Примеры использования

### Пример 1: Фильтрация по дате для обычной таблицы

```python
# Пользователь включает фильтрацию и выбирает дату "2025-01-15"
# SQL запрос:
SELECT * FROM budgetclascostsmo
WHERE (startdate IS NULL OR date(startdate) <= date('2025-01-15'))
  AND (enddate IS NULL OR enddate = '' OR date(enddate) >= date('2025-01-15'))
LIMIT 1000 OFFSET 0
```

### Пример 2: Поиск с фильтрацией по дате

```python
# Пользователь вводит "доход" и включает фильтрацию по дате "2025-01-15"
# SQL запрос:
SELECT * FROM budgetclastypeinc
WHERE (startdate IS NULL OR date(startdate) <= date('2025-01-15'))
  AND (enddate IS NULL OR enddate = '' OR date(enddate) >= date('2025-01-15'))
  AND (name LIKE '%доход%' OR code LIKE '%доход%' OR ...)
LIMIT 1000 OFFSET 0
```

### Пример 3: VIEW с фильтрацией (текущая реализация)

```python
# Загружаются ВСЕ данные из VIEW:
df = get_filtered_view(conn, 'v_budgetclastypeinc_merged', '2025-01-15')
# Затем фильтруются в pandas:
df = df[df['name'].str.contains('доход', case=False, na=False)]
# Затем применяется пагинация:
df = df.iloc[0:1000]
```

---

**Дата анализа:** 2026-01-28
**Версия кода:** Текущая версия из `references_management_dialog.py`

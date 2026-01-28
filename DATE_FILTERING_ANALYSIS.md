# Анализ использования фильтрации по дате в приложении

## Обзор

В приложении используется фильтрация данных справочников по дате для получения актуальных записей на указанную дату. Все таблицы бюджетных справочников содержат поля `startdate` и `enddate` для определения периода действия записи.

## Основные компоненты

### 1. Утилита фильтрации: `utils/db_utils.py`

**Функция:** `get_filtered_view()`

**Назначение:** Фильтрует данные из VIEW по указанной дате с дедупликацией

**Параметры:**
- `conn`: SQLite соединение
- `view_name`: Имя VIEW (например, 'v_budgetclastypeinc_merged')
- `filter_date`: Дата в формате 'YYYY-MM-DD' (опционально, по умолчанию текущая дата)
- `partition_by`: Поля для группировки при дедупликации (опционально)

**Логика работы:**
1. Определяет поля для PARTITION BY (исключая 'startdate', 'enddate', 'year')
2. Фильтрует записи: `startdate <= filter_date <= enddate` (или enddate пустой)
3. Применяет дедупликацию: выбирает запись с максимальным `startdate` для каждой комбинации полей

**SQL запрос:**
```sql
SELECT * FROM (
    SELECT *, ROW_NUMBER() OVER (
        PARTITION BY {partition_fields}
        ORDER BY startdate DESC
    ) AS rn
    FROM {view_name}
    WHERE (startdate IS NULL OR startdate <= {date_filter})
      AND (enddate IS NULL OR enddate = '' OR enddate >= {date_filter})
) WHERE rn = 1
```

**Использование:**
- Вызывается из `BudgetLevelProcessor.export_filtered_data()`
- Используется в скриптах Osnova для экспорта данных

---

### 2. Процессор уровней: `services/budget_level_processor.py`

**Класс:** `BudgetLevelProcessor`

#### Метод: `process_and_update_levels()`

**Параметры:**
- `filter_date`: Дата в формате 'YYYY-MM-DD' (опционально, по умолчанию текущая дата)
- `base_ppocode`: Код базового участника БП (по умолчанию '00000000')

**Проблема:** 
⚠️ **Метод НЕ использует фильтрацию по дате!** Загружает ВСЕ данные из VIEW без фильтрации:
```python
df = pd.read_sql_query(f"SELECT * FROM {VIEW_MERGED}", conn)
```

**Вызовы:**
- `views/budget_references_update_dialog.py:140` - вызывается БЕЗ параметра `filter_date`
  ```python
  processor.process_and_update_levels()  # Всегда использует текущую дату (но не фильтрует!)
  ```

**Рекомендация:** 
- Добавить фильтрацию по дате в `process_and_update_levels()` или использовать `get_filtered_view()`

---

### 3. Структура базы данных

**Таблицы с полями дат:**
- `budgetclastypeinc` - startdate, enddate
- `budgetclassubtypincmo` - startdate, enddate
- `budgetclasgabs` - startdate, enddate
- `budgetclasgabsmo` - startdate, enddate
- `budgetclasgrbs` - startdate, enddate
- `budgetclasgrbsmo` - startdate, enddate
- `budgetclascosts` - startdate, enddate
- `budgetclascostsmo` - startdate, enddate
- `budgetclasgaiffb` - startdate, enddate
- `budgetclasgaifmo` - startdate, enddate
- `budgetclassources` - startdate, enddate
- `budgetclassourcesmo` - startdate, enddate
- `oktmo` - startdate, enddate

**VIEW с объединенными данными:**
- `v_budgetclastypeinc_merged` - объединяет ФУ и МО доходы
- `v_budgetclascosts_merged` - объединяет ФУ и МО расходы
- `v_budgetclasgrbs_merged` - объединяет ФУ и МО распорядители
- `v_budgetclasgabs_merged` - объединяет ФУ и МО администраторы
- `v_budgetclasgaiffb_merged` - объединяет ФУ и МО источники дефицита
- `v_budgetclassources_merged` - объединяет ФУ и МО источники финансирования

**Индексы:**
- Все таблицы имеют индексы на `(startdate, enddate)` для оптимизации фильтрации

---

### 4. Места использования

#### 4.1. Обновление уровней после загрузки справочника

**Файл:** `views/budget_references_update_dialog.py`

**Метод:** `_update_levels_after_mo_update()`

**Вызов:**
```python
processor = BudgetLevelProcessor(self.db_path)
processor.process_and_update_levels()  # БЕЗ filter_date
```

**Проблема:** 
- Не передается `filter_date`, используется текущая дата по умолчанию
- Но метод не фильтрует данные, загружает все записи

---

#### 4.2. Экспорт данных (не используется в основном коде)

**Метод:** `BudgetLevelProcessor.export_filtered_data()`

**Статус:** 
- ✅ Правильно использует `get_filtered_view()` с фильтрацией по дате
- ⚠️ Не вызывается из основного кода приложения

---

### 5. Скрипты Osnova (вне основного приложения)

#### 5.1. `Osnova/app_budgetclastypeinc_merged.py`

**Функция:** `process_and_merge_data(filter_date: str = '2025-12-29')`

**Использование:**
- Использует `get_filtered_view()` для экспорта
- Принимает `filter_date` как параметр

#### 5.2. `Osnova/compare_texts.py`

**Использование:**
- Использует `get_filtered_view()` с опциональным `filter_date`

#### 5.3. `Osnova/text_diff_tool/data_loader.py`

**Функция:** `load_reference_data(filter_date: Optional[datetime] = None)`

**Логика:**
- Фильтрует справочник по дате в pandas
- Выбирает актуальные записи: `startdate <= filter_date <= enddate`
- Дедуплицирует по коду, выбирая запись с максимальным `startdate`

---

## Проблемы и рекомендации

### Проблема 1: `process_and_update_levels()` не фильтрует по дате

**Текущее состояние:**
```python
df = pd.read_sql_query(f"SELECT * FROM {VIEW_MERGED}", conn)  # Загружает ВСЕ записи
```

**Рекомендация:**
1. Использовать `get_filtered_view()` вместо прямого запроса
2. Или добавить фильтрацию в SQL запрос

### Проблема 2: Нет возможности указать дату при обновлении уровней

**Текущее состояние:**
```python
processor.process_and_update_levels()  # Всегда текущая дата
```

**Рекомендация:**
- Добавить параметр даты в UI или использовать дату проекта

### Проблема 3: `export_filtered_data()` не используется

**Рекомендация:**
- Добавить возможность экспорта отфильтрованных данных через UI

---

## Выводы

1. **Основная функция фильтрации:** `utils/db_utils.py::get_filtered_view()` - работает корректно
2. **Процессор уровней:** `process_and_update_levels()` - НЕ использует фильтрацию по дате
3. **Экспорт данных:** `export_filtered_data()` - использует фильтрацию, но не вызывается из UI
4. **Все таблицы справочников** имеют поля `startdate` и `enddate` для фильтрации
5. **VIEW в БД** созданы для объединения данных, но фильтрация выполняется программно

## Рекомендации по улучшению

1. ✅ Исправить `process_and_update_levels()` - добавить фильтрацию по дате
2. ✅ Добавить возможность указать дату при обновлении уровней
3. ✅ Добавить UI для экспорта отфильтрованных данных
4. ✅ Сохранять выбранную дату в конфигурации приложения

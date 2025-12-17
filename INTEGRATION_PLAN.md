# План интеграции нового функционала с данными

## Анализ текущего состояния

### 1. Структура данных формы 0503317

Данные формы хранятся в двух форматах:

#### В памяти (объект Form0503317):
```python
{
    'доходы_data': [
        {
            'код_классификации': '10000000000000000',
            'наименование_показателя': 'Налоговые и неналоговые доходы',
            'уровень': 1,
            'утвержденный': {
                'бюджет субъекта Российской Федерации': 1000000.0,
                'бюджет муниципального образования': 500000.0,
                ...
            },
            'исполненный': {
                'бюджет субъекта Российской Федерации': 950000.0,
                ...
            },
            'расчетный_утвержденный_бюджет субъекта Российской Федерации': 1000000.0,
            'расчетный_исполненный_бюджет субъекта Российской Федерации': 950000.0,
            'исходная_строка': 15  # номер строки в Excel
        },
        ...
    ],
    'расходы_data': [...],
    'источники_финансирования_data': [...],
    'консолидируемые_расчеты_data': [...],
    'meta_info': {...},
    'calculated_deficit_proficit': {
        'утвержденный': {...},
        'исполненный': {...}
    }
}
```

#### В БД (нормализованные таблицы):
- `income_values` - доходы
- `expense_values` - расходы
- `source_values` - источники финансирования
- `consolidated_values` - консолидируемые расчеты

Структура таблиц:
- `project_id`, `revision_id`
- `classification_code` (код классификации)
- `indicator_name` (наименование показателя)
- `level` (уровень)
- `budget_type` ('утвержденный' или 'исполненный')
- `data_type` ('оригинальные' или 'вычисленные')
- `v1`, `v2`, ... `vN` (значения по бюджетным колонкам)

### 2. Текущая реализация контроллеров

#### DocumentController
- ✅ Базовая структура реализована
- ✅ Извлечение данных из ревизии (`_extract_form_data`)
- ✅ Формирование Таблицы 2 (доходы) - базовая версия
- ✅ Формирование Таблицы 3 (расходы) - реализовано
- ⚠️ Требует доработки:
  - Фильтрация по уровням через справочник `ref_income_levels`
  - Вычисление процента исполнения для каждой строки
  - Замена всех меток из кода 1С
  - Работа с диаграммами в Word

#### SolutionController
- ✅ Базовая структура реализована
- ✅ Обработка Приложения 1 (доходы)
- ✅ Обработка Приложения 2 (расходы общие)
- ✅ Обработка Приложения 3 (расходы по ГРБС)
- ❌ Не реализовано:
  - Сохранение данных в БД
  - Агрегация данных (свертка по уровням)
  - Создание записей в справочниках кодов, если код не найден

## План интеграции

### Этап 1: Доработка DocumentController

#### 1.1. Улучшение извлечения данных из формы

**Проблема**: Метод `_extract_form_data` использует `load_revision_data`, который возвращает данные в старом формате JSON. Нужно использовать нормализованные данные.

**Решение**:
```python
def _extract_form_data(self, project_id: int, revision_id: int) -> Dict[str, Any]:
    """Извлечение данных из формы для формирования заключения"""
    # Используем нормализованные данные из БД
    income_df = self.db_manager.load_income_values_df(project_id, revision_id)
    expense_df = self.db_manager.load_expense_values_df(project_id, revision_id)
    
    # Преобразуем DataFrame в формат для работы
    доходы_data = self._convert_df_to_form_format(income_df, 'доходы')
    расходы_data = self._convert_df_to_form_format(expense_df, 'расходы')
    
    # Находим итоговые строки
    доходы_всего = self._find_total_row(доходы_data, 'доходы бюджета - всего')
    расходы_всего = self._find_total_row(расходы_data, 'расходы бюджета - всего')
    
    # Вычисляем агрегированные значения
    budget_col = 'бюджет субъекта Российской Федерации'
    доходы_утвержденный = доходы_всего.get('утвержденный', {}).get(budget_col, 0) if доходы_всего else 0
    # ... и т.д.
```

**Зависимости**:
- Метод `_convert_df_to_form_format` для преобразования DataFrame в формат словарей
- Метод `_find_total_row` для поиска итоговых строк

#### 1.2. Доработка Таблицы 2 (доходы)

**Проблема**: Фильтрация по уровням работает частично, нужно использовать справочник `ref_income_levels`.

**Решение**:
```python
def _insert_table2(self, doc: docx.Document, form_data: Dict[str, Any], protocol_date: datetime):
    """Вставка Таблицы 2 (доходы) с фильтрацией по уровням"""
    # Загружаем справочник уровней доходов
    income_levels_df = self.db_manager.load_income_levels_df()
    
    доходы_data = form_data['доходы_data']
    table_data = []
    budget_col = 'бюджет субъекта Российской Федерации'
    
    for item in доходы_data:
        код = item.get('код_классификации', '').replace(' ', '')
        уровень = item.get('уровень', 0)
        
        # Проверяем, нужно ли включать этот уровень в таблицу
        # Используем справочник для определения допустимых уровней
        if not self._should_include_level(уровень, код, income_levels_df):
            continue
        
        # Вычисляем процент исполнения
        убн = item.get('утвержденный', {}).get(budget_col, 0) / 1000
        исполнение = item.get('исполненный', {}).get(budget_col, 0) / 1000
        план = (исполнение / убн * 100) if убн != 0 else 0
        
        table_data.append({
            'наименование': item.get('наименование_показателя', ''),
            'убн': убн,
            'исполнение': исполнение,
            'план': round(план, 1),
            'уровень': уровень
        })
    
    # Сортируем по уровню и коду
    table_data.sort(key=lambda x: (x['уровень'], x.get('код', '')))
    
    # Вставляем таблицу в документ
    # ...
```

**Зависимости**:
- Метод `load_income_levels_df()` в DatabaseManager
- Метод `_should_include_level()` для проверки уровня

#### 1.3. Добавление всех меток из кода 1С

**Проблема**: Многие метки из кода 1С не обрабатываются.

**Решение**: Расширить словарь `replacements` в методе `_replace_placeholders`, добавив все метки из кода 1С.

**Метки, которые нужно добавить**:
- `<IzmVs>` - изменение в связи с...
- `<UvOst>`, `<UmOst>` - удельные веса остатков
- `<Vibor1>`, `<Vibor2>` - варианты текста
- И другие метки из кода 1С

### Этап 2: Доработка SolutionController

#### 2.1. Сохранение данных решений в БД

**Проблема**: Данные решений парсятся, но не сохраняются в БД.

**Решение**: Создать таблицы для хранения данных решений:
```sql
CREATE TABLE IF NOT EXISTS solution_data (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    project_id INTEGER NOT NULL,
    solution_file_path TEXT,
    parsed_at TEXT NOT NULL,
    FOREIGN KEY (project_id) REFERENCES projects(id)
);

CREATE TABLE IF NOT EXISTS solution_income_data (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    solution_id INTEGER NOT NULL,
    код TEXT,
    наименование TEXT,
    уровень INTEGER,
    ТТ INTEGER,
    сумма1 REAL,
    сумма2 REAL,
    сумма3 REAL,
    FOREIGN KEY (solution_id) REFERENCES solution_data(id)
);

CREATE TABLE IF NOT EXISTS solution_expense_data (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    solution_id INTEGER NOT NULL,
    код_Р TEXT,
    код_ПР TEXT,
    код_ЦС TEXT,
    код_ВР TEXT,
    уровень INTEGER,
    сумма1 REAL,
    сумма2 REAL,
    сумма3 REAL,
    FOREIGN KEY (solution_id) REFERENCES solution_data(id)
);

CREATE TABLE IF NOT EXISTS solution_expense_grbs_data (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    solution_id INTEGER NOT NULL,
    ГРБС TEXT,
    код_Р TEXT,
    код_ПР TEXT,
    код_ЦС TEXT,
    код_ВР TEXT,
    уровень INTEGER,
    сумма1 REAL,
    FOREIGN KEY (solution_id) REFERENCES solution_data(id)
);
```

**Метод сохранения**:
```python
def save_solution_data(self, project_id: int, solution_data: Dict[str, Any], file_path: str) -> int:
    """Сохранение данных решения в БД"""
    with sqlite3.connect(self.db_manager.db_path) as conn:
        cursor = conn.cursor()
        
        # Сохраняем основную запись решения
        cursor.execute('''
            INSERT INTO solution_data (project_id, solution_file_path, parsed_at)
            VALUES (?, ?, ?)
        ''', (project_id, file_path, datetime.now().isoformat()))
        solution_id = cursor.lastrowid
        
        # Сохраняем данные приложений
        self._save_appendix1_data(cursor, solution_id, solution_data.get('приложение1', []))
        self._save_appendix2_data(cursor, solution_id, solution_data.get('приложение2', []))
        self._save_appendix3_data(cursor, solution_id, solution_data.get('приложение3', []))
        
        conn.commit()
        return solution_id
```

#### 2.2. Агрегация данных решений

**Проблема**: Данные группируются, но не агрегируются по уровням.

**Решение**: Добавить методы агрегации:
```python
def aggregate_solution_data(self, solution_id: int) -> Dict[str, Any]:
    """Агрегация данных решения по уровням"""
    # Загружаем данные из БД
    income_data = self._load_solution_income_data(solution_id)
    expense_data = self._load_solution_expense_data(solution_id)
    
    # Агрегируем по уровням
    aggregated_income = self._aggregate_by_level(income_data)
    aggregated_expense = self._aggregate_by_level(expense_data)
    
    return {
        'доходы': aggregated_income,
        'расходы': aggregated_expense
    }
```

#### 2.3. Создание записей в справочниках

**Проблема**: Если код не найден в справочнике, нужно создать новую запись.

**Решение**:
```python
def _find_or_create_income_code(self, code: str) -> Dict[str, Any]:
    """Поиск кода в справочнике или создание новой записи"""
    код_д = self._find_income_code(code)
    
    if not код_д:
        # Создаем новую запись в справочнике
        код_д = self._create_income_code(code)
        logger.info(f"Создана новая запись в справочнике доходов: {code}")
    
    return код_д

def _create_income_code(self, code: str) -> Dict[str, Any]:
    """Создание новой записи в справочнике доходов"""
    # Определяем уровень по структуре кода
    уровень = self._determine_income_level(code)
    
    # Сохраняем в БД
    with sqlite3.connect(self.db_manager.db_path) as conn:
        cursor = conn.cursor()
        cursor.execute('''
            INSERT INTO ref_income_codes (код, наименование, уровень)
            VALUES (?, ?, ?)
        ''', (code, '', уровень))
        conn.commit()
    
    return {
        'код': code,
        'наименование': '',
        'уровень': уровень
    }
```

### Этап 3: Интеграция с UI

#### 3.1. Диалог для формирования заключений

**Структура диалога**:
- Выбор проекта и ревизии
- Ввод даты и номера протокола
- Ввод дополнительных данных (дата письма, номер письма, дата постановления администрации)
- Кнопка "Сформировать заключение"

**Интеграция с DocumentController**:
```python
def on_generate_conclusion(self):
    """Обработчик кнопки формирования заключения"""
    project_id = self.project_combo.currentData()
    revision_id = self.revision_combo.currentData()
    protocol_date = self.protocol_date.date().toPyDate()
    protocol_number = self.protocol_number.text()
    
    result = self.main_controller.document_controller.generate_conclusion(
        project_id=project_id,
        revision_id=revision_id,
        protocol_date=protocol_date,
        protocol_number=protocol_number,
        letter_date=self.letter_date.date().toPyDate() if self.letter_date.date() else None,
        letter_number=self.letter_number.text() or None,
        admin_date=self.admin_date.date().toPyDate() if self.admin_date.date() else None,
        admin_number=self.admin_number.text() or None
    )
    
    if result:
        QMessageBox.information(self, "Успех", f"Заключение сформировано: {result}")
```

#### 3.2. Диалог для загрузки решений

**Структура диалога**:
- Выбор проекта
- Выбор файла Word с решением
- Кнопка "Загрузить и обработать"
- Отображение результатов обработки (количество доходов, расходов)

**Интеграция с SolutionController**:
```python
def on_load_solution(self):
    """Обработчик кнопки загрузки решения"""
    project_id = self.project_combo.currentData()
    file_path, _ = QFileDialog.getOpenFileName(
        self, "Выберите файл решения", "", "Word Documents (*.docx)"
    )
    
    if not file_path:
        return
    
    result = self.main_controller.solution_controller.parse_solution_document(
        file_path=file_path,
        project_id=project_id
    )
    
    if result:
        # Сохраняем данные в БД
        solution_id = self.main_controller.solution_controller.save_solution_data(
            project_id=project_id,
            solution_data=result,
            file_path=file_path
        )
        
        # Показываем результаты
        self.show_results(result, solution_id)
```

### Этап 4: Работа со справочниками

#### 4.1. Методы загрузки справочников из Excel

**Проблема**: Справочники нужно загружать из Excel файлов (из приказов Минфина).

**Решение**: Создать методы загрузки:
```python
def load_income_codes_from_excel(self, file_path: str) -> pd.DataFrame:
    """Загрузка справочника кодов доходов из Excel"""
    df = pd.read_excel(file_path, sheet_name='Коды доходов')
    
    # Преобразуем в формат для БД
    codes_df = pd.DataFrame({
        'код': df['Код'],
        'наименование': df['Наименование'],
        'уровень': df['Уровень']
    })
    
    # Сохраняем в БД
    self._save_income_codes_to_db(codes_df)
    
    return codes_df
```

#### 4.2. UI для управления справочниками

**Структура UI**:
- Вкладка "Справочники" в главном окне
- Список справочников (доходы, расходы, ГРБС, и т.д.)
- Кнопка "Загрузить из Excel" для каждого справочника
- Таблица с данными справочника
- Возможность редактирования записей

## Приоритеты реализации

### Высокий приоритет:
1. ✅ Доработка `_extract_form_data` для использования нормализованных данных
2. ✅ Доработка Таблицы 2 с фильтрацией по уровням
3. ✅ Реализация сохранения данных решений в БД
4. ✅ Создание диалогов для работы с документами

### Средний приоритет:
5. Добавление всех меток из кода 1С
6. Реализация агрегации данных решений
7. Создание записей в справочниках при отсутствии кода

### Низкий приоритет:
8. Методы загрузки справочников из Excel
9. UI для управления справочниками
10. Работа с диаграммами в Word

## Зависимости между компонентами

```
DocumentController
    ├── DatabaseManager.load_income_values_df()
    ├── DatabaseManager.load_expense_values_df()
    ├── DatabaseManager.load_income_levels_df()  [НУЖНО СОЗДАТЬ]
    └── DatabaseManager.get_municipality_by_id()

SolutionController
    ├── DatabaseManager.load_income_reference_df()
    ├── DatabaseManager.load_expense_reference_df()
    ├── DatabaseManager.save_solution_data()  [НУЖНО СОЗДАТЬ]
    └── DatabaseManager.create_income_code()  [НУЖНО СОЗДАТЬ]

UI (MainWindow)
    ├── DocumentController.generate_conclusion()
    ├── DocumentController.generate_letters()
    ├── SolutionController.parse_solution_document()
    └── SolutionController.save_solution_data()
```

## Следующие шаги

1. **Создать методы в DatabaseManager**:
   - `load_income_levels_df()` - загрузка справочника уровней доходов
   - `save_solution_data()` - сохранение данных решений
   - `create_income_code()` - создание записи в справочнике доходов
   - `create_expense_code()` - создание записи в справочнике расходов

2. **Доработать DocumentController**:
   - Переписать `_extract_form_data` для использования нормализованных данных
   - Доработать `_insert_table2` с фильтрацией по уровням
   - Добавить все метки из кода 1С

3. **Доработать SolutionController**:
   - Реализовать сохранение данных в БД
   - Реализовать агрегацию данных
   - Реализовать создание записей в справочниках

4. **Создать UI диалоги**:
   - Диалог формирования заключений
   - Диалог формирования писем
   - Диалог загрузки решений

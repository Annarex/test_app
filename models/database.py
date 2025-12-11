import sqlite3
import json
from pathlib import Path
from typing import List, Optional, Dict, Any, Iterable, Tuple
from datetime import datetime
import pandas as pd
import os
from logger import logger
from .base_models import (
    Project,
    Reference,
    ProjectStatus,
    FormType,
    YearRef,
    MunicipalityRef,
    FormTypeMeta,
    PeriodRef,
    ProjectForm,
    FormRevisionRecord,
)
from .form_0503317 import Form0503317Constants

class DatabaseManager:
    """Менеджер базы данных"""
    
    def __init__(self, db_path: str = "budget_forms.db"):
        self.db_path = db_path
        self._init_database()
    
    def _init_database(self):
        """Инициализация базы данных"""
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            
            # --------------------------------------------------
            # Базовая таблица проектов
            # Проект содержит только базовую информацию:
            # - название, год (из справочника), МО (из справочника)
            # Формы, периоды, ревизии хранятся в project_forms и form_revisions
            # --------------------------------------------------
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS projects (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    name TEXT NOT NULL,
                    year_id INTEGER,
                    municipality_id INTEGER,
                    created_at TEXT NOT NULL
                )
            ''')
            
            # Добавляем новые поля, если их нет (миграция)
            try:
                cursor.execute('ALTER TABLE projects ADD COLUMN year_id INTEGER')
            except sqlite3.OperationalError:
                pass  # Колонка уже существует
            try:
                cursor.execute('ALTER TABLE projects ADD COLUMN municipality_id INTEGER')
            except sqlite3.OperationalError:
                pass  # Колонка уже существует
            
            # Таблица справочников (метаданные файлов справочников)
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS reference_data (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    name TEXT NOT NULL,
                    reference_type TEXT NOT NULL,
                    file_path TEXT NOT NULL,
                    loaded_at TEXT NOT NULL,
                    data TEXT
                )
            ''')

            # Таблица записей справочника доходов
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS income_reference_records (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    code TEXT NOT NULL UNIQUE,
                    name TEXT,
                    level INTEGER,
                    doc TEXT
                )
            ''')

            # Таблица записей справочника источников финансирования
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS source_reference_records (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    code TEXT NOT NULL UNIQUE,
                    name TEXT,
                    level INTEGER,
                    doc TEXT
                )
            ''')

            # --------------------------------------------------
            # Новая архитектура справочников и форм проекта
            # --------------------------------------------------

            # Справочник годов (для явного выбора года проекта)
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS ref_years (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    year INTEGER NOT NULL UNIQUE,
                    is_active INTEGER NOT NULL DEFAULT 1
                )
            ''')

            # Справочник видов муниципальных образований
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS ref_municipality_types (
                    код_вида_МО VARCHAR(1) PRIMARY KEY,
                    наименование TEXT NOT NULL
                )
            ''')
            
            # Справочник муниципальных образований (расширенный)
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS ref_municipalities (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    code VARCHAR(3) UNIQUE,
                    name TEXT NOT NULL,
                    код_вида_МО VARCHAR(1) REFERENCES ref_municipality_types(код_вида_МО),
                    код_МО VARCHAR(3),  -- для обратной совместимости
                    родительный_падеж TEXT,
                    адрес_совет TEXT,
                    адрес_администрация TEXT,
                    совет_почта VARCHAR(50),
                    администрация_почта VARCHAR(50),
                    должность_совет VARCHAR(30),
                    фамилия_совет VARCHAR(30),
                    имя_совет VARCHAR(30),
                    отчество_совет VARCHAR(30),
                    должность_администрация VARCHAR(30),
                    фамилия_администрация VARCHAR(30),
                    имя_администрация VARCHAR(30),
                    отчество_администрация VARCHAR(30),
                    дата_соглашения DATE,
                    дата_решения DATE,
                    номер_решения VARCHAR(50),
                    начальная_доходы REAL,
                    начальная_расходы REAL,
                    начальная_дефицит REAL,
                    is_active INTEGER NOT NULL DEFAULT 1
                )
            ''')
            
            # Миграция: добавляем новые поля в существующую таблицу
            for col_name, col_type in [
                ('код_МО', 'VARCHAR(3)'), ('код_вида_МО', 'VARCHAR(1)'),
                ('родительный_падеж', 'TEXT'),
                ('адрес_совет', 'TEXT'), ('адрес_администрация', 'TEXT'),
                ('совет_почта', 'VARCHAR(50)'), ('администрация_почта', 'VARCHAR(50)'),
                ('должность_совет', 'VARCHAR(30)'), ('фамилия_совет', 'VARCHAR(30)'),
                ('имя_совет', 'VARCHAR(30)'), ('отчество_совет', 'VARCHAR(30)'),
                ('должность_администрация', 'VARCHAR(30)'), ('фамилия_администрация', 'VARCHAR(30)'),
                ('имя_администрация', 'VARCHAR(30)'), ('отчество_администрация', 'VARCHAR(30)'),
                ('дата_соглашения', 'DATE'), ('дата_решения', 'DATE'),
                ('номер_решения', 'VARCHAR(50)'),
                ('начальная_доходы', 'REAL'), ('начальная_расходы', 'REAL'),
                ('начальная_дефицит', 'REAL')
            ]:
                try:
                    cursor.execute(f'ALTER TABLE ref_municipalities ADD COLUMN {col_name} {col_type}')
                except sqlite3.OperationalError:
                    pass  # Колонка уже существует

            # Справочник типов форм (0503317, 0503314 и т.д.)
            # ID задаём вручную в коде (не полагаемся на AUTOINCREMENT),
            # чтобы иметь стабильные идентификаторы типов форм.
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS ref_form_types (
                    id INTEGER PRIMARY KEY,
                    code TEXT NOT NULL UNIQUE,
                    name TEXT,
                    periodicity TEXT,         -- годовая, квартальная, полугодовая и т.п.
                    column_mapping TEXT,       -- JSON с mapping колонок для экспорта/валидации
                    is_active INTEGER NOT NULL DEFAULT 1
                )
            ''')
            
            # Миграция: добавляем поле column_mapping в существующую таблицу
            try:
                cursor.execute('ALTER TABLE ref_form_types ADD COLUMN column_mapping TEXT')
            except sqlite3.OperationalError:
                pass  # Колонка уже существует

            # Справочник периодов (расширенный)
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS ref_periods (
                    код_периода VARCHAR(2),
                    наименование VARCHAR(30),
                    отчет_на_дату DATE,
                    id INTEGER PRIMARY KEY,  -- для обратной совместимости
                    code TEXT NOT NULL,       -- Y, Q1, Q2, Q3, Q4, H1, H2 и т.п.
                    sort_order INTEGER NOT NULL DEFAULT 0,
                    form_type_code TEXT,      -- опциональная привязка к форме
                    is_active INTEGER NOT NULL DEFAULT 1
                )
            ''')
            
            # Миграция: добавляем новые поля
            for col_name, col_type in [
                ('код_периода', 'VARCHAR(2)'), ('наименование', 'VARCHAR(30)'),
                ('отчет_на_дату', 'DATE')
            ]:
                try:
                    cursor.execute(f'ALTER TABLE ref_periods ADD COLUMN {col_name} {col_type}')
                except sqlite3.OperationalError:
                    pass

            # Связка Проект ↔ Форма ↔ Период
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS project_forms (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    project_id INTEGER NOT NULL,
                    form_type_id INTEGER NOT NULL,
                    period_id INTEGER,
                    UNIQUE(project_id, form_type_id, period_id)
                )
            ''')

            # Ревизии форм в рамках project_form
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS form_revisions (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    project_form_id INTEGER NOT NULL,
                    revision TEXT NOT NULL,
                    status TEXT,
                    file_path TEXT,
                    created_at TEXT NOT NULL,
                    UNIQUE(project_form_id, revision)
                )
            ''')

            # Метаданные ревизий (отдельная таблица для каждой ревизии)
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS revision_metadata (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    revision_id INTEGER NOT NULL UNIQUE,
                    meta_info TEXT,
                    результат_исполнения_data TEXT,
                    FOREIGN KEY (revision_id) REFERENCES form_revisions(id) ON DELETE CASCADE
                )
            ''')

            # Справочники для классификаций расходов бюджетов
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS ref_grbs (
                    код_ГРБС VARCHAR(3) PRIMARY KEY,
                    наименование TEXT NOT NULL
                )
            ''')
            
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS ref_expense_sections (
                    код_РП VARCHAR(4) PRIMARY KEY,
                    наименование TEXT NOT NULL,
                    утверждающий_документ TEXT
                )
            ''')
            
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS ref_target_articles (
                    код_ЦСР VARCHAR(10) PRIMARY KEY,
                    наименование TEXT NOT NULL
                )
            ''')
            
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS ref_expense_types (
                    код_вида_СР VARCHAR(5) PRIMARY KEY,
                    наименование TEXT NOT NULL
                )
            ''')
            
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS ref_program_nonprogram (
                    код_ПНС VARCHAR(5) PRIMARY KEY,
                    наименование TEXT NOT NULL
                )
            ''')
            
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS ref_expense_kinds (
                    код_ВР VARCHAR(3) PRIMARY KEY,
                    наименование TEXT NOT NULL,
                    утверждающий_документ TEXT
                )
            ''')
            
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS ref_national_projects (
                    код_НПЦСР VARCHAR(1) PRIMARY KEY,
                    наименование TEXT NOT NULL,
                    утверждающий_документ TEXT
                )
            ''')
            
            # Справочники для классификаций доходов бюджетов
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS ref_gadb (
                    код_ГАДБ VARCHAR(3) PRIMARY KEY,
                    наименование TEXT
                )
            ''')
            
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS ref_income_groups (
                    код_группы_ДБ VARCHAR(1) PRIMARY KEY,
                    наименование TEXT
                )
            ''')
            
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS ref_income_subgroups (
                    код_подгруппы_ДБ VARCHAR(2) PRIMARY KEY,
                    наименование TEXT
                )
            ''')
            
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS ref_income_articles (
                    код_статьи_подстатьи_ДБ VARCHAR(5) PRIMARY KEY,
                    наименование TEXT
                )
            ''')
            
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS ref_income_elements (
                    код_элемента_ДБ VARCHAR(2) PRIMARY KEY,
                    наименование TEXT
                )
            ''')
            
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS ref_income_subkind_groups (
                    код_группы_ПДБ VARCHAR(4) PRIMARY KEY,
                    наименование TEXT
                )
            ''')
            
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS ref_income_analytical_groups (
                    код_группы_АПДБ VARCHAR(3) PRIMARY KEY,
                    наименование TEXT
                )
            ''')
            
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS ref_income_levels (
                    код_уровня VARCHAR(2) PRIMARY KEY,
                    наименование TEXT,
                    цвет VARCHAR(10)
                )
            ''')
            
            # Справочники кодов доходов и расходов (для работы с решениями)
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS ref_income_codes (
                    код VARCHAR(20) PRIMARY KEY,
                    название TEXT,
                    уровень INTEGER,
                    наименование TEXT
                )
            ''')
            
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS ref_expense_codes (
                    код VARCHAR(20) PRIMARY KEY,
                    название TEXT,
                    уровень INTEGER,
                    код_Р VARCHAR(4),
                    код_ПР VARCHAR(2),
                    код_ЦС VARCHAR(10),
                    код_ВР VARCHAR(3)
                )
            ''')
            
            # Миграция: переименовываем ЛВЛ в уровень, если таблицы уже существуют
            try:
                cursor.execute('ALTER TABLE ref_income_codes RENAME COLUMN ЛВЛ TO уровень')
            except sqlite3.OperationalError:
                pass  # Колонка уже переименована или не существует
            
            try:
                cursor.execute('ALTER TABLE ref_expense_codes RENAME COLUMN ЛВЛ TO уровень')
            except sqlite3.OperationalError:
                pass  # Колонка уже переименована или не существует

            # Первичное заполнение справочников (если они пустые)
            self._seed_config_dictionaries(cursor)

            # --------------------------------------------------
            # НОВЫЕ НОРМАЛИЗОВАННЫЕ ТАБЛИЦЫ ДЛЯ ЗНАЧЕНИЙ РАЗДЕЛОВ
            # --------------------------------------------------
            budget_cols = Form0503317Constants.BUDGET_COLUMNS
            consolidated_cols = Form0503317Constants.CONSOLIDATED_COLUMNS

            def _build_value_columns_sql(col_count: int) -> str:
                # v1..vN – значения по столбцам бюджета; отображение в человекочитаемые
                # названия хранится только в коде, а не в БД (mapping в программной части).
                return ", ".join(f"v{i+1} REAL" for i in range(col_count))

            budget_values_columns_sql = _build_value_columns_sql(len(budget_cols))
            consolidated_values_columns_sql = _build_value_columns_sql(len(consolidated_cols))

            # Доходы / Расходы / Источники – общая схема
            for table_name in ("income_values", "expense_values", "source_values"):
                cursor.execute(
                    f'''
                    CREATE TABLE IF NOT EXISTS {table_name} (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    project_id INTEGER NOT NULL,
                    revision_id INTEGER,
                        classification_code TEXT,
                        indicator_name TEXT,
                        line_code TEXT,
                        budget_type TEXT NOT NULL,   -- 'утвержденный' / 'исполненный'
                        data_type TEXT NOT NULL,     -- 'оригинальные' / 'вычисленные'
                        level INTEGER,                -- уровень строки (0-6), кэшируется для ускорения расчетов
                        source_row INTEGER,          -- исходная строка в Excel (для экспорта)
                        {budget_values_columns_sql}
                    )
                    '''
                )
                # Миграция: добавляем новые поля в существующие таблицы
                for col_name, col_type in [('level', 'INTEGER'), ('source_row', 'INTEGER')]:
                    try:
                        cursor.execute(f'ALTER TABLE {table_name} ADD COLUMN {col_name} {col_type}')
                    except sqlite3.OperationalError:
                        pass  # Колонка уже существует

            # Консолидируемые расчёты – отдельная таблица
            cursor.execute(
                f'''
                CREATE TABLE IF NOT EXISTS consolidated_values (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    project_id INTEGER NOT NULL,
                    revision_id INTEGER,
                    classification_code TEXT,
                    indicator_name TEXT,
                    line_code TEXT,
                    budget_type TEXT NOT NULL,   -- всегда 'поступления'
                    data_type TEXT NOT NULL,     -- 'оригинальные' / 'вычисленные'
                    level INTEGER,                -- уровень строки (0-2), кэшируется для ускорения расчетов
                    source_row INTEGER,          -- исходная строка в Excel (для экспорта)
                    {consolidated_values_columns_sql}
                )
                '''
            )
            # Миграция: добавляем новые поля в существующую таблицу
            for col_name, col_type in [('level', 'INTEGER'), ('source_row', 'INTEGER')]:
                try:
                    cursor.execute(f'ALTER TABLE consolidated_values ADD COLUMN {col_name} {col_type}')
                except sqlite3.OperationalError:
                    pass  # Колонка уже существует

            # Индексы для ускорения выборок по проекту/ревизии/коду
            for tbl in ('income_values', 'expense_values', 'source_values', 'consolidated_values'):
                try:
                    cursor.execute(
                        f'CREATE INDEX IF NOT EXISTS idx_{tbl}_proj_rev '
                        f'ON {tbl} (project_id, revision_id)'
                    )
                except sqlite3.OperationalError:
                    pass  # Индекс уже существует
                try:
                    cursor.execute(
                        f'CREATE INDEX IF NOT EXISTS idx_{tbl}_class_code '
                        f'ON {tbl} (classification_code)'
                    )
                except sqlite3.OperationalError:
                    pass
                # Индекс по уровню для агрегирования и фильтрации
                try:
                    cursor.execute(
                        f'CREATE INDEX IF NOT EXISTS idx_{tbl}_level '
                        f'ON {tbl} (level)'
                    )
                except sqlite3.OperationalError:
                    pass
                # Индекс по исходной строке для экспорта
                try:
                    cursor.execute(
                        f'CREATE INDEX IF NOT EXISTS idx_{tbl}_source_row '
                        f'ON {tbl} (source_row)'
                    )
                except sqlite3.OperationalError:
                    pass  # Индекс уже существует

            conn.commit()
            
            # Автозагрузка справочников при создании новой БД
            self._auto_load_references_if_new()
    
    def _auto_load_references_if_new(self):
        """
        Автоматическая загрузка справочников при создании новой БД.
        Проверяет, что БД только создана (нет записей в reference_data),
        и если есть файлы по стандартным путям, загружает их.
        """
        try:
            with sqlite3.connect(self.db_path) as conn:
                cursor = conn.cursor()
                cursor.execute('SELECT COUNT(*) FROM reference_data')
                count = cursor.fetchone()[0]
                
                # Если справочники уже есть, не загружаем автоматически
                if count > 0:
                    return
                
                # Пути к стандартным файлам справочников
                ref_paths = {
                    'доходы': Path('data/references/Классификация_доходов_бюджетов_с_полным_кодом.xls'),
                    'источники': Path('data/references/Классификация_источников_финансирования_дифицитов.xls')
                }
                
                # Пробуем также с расширением .xlsx
                ref_paths_xlsx = {
                    'доходы': Path('data/references/Классификация_доходов_бюджетов_с_полным_кодом.xlsx'),
                    'источники': Path('data/references/Классификация_источников_финансирования_дифицитов.xlsx')
                }
                
                # Загружаем справочники, если файлы существуют
                for ref_type, file_path in ref_paths.items():
                    if not file_path.exists():
                        # Пробуем с .xlsx
                        file_path = ref_paths_xlsx.get(ref_type)
                        if not file_path or not file_path.exists():
                            logger.warning(f"Автозагрузка справочника '{ref_type}': файл не найден по пути {file_path}")
                            continue
                    
                    try:
                        logger.info(f"Автозагрузка справочника '{ref_type}' из {file_path}")
                        
                        # Определяем имя справочника
                        name = file_path.stem
                        if ref_type == 'доходы':
                            name = 'Классификация доходов бюджетов'
                        elif ref_type == 'источники':
                            name = 'Классификация источников финансирования дефицитов'
                        
                        # Загружаем справочник напрямую через методы БД
                        import pandas as pd
                        df = pd.read_excel(str(file_path))
                        df.columns = [str(c).strip() for c in df.columns]
                        
                        # Определяем колонку с кодом классификации
                        code_column = None
                        if ref_type == 'доходы' and 'код_классификации_ДБ' in df.columns:
                            code_column = 'код_классификации_ДБ'
                        elif ref_type == 'источники' and 'код_классификации_ИФДБ' in df.columns:
                            code_column = 'код_классификации_ИФДБ'
                        
                        if code_column:
                            df[code_column] = (
                                df[code_column]
                                .astype(str)
                                .str.strip()
                                .str.replace(' ', '', regex=False)
                                .str.replace('\u00A0', '', regex=False)
                                .str.zfill(20)
                            )
                        
                        # Создаем объект справочника
                        reference = Reference()
                        reference.name = name
                        reference.reference_type = ref_type
                        reference.file_path = str(file_path)
                        
                        # Сохраняем в БД (метаданные)
                        self.save_reference(reference)
                        
                        # Сохраняем строки справочника
                        reference_data = df.to_dict('records')
                        self.save_reference_records(ref_type, reference_data)
                        
                        logger.info(f"Справочник '{ref_type}' успешно загружен автоматически")
                    except Exception as e:
                        logger.error(f"Ошибка автозагрузки справочника '{ref_type}': {e}", exc_info=True)
        except Exception as e:
            logger.error(f"Ошибка при проверке необходимости автозагрузки справочников: {e}", exc_info=True)
            # Не блокируем работу приложения из-за ошибки автозагрузки
    
    def _seed_config_dictionaries(self, cursor: sqlite3.Cursor) -> None:
        """
        Первичное заполнение справочников годов, типов форм и периодов,
        если таблицы пусты. Это позволяет сразу работать с типовой конфигурацией
        (форма 0503317, годовой и квартальные периоды).
        """
        # ref_years
        cursor.execute('SELECT COUNT(*) FROM ref_years')
        count_years = cursor.fetchone()[0]
        if count_years == 0:
            current_year = datetime.now().year
            years = [(current_year - 1, 1), (current_year, 1), (current_year + 1, 1)]
            cursor.executemany(
                'INSERT INTO ref_years (year, is_active) VALUES (?, ?)',
                years,
            )

        # ref_form_types
        cursor.execute('SELECT COUNT(*) FROM ref_form_types')
        count_forms = cursor.fetchone()[0]
        if count_forms == 0:
            # Базовая форма 0503317 (годовая/квартальная)
            # Сохраняем mapping колонок для формы 0503317
            from models.form_0503317 import Form0503317Constants
            mapping_json = json.dumps(Form0503317Constants.COLUMN_MAPPING, ensure_ascii=False)
            forms = [
                (503317, "0503317", "Форма 0503317", "Квартальная/6М/9М/12М", mapping_json, 1),
            ]
            cursor.executemany(
                'INSERT INTO ref_form_types (id, code, name, periodicity, column_mapping, is_active) '
                'VALUES (?, ?, ?, ?, ?, ?)',
                forms,
            )

        # ref_periods
        cursor.execute('SELECT COUNT(*) FROM ref_periods')
        count_periods = cursor.fetchone()[0]
        if count_periods == 0:
            # Общие периоды: год и кварталы
            periods = [
                ("Y", "Год", 0, None, 1),
                ("Q1", "I квартал", 1, None, 1),
                ("Q2", "II квартал", 2, None, 1),
                ("Q3", "III квартал", 3, None, 1),
                ("Q4", "IV квартал", 4, None, 1),
                ("M6", "6 месяцев", 5, None, 1),
                ("M9", "9 месяцев", 6, None, 1),
            ]
            cursor.executemany(
                'INSERT INTO ref_periods (code, наименование, sort_order, form_type_code, is_active) '
                'VALUES (?, ?, ?, ?, ?)',
                periods,
            )
        
        # ref_municipalities
        cursor.execute('SELECT COUNT(*) FROM ref_municipalities')
        count_municipalities = cursor.fetchone()[0]
        if count_municipalities == 0:
            # Предзагруженные муниципальные образования
            municipalities = [
                (None, "Амвросиевка", 1),
                (None, "Волноваха", 1),
                (None, "Володарка", 1),
                (None, "Горловка", 1),
                (None, "Дебальцево", 1),
                (None, "Докучаевск", 1),
                (None, "Донецк", 1),
                (None, "Енакиево", 1),
                (None, "Иловайск", 1),
                (None, "Красный лиман", 1),
                (None, "Макеевка", 1),
                (None, "Мангуш", 1),
                (None, "Мариуполь", 1),
                (None, "Новозаовск", 1),
                (None, "Снежное", 1),
                (None, "Старобешево", 1),
                (None, "Тельманово", 1),
                (None, "Торез", 1),
                (None, "Харцызск", 1),
                (None, "Шахтерск", 1),
                (None, "Ясиноватая", 1),
            ]
            cursor.executemany(
                'INSERT INTO ref_municipalities (code, name, is_active) VALUES (?, ?, ?)',
                municipalities,
            )
    
    def save_project(self, project: Project) -> int:
        """Сохранение проекта в БД (новая архитектура).

        В таблице projects теперь храним только базовые поля проекта:
        - id, name, year_id, municipality_id, created_at.
        Вся информация о формах, периодах и ревизиях хранится в project_forms / form_revisions.
        """
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()

            year_id = project.year_id
            municipality_id = project.municipality_id

            if project.id is None:
                cursor.execute(
                    '''
                    INSERT INTO projects (name, year_id, municipality_id, created_at)
                    VALUES (?, ?, ?, ?)
                    ''',
                    (
                        project.name,
                        year_id,
                        municipality_id,
                        project.created_at.isoformat(),
                    ),
                )
                project.id = cursor.lastrowid
            else:
                cursor.execute(
                    '''
                    UPDATE projects SET
                        name=?,
                        year_id=?,
                        municipality_id=?,
                        created_at=?
                    WHERE id=?
                    ''',
                    (
                        project.name,
                        year_id,
                        municipality_id,
                        project.created_at.isoformat(),
                        project.id,
                    ),
                )

            conn.commit()
            return project.id
    
    def _save_project_data(self, cursor, project_id: int, data: Dict[str, Any], revision_id: Optional[int] = None):
        """
        Сохранение данных проекта в нормализованные таблицы *_values.

        Старые JSON‑таблицы *_rows больше не используются для новых/обновлённых ревизий.
        """
        # Удаляем старые данные для этой ревизии (если указана) или для всего проекта
        if revision_id:
            params = (project_id, revision_id)
            cursor.execute('DELETE FROM income_values WHERE project_id=? AND revision_id=?', params)
            cursor.execute('DELETE FROM expense_values WHERE project_id=? AND revision_id=?', params)
            cursor.execute('DELETE FROM source_values WHERE project_id=? AND revision_id=?', params)
            cursor.execute('DELETE FROM consolidated_values WHERE project_id=? AND revision_id=?', params)
        else:
            params = (project_id,)
            cursor.execute('DELETE FROM income_values WHERE project_id=?', params)
            cursor.execute('DELETE FROM expense_values WHERE project_id=?', params)
            cursor.execute('DELETE FROM source_values WHERE project_id=?', params)
            cursor.execute('DELETE FROM consolidated_values WHERE project_id=?', params)
        
        # Сохраняем нормализованные значения в *_values
        budget_cols = Form0503317Constants.BUDGET_COLUMNS
        consolidated_cols = Form0503317Constants.CONSOLIDATED_COLUMNS

        self._save_section_values(
            cursor=cursor,
            project_id=project_id,
            revision_id=revision_id,
            section_rows=data.get('доходы_data') or [],
            table_name='income_values',
            budget_columns=budget_cols,
        )
        self._save_section_values(
            cursor=cursor,
            project_id=project_id,
            revision_id=revision_id,
            section_rows=data.get('расходы_data') or [],
            table_name='expense_values',
            budget_columns=budget_cols,
        )
        self._save_section_values(
            cursor=cursor,
            project_id=project_id,
            revision_id=revision_id,
            section_rows=data.get('источники_финансирования_data') or [],
            table_name='source_values',
            budget_columns=budget_cols,
        )
        self._save_consolidated_values(
            cursor=cursor,
            project_id=project_id,
            revision_id=revision_id,
            section_rows=data.get('консолидируемые_расчеты_data') or [],
            table_name='consolidated_values',
            consolidated_columns=consolidated_cols,
        )

    # ------------------------------------------------------------------
    # ВСПОМОГАТЕЛЬНЫЕ МЕТОДЫ ДЛЯ *_values И МИГРАЦИИ
    # ------------------------------------------------------------------

    def _iter_value_rows_for_budget_section(
        self,
        project_id: int,
        revision_id: Optional[int],
        section_rows: Iterable[Dict[str, Any]],
        budget_columns: List[str],
        only_calculated: bool = False,
    ) -> Iterable[Tuple]:
        """
        Преобразует одну логическую строку раздела (доходы/расходы/источники)
        в набор строк для *_values в формате:
        (project_id, revision_id, classification_code, indicator_name, line_code,
         budget_type, data_type, v1..vN)
        
        Args:
            only_calculated: Если True, возвращает только вычисленные значения
        """
        for row in section_rows:
            if not isinstance(row, dict):
                continue

            classification_code = row.get('код_классификации') or None
            indicator_name = row.get('наименование_показателя') or None
            line_code = row.get('код_строки') or None
            level = row.get('уровень')  # Кэшированный уровень
            source_row = row.get('исходная_строка')  # Исходная строка для экспорта

            # ОРИГИНАЛЬНЫЕ значения (пропускаем, если only_calculated=True)
            if not only_calculated:
                for budget_type in ('утвержденный', 'исполненный'):
                    values_dict = row.get(budget_type) or {}
                    vec = []
                    has_value = False
                    for col in budget_columns:
                        v = values_dict.get(col)
                        # 'x' трактуем как отсутствие значения
                        if isinstance(v, str) and v.lower() == 'x':
                            v = None
                        if v is not None:
                            has_value = True
                        vec.append(v)
                    if has_value:
                        yield (
                            project_id,
                            revision_id,
                            classification_code,
                            indicator_name,
                            line_code,
                            budget_type,
                            'оригинальные',
                            level,
                            source_row,
                            *vec,
                        )

            # ВЫЧИСЛЕННЫЕ значения (если есть расчетные колонки)
            # Ключи формата: 'расчетный_утвержденный_{budget_col}'
            for budget_type, prefix in (('утвержденный', 'расчетный_утвержденный_'),
                                        ('исполненный', 'расчетный_исполненный_')):
                vec = []
                has_value = False
                for col in budget_columns:
                    key = f'{prefix}{col}'
                    v = row.get(key)
                    if isinstance(v, str) and v.lower() == 'x':
                        v = None
                    if v is not None:
                        has_value = True
                    vec.append(v)
                if has_value:
                    yield (
                        project_id,
                        revision_id,
                        classification_code,
                        indicator_name,
                        line_code,
                        budget_type,
                        'вычисленные',
                        level,
                        source_row,
                        *vec,
                    )

    def _iter_value_rows_for_consolidated_section(
        self,
        project_id: int,
        revision_id: Optional[int],
        section_rows: Iterable[Dict[str, Any]],
        consolidated_columns: List[str],
        only_calculated: bool = False,
    ) -> Iterable[Tuple]:
        """
        Преобразует одну логическую строку консолидируемых расчётов
        в набор строк для consolidated_values:
        (project_id, revision_id, classification_code, indicator_name, line_code,
         budget_type, data_type, level, source_row, v1..vN)
        
        Args:
            only_calculated: Если True, возвращает только вычисленные значения
        """
        for row in section_rows:
            if not isinstance(row, dict):
                continue

            classification_code = row.get('код_классификации') or None
            indicator_name = row.get('наименование_показателя') or None
            line_code = row.get('код_строки') or None
            level = row.get('уровень')  # Кэшированный уровень
            source_row = row.get('исходная_строка')  # Исходная строка для экспорта

            # ОРИГИНАЛЬНЫЕ значения (словарь 'поступления')
            if not only_calculated:
                vec = []
                has_value = False
                source_dict = row.get('поступления') or {}
                for col in consolidated_columns:
                    v = source_dict.get(col)
                    if isinstance(v, str) and v.lower() == 'x':
                        v = None
                    if v is not None:
                        has_value = True
                    vec.append(v)
                if has_value:
                    yield (
                        project_id,
                        revision_id,
                        classification_code,
                        indicator_name,
                        line_code,
                        'поступления',
                        'оригинальные',
                        level,
                        source_row,
                        *vec,
                    )

            # ВЫЧИСЛЕННЫЕ значения
            # Для консолидированных расчетов сохраняем вычисленные значения, если они есть
            # (даже если они равны оригинальным - это важно для сравнения)
            vec = []
            has_calculated_field = False
            for col in consolidated_columns:
                key = f'расчетный_поступления_{col}'
                v = row.get(key)
                # Проверяем наличие ключа в словаре (не только значение)
                if key in row:
                    has_calculated_field = True
                if isinstance(v, str) and v.lower() == 'x':
                    v = None
                vec.append(v)
            
            # Сохраняем вычисленные значения, если поле 'расчетный_поступления_*' присутствует в словаре
            # Это важно для отображения сравнения оригинальных и расчетных значений
            if has_calculated_field:
                yield (
                    project_id,
                    revision_id,
                    classification_code,
                    indicator_name,
                    line_code,
                    'поступления',
                    'вычисленные',
                    level,
                    source_row,
                    *vec,
                )

    def _save_section_values(
        self,
        cursor: sqlite3.Cursor,
        project_id: int,
        revision_id: Optional[int],
        section_rows: List[Dict[str, Any]],
        table_name: str,
        budget_columns: List[str],
        only_calculated: bool = False,
    ) -> None:
        """Сохраняет нормализованные значения для разделов доходы/расходы/источники.
        
        Args:
            only_calculated: Если True, сохраняет только вычисленные значения (data_type='вычисленные')
        """
        # Данные раздела могут отсутствовать (например, если форма без этого блока)
        # В таком случае просто ничего не делаем.
        # Очищение по project_id/revision_id уже сделано в _save_project_data.
        from itertools import islice

        if not section_rows:
            return

        value_rows_iter = self._iter_value_rows_for_budget_section(
            project_id=project_id,
            revision_id=revision_id,
            section_rows=section_rows,
            budget_columns=budget_columns,
            only_calculated=only_calculated,
        )

        # Предварительно берём несколько строк, чтобы не выполнять пустой executemany
        first_batch = list(islice(value_rows_iter, 1000))
        if not first_batch:
            return

        placeholders = ", ".join(["?"] * (9 + len(budget_columns)))
        cursor.executemany(
            f'''
            INSERT INTO {table_name} (
                project_id, revision_id, classification_code, indicator_name,
                line_code, budget_type, data_type, level, source_row,
                {", ".join(f"v{i+1}" for i in range(len(budget_columns)))}
            )
            VALUES ({placeholders})
            ''',
            first_batch,
        )

        # Дозагружаем остальные строки батчами
        batch_size = 1000
        batch = list(islice(value_rows_iter, batch_size))
        while batch:
            cursor.executemany(
                f'''
                INSERT INTO {table_name} (
                    project_id, revision_id, classification_code, indicator_name,
                    line_code, budget_type, data_type, level, source_row,
                    {", ".join(f"v{i+1}" for i in range(len(budget_columns)))}
                )
                VALUES ({placeholders})
                ''',
                batch,
            )
            batch = list(islice(value_rows_iter, batch_size))

    def _save_consolidated_values(
        self,
        cursor: sqlite3.Cursor,
        project_id: int,
        revision_id: Optional[int],
        section_rows: List[Dict[str, Any]],
        table_name: str,
        consolidated_columns: List[str],
        only_calculated: bool = False,
    ) -> None:
        """Сохраняет нормализованные значения для консолидируемых расчётов.
        
        Args:
            only_calculated: Если True, сохраняет только вычисленные значения (data_type='вычисленные')
        """
        from itertools import islice
        if not section_rows:
            return
        value_rows_iter = self._iter_value_rows_for_consolidated_section(
            project_id=project_id,
            revision_id=revision_id,
            section_rows=section_rows,
            consolidated_columns=consolidated_columns,
            only_calculated=only_calculated,
        )

        first_batch = list(islice(value_rows_iter, 1000))
        if not first_batch:
            return

        placeholders = ", ".join(["?"] * (9 + len(consolidated_columns)))
        cursor.executemany(
            f'''
            INSERT INTO {table_name} (
                project_id, revision_id, classification_code, indicator_name,
                line_code, budget_type, data_type, level, source_row,
                {", ".join(f"v{i+1}" for i in range(len(consolidated_columns)))}
            )
            VALUES ({placeholders})
            ''',
            first_batch,
        )

        batch_size = 1000
        batch = list(islice(value_rows_iter, batch_size))
        while batch:
            cursor.executemany(
                f'''
                INSERT INTO {table_name} (
                    project_id, revision_id, classification_code, indicator_name,
                    line_code, budget_type, data_type, level, source_row,
                    {", ".join(f"v{i+1}" for i in range(len(consolidated_columns)))}
                )
                VALUES ({placeholders})
                ''',
                batch,
            )
            batch = list(islice(value_rows_iter, batch_size))

    # Легаси-миграция удалена: старые таблицы *_rows больше не поддерживаются.
    
    def load_projects(self) -> List[Project]:
        """Загрузка всех проектов (новая архитектура)."""
        projects: List[Project] = []
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute(
                '''
                SELECT id, name, year_id, municipality_id, created_at
                FROM projects ORDER BY created_at DESC
                '''
            )

            for row in cursor.fetchall():
                project_id = row[0]
                project_data = {
                    'id': project_id,
                    'name': row[1],
                    'year_id': row[2],
                    'municipality_id': row[3],
                    'created_at': row[4],
                    # Данные по умолчанию - пустые, данные загружаются только при загрузке ревизии
                    'data': {},
                }
                projects.append(Project.from_dict(project_data))

        return projects
    
    def _load_project_data(self, cursor, project_id: int, revision_id: Optional[int] = None) -> Dict[str, Any]:
        """
        Загрузка данных проекта.

        Данные собираются только из новых нормализованных таблиц *_values.
        Для старых проектов миграция из *_rows выполняется один раз при инициализации БД.
        """
        data_from_values = self._load_project_data_from_values(cursor, project_id, revision_id)
        return data_from_values

    def load_project_data_values(self, project_id: int, revision_id: Optional[int] = None) -> Dict[str, Any]:
        """Публичный метод загрузки данных проекта из нормализованных таблиц *_values."""
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            return self._load_project_data_from_values(cursor, project_id, revision_id)

    def _load_project_data_from_values(
        self,
        cursor: sqlite3.Cursor,
        project_id: int,
        revision_id: Optional[int] = None,
    ) -> Dict[str, Any]:
        """
        Загрузка данных проекта из нормализованных таблиц *_values.

        Восстанавливает структуру:
        - доходы_data / расходы_data / источники_финансирования_data
          с полями 'утвержденный' / 'исполненный' и, при наличии, 'расчетный_*';
          Уникальность: (classification_code, indicator_name, line_code)
        - консолидируемые_расчеты_data с полями 'поступления' и 'расчетный_поступления_*'.
          Уникальность: (indicator_name, line_code) - только наименование и код строки

        Служебные поля (mapping, исходная_строка и т.п.) здесь не восстанавливаются
        и при необходимости могут быть дополнительно подчитаны из JSON‑таблиц.
        """
        data: Dict[str, Any] = {}

        # Доходы / Расходы / Источники
        def load_budget_section_values(
            table_name: str,
            section_name: str,
        ) -> List[Dict[str, Any]]:
            where_clause = 'project_id=? AND revision_id IS ?' if revision_id is None else 'project_id=? AND revision_id=?'
            params = (project_id, revision_id)
            cursor.execute(
                f'''
                SELECT classification_code, indicator_name, line_code,
                       budget_type, data_type, level, source_row,
                       {", ".join(f"v{i+1}" for i in range(len(Form0503317Constants.BUDGET_COLUMNS)))}
                FROM {table_name}
                WHERE {where_clause}
                ORDER BY id
                ''',
                params,
            )
            rows = cursor.fetchall()
            if not rows:
                return []

            # Группируем по (classification_code, indicator_name, line_code)
            grouped: Dict[Tuple[Optional[str], Optional[str], Optional[str]], Dict[str, Any]] = {}
            for row in rows:
                classification_code, indicator_name, line_code, budget_type, data_type, level, source_row, *values = row
                key = (classification_code, indicator_name, line_code)
                if key not in grouped:
                    grouped[key] = {
                        'код_классификации': classification_code or '',
                        'наименование_показателя': indicator_name or '',
                        'код_строки': line_code or '',
                        'раздел': section_name,
                        'уровень': level,  # Используем кэшированный уровень из БД
                        'исходная_строка': source_row,  # Используем сохраненную исходную строку
                    }

                target = grouped[key]
                cols = Form0503317Constants.BUDGET_COLUMNS

                if data_type == 'оригинальные':
                    bucket = target.setdefault(budget_type, {})
                    for idx, col_name in enumerate(cols):
                        bucket[col_name] = values[idx]
                elif data_type == 'вычисленные':
                    prefix = 'расчетный_утвержденный_' if budget_type == 'утвержденный' else 'расчетный_исполненный_'
                    for idx, col_name in enumerate(cols):
                        target[f'{prefix}{col_name}'] = values[idx]

            return list(grouped.values())

        доходы = load_budget_section_values('income_values', 'доходы')
        расходы = load_budget_section_values('expense_values', 'расходы')
        источники = load_budget_section_values('source_values', 'источники_финансирования')

        if доходы:
            data['доходы_data'] = доходы
        if расходы:
            data['расходы_data'] = расходы
        if источники:
            data['источники_финансирования_data'] = источники

        # Консолидируемые расчёты
        where_clause = 'project_id=? AND revision_id IS ?' if revision_id is None else 'project_id=? AND revision_id=?'
        params = (project_id, revision_id)
        cursor.execute(
            f'''
            SELECT classification_code, indicator_name, line_code,
                   budget_type, data_type, level, source_row,
                   {", ".join(f"v{i+1}" for i in range(len(Form0503317Constants.CONSOLIDATED_COLUMNS)))}
            FROM consolidated_values
            WHERE {where_clause}
            ORDER BY id
            ''',
            params,
        )
        rows = cursor.fetchall()
        if rows:
            # Для консолидированных расчетов уникальность определяется только по наименованию и коду строки
            grouped_cons: Dict[Tuple[Optional[str], Optional[str]], Dict[str, Any]] = {}
            cols_cons = Form0503317Constants.CONSOLIDATED_COLUMNS

            for row in rows:
                classification_code, indicator_name, line_code, budget_type, data_type, level, source_row, *values = row
                # Уникальный ключ: только наименование и код строки (как для доходов/расходов/источников)
                key = (indicator_name, line_code)
                if key not in grouped_cons:
                    grouped_cons[key] = {
                        'код_классификации': classification_code or '',
                        'наименование_показателя': indicator_name or '',
                        'код_строки': line_code or '',
                        'раздел': 'консолидируемые_расчеты',
                        'уровень': level,  # Используем кэшированный уровень из БД
                        'исходная_строка': source_row,  # Используем сохраненную исходную строку
                    }

                target = grouped_cons[key]

                if data_type == 'оригинальные':
                    bucket = target.setdefault('поступления', {})
                    for idx, col_name in enumerate(cols_cons):
                        bucket[col_name] = values[idx]
                elif data_type == 'вычисленные':
                    for idx, col_name in enumerate(cols_cons):
                        target[f'расчетный_поступления_{col_name}'] = values[idx]

            data['консолидируемые_расчеты_data'] = list(grouped_cons.values())
        
        return data

    # ------------------------------------------------------------------
    # Методы работы со справочниками и новой архитектурой форм/ревизий
    # ------------------------------------------------------------------

    # ----- Справочник лет -----

    def get_or_create_year(self, year: int) -> YearRef:
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute('SELECT id, year, is_active FROM ref_years WHERE year=?', (year,))
            row = cursor.fetchone()
            if row:
                return YearRef.from_row({'id': row[0], 'year': row[1], 'is_active': row[2]})

            cursor.execute(
                'INSERT INTO ref_years (year, is_active) VALUES (?, 1)',
                (year,)
            )
            year_id = cursor.lastrowid
            conn.commit()
            return YearRef.from_row({'id': year_id, 'year': year, 'is_active': 1})

    def load_years(self) -> List[YearRef]:
        years: List[YearRef] = []
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute('SELECT id, year, is_active FROM ref_years ORDER BY year DESC')
            for row in cursor.fetchall():
                years.append(YearRef.from_row({'id': row[0], 'year': row[1], 'is_active': row[2]}))
        return years

    def save_years_bulk(self, years: List[YearRef]) -> None:
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute('DELETE FROM ref_years')
            if years:
                cursor.executemany(
                    'INSERT INTO ref_years (year, is_active) VALUES (?, ?)',
                    [(y.year, 1 if y.is_active else 0) for y in years],
                )
            conn.commit()

    # ----- Справочник МО -----

    def get_or_create_municipality(self, name: str, code: Optional[str] = None) -> MunicipalityRef:
        name = (name or "").strip()
        code = (code or "").strip() or None
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            if code:
                cursor.execute(
                    'SELECT id, code, name, is_active FROM ref_municipalities WHERE code=?',
                    (code,)
                )
            else:
                cursor.execute(
                    'SELECT id, code, name, is_active FROM ref_municipalities WHERE name=?',
                    (name,)
                )
            row = cursor.fetchone()
            if row:
                return MunicipalityRef.from_row(
                    {'id': row[0], 'code': row[1], 'name': row[2], 'is_active': row[3]}
                )

            cursor.execute(
                'INSERT INTO ref_municipalities (code, name, is_active) VALUES (?, ?, 1)',
                (code, name)
            )
            m_id = cursor.lastrowid
            conn.commit()
            return MunicipalityRef.from_row({'id': m_id, 'code': code, 'name': name, 'is_active': 1})

    def load_municipalities(self) -> List[MunicipalityRef]:
        result: List[MunicipalityRef] = []
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute('SELECT id, code, name, is_active FROM ref_municipalities ORDER BY name')
            for row in cursor.fetchall():
                result.append(
                    MunicipalityRef.from_row(
                        {'id': row[0], 'code': row[1], 'name': row[2], 'is_active': row[3]}
                    )
                )
        return result
    
    def get_municipality_by_id(self, municipality_id: int):
        """Получение расширенной информации о МО по ID"""
        from models.base_models import ExtendedMunicipalityRef
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute('''
                SELECT id, code, name, код_вида_МО, код_МО, родительный_падеж,
                       адрес_совет, адрес_администрация, совет_почта, администрация_почта,
                       должность_совет, фамилия_совет, имя_совет, отчество_совет,
                       должность_администрация, фамилия_администрация, имя_администрация, отчество_администрация,
                       дата_соглашения, дата_решения, номер_решения,
                       начальная_доходы, начальная_расходы, начальная_дефицит, is_active
                FROM ref_municipalities WHERE id=?
            ''', (municipality_id,))
            row = cursor.fetchone()
            if row:
                return ExtendedMunicipalityRef.from_row({
                    'id': row[0],
                    'code': row[1],
                    'name': row[2],
                    'код_вида_МО': row[3],
                    'код_МО': row[4],
                    'родительный_падеж': row[5],
                    'адрес_совет': row[6],
                    'адрес_администрация': row[7],
                    'совет_почта': row[8],
                    'администрация_почта': row[9],
                    'должность_совет': row[10],
                    'фамилия_совет': row[11],
                    'имя_совет': row[12],
                    'отчество_совет': row[13],
                    'должность_администрация': row[14],
                    'фамилия_администрация': row[15],
                    'имя_администрация': row[16],
                    'отчество_администрация': row[17],
                    'дата_соглашения': row[18],
                    'дата_решения': row[19],
                    'номер_решения': row[20],
                    'начальная_доходы': row[21],
                    'начальная_расходы': row[22],
                    'начальная_дефицит': row[23],
                    'is_active': row[24]
                })
        return None

    def save_municipalities_bulk(self, municip_list: List[MunicipalityRef]) -> None:
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute('DELETE FROM ref_municipalities')
            if municip_list:
                cursor.executemany(
                    'INSERT INTO ref_municipalities (code, name, is_active) VALUES (?, ?, ?)',
                    [
                        (m.code or None, m.name, 1 if m.is_active else 0)
                        for m in municip_list
                    ],
                )
            conn.commit()

    # ----- Справочник типов форм -----

    def get_form_type_meta_by_code(self, code: str) -> Optional[FormTypeMeta]:
        """Получение мета‑информации о типе формы по коду (без автосоздания)."""
        code = (code or "").strip()
        if not code:
            return None
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute(
                'SELECT id, code, name, periodicity, column_mapping, is_active FROM ref_form_types WHERE code=?',
                (code,)
            )
            row = cursor.fetchone()
            if not row:
                return None
            return FormTypeMeta.from_row(
                {'id': row[0], 'code': row[1], 'name': row[2],
                 'periodicity': row[3], 'column_mapping': row[4], 'is_active': row[5]}
            )

    def load_form_types_meta(self) -> List[FormTypeMeta]:
        result: List[FormTypeMeta] = []
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute('SELECT id, code, name, periodicity, column_mapping, is_active FROM ref_form_types ORDER BY code')
            for row in cursor.fetchall():
                result.append(
                    FormTypeMeta.from_row(
                        {'id': row[0], 'code': row[1], 'name': row[2],
                         'periodicity': row[3], 'column_mapping': row[4], 'is_active': row[5]}
                    )
                )
        return result

    def save_form_types_bulk(self, forms_list: List[FormTypeMeta]) -> None:
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute('DELETE FROM ref_form_types')
            if forms_list:
                # ID типов форм задаём вручную (стабильные идентификаторы),
                # а не используем AUTOINCREMENT SQLite.
                for f in forms_list:
                    # Если ID явно задан в модели – используем его,
                    # иначе пытаемся вывести ID из кода формы (например, '0503317' → 503317).
                    form_id = getattr(f, "id", None)
                    if not form_id:
                        try:
                            form_id = int(str(f.code).lstrip("0") or "0")
                        except ValueError:
                            # На крайний случай – не сохраняем такую строку, чтобы не ломать связи
                            logger.warning(f"Невозможно определить ID для типа формы с кодом '{f.code}', запись пропущена")
                            continue

                    column_mapping_json = json.dumps(f.column_mapping, ensure_ascii=False) if f.column_mapping else None
                    cursor.execute(
                        'INSERT INTO ref_form_types (id, code, name, periodicity, column_mapping, is_active) '
                        'VALUES (?, ?, ?, ?, ?, ?)',
                        (
                            form_id,
                            f.code,
                            f.name,
                            f.periodicity or None,
                            column_mapping_json,
                            1 if f.is_active else 0,
                        ),
                    )
            conn.commit()

    # ----- Справочник периодов -----

    def load_periods(self, form_type_code: Optional[str] = None) -> List[PeriodRef]:
        result: List[PeriodRef] = []
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            if form_type_code:
                cursor.execute(
                    'SELECT id, code, наименование, sort_order, form_type_code, is_active '
                    'FROM ref_periods WHERE form_type_code=? ORDER BY sort_order, code',
                    (form_type_code,)
                )
            else:
                cursor.execute(
                    'SELECT id, code, наименование, sort_order, form_type_code, is_active '
                    'FROM ref_periods ORDER BY sort_order, code'
                )
            for row in cursor.fetchall():
                result.append(
                    PeriodRef.from_row(
                        {'id': row[0], 'code': row[1], 'name': row[2],  # маппинг: наименование -> name
                         'sort_order': row[3], 'form_type_code': row[4], 'is_active': row[5]}
                    )
                )
        return result

    def get_period_by_code(self, code: str, form_type_code: Optional[str] = None) -> Optional[PeriodRef]:
        """Получение периода по коду (и, опционально, коду формы), без автосоздания."""
        code = (code or "").strip()
        if not code:
            return None
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            form_type_code = (form_type_code or "").strip() or None

            # 1) Пытаемся найти период, привязанный к конкретному типу формы
            if form_type_code:
                cursor.execute(
                    'SELECT id, code, наименование, sort_order, form_type_code, is_active '
                    'FROM ref_periods WHERE code=? AND form_type_code=?',
                    (code, form_type_code)
                )
                row = cursor.fetchone()
                if row:
                    return PeriodRef.from_row(
                        {'id': row[0], 'code': row[1], 'name': row[2],  # маппинг: наименование -> name
                         'sort_order': row[3], 'form_type_code': row[4], 'is_active': row[5]}
                    )

            # 2) Если не нашли — пробуем общий период (form_type_code IS NULL)
            cursor.execute(
                'SELECT id, code, наименование, sort_order, form_type_code, is_active '
                'FROM ref_periods WHERE code=? AND form_type_code IS NULL',
                (code,)
            )
            row = cursor.fetchone()
            if not row:
                return None
            return PeriodRef.from_row(
                {'id': row[0], 'code': row[1], 'name': row[2],  # маппинг: наименование -> name
                 'sort_order': row[3], 'form_type_code': row[4], 'is_active': row[5]}
            )

    def save_periods_bulk(self, periods_list: List[PeriodRef]) -> None:
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute('DELETE FROM ref_periods')
            if periods_list:
                # ID периодов также задаются/сохраняются вручную через интерфейс.
                # Если ID есть — вставляем его явно; если нет — даём SQLite сгенерировать.
                for p in periods_list:
                    period_id = getattr(p, "id", None)
                    if period_id:
                        cursor.execute(
                            'INSERT INTO ref_periods (id, code, наименование, sort_order, form_type_code, is_active) '
                            'VALUES (?, ?, ?, ?, ?, ?)',
                            (
                                period_id,
                                p.code,
                                p.name,  # маппинг: name -> наименование
                                p.sort_order,
                                p.form_type_code or None,
                                1 if p.is_active else 0,
                            ),
                        )
                    else:
                        cursor.execute(
                            'INSERT INTO ref_periods (code, наименование, sort_order, form_type_code, is_active) '
                            'VALUES (?, ?, ?, ?, ?)',
                            (
                                p.code,
                                p.name,  # маппинг: name -> наименование
                                p.sort_order,
                                p.form_type_code or None,
                                1 if p.is_active else 0,
                            ),
                        )
            conn.commit()

    # ----- ProjectForm и FormRevisionRecord -----

    def get_or_create_project_form(self, project_id: int, form_type_id: int,
                                   period_id: Optional[int]) -> ProjectForm:
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute(
                'SELECT id, project_id, form_type_id, period_id '
                'FROM project_forms WHERE project_id=? AND form_type_id=? AND '
                '(period_id IS ? OR period_id = ?)',
                (project_id, form_type_id, period_id, period_id)
            )
            row = cursor.fetchone()
            if row:
                return ProjectForm.from_row(
                    {'id': row[0], 'project_id': row[1],
                     'form_type_id': row[2], 'period_id': row[3]}
                )

            cursor.execute(
                'INSERT INTO project_forms (project_id, form_type_id, period_id) VALUES (?, ?, ?)',
                (project_id, form_type_id, period_id)
            )
            pf_id = cursor.lastrowid
            conn.commit()
            return ProjectForm.from_row(
                {'id': pf_id, 'project_id': project_id,
                 'form_type_id': form_type_id, 'period_id': period_id}
            )

    def load_project_forms(self, project_id: int) -> List[ProjectForm]:
        result: List[ProjectForm] = []
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute(
                'SELECT id, project_id, form_type_id, period_id '
                'FROM project_forms WHERE project_id=? ORDER BY id',
                (project_id,)
            )
            for row in cursor.fetchall():
                result.append(
                    ProjectForm.from_row(
                        {'id': row[0], 'project_id': row[1],
                         'form_type_id': row[2], 'period_id': row[3]}
                    )
                )
        return result

    def get_project_form_by_id(self, project_form_id: int) -> Optional[ProjectForm]:
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute(
                'SELECT id, project_id, form_type_id, period_id FROM project_forms WHERE id=?',
                (project_form_id,),
            )
            row = cursor.fetchone()
            if row:
                return ProjectForm.from_row(
                    {'id': row[0], 'project_id': row[1], 'form_type_id': row[2], 'period_id': row[3]}
                )
        return None

    def get_form_type_meta_by_id(self, form_type_id: int) -> Optional[FormTypeMeta]:
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute(
                'SELECT id, code, name, periodicity, column_mapping, is_active FROM ref_form_types WHERE id=?',
                (form_type_id,),
            )
            row = cursor.fetchone()
            if row:
                return FormTypeMeta.from_row(
                    {
                        'id': row[0],
                        'code': row[1],
                        'name': row[2],
                        'periodicity': row[3],
                        'column_mapping': row[4],
                        'is_active': row[5],
                    }
                )
        return None

    def get_period_by_id(self, period_id: int) -> Optional[PeriodRef]:
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute(
                'SELECT id, code, наименование, sort_order, form_type_code, is_active FROM ref_periods WHERE id=?',
                (period_id,),
            )
            row = cursor.fetchone()
            if row:
                return PeriodRef.from_row(
                    {
                        'id': row[0],
                        'code': row[1],
                        'name': row[2],  # маппинг: наименование -> name
                        'sort_order': row[3],
                        'form_type_code': row[4],
                        'is_active': row[5],
                    }
                )
        return None

    def create_or_update_form_revision(
        self,
        project_form_id: int,
        revision: str,
        status: ProjectStatus,
        file_path: str,
    ) -> FormRevisionRecord:
        """Создать или обновить ревизию формы по ключу (project_form_id, revision)."""
        revision = (revision or "").strip()
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute(
                'SELECT id, project_form_id, revision, status, file_path, created_at '
                'FROM form_revisions WHERE project_form_id=? AND revision=?',
                (project_form_id, revision)
            )
            row = cursor.fetchone()
            now_iso = datetime.now().isoformat()
            if row:
                # Обновляем существующую ревизию
                cursor.execute(
                    'UPDATE form_revisions SET status=?, file_path=? WHERE id=?',
                    (status.value, file_path, row[0])
                )
                conn.commit()
                return FormRevisionRecord.from_row(
                    {
                        'id': row[0],
                        'project_form_id': row[1],
                        'revision': row[2],
                        'status': status.value,
                        'file_path': file_path,
                        'created_at': row[5] or now_iso,
                    }
                )

            # Создаём новую ревизию
            cursor.execute(
                'INSERT INTO form_revisions (project_form_id, revision, status, file_path, created_at) '
                'VALUES (?, ?, ?, ?, ?)',
                (project_form_id, revision, status.value, file_path, now_iso)
            )
            fr_id = cursor.lastrowid
            conn.commit()
            return FormRevisionRecord.from_row(
                {
                    'id': fr_id,
                    'project_form_id': project_form_id,
                    'revision': revision,
                    'status': status.value,
                    'file_path': file_path,
                    'created_at': now_iso,
                }
            )

    def load_form_revisions(self, project_form_id: int) -> List[FormRevisionRecord]:
        result: List[FormRevisionRecord] = []
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute(
                'SELECT id, project_form_id, revision, status, file_path, created_at '
                'FROM form_revisions WHERE project_form_id=? ORDER BY id',
                (project_form_id,)
            )
            for row in cursor.fetchall():
                result.append(
                    FormRevisionRecord.from_row(
                        {
                            'id': row[0],
                            'project_form_id': row[1],
                            'revision': row[2],
                            'status': row[3],
                            'file_path': row[4],
                            'created_at': row[5],
                        }
                    )
                )
        return result

    def get_form_revision_by_id(self, revision_id: int) -> Optional[FormRevisionRecord]:
        """Получение ревизии формы по ID."""
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute(
                'SELECT id, project_form_id, revision, status, file_path, created_at '
                'FROM form_revisions WHERE id=?',
                (revision_id,)
            )
            row = cursor.fetchone()
            if row:
                return FormRevisionRecord.from_row({
                    'id': row[0],
                    'project_form_id': row[1],
                    'revision': row[2],
                    'status': row[3],
                    'file_path': row[4],
                    'created_at': row[5],
                })
            return None
    
    def update_form_revision(
        self,
        revision_id: int,
        revision: str,
        status: ProjectStatus,
        file_path: str,
    ) -> bool:
        """Обновление ревизии формы по ID."""
        revision = (revision or "").strip()
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute(
                'UPDATE form_revisions SET revision=?, status=?, file_path=? WHERE id=?',
                (revision, status.value, file_path, revision_id)
            )
            conn.commit()
            return cursor.rowcount > 0
    
    def delete_form_revision(self, revision_id: int) -> None:
        """
        Удаление одной ревизии формы и всех связанных нормализованных данных.
        Удаляются:
        - записи в *_values (income/expense/source/consolidated)
        - revision_metadata
        - сама запись form_revisions
        - исходный файл ревизии (если существует)
        """
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            # Сначала пытаемся прочитать путь к файлу ревизии, чтобы удалить файл после транзакции
            cursor.execute('SELECT file_path FROM form_revisions WHERE id=?', (revision_id,))
            row = cursor.fetchone()
            file_path = row[0] if row else None

            tables_with_revision = [
                'income_values',
                'expense_values',
                'source_values',
                'consolidated_values',
                'revision_metadata',
            ]
            for table in tables_with_revision:
                cursor.execute(f'DELETE FROM {table} WHERE revision_id=?', (revision_id,))
            cursor.execute('DELETE FROM form_revisions WHERE id=?', (revision_id,))
            conn.commit()

        # Удаляем файл ревизии вне транзакции БД
        try:
            if file_path and os.path.exists(file_path):
                os.remove(file_path)
        except Exception as e:
            # Не блокируем удаление ревизии из-за ошибки удаления файла
            logger.warning(f"Не удалось удалить файл ревизии {file_path}: {e}", exc_info=True)
    
    def delete_project(self, project_id: int):
        """Удаление проекта"""
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute('DELETE FROM projects WHERE id=?', (project_id,))
            conn.commit()
    
    def save_reference(self, reference: Reference) -> int:
        """Сохранение справочника"""
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            
            if reference.id is None:
                cursor.execute('''
                    INSERT INTO reference_data 
                    (name, reference_type, file_path, loaded_at, data)
                    VALUES (?, ?, ?, ?, NULL)
                ''', (
                    reference.name,
                    reference.reference_type,
                    reference.file_path,
                    reference.loaded_at.isoformat()
                ))
                reference.id = cursor.lastrowid
            else:
                cursor.execute('''
                    UPDATE reference_data SET
                    name=?, reference_type=?, file_path=?, loaded_at=?, data=NULL
                    WHERE id=?
                ''', (
                    reference.name,
                    reference.reference_type,
                    reference.file_path,
                    reference.loaded_at.isoformat(),
                    reference.id
                ))
            
            conn.commit()
            return reference.id

    def save_reference_records(self, reference_type: str, records: list):
        """
        Сохранение строк справочника в отдельные SQL-таблицы.
        Ожидается список словарей с ключами:
        - для доходов: 'код_классификации_ДБ', 'наименование', 'уровень_кода', 'Утверждающий документ'
        - для источников: 'код_классификации_ИФДБ', 'наименование', 'уровень_кода', 'Утверждающий документ'
        """
        if not records:
            return

        table_name = None
        code_field = None

        if reference_type == 'доходы':
            table_name = 'income_reference_records'
            code_field = 'код_классификации_ДБ'
        elif reference_type == 'источники':
            table_name = 'source_reference_records'
            code_field = 'код_классификации_ИФДБ'

        if not table_name or not code_field:
            return

        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()

            # Полностью очищаем таблицу перед загрузкой нового справочника
            cursor.execute(f'DELETE FROM {table_name}')

            # Готовим вставку новых строк, контролируя уникальность кода
            rows_to_insert = []
            seen_codes = set()
            for rec in records:
                code = str(rec.get(code_field, '')).replace(' ', '')
                if not code:
                    continue
                if code in seen_codes:
                    # Пропускаем дубликаты после нормализации кода,
                    # чтобы не нарушать UNIQUE-ограничение
                    continue
                seen_codes.add(code)

                name = rec.get('наименование')
                level = rec.get('уровень_кода')
                doc = rec.get('Утверждающий документ')
                rows_to_insert.append((code, name, level, doc))

            if rows_to_insert:
                cursor.executemany(
                    f'''
                    INSERT INTO {table_name} (code, name, level, doc)
                    VALUES (?, ?, ?, ?)
                    ''',
                    rows_to_insert
                )

            conn.commit()
    
    def load_references(self) -> List[Reference]:
        """Загрузка всех справочников"""
        references = []
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute('SELECT * FROM reference_data ORDER BY loaded_at DESC')
            
            for row in cursor.fetchall():
                ref_data = {
                    'id': row[0],
                    'name': row[1],
                    'reference_type': row[2],
                    'file_path': row[3],
                    'loaded_at': row[4],
                    'data': None  # данные строк теперь берём из отдельных таблиц
                }
                references.append(Reference.from_dict(ref_data))
        
        return references
    
    def load_income_reference_df(self) -> pd.DataFrame:
        """Загрузка справочника доходов как DataFrame из SQL-таблицы income_reference_records"""
        with sqlite3.connect(self.db_path) as conn:
            query = '''
                SELECT code AS код_классификации_ДБ,
                       name AS наименование,
                       level AS уровень_кода,
                       doc AS Утверждающий_документ
                FROM income_reference_records
            '''
            df = pd.read_sql_query(query, conn)
        return df

    def load_sources_reference_df(self) -> pd.DataFrame:
        """Загрузка справочника источников финансирования как DataFrame из SQL-таблицы source_reference_records"""
        with sqlite3.connect(self.db_path) as conn:
            query = '''
                SELECT code AS код_классификации_ИФДБ,
                       name AS наименование,
                       level AS уровень_кода,
                       doc AS Утверждающий_документ
                FROM source_reference_records
            '''
            df = pd.read_sql_query(query, conn)
        return df

    # ----- Нормализованные данные форм (values) -----

    def load_income_values_df(self, project_id: Optional[int] = None, revision_id: Optional[int] = None) -> pd.DataFrame:
        """
        Загрузка данных раздела «Доходы» из таблицы income_values в виде DataFrame.
        Удобно для аналитики и внешних расчётов.
        """
        with sqlite3.connect(self.db_path) as conn:
            base_query = '''
                SELECT
                    project_id,
                    revision_id,
                    classification_code AS код_классификации,
                    indicator_name     AS наименование_показателя,
                    line_code          AS код_строки,
                    budget_type        AS тип_бюджета,
                    data_type          AS тип_данных,
                    level              AS level,
                    source_row         AS source_row,
                    {value_cols}
                FROM income_values
            '''
            value_cols = ", ".join(f"v{i+1} AS v{i+1}" for i in range(len(Form0503317Constants.BUDGET_COLUMNS)))
            query = base_query.format(value_cols=value_cols)

            params: list = []
            where_clauses: list = []
            if project_id is not None:
                where_clauses.append("project_id = ?")
                params.append(project_id)
            if revision_id is not None:
                where_clauses.append("revision_id = ?")
                params.append(revision_id)

            if where_clauses:
                query += " WHERE " + " AND ".join(where_clauses)

            df = pd.read_sql_query(query, conn, params=params or None)
        return df

    def load_expense_values_df(self, project_id: Optional[int] = None, revision_id: Optional[int] = None) -> pd.DataFrame:
        """Загрузка данных раздела «Расходы» из expense_values как DataFrame."""
        with sqlite3.connect(self.db_path) as conn:
            base_query = '''
                SELECT
                    project_id,
                    revision_id,
                    classification_code AS код_классификации,
                    indicator_name     AS наименование_показателя,
                    line_code          AS код_строки,
                    budget_type        AS тип_бюджета,
                    data_type          AS тип_данных,
                    level              AS level,
                    source_row         AS source_row,
                    {value_cols}
                FROM expense_values
            '''
            value_cols = ", ".join(f"v{i+1} AS v{i+1}" for i in range(len(Form0503317Constants.BUDGET_COLUMNS)))
            query = base_query.format(value_cols=value_cols)

            params: list = []
            where_clauses: list = []
            if project_id is not None:
                where_clauses.append("project_id = ?")
                params.append(project_id)
            if revision_id is not None:
                where_clauses.append("revision_id = ?")
                params.append(revision_id)

            if where_clauses:
                query += " WHERE " + " AND ".join(where_clauses)

            df = pd.read_sql_query(query, conn, params=params or None)
        return df

    def load_source_values_df(self, project_id: Optional[int] = None, revision_id: Optional[int] = None) -> pd.DataFrame:
        """Загрузка данных раздела «Источники финансирования» из source_values как DataFrame."""
        with sqlite3.connect(self.db_path) as conn:
            base_query = '''
                SELECT
                    project_id,
                    revision_id,
                    classification_code AS код_классификации,
                    indicator_name     AS наименование_показателя,
                    line_code          AS код_строки,
                    budget_type        AS тип_бюджета,
                    data_type          AS тип_данных,
                    level              AS level,
                    source_row         AS source_row,
                    {value_cols}
                FROM source_values
            '''
            value_cols = ", ".join(f"v{i+1} AS v{i+1}" for i in range(len(Form0503317Constants.BUDGET_COLUMNS)))
            query = base_query.format(value_cols=value_cols)

            params: list = []
            where_clauses: list = []
            if project_id is not None:
                where_clauses.append("project_id = ?")
                params.append(project_id)
            if revision_id is not None:
                where_clauses.append("revision_id = ?")
                params.append(revision_id)

            if where_clauses:
                query += " WHERE " + " AND ".join(where_clauses)

            df = pd.read_sql_query(query, conn, params=params or None)
        return df

    def load_consolidated_values_df(self, project_id: Optional[int] = None, revision_id: Optional[int] = None) -> pd.DataFrame:
        """Загрузка данных раздела «Консолидируемые расчёты» из consolidated_values как DataFrame."""
        with sqlite3.connect(self.db_path) as conn:
            base_query = '''
                SELECT
                    project_id,
                    revision_id,
                    classification_code AS код_классификации,
                    indicator_name     AS наименование_показателя,
                    line_code          AS код_строки,
                    budget_type        AS тип_бюджета,
                    data_type          AS тип_данных,
                    level              AS level,
                    source_row         AS source_row,
                    {value_cols}
                FROM consolidated_values
            '''
            value_cols = ", ".join(f"v{i+1} AS v{i+1}" for i in range(len(Form0503317Constants.CONSOLIDATED_COLUMNS)))
            query = base_query.format(value_cols=value_cols)

            params: list = []
            where_clauses: list = []
            if project_id is not None:
                where_clauses.append("project_id = ?")
                params.append(project_id)
            if revision_id is not None:
                where_clauses.append("revision_id = ?")
                params.append(revision_id)

            if where_clauses:
                query += " WHERE " + " AND ".join(where_clauses)

            df = pd.read_sql_query(query, conn, params=params or None)
        return df

    # ----- Аналитика по нормализованным данным -----

    def summarize_budget_by_level(
        self,
        section: str,
        project_id: int,
        revision_id: Optional[int] = None,
        budget_type: Optional[str] = None,
        data_type: str = 'вычисленные',
    ) -> pd.DataFrame:
        """
        Агрегация бюджетных разделов (доходы/расходы/источники) по уровню.

        Возвращает DataFrame с колонками:
        уровень, v1..vN (сумма по выбранным фильтрам).
        """
        table_map = {
            'доходы': 'income_values',
            'расходы': 'expense_values',
            'источники_финансирования': 'source_values',
        }
        if section not in table_map:
            raise ValueError(f'Неизвестный раздел: {section}')

        table = table_map[section]
        value_cols = Form0503317Constants.BUDGET_COLUMNS

        with sqlite3.connect(self.db_path) as conn:
            where_clauses = ["project_id = ?"]
            params: list = [project_id]
            if revision_id is None:
                where_clauses.append("revision_id IS NULL")
            else:
                where_clauses.append("revision_id = ?")
                params.append(revision_id)

            if budget_type:
                where_clauses.append("budget_type = ?")
                params.append(budget_type)
            if data_type:
                where_clauses.append("data_type = ?")
                params.append(data_type)

            sums = ", ".join(f"SUM(v{i+1}) AS {col}" for i, col in enumerate(value_cols))
            query = f'''
                SELECT level AS уровень, {sums}
                FROM {table}
                WHERE {' AND '.join(where_clauses)}
                GROUP BY level
                ORDER BY level
            '''
            return pd.read_sql_query(query, conn, params=params)

    def summarize_consolidated_by_level(
        self,
        project_id: int,
        revision_id: Optional[int] = None,
        data_type: str = 'вычисленные',
    ) -> pd.DataFrame:
        """
        Агрегация консолидируемых расчётов по уровню.

        Возвращает DataFrame с колонками:
        уровень, v1..vN (сумма по выбранным фильтрам).
        """
        value_cols = Form0503317Constants.CONSOLIDATED_COLUMNS
        with sqlite3.connect(self.db_path) as conn:
            where_clauses = ["project_id = ?"]
            params: list = [project_id]
            if revision_id is None:
                where_clauses.append("revision_id IS NULL")
            else:
                where_clauses.append("revision_id = ?")
                params.append(revision_id)

            if data_type:
                where_clauses.append("data_type = ?")
                params.append(data_type)

            sums = ", ".join(f"SUM(v{i+1}) AS {col}" for i, col in enumerate(value_cols))
            query = f'''
                SELECT level AS уровень, {sums}
                FROM consolidated_values
                WHERE {' AND '.join(where_clauses)}
                GROUP BY level
                ORDER BY level
            '''
            return pd.read_sql_query(query, conn, params=params)
    
    def calculate_sums_from_values(
        self,
        project_id: int,
        revision_id: int,
        reference_data_доходы: Optional[pd.DataFrame] = None,
        reference_data_источники: Optional[pd.DataFrame] = None,
    ) -> Dict[str, List[Dict[str, Any]]]:
        """
        Расчет сумм напрямую из нормализованных данных *_values без преобразования в старый формат.
        
        Возвращает словарь с ключами:
        - 'доходы_data', 'расходы_data', 'источники_финансирования_data', 'консолидируемые_расчеты_data'
        
        Каждый раздел содержит список словарей с полями:
        - 'код_классификации', 'наименование_показателя', 'код_строки', 'раздел'
        - 'утвержденный', 'исполненный' (словари со значениями по бюджетным колонкам)
        - 'расчетный_утвержденный_...', 'расчетный_исполненный_...' (вычисленные значения)
        - 'уровень' (пересчитанный на основе справочников)
        """
        from models.form_0503317 import Form0503317
        
        # Загружаем данные из нормализованных таблиц как DataFrame
        income_df = self.load_income_values_df(project_id, revision_id)
        expense_df = self.load_expense_values_df(project_id, revision_id)
        source_df = self.load_source_values_df(project_id, revision_id)
        consolidated_df = self.load_consolidated_values_df(project_id, revision_id)
        
        # Создаем временную форму для расчетов
        form = Form0503317()
        form.reference_data_доходы = reference_data_доходы
        form.reference_data_источники = reference_data_источники
        
        result = {}
        
        # Обрабатываем каждый раздел
        for section_name, df, table_type in [
            ('доходы_data', income_df, 'budget'),
            ('расходы_data', expense_df, 'budget'),
            ('источники_финансирования_data', source_df, 'budget'),
            ('консолидируемые_расчеты_data', consolidated_df, 'consolidated'),
        ]:
            if df.empty:
                result[section_name] = []
                continue
            
            # Преобразуем DataFrame в формат для расчетов
            if table_type == 'budget':
                section_data = self._convert_budget_df_to_calculation_format(df, section_name, form)
            else:
                section_data = self._convert_consolidated_df_to_calculation_format(df, form)
            
            if not section_data:
                result[section_name] = []
                continue
            
            # Выполняем расчеты
            if table_type == 'budget':
                df_calc = form._prepare_dataframe_for_calculation(section_data, form.constants.BUDGET_COLUMNS)
                if section_name == 'источники_финансирования_data':
                    df_with_sums = form._calculate_sources_sums(df_calc, form.constants.BUDGET_COLUMNS)
                else:
                    df_with_sums = form._calculate_standard_sums(df_calc, form.constants.BUDGET_COLUMNS)
                result[section_name] = df_with_sums.to_dict('records')
            else:
                df_calc = form._prepare_consolidated_dataframe_for_calculation(
                    section_data, form.constants.CONSOLIDATED_COLUMNS)
                
                # Отладочный вывод: проверяем колонки DataFrame до расчета
                calc_cols_before = [c for c in df_calc.columns if c.startswith('расчетный_поступления_')]
                logger.debug(f"Консолидированные расчеты: DataFrame имеет {len(calc_cols_before)} расчетных колонок до расчета")
                
                df_with_sums = form._calculate_consolidated_sums(df_calc)
                
                # Отладочный вывод: проверяем колонки DataFrame после расчета
                calc_cols_after = [c for c in df_with_sums.columns if c.startswith('расчетный_поступления_')]
                logger.debug(f"Консолидированные расчеты: DataFrame имеет {len(calc_cols_after)} расчетных колонок после расчета")
                
                result_rows = df_with_sums.to_dict('records')
                
                # Отладочный вывод: проверяем наличие расчетных полей в словаре
                if result_rows:
                    sample_row = result_rows[0]
                    calc_keys = [k for k in sample_row.keys() if k.startswith('расчетный_поступления_')]
                    logger.debug(f"Консолидированные расчеты: {len(result_rows)} строк в словаре, "
                          f"найдено расчетных полей в первой строке: {len(calc_keys)}")
                    if calc_keys:
                        logger.debug(f"  Примеры ключей: {calc_keys[:3]}")
                        # Проверяем, есть ли не-None значения
                        non_none_count = sum(1 for k in calc_keys if sample_row.get(k) is not None)
                        logger.debug(f"  Не-None значений в первой строке: {non_none_count}")
                        # Проверяем несколько строк
                        rows_with_calc = sum(1 for r in result_rows[:10] if any(k.startswith('расчетный_поступления_') for k in r.keys()))
                        logger.debug(f"  Строк с расчетными полями (первые 10): {rows_with_calc}")
                
                result[section_name] = result_rows
        
        return result
    
    def _convert_budget_df_to_calculation_format(
        self,
        df: pd.DataFrame,
        section_name: str,
        form: Any,
    ) -> List[Dict[str, Any]]:
        """Преобразует DataFrame из *_values в формат для расчетов."""
        if df.empty:
            return []
        
        # Группируем по уникальным строкам (classification_code, indicator_name, line_code)
        grouped = df.groupby(['код_классификации', 'наименование_показателя', 'код_строки'])
        
        result = []
        budget_cols = form.constants.BUDGET_COLUMNS
        
        for (code, name, line_code), group in grouped:
            row_data = {
                'код_классификации': code or '',
                'наименование_показателя': name or '',
                'код_строки': line_code or '',
                'раздел': section_name.replace('_data', ''),
                'утвержденный': {},
                'исполненный': {},
            }
            
            # Используем кэшированный уровень из БД, если есть
            # Иначе определяем заново
            level = None
            source_row = None
            for _, r in group.iterrows():
                if pd.notna(r.get('level')):
                    level = int(r['level'])
                if pd.notna(r.get('source_row')):
                    source_row = int(r['source_row'])
                break  # Берем из первой строки группы
            
            if level is None:
                level = form._determine_level(
                    row_data['код_классификации'],
                    row_data['раздел'],
                    row_data['наименование_показателя']
                )
            row_data['уровень'] = level
            if source_row is not None:
                row_data['исходная_строка'] = source_row
            
            # Собираем оригинальные и вычисленные значения
            for _, r in group.iterrows():
                budget_type = r['тип_бюджета']
                data_type = r['тип_данных']
                
                values_dict = {}
                for i, col in enumerate(budget_cols):
                    val = r.get(f'v{i+1}')
                    if val is not None:
                        values_dict[col] = val
                
                if data_type == 'оригинальные':
                    row_data[budget_type] = values_dict
                elif data_type == 'вычисленные':
                    prefix = 'расчетный_утвержденный_' if budget_type == 'утвержденный' else 'расчетный_исполненный_'
                    for col, val in values_dict.items():
                        row_data[f'{prefix}{col}'] = val
            
            result.append(row_data)
        
        # Сортируем: сначала строки "всего", затем по исходной строке, потом по коду строки
        def sort_key(item: Dict[str, Any]):
            name_lower = str(item.get('наименование_показателя', '')).lower()
            is_total = 'всего' in name_lower
            source_row_val = item.get('исходная_строка')
            line_code_val = item.get('код_строки') or ''
            return (0 if is_total else 1, source_row_val if source_row_val is not None else 10**9, line_code_val)

        result.sort(key=sort_key)
        return result
    
    def _convert_consolidated_df_to_calculation_format(
        self,
        df: pd.DataFrame,
        form: Any,
    ) -> List[Dict[str, Any]]:
        """
        Преобразует DataFrame из consolidated_values в формат для расчетов.
        
        Уникальность записей определяется только по (indicator_name, line_code),
        без учета classification_code (в отличие от доходов/расходов/источников).
        """
        if df.empty:
            return []
        
        # Группируем по уникальным строкам: только наименование и код строки
        # (как для доходов/расходов/источников, но без кода классификации)
        grouped = df.groupby(['наименование_показателя', 'код_строки'])
        
        result = []
        consolidated_cols = form.constants.CONSOLIDATED_COLUMNS
        
        for (name, line_code), group in grouped:
            # Берем код классификации из первой строки группы (если есть)
            classification_code = None
            for _, r in group.iterrows():
                code_val = r.get('код_классификации')
                if code_val:
                    classification_code = code_val
                    break
            
            row_data = {
                'код_классификации': classification_code or '',
                'наименование_показателя': name or '',
                'код_строки': line_code or '',
                'раздел': 'консолидируемые_расчеты',
                'поступления': {},
            }
            
            # Используем кэшированный уровень из БД, если есть
            # Иначе определяем заново
            level = None
            source_row = None
            for _, r in group.iterrows():
                if r.get('level') is not None:
                    level = int(r['level'])
                if r.get('source_row') is not None:
                    source_row = int(r['source_row'])
                break  # Берем из первой строки группы
            
            if level is None:
                level = form._determine_consolidated_level(row_data['код_строки'])
            row_data['уровень'] = level
            if source_row is not None:
                row_data['исходная_строка'] = source_row
            
            # Собираем оригинальные и вычисленные значения
            for _, r in group.iterrows():
                data_type = r['тип_данных']
                
                values_dict = {}
                for i, col in enumerate(consolidated_cols):
                    val = r.get(f'v{i+1}')
                    if val is not None:
                        values_dict[col] = val
                
                if data_type == 'оригинальные':
                    row_data['поступления'] = values_dict
                elif data_type == 'вычисленные':
                    for col, val in values_dict.items():
                        row_data[f'расчетный_поступления_{col}'] = val
            
            result.append(row_data)
        
        # Сортируем: строки "всего/итого" наверх, затем по исходной строке, потом по коду
        def sort_key(item: Dict[str, Any]):
            name_lower = str(item.get('наименование_показателя', '')).lower()
            is_total = ('всего' in name_lower) or ('итого' in name_lower)
            source_row_val = item.get('исходная_строка')
            line_code_val = item.get('код_строки') or ''
            return (0 if is_total else 1, source_row_val if source_row_val is not None else 10**9, line_code_val)

        result.sort(key=sort_key)
        return result
    
    def load_revision_metadata(self, revision_id: int) -> Dict[str, Any]:
        """Загрузка метаданных ревизии из отдельной таблицы"""
        result = {}
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute(
                'SELECT meta_info, результат_исполнения_data FROM revision_metadata WHERE revision_id=?',
                (revision_id,)
            )
            meta_row = cursor.fetchone()
            if meta_row:
                if meta_row[0]:  # meta_info
                    try:
                        result['meta_info'] = json.loads(meta_row[0])
                    except (json.JSONDecodeError, TypeError) as e:
                        logger.warning(f"Ошибка загрузки meta_info для ревизии {revision_id}: {e}", exc_info=True)
                
                if meta_row[1]:  # результат_исполнения_data
                    try:
                        result['результат_исполнения_data'] = json.loads(meta_row[1])
                    except (json.JSONDecodeError, TypeError) as e:
                        logger.warning(f"Ошибка загрузки результат_исполнения_data для ревизии {revision_id}: {e}", exc_info=True)
        return result
    
    def save_revision_data(self, project_id: int, revision_id: int, data: Dict[str, Any]):
        """Сохранение данных ревизии (разделы + метаданные)"""
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            # Сохраняем данные разделов
            self._save_project_data(cursor, project_id, data, revision_id)
            # Сохраняем метаданные отдельно
            meta_info = data.get('meta_info')
            результат_исполнения_data = data.get('результат_исполнения_data')
            if meta_info or результат_исполнения_data:
                # Удаляем старые метаданные
                cursor.execute('DELETE FROM revision_metadata WHERE revision_id=?', (revision_id,))
                # Сохраняем новые метаданные
                meta_info_json = json.dumps(meta_info, ensure_ascii=False, default=str) if meta_info else None
                результат_исполнения_json = json.dumps(результат_исполнения_data, ensure_ascii=False, default=str) if результат_исполнения_data else None
                if meta_info_json or результат_исполнения_json:
                    cursor.execute(
                        'INSERT INTO revision_metadata (revision_id, meta_info, результат_исполнения_data) VALUES (?, ?, ?)',
                        (revision_id, meta_info_json, результат_исполнения_json)
                    )
            conn.commit()
    
    def load_revision_data(self, project_id: int, revision_id: int) -> Dict[str, Any]:
        """Загрузка данных ревизии (разделы + метаданные)"""
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            # Загружаем данные разделов
            data = self._load_project_data(cursor, project_id, revision_id)
            # Загружаем метаданные отдельно
            metadata = self.load_revision_metadata(revision_id)
            data.update(metadata)
            return data
    
    def update_calculated_values(self, project_id: int, revision_id: int, calculated_data: Dict[str, List[Dict[str, Any]]]):
        """
        Обновление только вычисленных значений в таблицах *_values.
        calculated_data должен содержать ключи: 'доходы_data', 'расходы_data', 
        'источники_финансирования_data', 'консолидируемые_расчеты_data'
        с данными, содержащими поля вида 'расчетный_утвержденный_...', 'расчетный_исполненный_...'
        """
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            budget_cols = Form0503317Constants.BUDGET_COLUMNS
            consolidated_cols = Form0503317Constants.CONSOLIDATED_COLUMNS
            
            # Обновляем вычисленные значения для бюджетных разделов
            for section_key, table_name in [
                ('доходы_data', 'income_values'),
                ('расходы_data', 'expense_values'),
                ('источники_финансирования_data', 'source_values'),
            ]:
                if section_key not in calculated_data:
                    continue
                
                section_rows = calculated_data[section_key]
                if not section_rows:
                    continue
                
                # Удаляем старые вычисленные значения для этой ревизии
                cursor.execute(
                    f'DELETE FROM {table_name} WHERE project_id=? AND revision_id=? AND data_type=?',
                    (project_id, revision_id, 'вычисленные')
                )
                
                # Сохраняем новые вычисленные значения
                self._save_section_values(
                    cursor=cursor,
                    project_id=project_id,
                    revision_id=revision_id,
                    section_rows=section_rows,
                    table_name=table_name,
                    budget_columns=budget_cols,
                    only_calculated=True,  # Сохраняем только вычисленные значения
                )
            
            # Обновляем вычисленные значения для консолидируемых расчетов
            if 'консолидируемые_расчеты_data' in calculated_data:
                consolidated_rows = calculated_data['консолидируемые_расчеты_data']
                if consolidated_rows:
                    # Отладочный вывод: проверяем наличие расчетных полей
                    sample_row = consolidated_rows[0] if consolidated_rows else None
                    if sample_row:
                        calc_keys = [k for k in sample_row.keys() if k.startswith('расчетный_поступления_')]
                        logger.debug(f"Сохранение консолидированных расчетов: {len(consolidated_rows)} строк, "
                              f"найдено расчетных полей в первой строке: {len(calc_keys)}")
                        if calc_keys:
                            logger.debug(f"  Примеры ключей: {calc_keys[:3]}")
                    
                    # Удаляем старые вычисленные значения
                    cursor.execute(
                        'DELETE FROM consolidated_values WHERE project_id=? AND revision_id=? AND data_type=?',
                        (project_id, revision_id, 'вычисленные')
                    )
                    
                    # Сохраняем новые вычисленные значения
                    self._save_consolidated_values(
                        cursor=cursor,
                        project_id=project_id,
                        revision_id=revision_id,
                        section_rows=consolidated_rows,
                        table_name='consolidated_values',
                        consolidated_columns=consolidated_cols,
                        only_calculated=True,
                    )
            
            conn.commit()
    
    # ------------------------------------------------------------------
    # Методы загрузки справочников из Excel
    # ------------------------------------------------------------------
    
    def load_reference_from_excel(
        self,
        file_path: str,
        table_name: str,
        column_mapping: Dict[str, str],
        primary_key_column: str
    ) -> int:
        """
        Универсальный метод загрузки справочника из Excel файла.
        
        Args:
            file_path: Путь к Excel файлу
            table_name: Имя таблицы в БД для загрузки
            column_mapping: Словарь {имя_колонки_в_excel: имя_колонки_в_бд}
            primary_key_column: Имя колонки, которая является первичным ключом
        
        Returns:
            Количество загруженных записей
        """
        try:
            df = pd.read_excel(file_path)
            df.columns = [str(c).strip() for c in df.columns]
            
            # Нормализуем названия колонок
            df = df.rename(columns=column_mapping)
            
            # Удаляем пустые строки
            df = df.dropna(subset=[primary_key_column])
            
            with sqlite3.connect(self.db_path) as conn:
                cursor = conn.cursor()
                
                # Очищаем таблицу перед загрузкой
                cursor.execute(f'DELETE FROM {table_name}')
                
                # Подготавливаем данные для вставки
                columns = list(column_mapping.values())
                placeholders = ', '.join(['?'] * len(columns))
                
                count = 0
                for _, row in df.iterrows():
                    values = []
                    for col in columns:
                        val = row.get(col)
                        if pd.isna(val):
                            val = None
                        else:
                            val = str(val).strip()
                            if val == '':
                                val = None
                        values.append(val)
                    
                    # Пропускаем строки без первичного ключа
                    if values[columns.index(primary_key_column)] is None:
                        continue
                    
                    try:
                        cursor.execute(
                            f'INSERT INTO {table_name} ({", ".join(columns)}) VALUES ({placeholders})',
                            values
                        )
                        count += 1
                    except sqlite3.IntegrityError as e:
                        logger.warning(f"Пропущена дублирующая запись в {table_name}: {e}")
                        continue
                
                conn.commit()
                logger.info(f"Загружено {count} записей в {table_name} из {file_path}")
                return count
                
        except Exception as e:
            logger.error(f"Ошибка загрузки справочника из {file_path}: {e}", exc_info=True)
            raise
    
    def load_grbs_from_excel(self, file_path: str) -> int:
        """Загрузка справочника ГРБС из Excel"""
        return self.load_reference_from_excel(
            file_path=file_path,
            table_name='ref_grbs',
            column_mapping={'код_ГРБС': 'код_ГРБС', 'наименование': 'наименование'},
            primary_key_column='код_ГРБС'
        )
    
    def load_expense_sections_from_excel(self, file_path: str) -> int:
        """Загрузка справочника разделов/подразделов расходов из Excel"""
        return self.load_reference_from_excel(
            file_path=file_path,
            table_name='ref_expense_sections',
            column_mapping={
                'код_РП': 'код_РП',
                'наименование': 'наименование',
                'утверждающий_документ': 'утверждающий_документ'
            },
            primary_key_column='код_РП'
        )
    
    def load_target_articles_from_excel(self, file_path: str) -> int:
        """Загрузка справочника целевых статей расходов из Excel"""
        return self.load_reference_from_excel(
            file_path=file_path,
            table_name='ref_target_articles',
            column_mapping={'код_ЦСР': 'код_ЦСР', 'наименование': 'наименование'},
            primary_key_column='код_ЦСР'
        )
    
    def load_expense_types_from_excel(self, file_path: str) -> int:
        """Загрузка справочника видов статей расходов из Excel"""
        return self.load_reference_from_excel(
            file_path=file_path,
            table_name='ref_expense_types',
            column_mapping={'код_вида_СР': 'код_вида_СР', 'наименование': 'наименование'},
            primary_key_column='код_вида_СР'
        )
    
    def load_program_nonprogram_from_excel(self, file_path: str) -> int:
        """Загрузка справочника программных/непрограммных статей из Excel"""
        return self.load_reference_from_excel(
            file_path=file_path,
            table_name='ref_program_nonprogram',
            column_mapping={'код_ПНС': 'код_ПНС', 'наименование': 'наименование'},
            primary_key_column='код_ПНС'
        )
    
    def load_expense_kinds_from_excel(self, file_path: str) -> int:
        """Загрузка справочника видов расходов из Excel"""
        return self.load_reference_from_excel(
            file_path=file_path,
            table_name='ref_expense_kinds',
            column_mapping={
                'код_ВР': 'код_ВР',
                'наименование': 'наименование',
                'утверждающий_документ': 'утверждающий_документ'
            },
            primary_key_column='код_ВР'
        )
    
    def load_national_projects_from_excel(self, file_path: str) -> int:
        """Загрузка справочника национальных проектов из Excel"""
        return self.load_reference_from_excel(
            file_path=file_path,
            table_name='ref_national_projects',
            column_mapping={
                'код_НПЦСР': 'код_НПЦСР',
                'наименование': 'наименование',
                'утверждающий_документ': 'утверждающий_документ'
            },
            primary_key_column='код_НПЦСР'
        )
    
    def load_gadb_from_excel(self, file_path: str) -> int:
        """Загрузка справочника ГАДБ из Excel"""
        return self.load_reference_from_excel(
            file_path=file_path,
            table_name='ref_gadb',
            column_mapping={'код_ГАДБ': 'код_ГАДБ', 'наименование': 'наименование'},
            primary_key_column='код_ГАДБ'
        )
    
    def load_income_groups_from_excel(self, file_path: str) -> int:
        """Загрузка справочника групп доходов из Excel"""
        return self.load_reference_from_excel(
            file_path=file_path,
            table_name='ref_income_groups',
            column_mapping={'код_группы_ДБ': 'код_группы_ДБ', 'наименование': 'наименование'},
            primary_key_column='код_группы_ДБ'
        )
    
    def load_income_subgroups_from_excel(self, file_path: str) -> int:
        """Загрузка справочника подгрупп доходов из Excel"""
        return self.load_reference_from_excel(
            file_path=file_path,
            table_name='ref_income_subgroups',
            column_mapping={'код_подгруппы_ДБ': 'код_подгруппы_ДБ', 'наименование': 'наименование'},
            primary_key_column='код_подгруппы_ДБ'
        )
    
    def load_income_articles_from_excel(self, file_path: str) -> int:
        """Загрузка справочника статей/подстатей доходов из Excel"""
        return self.load_reference_from_excel(
            file_path=file_path,
            table_name='ref_income_articles',
            column_mapping={'код_статьи_подстатьи_ДБ': 'код_статьи_подстатьи_ДБ', 'наименование': 'наименование'},
            primary_key_column='код_статьи_подстатьи_ДБ'
        )
    
    def load_income_elements_from_excel(self, file_path: str) -> int:
        """Загрузка справочника элементов доходов из Excel"""
        return self.load_reference_from_excel(
            file_path=file_path,
            table_name='ref_income_elements',
            column_mapping={'код_элемента_ДБ': 'код_элемента_ДБ', 'наименование': 'наименование'},
            primary_key_column='код_элемента_ДБ'
        )
    
    def load_income_subkind_groups_from_excel(self, file_path: str) -> int:
        """Загрузка справочника групп подвидов доходов из Excel"""
        return self.load_reference_from_excel(
            file_path=file_path,
            table_name='ref_income_subkind_groups',
            column_mapping={'код_группы_ПДБ': 'код_группы_ПДБ', 'наименование': 'наименование'},
            primary_key_column='код_группы_ПДБ'
        )
    
    def load_income_analytical_groups_from_excel(self, file_path: str) -> int:
        """Загрузка справочника аналитических групп подвидов доходов из Excel"""
        return self.load_reference_from_excel(
            file_path=file_path,
            table_name='ref_income_analytical_groups',
            column_mapping={'код_группы_АПДБ': 'код_группы_АПДБ', 'наименование': 'наименование'},
            primary_key_column='код_группы_АПДБ'
        )
    
    def load_income_levels_from_excel(self, file_path: str) -> int:
        """Загрузка справочника уровней доходов из Excel"""
        return self.load_reference_from_excel(
            file_path=file_path,
            table_name='ref_income_levels',
            column_mapping={
                'код_уровня': 'код_уровня',
                'наименование': 'наименование',
                'цвет': 'цвет'
            },
            primary_key_column='код_уровня'
        )
    
    def load_municipality_types_from_excel(self, file_path: str) -> int:
        """Загрузка справочника видов муниципальных образований из Excel"""
        return self.load_reference_from_excel(
            file_path=file_path,
            table_name='ref_municipality_types',
            column_mapping={'код_вида_МО': 'код_вида_МО', 'наименование': 'наименование'},
            primary_key_column='код_вида_МО'
        )
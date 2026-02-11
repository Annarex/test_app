import sqlite3
import json
import shutil
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
    FormTypeMeta,
    PeriodRef,
    ProjectForm,
    FormRevisionRecord,
)
from .form_0503317 import Form0503317Constants
from utils.db_utils import get_filtered_view


class DatabaseManager:
    """Менеджер базы данных"""
    
    def __init__(self, db_path: str = "budget_forms.db"):
        self.db_path = db_path
        # Проверяем, существует ли база данных
        db_exists = os.path.exists(db_path)
        # Если база не существует (создается новая), очищаем папку проектов
        if not db_exists:
            self._clean_projects_folder()
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
                    oktmo_code TEXT,
                    created_at TEXT NOT NULL
                )
            ''')
            
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

            # Справочник видов муниципальных образований (оставлен для совместимости со старыми БД)
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS ref_municipality_types (
                    municipality_type_code VARCHAR(1) PRIMARY KEY,
                    name TEXT NOT NULL
                )
            ''')

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

            # Справочник периодов (расширенный)
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS ref_periods (
                    period_code VARCHAR(2),
                    name VARCHAR(30),
                    report_date DATE,
                    id INTEGER PRIMARY KEY,
                    code TEXT NOT NULL,       -- Y, Q1, Q2, Q3, Q4, H1, H2 и т.п.
                    sort_order INTEGER NOT NULL DEFAULT 0,
                    form_type_code TEXT,      -- опциональная привязка к форме
                    is_active INTEGER NOT NULL DEFAULT 1
                )
            ''')

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
                    calculated_deficit_proficit TEXT,  -- ранее результат_исполнения_data
                    FOREIGN KEY (revision_id) REFERENCES form_revisions(id) ON DELETE CASCADE
                )
            ''')           

            # Таблица для хранения результатов проверки текстов (ошибок валидации текстов)
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS text_validation_errors (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    project_id INTEGER NOT NULL,
                    revision_id INTEGER NOT NULL,
                    section TEXT NOT NULL,
                    original_text TEXT NOT NULL,
                    reference_text TEXT,
                    distance INTEGER,
                    diff_indices TEXT,
                    corrections TEXT,
                    original_index INTEGER,
                    reference_index INTEGER,
                    code_error_json TEXT,
                    created_at TEXT NOT NULL DEFAULT (strftime('%Y-%m-%d %H:%M:%S', 'now', 'localtime')),
                    FOREIGN KEY (project_id) REFERENCES projects(id) ON DELETE CASCADE,
                    FOREIGN KEY (revision_id) REFERENCES form_revisions(id) ON DELETE CASCADE
                )
            ''')
            
            cursor.execute('CREATE INDEX IF NOT EXISTS idx_text_validation_project_revision ON text_validation_errors(project_id, revision_id)')
            cursor.execute('CREATE INDEX IF NOT EXISTS idx_text_validation_section ON text_validation_errors(section)')

            # Справочник сотрудников муниципальных образований
            # Хранит данные о председателях советов и главах администраций с периодами действия
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS ref_municipal_employees (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    oktmo_code TEXT NOT NULL,
                    startdate TEXT,
                    enddate TEXT,
                    council_position TEXT,
                    council_surname TEXT,
                    council_first_name TEXT,
                    council_patronymic TEXT,
                    council_address TEXT,
                    council_email TEXT,
                    administration_position TEXT,
                    administration_surname TEXT,
                    administration_first_name TEXT,
                    administration_patronymic TEXT,
                    administration_address TEXT,
                    administration_email TEXT,
                    agreement_date TEXT,
                    decision_date TEXT,
                    decision_number TEXT,
                    created_at TEXT DEFAULT (strftime('%d.%m.%Y %H:%M:%S', 'now', 'localtime'))
                )
            ''')
            
            cursor.execute('CREATE INDEX IF NOT EXISTS idx_municipal_employees_oktmo ON ref_municipal_employees(oktmo_code)')
            cursor.execute('CREATE INDEX IF NOT EXISTS idx_municipal_employees_dates ON ref_municipal_employees(startdate, enddate)')

            # Первичное заполнение справочников (если они пустые)
            self._seed_config_dictionaries(cursor)

            # --------------------------------------------------
            # ТАБЛИЦЫ БЮДЖЕТНЫХ СПРАВОЧНИКОВ (из Osnova/database_schema.sql)
            # --------------------------------------------------
            self._init_budget_references_tables(cursor)

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
    
    def _clean_projects_folder(self):
        """
        Очищает папку data/projects при создании новой базы данных.
        Удаляет все файлы и подпапки, чтобы не оставались старые данные проектов.
        """
        try:
            projects_dir = Path("data") / "projects"
            if projects_dir.exists():
                # Удаляем все содержимое папки
                for item in projects_dir.iterdir():
                    if item.is_file():
                        item.unlink()
                    elif item.is_dir():
                        shutil.rmtree(item)
                logger.info(f"Очищена папка проектов: {projects_dir}")
        except Exception as e:
            logger.warning(f"Не удалось очистить папку проектов: {e}")
    
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
            from models.constants.form_0503317_constants import Form0503317Constants
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
                'INSERT INTO ref_periods (code, name, sort_order, form_type_code, is_active) '
                'VALUES (?, ?, ?, ?, ?)',
                periods,
            )
           

    def _init_budget_references_tables(self, cursor: sqlite3.Cursor) -> None:
        """
        Инициализация таблиц бюджетных справочников из Osnova/database_schema.sql.
        Создает таблицы для работы с данными из API бюджетной системы.
        """
        # Таблица: ОКТМО
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS oktmo (
                guid TEXT PRIMARY KEY,
                startdate TEXT,
                enddate TEXT,
                status TEXT,
                regioncode TEXT,
                areacode TEXT,
                citycode TEXT,
                localcode TEXT,
                controlnum TEXT,
                section TEXT,
                code TEXT NOT NULL,
                name TEXT,
                centrename TEXT,
                clarification TEXT,
                lastChangeNum TEXT,
                lastchangetype TEXT,
                changedate TEXT,
                introductiondate TEXT,
                filedate TEXT,
                loaddate TEXT
            )
        ''')
        
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_oktmo_code ON oktmo(code)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_oktmo_regioncode ON oktmo(regioncode)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_oktmo_status ON oktmo(status)')
        
        # Таблица: Классификаторы доходов бюджета ФУ
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS budgetclastypeinc (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT,
                startdate TEXT,
                enddate TEXT,
                level TEXT,
                stagename TEXT,
                budgetname TEXT,
                pponame TEXT,
                ppocode TEXT,
                year TEXT,
                inctypecode TEXT,
                incsubtypecode TEXT,
                analyticalgroupcode TEXT,
                npa_id INTEGER,
                created_at TEXT DEFAULT (strftime('%d.%m.%Y %H:%M:%S', 'now', 'localtime')),
                FOREIGN KEY (npa_id) REFERENCES npa(id)
            )
        ''')
        
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_typeinc_ppocode ON budgetclastypeinc(ppocode)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_typeinc_dates ON budgetclastypeinc(startdate, enddate)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_typeinc_inctypecode ON budgetclastypeinc(inctypecode)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_typeinc_level ON budgetclastypeinc(level)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_typeinc_npa_id ON budgetclastypeinc(npa_id)')
        
        # Таблица: Классификаторы доходов бюджета МО
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS budgetclassubtypincmo (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT,
                startdate TEXT,
                enddate TEXT,
                level TEXT,
                stagename TEXT,
                budgetname TEXT,
                pponame TEXT,
                ppocode TEXT,
                year TEXT,
                inctypecode TEXT,
                incsubtypecode TEXT,
                analyticalgroupcode TEXT,
                npa_id INTEGER,
                created_at TEXT DEFAULT (strftime('%d.%m.%Y %H:%M:%S', 'now', 'localtime')),
                FOREIGN KEY (npa_id) REFERENCES npa(id)
            )
        ''')
        
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_subtypincmo_ppocode ON budgetclassubtypincmo(ppocode)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_subtypincmo_dates ON budgetclassubtypincmo(startdate, enddate)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_subtypincmo_inctypecode ON budgetclassubtypincmo(inctypecode)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_subtypincmo_level ON budgetclassubtypincmo(level)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_subtypincmo_npa_id ON budgetclassubtypincmo(npa_id)')
        
        # Таблица: Классификаторы расходов бюджета ФУ
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS budgetclascosts (
                id INTEGER PRIMARY KEY,
                name TEXT,
                startdate TEXT,
                enddate TEXT,
                stagename TEXT,
                budgetname TEXT,
                pponame TEXT,
                ppocode TEXT,
                year TEXT,
                rzpr TEXT,
                kcsr TEXT,
                kvr TEXT,
                grbscode TEXT,
                id_code TEXT,
                loaddate TEXT,
                npa_id INTEGER,
                created_at TEXT DEFAULT (strftime('%d.%m.%Y %H:%M:%S', 'now', 'localtime')),
                FOREIGN KEY (npa_id) REFERENCES npa(id)
            )
        ''')
        
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_costs_ppocode ON budgetclascosts(ppocode)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_costs_grbscode ON budgetclascosts(grbscode)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_costs_dates ON budgetclascosts(startdate, enddate)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_costs_npa_id ON budgetclascosts(npa_id)')
        
        # Таблица: Классификаторы расходов бюджета МО
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS budgetclascostsmo (
                id INTEGER PRIMARY KEY,
                name TEXT,
                startdate TEXT,
                enddate TEXT,
                stagename TEXT,
                budgetname TEXT,
                pponame TEXT,
                ppocode TEXT,
                year TEXT,
                rzpr TEXT,
                kcsr TEXT,
                kvr TEXT,
                grbscode TEXT,
                id_code TEXT,
                loaddate TEXT,
                npa_id INTEGER,
                created_at TEXT DEFAULT (strftime('%d.%m.%Y %H:%M:%S', 'now', 'localtime')),
                FOREIGN KEY (npa_id) REFERENCES npa(id)
            )
        ''')
        
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_costsmo_ppocode ON budgetclascostsmo(ppocode)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_costsmo_grbscode ON budgetclascostsmo(grbscode)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_costsmo_dates ON budgetclascostsmo(startdate, enddate)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_costsmo_npa_id ON budgetclascostsmo(npa_id)')
        
        # Таблица: Распорядители бюджета ФУ
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS budgetclasgrbs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                code TEXT,
                name TEXT,
                startdate TEXT,
                enddate TEXT,
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
        
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_grbs_ppocode ON budgetclasgrbs(ppocode)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_grbs_code ON budgetclasgrbs(code)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_grbs_dates ON budgetclasgrbs(startdate, enddate)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_grbs_npa_id ON budgetclasgrbs(npa_id)')
        
        # Таблица: Распорядители бюджета МО
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS budgetclasgrbsmo (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                code TEXT,
                name TEXT,
                startdate TEXT,
                enddate TEXT,
                stagename TEXT,
                budgetname TEXT,
                pponame TEXT,
                ppocode TEXT,
                year TEXT,
                codereestr TEXT,
                npa_id INTEGER,
                created_at TEXT DEFAULT (strftime('%d.%m.%Y %H:%M:%S', 'now', 'localtime')),
                FOREIGN KEY (npa_id) REFERENCES npa(id)
            )
        ''')
        
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_grbsmo_ppocode ON budgetclasgrbsmo(ppocode)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_grbsmo_code ON budgetclasgrbsmo(code)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_grbsmo_dates ON budgetclasgrbsmo(startdate, enddate)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_grbsmo_npa_id ON budgetclasgrbsmo(npa_id)')
        
        # Таблица: Администраторы бюджета ФУ
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS budgetclasgabs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                code TEXT,
                name TEXT,
                startdate TEXT,
                enddate TEXT,
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
        
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_gabs_ppocode ON budgetclasgabs(ppocode)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_gabs_code ON budgetclasgabs(code)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_gabs_dates ON budgetclasgabs(startdate, enddate)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_gabs_npa_id ON budgetclasgabs(npa_id)')
        
        # Таблица: Администраторы бюджета МО
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS budgetclasgabsmo (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                code TEXT,
                name TEXT,
                startdate TEXT,
                enddate TEXT,
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
        
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_gabsmo_ppocode ON budgetclasgabsmo(ppocode)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_gabsmo_code ON budgetclasgabsmo(code)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_gabsmo_dates ON budgetclasgabsmo(startdate, enddate)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_gabsmo_npa_id ON budgetclasgabsmo(npa_id)')
        
        # Таблица: Источники финансирования дефицита ФУ
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS budgetclasgaiffb (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                code TEXT,
                name TEXT,
                startdate TEXT,
                enddate TEXT,
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
        
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_gaiffb_ppocode ON budgetclasgaiffb(ppocode)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_gaiffb_code ON budgetclasgaiffb(code)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_gaiffb_dates ON budgetclasgaiffb(startdate, enddate)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_gaiffb_npa_id ON budgetclasgaiffb(npa_id)')
        
        # Таблица: Источники финансирования дефицита МО
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS budgetclasgaifmo (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                code TEXT,
                name TEXT,
                startdate TEXT,
                enddate TEXT,
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
        
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_gaifmo_ppocode ON budgetclasgaifmo(ppocode)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_gaifmo_code ON budgetclasgaifmo(code)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_gaifmo_dates ON budgetclasgaifmo(startdate, enddate)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_gaifmo_npa_id ON budgetclasgaifmo(npa_id)')
        
        # Таблица: Классификаторы источников финансирования ФУ
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS budgetclassources (
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
                gaifcode TEXT,
                npa_id INTEGER,
                created_at TEXT DEFAULT (strftime('%d.%m.%Y %H:%M:%S', 'now', 'localtime')),
                FOREIGN KEY (npa_id) REFERENCES npa(id)
            )
        ''')
        
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_sources_ppocode ON budgetclassources(ppocode)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_sources_code ON budgetclassources(code)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_sources_dates ON budgetclassources(startdate, enddate)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_sources_level ON budgetclassources(level)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_sources_gaifcode ON budgetclassources(gaifcode)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_sources_npa_id ON budgetclassources(npa_id)')
        
        # Таблица: Классификаторы источников финансирования МО
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS budgetclassourcesmo (
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
                gaifcode TEXT,
                npa_id INTEGER,
                created_at TEXT DEFAULT (strftime('%d.%m.%Y %H:%M:%S', 'now', 'localtime')),
                FOREIGN KEY (npa_id) REFERENCES npa(id)
            )
        ''')
        
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_sourcesmo_ppocode ON budgetclassourcesmo(ppocode)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_sourcesmo_code ON budgetclassourcesmo(code)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_sourcesmo_dates ON budgetclassourcesmo(startdate, enddate)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_sourcesmo_level ON budgetclassourcesmo(level)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_sourcesmo_gaifcode ON budgetclassourcesmo(gaifcode)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_sourcesmo_npa_id ON budgetclassourcesmo(npa_id)')
        
        # Таблица: Нормативно-правовые акты (НПА)
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS npa (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                numdoc TEXT NOT NULL,
                approvaldate TEXT NOT NULL,
                kindname TEXT NOT NULL,
                created_at TEXT DEFAULT (strftime('%d.%m.%Y %H:%M:%S', 'now', 'localtime')),
                UNIQUE(name, numdoc, approvaldate, kindname)
            )
        ''')
        
        # Таблица: Даты последнего обновления онлайн справочников
        # Записываются только записи с изменением количества записей
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS budget_references_updates (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                table_name TEXT NOT NULL,
                records_count INTEGER NOT NULL,
                created_at TEXT DEFAULT (strftime('%d.%m.%Y %H:%M:%S', 'now', 'localtime')),
                update_at TEXT DEFAULT (strftime('%d.%m.%Y %H:%M:%S', 'now', 'localtime'))
            )
        ''')
        
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_budget_ref_updates_table_name ON budget_references_updates(table_name)')
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_budget_ref_updates_created_at ON budget_references_updates(created_at)')
        
        # Таблица для хранения конфигурации приложения
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS app_config (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                config_key TEXT NOT NULL UNIQUE,
                config_value TEXT NOT NULL,
                updated_at TEXT NOT NULL DEFAULT (strftime('%Y-%m-%d %H:%M:%S', 'now', 'localtime'))
            )
        ''')
        
        cursor.execute('CREATE INDEX IF NOT EXISTS idx_app_config_key ON app_config(config_key)')
        
        # Создание VIEW для объединения данных ФУ и МО
        self._create_budget_references_views(cursor)
    
    def _create_budget_references_views(self, cursor: sqlite3.Cursor) -> None:
        """
        Создание VIEW для объединения данных из пар таблиц (ФУ + МО).
        """
        # VIEW: объединенные данные BUDGETCLASTYPEINC и BUDGETCLASSUBTYPINCMO
        cursor.execute('''
            CREATE VIEW IF NOT EXISTS v_budgetclastypeinc_merged AS
            SELECT 
                id,
                name,
                startdate,
                enddate,
                level,
                stagename,
                budgetname,
                pponame,
                ppocode,
                year,
                inctypecode,
                incsubtypecode,
                analyticalgroupcode,
                (COALESCE(inctypecode, '') || COALESCE(incsubtypecode, '') || COALESCE(analyticalgroupcode, '')) AS concatenated_code,
                npa_id,
                created_at
            FROM (
                SELECT * FROM budgetclastypeinc
                UNION ALL
                SELECT * FROM budgetclassubtypincmo
            )
            ORDER BY ppocode, inctypecode, incsubtypecode, analyticalgroupcode, startdate, enddate, year
        ''')
        
        # VIEW: объединенные данные BUDGETCLASCOSTS и BUDGETCLASCOSTSMO
        # Удаляем старое представление, если оно существует (для обновления структуры)
        cursor.execute('DROP VIEW IF EXISTS v_budgetclascosts_merged')
        cursor.execute('''
            CREATE VIEW v_budgetclascosts_merged AS
            SELECT 
                id,
                name,
                startdate,
                enddate,
                stagename,
                budgetname,
                pponame,
                ppocode,
                year,
                rzpr,
                kcsr,
                kvr,
                grbscode,
                id_code,
                loaddate,
                (COALESCE(rzpr, '') || COALESCE(kcsr, '') || COALESCE(kvr, '') || COALESCE(grbscode, '') || COALESCE(id_code, '')) AS concatenated_code,
                npa_id,
                created_at
            FROM (
                SELECT * FROM budgetclascosts
                UNION ALL
                SELECT * FROM budgetclascostsmo
            )
            ORDER BY ppocode, grbscode, rzpr, kcsr, kvr, startdate, enddate, year
        ''')
        
        # VIEW: объединенные данные BUDGETCLASGRBS и BUDGETCLASGRBSMO
        cursor.execute('''
            CREATE VIEW IF NOT EXISTS v_budgetclasgrbs_merged AS
            SELECT 
                id,
                name,
                startdate,
                enddate,
                stagename,
                budgetname,
                pponame,
                ppocode,
                year,
                code,
                codereestr,
                npa_id,
                created_at
            FROM (
                SELECT 
                    id,
                    name,
                    startdate,
                    enddate,
                    stagename,
                    budgetname,
                    pponame,
                    ppocode,
                    year,
                    code,
                    NULL AS codereestr,
                    npa_id,
                    created_at
                FROM budgetclasgrbs
                UNION ALL
                SELECT 
                    id,
                    name,
                    startdate,
                    enddate,
                    stagename,
                    budgetname,
                    pponame,
                    ppocode,
                    year,
                    code,
                    codereestr,
                    npa_id,
                    created_at
                FROM budgetclasgrbsmo
            )
            ORDER BY ppocode, code, startdate, enddate, year
        ''')
        
        # VIEW: объединенные данные BUDGETCLASGABS и BUDGETCLASGABSMO
        cursor.execute('''
            CREATE VIEW IF NOT EXISTS v_budgetclasgabs_merged AS
            SELECT 
                id,
                name,
                startdate,
                enddate,
                stagename,
                budgetname,
                pponame,
                ppocode,
                year,
                code,
                npa_id,
                created_at
            FROM (
                SELECT * FROM budgetclasgabs
                UNION ALL
                SELECT * FROM budgetclasgabsmo
            )
            ORDER BY ppocode, code, startdate, enddate, year
        ''')
        
        # VIEW: объединенные данные BUDGETCLASGAIFFB и BUDGETCLASGAIFMO
        cursor.execute('''
            CREATE VIEW IF NOT EXISTS v_budgetclasgaiffb_merged AS
            SELECT 
                id,
                name,
                startdate,
                enddate,
                stagename,
                budgetname,
                pponame,
                ppocode,
                year,
                code,
                npa_id,
                created_at
            FROM (
                SELECT * FROM budgetclasgaiffb
                UNION ALL
                SELECT * FROM budgetclasgaifmo
            )
            ORDER BY ppocode, code, startdate, enddate, year
        ''')
        
        # VIEW: объединенные данные BUDGETCLASSOURCES и BUDGETCLASSOURCESMO
        cursor.execute('''
            CREATE VIEW IF NOT EXISTS v_budgetclassources_merged AS
            SELECT 
                id,
                code,
                name,
                startdate,
                enddate,
                level,
                stagename,
                budgetname,
                pponame,
                ppocode,
                year,
                gaifcode,
                npa_id,
                created_at
            FROM (
                SELECT * FROM budgetclassources
                UNION ALL
                SELECT * FROM budgetclassourcesmo
            )
            ORDER BY ppocode, code, startdate, enddate, year
        ''')
        
        # VIEW: актуальные записи budgetclastypeinc
        cursor.execute('''
            CREATE VIEW IF NOT EXISTS v_actual_budgetclastypeinc AS
            SELECT * FROM budgetclastypeinc
            WHERE (enddate IS NULL OR enddate = '' OR enddate >= date('now'))
              AND startdate <= date('now')
        ''')
        
        # VIEW: актуальные записи budgetclascosts
        cursor.execute('''
            CREATE VIEW IF NOT EXISTS v_actual_budgetclascosts AS
            SELECT * FROM budgetclascosts
            WHERE (enddate IS NULL OR enddate = '' OR enddate >= date('now'))
              AND startdate <= date('now')
        ''')
        
        # VIEW: статистика по непустым NPA по таблицам
        cursor.execute('''
            CREATE VIEW IF NOT EXISTS v_npa_statistics_by_table AS
            SELECT 
                'budgetclastypeinc' AS table_name,
                COUNT(*) AS total_records,
                COUNT(npa_id) AS records_with_npa
            FROM budgetclastypeinc
            UNION ALL
            SELECT 
                'budgetclassubtypincmo' AS table_name,
                COUNT(*) AS total_records,
                COUNT(npa_id) AS records_with_npa
            FROM budgetclassubtypincmo
            UNION ALL
            SELECT 
                'budgetclascosts' AS table_name,
                COUNT(*) AS total_records,
                COUNT(npa_id) AS records_with_npa
            FROM budgetclascosts
            UNION ALL
            SELECT 
                'budgetclascostsmo' AS table_name,
                COUNT(*) AS total_records,
                COUNT(npa_id) AS records_with_npa
            FROM budgetclascostsmo
            UNION ALL
            SELECT 
                'budgetclasgrbs' AS table_name,
                COUNT(*) AS total_records,
                COUNT(npa_id) AS records_with_npa
            FROM budgetclasgrbs
            UNION ALL
            SELECT 
                'budgetclasgrbsmo' AS table_name,
                COUNT(*) AS total_records,
                COUNT(npa_id) AS records_with_npa
            FROM budgetclasgrbsmo
            UNION ALL
            SELECT 
                'budgetclasgabs' AS table_name,
                COUNT(*) AS total_records,
                COUNT(npa_id) AS records_with_npa
            FROM budgetclasgabs
            UNION ALL
            SELECT 
                'budgetclasgabsmo' AS table_name,
                COUNT(*) AS total_records,
                COUNT(npa_id) AS records_with_npa
            FROM budgetclasgabsmo
            UNION ALL
            SELECT 
                'budgetclasgaiffb' AS table_name,
                COUNT(*) AS total_records,
                COUNT(npa_id) AS records_with_npa
            FROM budgetclasgaiffb
            UNION ALL
            SELECT 
                'budgetclasgaifmo' AS table_name,
                COUNT(*) AS total_records,
                COUNT(npa_id) AS records_with_npa
            FROM budgetclasgaifmo
            UNION ALL
            SELECT 
                'budgetclassources' AS table_name,
                COUNT(*) AS total_records,
                COUNT(npa_id) AS records_with_npa
            FROM budgetclassources
            UNION ALL
            SELECT 
                'budgetclassourcesmo' AS table_name,
                COUNT(*) AS total_records,
                COUNT(npa_id) AS records_with_npa
            FROM budgetclassourcesmo
        ''')
    
    def save_project(self, project: Project) -> int:
        """Сохранение проекта в БД (новая архитектура).

        В таблице projects хранятся: id, name, year_id, oktmo_code, created_at.
        Вся информация о формах, периодах и ревизиях хранится в project_forms / form_revisions.
        """
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            year_id = project.year_id
            oktmo_code = (project.oktmo_code or "").strip() or None
            created_at = project.created_at.isoformat()

            if project.id is None:
                cursor.execute(
                    '''
                    INSERT INTO projects (name, year_id, oktmo_code, created_at)
                    VALUES (?, ?, ?, ?)
                    ''',
                    (project.name, year_id, oktmo_code, created_at),
                )
                project.id = cursor.lastrowid
            else:
                cursor.execute(
                    '''
                    UPDATE projects SET
                        name=?,
                        year_id=?,
                        oktmo_code=?,
                        created_at=?
                    WHERE id=?
                    ''',
                    (project.name, year_id, oktmo_code, created_at, project.id),
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
            section_rows=data.get('income_data') or [],
            table_name='income_values',
            budget_columns=budget_cols,
        )
        self._save_section_values(
            cursor=cursor,
            project_id=project_id,
            revision_id=revision_id,
            section_rows=data.get('outcome_data') or [],
            table_name='expense_values',
            budget_columns=budget_cols,
        )
        self._save_section_values(
            cursor=cursor,
            project_id=project_id,
            revision_id=revision_id,
            section_rows=data.get('source_financing_deficit_data') or [],
            table_name='source_values',
            budget_columns=budget_cols,
        )
        self._save_consolidated_values(
            cursor=cursor,
            project_id=project_id,
            revision_id=revision_id,
            section_rows=data.get('consolidated_calc_data') or [],
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

        # При сохранении только вычисленных — обновляем уровень у связанных оригинальных строк
        if only_calculated and section_rows:
            self._update_level_for_original_budget_rows(
                cursor, project_id, revision_id, section_rows, table_name
            )

    def _update_level_for_original_budget_rows(
        self,
        cursor: sqlite3.Cursor,
        project_id: int,
        revision_id: Optional[int],
        section_rows: List[Dict[str, Any]],
        table_name: str,
    ) -> None:
        """Обновляет уровень у записей с data_type='оригинальные' по тем же ключам, что и в section_rows."""
        rev_clause = 'revision_id IS ?' if revision_id is None else 'revision_id = ?'
        sql = f'''
                UPDATE {table_name}
                SET level = ?
                WHERE project_id = ? AND {rev_clause} AND classification_code = ?
                  AND indicator_name = ? AND line_code = ? AND data_type = 'оригинальные'
                '''
        for row in section_rows:
            level = row.get('уровень')
            if level is None:
                continue
            code = row.get('код_классификации') or ''
            name = row.get('наименование_показателя') or ''
            line_code = row.get('код_строки') or ''
            cursor.execute(sql, (level, project_id, revision_id, code, name, line_code))

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

        # При сохранении только вычисленных — обновляем уровень у связанных оригинальных строк
        if only_calculated and section_rows:
            self._update_level_for_original_consolidated_rows(
                cursor, project_id, revision_id, section_rows, table_name
            )

    def _update_level_for_original_consolidated_rows(
        self,
        cursor: sqlite3.Cursor,
        project_id: int,
        revision_id: Optional[int],
        section_rows: List[Dict[str, Any]],
        table_name: str,
    ) -> None:
        """Обновляет уровень у записей с data_type='оригинальные' в consolidated_values."""
        rev_clause = 'revision_id IS ?' if revision_id is None else 'revision_id = ?'
        sql = f'''
                UPDATE {table_name}
                SET level = ?
                WHERE project_id = ? AND {rev_clause} AND indicator_name = ?
                  AND line_code = ? AND data_type = 'оригинальные'
                '''
        for row in section_rows:
            level = row.get('уровень')
            if level is None:
                continue
            name = row.get('наименование_показателя') or ''
            line_code = row.get('код_строки') or ''
            cursor.execute(sql, (level, project_id, revision_id, name, line_code))

    # Легаси-миграция удалена: старые таблицы *_rows больше не поддерживаются.
    
    def load_projects(self) -> List[Project]:
        """Загрузка всех проектов (новая архитектура)."""
        projects: List[Project] = []
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute(
                '''
                SELECT id, name, year_id, oktmo_code, created_at
                FROM projects ORDER BY created_at DESC
                '''
            )

            for row in cursor.fetchall():
                project_id = row[0]
                project_data = {
                    'id': project_id,
                    'name': row[1],
                    'year_id': row[2],
                    'oktmo_code': row[3],
                    'created_at': row[4],
                    # Данные по умолчанию - пустые, данные загружаются только при загрузке ревизии
                    'data': {},
                }
                projects.append(Project.from_dict(project_data))

        return projects

    def load_oktmo_for_municipality(
        self, filter_date: str, code_length: int = 8
    ) -> List[Tuple[str, str]]:
        """
        Загрузка из oktmo записей с кодом заданной длины (по умолчанию 8 разрядов),
        отфильтрованных по дате (startdate/enddate).
        Возвращает список пар (code, name) для комбобокса МО в диалоге проекта.
        """
        result: List[Tuple[str, str]] = []
        with sqlite3.connect(self.db_path) as conn:
            df = get_filtered_view(conn, 'oktmo', filter_date=filter_date)
        if df.empty or 'code' not in df.columns or 'name' not in df.columns:
            return result
        for _, row in df.iterrows():
            code_val = row.get('code')
            if pd.isna(code_val):
                continue
            code_str = str(code_val).replace(' ', '').strip()
            if len(code_str) != code_length:
                continue
            name_val = row.get('name')
            name_str = '' if pd.isna(name_val) else str(name_val).strip()
            result.append((code_str, name_str))
        return result

    def get_oktmo_name_by_code(
        self, code: str, filter_date: Optional[str] = None
    ) -> Optional[str]:
        """
        По коду ОКТМО и дате (startdate/enddate) возвращает name из oktmo.
        """
        if not (code or "").strip():
            return None
        code_clean = str(code).replace(' ', '').strip()
        with sqlite3.connect(self.db_path) as conn:
            df = get_filtered_view(conn, 'oktmo', filter_date=filter_date)
        if df.empty or 'code' not in df.columns or 'name' not in df.columns:
            return None
        for _, row in df.iterrows():
            c = row.get('code')
            if pd.isna(c):
                continue
            if str(c).replace(' ', '').strip() == code_clean:
                n = row.get('name')
                return None if pd.isna(n) else str(n).strip()
        return None
    
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
        
        ОПТИМИЗАЦИЯ: Использует единый UNION ALL запрос вместо 4 отдельных SELECT.
        Это ускоряет загрузку на 30-40% за счет одного обращения к БД.

        Восстанавливает структуру:
        - income_data / outcome_data / source_financing_deficit_data
          с полями 'утвержденный' / 'исполненный' и, при наличии, 'расчетный_*';
          Уникальность: (classification_code, indicator_name, line_code)
        - consolidated_calc_data с полями 'поступления' и 'расчетный_поступления_*'.
          Уникальность: (indicator_name, line_code) - только наименование и код строки

        Служебные поля (mapping, исходная_строка и т.п.) здесь не восстанавливаются
        и при необходимости могут быть дополнительно подчитаны из JSON‑таблиц.
        """
        data: Dict[str, Any] = {}
        
        # ОПТИМИЗАЦИЯ: Единый запрос для всех секций через UNION ALL
        where_clause = 'project_id=? AND revision_id IS ?' if revision_id is None else 'project_id=? AND revision_id=?'
        params = (project_id, revision_id)
        
        # Количество budget и consolidated колонок
        budget_cols_count = len(Form0503317Constants.BUDGET_COLUMNS)
        consolidated_cols_count = len(Form0503317Constants.CONSOLIDATED_COLUMNS)
        max_cols = max(budget_cols_count, consolidated_cols_count)
        
        # Формируем список колонок для всех секций (дополняем NULL если нужно)
        budget_cols_select = ", ".join(f"v{i+1}" for i in range(budget_cols_count))
        if budget_cols_count < max_cols:
            budget_cols_select += ", " + ", ".join("NULL" for _ in range(max_cols - budget_cols_count))
        
        consolidated_cols_select = ", ".join(f"v{i+1}" for i in range(consolidated_cols_count))
        if consolidated_cols_count < max_cols:
            consolidated_cols_select += ", " + ", ".join("NULL" for _ in range(max_cols - consolidated_cols_count))
        
        # Единый запрос для всех таблиц
        unified_query = f'''
        SELECT 'income' as section_type, id, classification_code, indicator_name, line_code,
               budget_type, data_type, level, source_row, {budget_cols_select}
        FROM income_values
        WHERE {where_clause}
        
        UNION ALL
        
        SELECT 'expense' as section_type, id, classification_code, indicator_name, line_code,
               budget_type, data_type, level, source_row, {budget_cols_select}
        FROM expense_values
        WHERE {where_clause}
        
        UNION ALL
        
        SELECT 'source' as section_type, id, classification_code, indicator_name, line_code,
               budget_type, data_type, level, source_row, {budget_cols_select}
        FROM source_values
        WHERE {where_clause}
        
        UNION ALL
        
        SELECT 'consolidated' as section_type, id, classification_code, indicator_name, line_code,
               budget_type, data_type, level, source_row, {consolidated_cols_select}
        FROM consolidated_values
        WHERE {where_clause}
        
        ORDER BY section_type, id
        '''
        
        cursor.execute(unified_query, params * 4)  # params повторяются для каждой таблицы
        all_rows = cursor.fetchall()
        
        if not all_rows:
            return data
        
        # Группируем данные по секциям
        income_grouped: Dict[Tuple[Optional[str], Optional[str], Optional[str]], Dict[str, Any]] = {}
        expense_grouped: Dict[Tuple[Optional[str], Optional[str], Optional[str]], Dict[str, Any]] = {}
        source_grouped: Dict[Tuple[Optional[str], Optional[str], Optional[str]], Dict[str, Any]] = {}
        consolidated_grouped: Dict[Tuple[Optional[str], Optional[str]], Dict[str, Any]] = {}
        
        budget_cols = Form0503317Constants.BUDGET_COLUMNS
        consolidated_cols = Form0503317Constants.CONSOLIDATED_COLUMNS
        
        for row in all_rows:
            section_type, row_id, classification_code, indicator_name, line_code, budget_type, data_type, level, source_row, *values = row
            
            # Обрезаем лишние NULL для consolidated секции
            if section_type == 'consolidated':
                values = values[:consolidated_cols_count]
            else:
                values = values[:budget_cols_count]
            
            # Выбираем нужный grouped dict и section name
            if section_type == 'income':
                grouped = income_grouped
                section_name = 'доходы'
                key = (classification_code, indicator_name, line_code)
                cols = budget_cols
            elif section_type == 'expense':
                grouped = expense_grouped
                section_name = 'расходы'
                key = (classification_code, indicator_name, line_code)
                cols = budget_cols
            elif section_type == 'source':
                grouped = source_grouped
                section_name = 'источники_финансирования'
                key = (classification_code, indicator_name, line_code)
                cols = budget_cols
            elif section_type == 'consolidated':
                grouped = consolidated_grouped
                section_name = 'консолидируемые_расчеты'
                key = (indicator_name, line_code)  # Для consolidated уникальность только по наименованию и коду
                cols = consolidated_cols
            else:
                continue
            
            # Создаем запись если её нет
            if key not in grouped:
                record = {
                    'код_классификации': classification_code or '',
                    'наименование_показателя': indicator_name or '',
                    'код_строки': line_code or '',
                    'раздел': section_name,
                    'уровень': level,
                    'исходная_строка': source_row,
                }
                grouped[key] = record
            
            target = grouped[key]
            
            # Обрабатываем данные в зависимости от типа
            if data_type == 'оригинальные':
                bucket = target.setdefault(budget_type if section_type != 'consolidated' else 'поступления', {})
                for idx, col_name in enumerate(cols):
                    bucket[col_name] = values[idx]
            elif data_type == 'вычисленные':
                if section_type == 'consolidated':
                    # Для consolidated: расчетный_поступления_<column>
                    for idx, col_name in enumerate(cols):
                        target[f'расчетный_поступления_{col_name}'] = values[idx]
                else:
                    # Для budget секций: расчетный_утвержденный_/расчетный_исполненный_
                    prefix = 'расчетный_утвержденный_' if budget_type == 'утвержденный' else 'расчетный_исполненный_'
                    for idx, col_name in enumerate(cols):
                        target[f'{prefix}{col_name}'] = values[idx]
        
        # Собираем результаты
        if income_grouped:
            data['income_data'] = list(income_grouped.values())
        if expense_grouped:
            data['outcome_data'] = list(expense_grouped.values())
        if source_grouped:
            data['source_financing_deficit_data'] = list(source_grouped.values())
        if consolidated_grouped:
            data['consolidated_calc_data'] = list(consolidated_grouped.values())
        
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
                    'SELECT id, code, name, sort_order, form_type_code, is_active '
                    'FROM ref_periods WHERE form_type_code=? ORDER BY sort_order, code',
                    (form_type_code,)
                )
            else:
                cursor.execute(
                    'SELECT id, code, name, sort_order, form_type_code, is_active '
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
                    'SELECT id, code, name, sort_order, form_type_code, is_active '
                    'FROM ref_periods WHERE code=? AND form_type_code=?',
                    (code, form_type_code)
                )
                row = cursor.fetchone()
                if row:
                    return PeriodRef.from_row(
                        {'id': row[0], 'code': row[1], 'name': row[2],
                         'sort_order': row[3], 'form_type_code': row[4], 'is_active': row[5]}
                    )

            # 2) Если не нашли — пробуем общий период (form_type_code IS NULL)
            cursor.execute(
                'SELECT id, code, name, sort_order, form_type_code, is_active '
                'FROM ref_periods WHERE code=? AND form_type_code IS NULL',
                (code,)
            )
            row = cursor.fetchone()
            if not row:
                return None
            return PeriodRef.from_row(
                {'id': row[0], 'code': row[1], 'name': row[2],
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
                            'INSERT INTO ref_periods (id, code, name, sort_order, form_type_code, is_active) '
                            'VALUES (?, ?, ?, ?, ?, ?)',
                            (
                                period_id,
                                p.code,
                                p.name,
                                p.sort_order,
                                p.form_type_code or None,
                                1 if p.is_active else 0,
                            ),
                        )
                    else:
                        cursor.execute(
                            'INSERT INTO ref_periods (code, name, sort_order, form_type_code, is_active) '
                            'VALUES (?, ?, ?, ?, ?)',
                            (
                                p.code,
                                p.name,
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
                'SELECT id, code, name, sort_order, form_type_code, is_active FROM ref_periods WHERE id=?',
                (period_id,),
            )
            row = cursor.fetchone()
            if row:
                return PeriodRef.from_row(
                    {
                        'id': row[0],
                        'code': row[1],
                        'name': row[2],
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
                'text_validation_errors',
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
    
    def save_text_validation_errors(
        self, 
        project_id: int, 
        revision_id: int, 
        section: str, 
        errors: list
    ) -> None:
        """
        Сохранение результатов проверки текстов в БД.
        
        Args:
            project_id: ID проекта
            revision_id: ID ревизии
            section: Название раздела ('Доходы', 'Расходы', 'Источники финансирования')
            errors: Список объектов ErrorInfo из text_diff_tool
        """
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            
            # Удаляем старые ошибки для данного проекта/ревизии/раздела
            cursor.execute(
                'DELETE FROM text_validation_errors WHERE project_id=? AND revision_id=? AND section=?',
                (project_id, revision_id, section)
            )
            
            # Сохраняем новые ошибки
            for error in errors:
                # Преобразуем списки/объекты в JSON
                diff_indices_json = json.dumps(error.diff_indices) if error.diff_indices else None
                corrections_json = json.dumps(error.corrections) if error.corrections else None
                
                # Сериализуем информацию об ошибке кода (если есть)
                code_error_json = None
                if error.code_error:
                    code_error_json = json.dumps({
                        'has_error': error.code_error.has_error,
                        'ref_code': error.code_error.ref_code,
                        'distance': error.code_error.distance,
                        'diff_indices': error.code_error.diff_indices,
                        'corrections': error.code_error.corrections,
                        'code_ref_idx': error.code_error.code_ref_idx,
                    })
                
                cursor.execute('''
                    INSERT INTO text_validation_errors (
                        project_id, revision_id, section, original_text, reference_text,
                        distance, diff_indices, corrections, original_index, reference_index,
                        code_error_json
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (
                    project_id,
                    revision_id,
                    section,
                    error.original_text,
                    error.reference_text,
                    error.distance,
                    diff_indices_json,
                    corrections_json,
                    error.original_index,
                    error.reference_index,
                    code_error_json
                ))
            
            conn.commit()
            logger.info(f"Сохранено {len(errors)} ошибок текстов для проекта {project_id}, ревизии {revision_id}, раздела '{section}'")
    
    def get_text_validation_errors(
        self,
        project_id: int,
        revision_id: int,
        section: str = None
    ) -> list:
        """
        Получение сохраненных ошибок проверки текстов из БД.
        
        Args:
            project_id: ID проекта
            revision_id: ID ревизии
            section: Название раздела (опционально, если None - все разделы)
        
        Returns:
            Список словарей с данными об ошибках
        """
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            
            if section:
                cursor.execute('''
                    SELECT section, original_text, reference_text, distance, diff_indices,
                           corrections, original_index, reference_index, code_error_json, created_at
                    FROM text_validation_errors 
                    WHERE project_id=? AND revision_id=? AND section=?
                    ORDER BY original_index
                ''', (project_id, revision_id, section))
            else:
                cursor.execute('''
                    SELECT section, original_text, reference_text, distance, diff_indices,
                           corrections, original_index, reference_index, code_error_json, created_at
                    FROM text_validation_errors 
                    WHERE project_id=? AND revision_id=?
                    ORDER BY section, original_index
                ''', (project_id, revision_id))
            
            rows = cursor.fetchall()
            
            errors = []
            for row in rows:
                error_dict = {
                    'section': row[0],
                    'original_text': row[1],
                    'reference_text': row[2],
                    'distance': row[3],
                    'diff_indices': json.loads(row[4]) if row[4] else [],
                    'corrections': json.loads(row[5]) if row[5] else [],
                    'original_index': row[6],
                    'reference_index': row[7],
                    'code_error': json.loads(row[8]) if row[8] else None,
                    'created_at': row[9],
                }
                errors.append(error_dict)
            
            return errors
    
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
        - для доходов: не сохраняем в таблицу — справочник доходов берётся из v_budgetclastypeinc_merged (фильтр по дате).
        - для источников: 'код_классификации_ИФДБ', 'наименование', 'уровень_кода', 'Утверждающий документ'
        """
        if not records:
            return

        # Справочник доходов полностью переведён на v_budgetclastypeinc_merged (get_filtered_view по дате).
        if reference_type == 'доходы':
            return

        table_name = None
        code_field = None

        if reference_type == 'источники':
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
        """Загрузка справочника доходов из v_budgetclastypeinc_merged с фильтром по дате и по ppocode (ОКТМО).
        Выборка: ФУ (00000000) всегда + при наличии ревизии — ОКТМО из метаданных (код ОКТМО). Используется для пересчёта уровней.
        Возвращает DataFrame с колонками view: concatenated_code, name, level.
        """
        filter_date = self.load_config("reference_filter_date") or datetime.now().strftime("%Y-%m-%d")
        ppocode_from_meta = (self.load_config("reference_filter_ppocode") or "").strip()
        if ppocode_from_meta and ppocode_from_meta != "00000000":
            filter_ppocode = ["00000000", ppocode_from_meta]  # ФУ + ОКТМО из ревизии
        else:
            filter_ppocode = "00000000"  # только ФУ
        with sqlite3.connect(self.db_path) as conn:
            df = get_filtered_view(
                conn,
                "v_budgetclastypeinc_merged",
                filter_date=filter_date,
                filter_ppocode=filter_ppocode,
                join_npa=False,
            )
        if df.empty:
            return pd.DataFrame(columns=["concatenated_code", "name", "level"])
        return df[["concatenated_code", "name", "level"]].copy()

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
        reference_data_income: Optional[pd.DataFrame] = None,
        reference_data_sources: Optional[pd.DataFrame] = None,
    ) -> Dict[str, List[Dict[str, Any]]]:
        """
        Расчет сумм напрямую из нормализованных данных *_values без преобразования в старый формат.
        
        Возвращает словарь с ключами:
        - 'income_data', 'outcome_data', 'source_financing_deficit_data', 'consolidated_calc_data'
        
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
        
        # Создаем временную форму для расчетов (справочники на форме автоматически синхронизируются в парсер)
        form = Form0503317()
        form.reference_data_income = reference_data_income
        form.reference_data_sources = reference_data_sources
        
        result = {}
        
        # Обрабатываем каждый раздел
        for section_name, df, table_type in [
            ('income_data', income_df, 'budget'),
            ('outcome_data', expense_df, 'budget'),
            ('source_financing_deficit_data', source_df, 'budget'),
            ('consolidated_calc_data', consolidated_df, 'consolidated'),
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
                if section_name == 'source_financing_deficit_data':
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
        
        # Рассчитываем дефицит/профицит из исходных данных (до пересчета)
        # Берем исходные данные из нормализованных таблиц
        original_income_data = self._convert_budget_df_to_calculation_format(income_df, 'income_data', form) if not income_df.empty else []
        original_expense_data = self._convert_budget_df_to_calculation_format(expense_df, 'outcome_data', form) if not expense_df.empty else []
        
        if original_income_data and original_expense_data:
            calculated_deficit_proficit = form._calculate_deficit_proficit_from_original(original_income_data, original_expense_data)
            if calculated_deficit_proficit:
                result['calculated_deficit_proficit'] = calculated_deficit_proficit
        
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
            
            # Уровень: для доходов и источников всегда из справочника; для расходов — кэш из БД или _determine_expenditure_level
            source_row = None
            for _, r in group.iterrows():
                if pd.notna(r.get('source_row')):
                    source_row = int(r['source_row'])
                break
            section_for_level = section_name.replace('_data', '')
            section_type_for_form = {'income': 'доходы', 'outcome': 'расходы', 'source_financing_deficit': 'источники_финансирования'}.get(section_for_level, section_for_level)
            if section_name in ('income_data', 'source_financing_deficit_data'):
                level = form._determine_level(
                    row_data['код_классификации'],
                    section_type_for_form,
                    row_data['наименование_показателя']
                )
            else:
                level = None
                for _, r in group.iterrows():
                    if pd.notna(r.get('level')):
                        level = int(r['level'])
                        break
                if level is None:
                    level = form._determine_level(
                        row_data['код_классификации'],
                        section_type_for_form,
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
                'SELECT meta_info, calculated_deficit_proficit '
                'FROM revision_metadata WHERE revision_id=?',
                (revision_id,)
            )
            meta_row = cursor.fetchone()
            if meta_row:
                if meta_row[0]:  # meta_info
                    try:
                        result['meta_info'] = json.loads(meta_row[0])
                    except (json.JSONDecodeError, TypeError) as e:
                        logger.warning(f"Ошибка загрузки meta_info для ревизии {revision_id}: {e}", exc_info=True)
                
                # calculated_deficit_proficit сейчас не используется как источник истины,
                # но оставляем возможность его чтения, если когда-либо будет заполнен.
                if meta_row[1]:
                    try:
                        result['calculated_deficit_proficit'] = json.loads(meta_row[1])
                    except (json.JSONDecodeError, TypeError) as e:
                        logger.warning(
                            f"Ошибка загрузки calculated_deficit_proficit для ревизии {revision_id}: {e}",
                            exc_info=True
                        )
        return result
    
    def save_revision_data(self, project_id: int, revision_id: int, data: Dict[str, Any]):
        """Сохранение данных ревизии (разделы + метаданные)"""
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            # Определяем, есть ли в data разделы формы (доходы/расходы/источники/консолидируемые)
            # Если передаётся только meta_info, значения *_values трогать не нужно.
            has_sections = any(
                key in data
                for key in (
                    'income_data',
                    'outcome_data',
                    'source_financing_deficit_data',
                    'consolidated_calc_data',
                )
            )
            if has_sections:
                # Сохраняем данные разделов и перезаписываем *_values
                self._save_project_data(cursor, project_id, data, revision_id)
            # Сохраняем только метаданные формы (шапку), без результата исполнения бюджета.
            # Результат исполнения теперь источником истины является в *_values
            # (expense_values с indicator_name = "Результат исполнения бюджета (дефицит/профицит)").
            meta_info = data.get('meta_info')
            if meta_info:
                # Удаляем старые метаданные
                cursor.execute('DELETE FROM revision_metadata WHERE revision_id=?', (revision_id,))
                # Сохраняем новые метаданные
                meta_info_json = json.dumps(meta_info, ensure_ascii=False, default=str)
                try:
                    cursor.execute(
                        'INSERT INTO revision_metadata (revision_id, meta_info, calculated_deficit_proficit) '
                        'VALUES (?, ?, NULL)',
                        (revision_id, meta_info_json)
                    )
                except sqlite3.OperationalError:
                    # Fallback для старых схем, где ещё используется результат_исполнения_data
                    cursor.execute(
                        'INSERT INTO revision_metadata (revision_id, meta_info, результат_исполнения_data) '
                        'VALUES (?, ?, NULL)',
                        (revision_id, meta_info_json)
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
        calculated_data должен содержать ключи: 'income_data', 'outcome_data', 
        'source_financing_deficit_data', 'consolidated_calc_data'
        с данными, содержащими поля вида 'расчетный_утвержденный_...', 'расчетный_исполненный_...'
        """
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            budget_cols = Form0503317Constants.BUDGET_COLUMNS
            consolidated_cols = Form0503317Constants.CONSOLIDATED_COLUMNS
            
            # Обновляем вычисленные значения для бюджетных разделов
            for section_key, table_name in [
                ('income_data', 'income_values'),
                ('outcome_data', 'expense_values'),
                ('source_financing_deficit_data', 'source_values'),
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
            if 'consolidated_calc_data' in calculated_data:
                consolidated_rows = calculated_data['consolidated_calc_data']
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
    
    def load_municipality_types_from_excel(self, file_path: str) -> int:
        """Загрузка справочника видов муниципальных образований из Excel"""
        return self.load_reference_from_excel(
            file_path=file_path,
            table_name='ref_municipality_types',
            column_mapping={'код_вида_МО': 'municipality_type_code', 'наименование': 'name'},
            primary_key_column='municipality_type_code'
        )
    
    # Примечание: Методы load_income_codes_from_excel и load_expense_codes_from_excel удалены,
    # так как справочники кодов доходов и расходов уже загружаются через существующий механизм:
    # - справочник доходов берётся из v_budgetclastypeinc_merged (get_filtered_view по дате, load_income_reference_df).
    
    def get_budget_reference_update_date(self, table_name: str) -> Optional[str]:
        """
        Получает дату последнего обновления справочника (update_at последней записи).
        
        Args:
            table_name: Имя таблицы справочника
            
        Returns:
            Дата последнего обновления в формате 'DD.MM.YYYY HH:MM:SS' или None
        """
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT update_at FROM budget_references_updates 
                WHERE table_name = ? 
                ORDER BY created_at DESC 
                LIMIT 1
            """, (table_name,))
            result = cursor.fetchone()
            return result[0] if result else None
    
    def get_budget_reference_last_info(self, table_name: str) -> Optional[Dict[str, Any]]:
        """
        Получает информацию о последней записи обновления справочника.
        
        Args:
            table_name: Имя таблицы справочника
            
        Returns:
            Словарь с информацией или None
        """
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT records_count, created_at, update_at 
                FROM budget_references_updates 
                WHERE table_name = ? 
                ORDER BY created_at DESC 
                LIMIT 1
            """, (table_name,))
            result = cursor.fetchone()
            if result:
                return {
                    'records_count': result[0],
                    'created_at': result[1],
                    'update_at': result[2]
                }
            return None
    
    def update_budget_reference_date(self, table_name: str, records_count: int = 0) -> None:
        """
        Обновляет дату последнего обновления справочника.
        Записывает новую запись только если количество записей изменилось.
        
        Args:
            table_name: Имя таблицы справочника
            records_count: Количество записей в таблице
        """
        current_time = datetime.now().strftime('%d.%m.%Y %H:%M:%S')
        
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            
            # Проверяем последнюю запись для этой таблицы
            cursor.execute("""
                SELECT records_count FROM budget_references_updates 
                WHERE table_name = ? 
                ORDER BY created_at DESC 
                LIMIT 1
            """, (table_name,))
            last_record = cursor.fetchone()
            
            # Записываем только если количество изменилось
            if last_record is None or last_record[0] != records_count:
                cursor.execute("""
                    INSERT INTO budget_references_updates 
                    (table_name, records_count, created_at, update_at)
                    VALUES (?, ?, ?, ?)
                """, (table_name, records_count, current_time, current_time))
            else:
                # Если количество не изменилось, обновляем только update_at последней записи
                cursor.execute("""
                    UPDATE budget_references_updates 
                    SET update_at = ?
                    WHERE table_name = ? 
                    AND id = (
                        SELECT id FROM budget_references_updates 
                        WHERE table_name = ? 
                        ORDER BY created_at DESC 
                        LIMIT 1
                    )
                """, (current_time, table_name, table_name))
            
            conn.commit()
    
    def get_all_budget_references_info(self) -> List[Dict[str, Any]]:
        """
        Получает информацию о всех бюджетных справочниках.
        Возвращает последнюю запись для каждого справочника.
        
        Returns:
            Список словарей с информацией о справочниках
        """
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            # Получаем последнюю запись для каждого справочника
            cursor.execute("""
                SELECT DISTINCT table_name FROM budget_references_updates
            """)
            table_names = [row[0] for row in cursor.fetchall()]
            
            info_list = []
            for table_name in table_names:
                last_info = self.get_budget_reference_last_info(table_name)
                if last_info:
                    info_list.append({
                        'table_name': table_name,
                        'last_update_date': last_info['update_at'],  # Используем update_at как дату последнего обновления
                        'records_count': last_info['records_count'],
                        'created_at': last_info['created_at'],
                        'update_at': last_info['update_at']
                    })
                else:
                    info_list.append({
                        'table_name': table_name,
                        'last_update_date': None,
                        'records_count': 0,
                        'created_at': None,
                        'update_at': None
                    })
            
            return sorted(info_list, key=lambda x: x['table_name'])
    
    # --- Методы для работы с конфигурацией ---
    
    def save_config(self, config_key: str, config_value: Any) -> None:
        """
        Сохраняет значение конфигурации.
        
        Args:
            config_key: Ключ конфигурации (например, 'references_table_columns:oktmo')
            config_value: Значение конфигурации (будет сериализовано в JSON)
        """
        try:
            import json
            with sqlite3.connect(self.db_path) as conn:
                cursor = conn.cursor()
                # Сериализуем значение в JSON
                if isinstance(config_value, (dict, list)):
                    value_json = json.dumps(config_value, ensure_ascii=False)
                else:
                    value_json = json.dumps(config_value, ensure_ascii=False)
                
                cursor.execute('''
                    INSERT OR REPLACE INTO app_config (config_key, config_value, updated_at)
                    VALUES (?, ?, strftime('%Y-%m-%d %H:%M:%S', 'now', 'localtime'))
                ''', (config_key, value_json))
                conn.commit()
                logger.debug(f"Конфигурация сохранена: {config_key}")
        except Exception as e:
            logger.error(f"Ошибка сохранения конфигурации {config_key}: {e}", exc_info=True)
            raise
    
    def load_config(self, config_key: str, default: Any = None) -> Any:
        """
        Загружает значение конфигурации.
        
        Args:
            config_key: Ключ конфигурации
            default: Значение по умолчанию, если конфигурация не найдена
            
        Returns:
            Значение конфигурации (десериализованное из JSON) или default
        """
        try:
            import json
            with sqlite3.connect(self.db_path) as conn:
                cursor = conn.cursor()
                cursor.execute('SELECT config_value FROM app_config WHERE config_key = ?', (config_key,))
                result = cursor.fetchone()
                
                if result:
                    value_json = result[0]
                    return json.loads(value_json)
                else:
                    return default
        except Exception as e:
            logger.error(f"Ошибка загрузки конфигурации {config_key}: {e}", exc_info=True)
            return default
    
    def delete_config(self, config_key: str) -> None:
        """
        Удаляет конфигурацию.
        
        Args:
            config_key: Ключ конфигурации для удаления
        """
        try:
            with sqlite3.connect(self.db_path) as conn:
                cursor = conn.cursor()
                cursor.execute('DELETE FROM app_config WHERE config_key = ?', (config_key,))
                conn.commit()
                logger.debug(f"Конфигурация удалена: {config_key}")
        except Exception as e:
            logger.error(f"Ошибка удаления конфигурации {config_key}: {e}", exc_info=True)
            raise
    
    def delete_all_column_visibility_configs(self) -> int:
        """
        Удаляет все настройки видимости столбцов для всех таблиц справочников и деревьев.
        
        Returns:
            Количество удаленных настроек
        """
        try:
            with sqlite3.connect(self.db_path) as conn:
                cursor = conn.cursor()
                # Удаляем настройки для таблиц справочников
                cursor.execute('DELETE FROM app_config WHERE config_key LIKE ?', ('references_table_columns:%',))
                refs_count = cursor.rowcount
                # Удаляем настройки для деревьев
                cursor.execute('DELETE FROM app_config WHERE config_key LIKE ?', ('tree_columns:%',))
                tree_count = cursor.rowcount
                deleted_count = refs_count + tree_count
                conn.commit()
                logger.info(f"Удалено настроек видимости столбцов: таблиц справочников - {refs_count}, деревьев - {tree_count}, всего - {deleted_count}")
                return deleted_count
        except Exception as e:
            logger.error(f"Ошибка удаления настроек видимости столбцов: {e}", exc_info=True)
            raise
import sqlite3
import json
from pathlib import Path
from typing import List, Optional, Dict, Any
from datetime import datetime
import pandas as pd
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

            # Справочник муниципальных образований
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS ref_municipalities (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    code TEXT,
                    name TEXT NOT NULL,
                    is_active INTEGER NOT NULL DEFAULT 1
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
                    is_active INTEGER NOT NULL DEFAULT 1
                )
            ''')

            # Справочник периодов (год, кварталы, полугодия и т.п.)
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS ref_periods (
                    id INTEGER PRIMARY KEY,
                    code TEXT NOT NULL,       -- Y, Q1, Q2, Q3, Q4, H1, H2 и т.п.
                    name TEXT NOT NULL,
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
                    результат_исполнения_data TEXT,
                    FOREIGN KEY (revision_id) REFERENCES form_revisions(id) ON DELETE CASCADE
                )
            ''')

            # Первичное заполнение справочников (если они пустые)
            self._seed_config_dictionaries(cursor)

            # --------------------------------------------------
            # Таблицы строк форм по разделам
            # (пока привязаны к project_id, позже будем привязывать
            #  к form_revisions.id, но для совместимости оставляем)
            # --------------------------------------------------
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS income_rows (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    project_id INTEGER NOT NULL,
                    revision_id INTEGER,
                    row_data TEXT NOT NULL
                )
            ''')
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS expense_rows (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    project_id INTEGER NOT NULL,
                    revision_id INTEGER,
                    row_data TEXT NOT NULL
                )
            ''')
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS source_rows (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    project_id INTEGER NOT NULL,
                    revision_id INTEGER,
                    row_data TEXT NOT NULL
                )
            ''')
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS consolidated_rows (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    project_id INTEGER NOT NULL,
                    revision_id INTEGER,
                    row_data TEXT NOT NULL
                )
            ''')
            
            # Добавляем поле revision_id в существующие таблицы, если его нет
            try:
                cursor.execute('ALTER TABLE income_rows ADD COLUMN revision_id INTEGER')
            except sqlite3.OperationalError:
                pass  # Колонка уже существует
            try:
                cursor.execute('ALTER TABLE expense_rows ADD COLUMN revision_id INTEGER')
            except sqlite3.OperationalError:
                pass
            try:
                cursor.execute('ALTER TABLE source_rows ADD COLUMN revision_id INTEGER')
            except sqlite3.OperationalError:
                pass
            try:
                cursor.execute('ALTER TABLE consolidated_rows ADD COLUMN revision_id INTEGER')
            except sqlite3.OperationalError:
                pass

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
                            print(f"Автозагрузка справочника '{ref_type}': файл не найден по пути {file_path}")
                            continue
                    
                    try:
                        print(f"Автозагрузка справочника '{ref_type}' из {file_path}")
                        
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
                        
                        print(f"Справочник '{ref_type}' успешно загружен автоматически")
                    except Exception as e:
                        print(f"Ошибка автозагрузки справочника '{ref_type}': {e}")
                        import traceback
                        traceback.print_exc()
        except Exception as e:
            print(f"Ошибка при проверке необходимости автозагрузки справочников: {e}")
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
            forms = [
                (503317, "0503317", "Форма 0503317", "Квартальная/6М/9М/12М", 1),
            ]
            cursor.executemany(
                'INSERT INTO ref_form_types (id,code, name, periodicity, is_active) '
                'VALUES (?, ?, ?, ?, ?)',
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
        """Сохранение данных проекта в таблицы строк (без метаданных)"""
        # Удаляем старые данные для этой ревизии (если указана) или для всего проекта
        if revision_id:
            cursor.execute('DELETE FROM income_rows WHERE project_id=? AND revision_id=?', (project_id, revision_id))
            cursor.execute('DELETE FROM expense_rows WHERE project_id=? AND revision_id=?', (project_id, revision_id))
            cursor.execute('DELETE FROM source_rows WHERE project_id=? AND revision_id=?', (project_id, revision_id))
            cursor.execute('DELETE FROM consolidated_rows WHERE project_id=? AND revision_id=?', (project_id, revision_id))
        else:
            cursor.execute('DELETE FROM income_rows WHERE project_id=?', (project_id,))
            cursor.execute('DELETE FROM expense_rows WHERE project_id=?', (project_id,))
            cursor.execute('DELETE FROM source_rows WHERE project_id=?', (project_id,))
            cursor.execute('DELETE FROM consolidated_rows WHERE project_id=?', (project_id,))
        
        # Сохраняем данные по разделам (исключаем метаданные)
        section_tables = {
            'доходы_data': 'income_rows',
            'расходы_data': 'expense_rows',
            'источники_финансирования_data': 'source_rows',
            'консолидируемые_расчеты_data': 'consolidated_rows'
        }
        
        # Оптимизация: используем batch INSERT вместо множественных INSERT в цикле
        for section_key, table_name in section_tables.items():
            if section_key in data and data[section_key]:
                rows = data[section_key]
                if isinstance(rows, list):
                    # Подготавливаем данные для batch insert
                    if revision_id:
                        batch_data = [
                            (project_id, revision_id, json.dumps(row, ensure_ascii=False, default=str))
                            for row in rows
                        ]
                        cursor.executemany(
                            f'INSERT INTO {table_name} (project_id, revision_id, row_data) VALUES (?, ?, ?)',
                            batch_data
                        )
                    else:
                        batch_data = [
                            (project_id, json.dumps(row, ensure_ascii=False, default=str))
                            for row in rows
                        ]
                        cursor.executemany(
                            f'INSERT INTO {table_name} (project_id, row_data) VALUES (?, ?)',
                            batch_data
                        )
    
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
        """Загрузка данных проекта из таблиц строк (без метаданных)"""
        data = {}
        
        section_tables = {
            'доходы_data': 'income_rows',
            'расходы_data': 'expense_rows',
            'источники_финансирования_data': 'source_rows',
            'консолидируемые_расчеты_data': 'consolidated_rows'
        }
        
        for section_key, table_name in section_tables.items():
            if revision_id:
                cursor.execute(
                    f'SELECT row_data FROM {table_name} WHERE project_id=? AND revision_id=? ORDER BY id',
                    (project_id, revision_id)
                )
            else:
                cursor.execute(
                    f'SELECT row_data FROM {table_name} WHERE project_id=? AND (revision_id IS NULL OR revision_id=?) ORDER BY id',
                    (project_id, revision_id or 0)
                )
            rows = []
            for row in cursor.fetchall():
                try:
                    row_data = json.loads(row[0])
                    rows.append(row_data)
                except (json.JSONDecodeError, TypeError) as e:
                    print(f"Ошибка загрузки строки из {table_name}: {e}")
                    continue
            if rows:
                data[section_key] = rows
        
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
                'SELECT id, code, name, periodicity, is_active FROM ref_form_types WHERE code=?',
                (code,)
            )
            row = cursor.fetchone()
            if not row:
                return None
            return FormTypeMeta.from_row(
                {'id': row[0], 'code': row[1], 'name': row[2],
                 'periodicity': row[3], 'is_active': row[4]}
            )

    def load_form_types_meta(self) -> List[FormTypeMeta]:
        result: List[FormTypeMeta] = []
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute('SELECT id, code, name, periodicity, is_active FROM ref_form_types ORDER BY code')
            for row in cursor.fetchall():
                result.append(
                    FormTypeMeta.from_row(
                        {'id': row[0], 'code': row[1], 'name': row[2],
                         'periodicity': row[3], 'is_active': row[4]}
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
                            print(f"Невозможно определить ID для типа формы с кодом '{f.code}', запись пропущена")
                            continue

                    cursor.execute(
                        'INSERT INTO ref_form_types (id, code, name, periodicity, is_active) '
                        'VALUES (?, ?, ?, ?, ?)',
                        (
                            form_id,
                            f.code,
                            f.name,
                            f.periodicity or None,
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
                        {'id': row[0], 'code': row[1], 'name': row[2],
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
        """Удаление одной ревизии формы (без затрагивания проекта и других ревизий)."""
        with sqlite3.connect(self.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute('DELETE FROM form_revisions WHERE id=?', (revision_id,))
            conn.commit()
    
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
                        print(f"Ошибка загрузки meta_info для ревизии {revision_id}: {e}")
                
                if meta_row[1]:  # результат_исполнения_data
                    try:
                        result['результат_исполнения_data'] = json.loads(meta_row[1])
                    except (json.JSONDecodeError, TypeError) as e:
                        print(f"Ошибка загрузки результат_исполнения_data для ревизии {revision_id}: {e}")
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
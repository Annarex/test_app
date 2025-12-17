"""
Диалог для управления справочниками (загрузка из Excel, просмотр, редактирование)
"""
from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QTabWidget, QWidget,
    QTableWidget, QTableWidgetItem, QPushButton, QFileDialog,
    QMessageBox, QLabel, QHeaderView, QLineEdit, QComboBox
)
from PyQt5.QtCore import Qt
from pathlib import Path
import pandas as pd
from logger import logger

from models.database import DatabaseManager


class ReferencesManagementDialog(QDialog):
    """Диалог для управления справочниками"""
    
    # Список справочников с их методами загрузки
    REFERENCE_TYPES = {
        'Коды доходов': {
            'table': 'income_reference_records',
            'load_method': None,  # Загружается через ReferenceController.load_reference_file()
            'columns': ['code', 'name', 'level', 'doc'],
            'display_columns': ['code AS код', 'name AS наименование', 'level AS уровень', 'doc AS документ']
        },
        'Коды расходов': {
            'table': 'ref_expense_codes',  # Используем существующую таблицу, если есть
            'load_method': None,  # Загружается через существующие методы или через ReferenceController
            'columns': ['код', 'название', 'уровень', 'код_Р', 'код_ПР', 'код_ЦС', 'код_ВР']
        },
        'ГРБС': {
            'table': 'ref_grbs',
            'load_method': 'load_grbs_from_excel',
            'columns': ['код_ГРБС', 'наименование']
        },
        'Разделы/подразделы расходов': {
            'table': 'ref_expense_sections',
            'load_method': 'load_expense_sections_from_excel',
            'columns': ['код_РП', 'наименование', 'утверждающий_документ']
        },
        'Целевые статьи расходов': {
            'table': 'ref_target_articles',
            'load_method': 'load_target_articles_from_excel',
            'columns': ['код_ЦСР', 'наименование']
        },
        'Виды статей расходов': {
            'table': 'ref_expense_types',
            'load_method': 'load_expense_types_from_excel',
            'columns': ['код_вида_СР', 'наименование']
        },
        'Программные/непрограммные статьи': {
            'table': 'ref_program_nonprogram',
            'load_method': 'load_program_nonprogram_from_excel',
            'columns': ['код_ПНС', 'наименование']
        },
        'Виды расходов': {
            'table': 'ref_expense_kinds',
            'load_method': 'load_expense_kinds_from_excel',
            'columns': ['код_ВР', 'наименование', 'утверждающий_документ']
        },
        'Национальные проекты': {
            'table': 'ref_national_projects',
            'load_method': 'load_national_projects_from_excel',
            'columns': ['код_НПЦСР', 'наименование', 'утверждающий_документ']
        },
        'ГАДБ': {
            'table': 'ref_gadb',
            'load_method': 'load_gadb_from_excel',
            'columns': ['код_ГАДБ', 'наименование']
        },
        'Группы доходов': {
            'table': 'ref_income_groups',
            'load_method': 'load_income_groups_from_excel',
            'columns': ['код_группы_ДБ', 'наименование']
        },
        'Подгруппы доходов': {
            'table': 'ref_income_subgroups',
            'load_method': 'load_income_subgroups_from_excel',
            'columns': ['код_подгруппы_ДБ', 'наименование']
        },
        'Статьи/подстатьи доходов': {
            'table': 'ref_income_articles',
            'load_method': 'load_income_articles_from_excel',
            'columns': ['код_статьи_подстатьи_ДБ', 'наименование']
        },
        'Элементы доходов': {
            'table': 'ref_income_elements',
            'load_method': 'load_income_elements_from_excel',
            'columns': ['код_элемента_ДБ', 'наименование']
        },
        'Группы подвидов доходов': {
            'table': 'ref_income_subkind_groups',
            'load_method': 'load_income_subkind_groups_from_excel',
            'columns': ['код_группы_ПДБ', 'наименование']
        },
        'Аналитические группы подвидов доходов': {
            'table': 'ref_income_analytical_groups',
            'load_method': 'load_income_analytical_groups_from_excel',
            'columns': ['код_группы_АПДБ', 'наименование']
        },
        'Уровни доходов': {
            'table': 'ref_income_levels',
            'load_method': 'load_income_levels_from_excel',
            'columns': ['код_уровня', 'наименование', 'цвет']
        }
    }
    
    def __init__(self, db_manager: DatabaseManager, parent=None):
        super().__init__(parent)
        self.db_manager = db_manager
        self.setWindowTitle("Управление справочниками")
        self.setMinimumSize(900, 600)
        
        self.current_reference_type = None
        self.current_table = None
        
        self.init_ui()
        self.load_reference_list()
    
    def init_ui(self):
        """Инициализация UI"""
        layout = QVBoxLayout(self)
        
        # Выбор справочника
        selection_layout = QHBoxLayout()
        selection_layout.addWidget(QLabel("Справочник:"))
        
        self.reference_combo = QComboBox()
        self.reference_combo.currentTextChanged.connect(self.on_reference_changed)
        selection_layout.addWidget(self.reference_combo)
        
        layout.addLayout(selection_layout)
        
        # Вкладки
        self.tabs = QTabWidget()
        
        # Вкладка просмотра
        self.view_tab = QWidget()
        self.view_table = QTableWidget()
        self.setup_view_tab()
        self.tabs.addTab(self.view_tab, "Просмотр")
        
        # Вкладка загрузки
        self.load_tab = QWidget()
        self.setup_load_tab()
        self.tabs.addTab(self.load_tab, "Загрузка из Excel")
        
        layout.addWidget(self.tabs)
        
        # Кнопки
        buttons_layout = QHBoxLayout()
        buttons_layout.addStretch()
        
        close_btn = QPushButton("Закрыть")
        close_btn.clicked.connect(self.accept)
        buttons_layout.addWidget(close_btn)
        
        layout.addLayout(buttons_layout)
    
    def setup_view_tab(self):
        """Настройка вкладки просмотра"""
        layout = QVBoxLayout(self.view_tab)
        
        # Кнопки управления
        buttons_layout = QHBoxLayout()
        
        refresh_btn = QPushButton("Обновить")
        refresh_btn.clicked.connect(self.load_current_reference)
        buttons_layout.addWidget(refresh_btn)
        
        buttons_layout.addStretch()
        layout.addLayout(buttons_layout)
        
        # Таблица
        self.view_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.view_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.view_table.horizontalHeader().setStretchLastSection(True)
        layout.addWidget(self.view_table)
        
        # Статус
        self.status_label = QLabel("Выберите справочник для просмотра")
        layout.addWidget(self.status_label)
    
    def setup_load_tab(self):
        """Настройка вкладки загрузки"""
        layout = QVBoxLayout(self.load_tab)
        
        # Выбор файла
        file_layout = QHBoxLayout()
        file_layout.addWidget(QLabel("Файл Excel:"))
        
        self.file_path_edit = QLineEdit()
        self.file_path_edit.setReadOnly(True)
        file_layout.addWidget(self.file_path_edit)
        
        browse_btn = QPushButton("Обзор...")
        browse_btn.clicked.connect(self.browse_excel_file)
        file_layout.addWidget(browse_btn)
        
        layout.addLayout(file_layout)
        
        # Кнопка загрузки
        self.load_btn = QPushButton("Загрузить справочник")
        self.load_btn.clicked.connect(self.load_reference_from_excel)
        self.load_btn.setEnabled(False)
        layout.addWidget(self.load_btn)
        
        # Статус загрузки
        self.load_status_label = QLabel("Выберите файл Excel для загрузки")
        layout.addWidget(self.load_status_label)
        
        layout.addStretch()
    
    def load_reference_list(self):
        """Загрузка списка справочников"""
        self.reference_combo.clear()
        for ref_name in self.REFERENCE_TYPES.keys():
            self.reference_combo.addItem(ref_name)
    
    def on_reference_changed(self, reference_name: str):
        """Обработчик изменения выбранного справочника"""
        if reference_name:
            self.current_reference_type = self.REFERENCE_TYPES.get(reference_name)
            self.current_table = reference_name
            self.load_current_reference()
    
    def load_current_reference(self):
        """Загрузка текущего справочника в таблицу"""
        if not self.current_reference_type:
            return
        
        try:
            table_name = self.current_reference_type['table']
            columns = self.current_reference_type.get('columns', [])
            display_columns = self.current_reference_type.get('display_columns', [])
            
            # Загружаем данные из БД
            import sqlite3
            conn = sqlite3.connect(self.db_manager.db_path)
            
            # Проверяем существование таблицы
            cursor = conn.cursor()
            cursor.execute(
                "SELECT name FROM sqlite_master WHERE type='table' AND name=?",
                (table_name,)
            )
            if not cursor.fetchone():
                self.status_label.setText(f"Таблица {table_name} не найдена в БД")
                conn.close()
                return
            
            # Для справочника доходов используем специальный запрос с алиасами
            if table_name == 'income_reference_records' and display_columns:
                query = f'SELECT {", ".join(display_columns)} FROM {table_name}'
                df = pd.read_sql_query(query, conn)
                # Используем реальные названия колонок из результата запроса (с алиасами)
                available_columns = list(df.columns)
            else:
                # Проверяем существование колонок
                cursor.execute(f"PRAGMA table_info({table_name})")
                existing_columns = [row[1] for row in cursor.fetchall()]
                available_columns = [col for col in columns if col in existing_columns]
                
                if not available_columns:
                    self.status_label.setText(f"Нет доступных колонок в таблице {table_name}")
                    conn.close()
                    return
                
                query = f'SELECT {", ".join(available_columns)} FROM {table_name}'
                df = pd.read_sql_query(query, conn)
            
            conn.close()
            
            # Заполняем таблицу
            self.view_table.setRowCount(len(df))
            self.view_table.setColumnCount(len(available_columns))
            self.view_table.setHorizontalHeaderLabels(available_columns)
            
            for row_idx, (_, row) in enumerate(df.iterrows()):
                for col_idx, col_name in enumerate(available_columns):
                    value = row.get(col_name, '')
                    item = QTableWidgetItem(str(value) if pd.notna(value) else '')
                    self.view_table.setItem(row_idx, col_idx, item)
            
            self.status_label.setText(f"Загружено записей: {len(df)}")
            
        except Exception as e:
            logger.error(f"Ошибка загрузки справочника: {e}", exc_info=True)
            QMessageBox.warning(self, "Ошибка", f"Не удалось загрузить справочник:\n{str(e)}")
            self.status_label.setText("Ошибка загрузки")
    
    def browse_excel_file(self):
        """Открыть диалог выбора Excel файла"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Выберите файл Excel",
            "",
            "Excel files (*.xlsx *.xls);;All files (*.*)"
        )
        
        if file_path:
            self.file_path_edit.setText(file_path)
            self.load_btn.setEnabled(True)
            self.load_status_label.setText(f"Выбран файл: {Path(file_path).name}")
    
    def load_reference_from_excel(self):
        """Загрузка справочника из Excel"""
        if not self.current_reference_type:
            QMessageBox.warning(self, "Ошибка", "Выберите справочник")
            return
        
        file_path = self.file_path_edit.text()
        if not file_path:
            QMessageBox.warning(self, "Ошибка", "Выберите файл Excel")
            return
        
        try:
            load_method_name = self.current_reference_type.get('load_method')
            
            # Для справочников доходов и расходов используем существующий механизм
            if self.current_reference_type['table'] == 'income_reference_records':
                QMessageBox.information(
                    self,
                    "Информация",
                    "Справочник доходов загружается через меню 'Справочники' → 'Загрузить справочник доходов...'\n"
                    "Используйте существующий функционал загрузки."
                )
                return
            
            if not load_method_name:
                QMessageBox.warning(
                    self,
                    "Ошибка",
                    "Для этого справочника загрузка из Excel не поддерживается.\n"
                    "Используйте существующие методы загрузки."
                )
                return
            
            load_method = getattr(self.db_manager, load_method_name)
            
            # Загружаем справочник
            self.load_btn.setEnabled(False)
            self.load_status_label.setText("Загрузка...")
            
            # Остальные методы возвращают количество записей
            count = load_method(file_path)
            
            self.load_status_label.setText(f"Загружено записей: {count}")
            QMessageBox.information(
                self,
                "Успех",
                f"Справочник успешно загружен.\nЗагружено записей: {count}"
            )
            
            # Обновляем таблицу просмотра
            self.load_current_reference()
            
        except Exception as e:
            logger.error(f"Ошибка загрузки справочника из Excel: {e}", exc_info=True)
            QMessageBox.critical(
                self,
                "Ошибка",
                f"Ошибка загрузки справочника:\n{str(e)}"
            )
            self.load_status_label.setText("Ошибка загрузки")
        finally:
            self.load_btn.setEnabled(True)

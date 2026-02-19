"""
Универсальный диалог для проверки текстов из Excel файла
Позволяет загружать данные из Excel, выбирать справочник и проверять на ошибки
"""
from PyQt5.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QPushButton,
                             QLabel, QTableWidget, QTableWidgetItem, QHeaderView,
                             QComboBox, QMessageBox, QFileDialog, QGroupBox,
                             QLineEdit, QDateEdit, QSpinBox, QCheckBox,
                             QProgressDialog, QTextBrowser, QSizePolicy, QWidget,
                             QSplitter, QTabWidget, QListWidget, QListWidgetItem)
from PyQt5.QtCore import Qt, QDate, QTimer
from PyQt5.QtGui import QColor
from datetime import datetime
import sqlite3
import pandas as pd
from logger import logger
from models.database import DatabaseManager
from utils.db_utils import get_filtered_view
from utils.text_validation.error_finder import find_errors
from utils.text_validation.text_comparator import find_differences
from utils.text_validation.display_helpers import richtext_to_html, find_alternative_variants


class ExcelPreviewDialog(QDialog):
    """Диалог предпросмотра данных Excel"""
    
    def __init__(self, excel_data, header_row, data_start_row, sheet_name=None, parent=None):
        super().__init__(parent)
        self.excel_data = excel_data
        self.header_row = header_row
        self.data_start_row = data_start_row
        self.sheet_name = sheet_name
        
        title = "Предпросмотр данных Excel"
        if sheet_name:
            title += f" [Лист: {sheet_name}]"
        self.setWindowTitle(title)
        self.resize(1200, 700)
        self.init_ui()
        self.load_data()
    
    def init_ui(self):
        layout = QVBoxLayout(self)
        
        # Информация
        info_parts = [
            f"Всего строк: <b>{len(self.excel_data)}</b>",
            f"Столбцов: <b>{len(self.excel_data.columns)}</b>",
        ]
        if self.sheet_name:
            info_parts.append(f"Лист: <b>{self.sheet_name}</b>")
        info_parts.extend([
            f"Заголовки: строка <b>{self.header_row}</b>",
            f"Данные: с строки <b>{self.data_start_row}</b>"
        ])
        info_text = " | ".join(info_parts)
        info_label = QLabel(info_text)
        info_label.setStyleSheet("padding: 5px; background-color: #f0f0f0; border-radius: 3px;")
        layout.addWidget(info_label)
        
        # Таблица
        self.table = QTableWidget()
        self.table.setAlternatingRowColors(True)
        self.table.horizontalHeader().setStretchLastSection(True)
        layout.addWidget(self.table)
        
        # Подсказка
        hint = QLabel("💡 Коды с ведущими нулями (007, 0123) подсвечиваются желтым")
        hint.setStyleSheet("color: #666; font-size: 8pt; font-style: italic;")
        layout.addWidget(hint)
        
        # Кнопка закрытия
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        close_btn = QPushButton("Закрыть")
        close_btn.clicked.connect(self.accept)
        button_layout.addWidget(close_btn)
        layout.addLayout(button_layout)
    
    def load_data(self):
        """Загрузка данных в таблицу"""
        df = self.excel_data
        rows_to_show = min(len(df), 1000)  # Показываем до 1000 строк
        
        # Добавляем столбец с номерами строк
        self.table.setRowCount(rows_to_show)
        self.table.setColumnCount(len(df.columns) + 1)
        
        # Заголовки (конвертируем в строки для PyQt)
        headers = ["№ строки"] + [str(col) for col in df.columns.tolist()]
        self.table.setHorizontalHeaderLabels(headers)
        
        for row_idx in range(rows_to_show):
            # Номер строки в файле
            actual_row = self.data_start_row + row_idx
            row_num_item = QTableWidgetItem(str(actual_row))
            row_num_item.setFlags(row_num_item.flags() & ~Qt.ItemIsEditable)
            row_num_item.setBackground(QColor("#f0f0f0"))
            self.table.setItem(row_idx, 0, row_num_item)
            
            # Данные
            for col_idx, col_name in enumerate(df.columns):
                value = df.iloc[row_idx][col_name]
                display_value = str(value) if value and str(value).strip() else ""
                
                item = QTableWidgetItem(display_value)
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                
                # Подсветка кодов с ведущими нулями
                if display_value and len(display_value) >= 2:
                    if display_value[0] == '0' and any(c.isdigit() for c in display_value[1:]):
                        is_decimal = (
                            len(display_value) >= 3 and 
                            display_value[1] in ',.;' and 
                            display_value[2:].replace(',', '').replace('.', '').replace(';', '').isdigit()
                        )
                        if not is_decimal:
                            item.setBackground(QColor("#fff3cd"))
                            item.setToolTip("Код с ведущим нулем (сохранен как текст)")
                
                self.table.setItem(row_idx, col_idx + 1, item)
        
        self.table.resizeColumnsToContents()
        self.table.setColumnWidth(0, 80)


class UniversalTextValidationDialog(QDialog):
    """Универсальный диалог для проверки текстов из Excel"""
    
    def __init__(self, db_path: str, parent=None):
        super().__init__(parent)
        self.db_path = db_path
        self.db_manager = DatabaseManager(db_path)
        self.excel_data = None
        self.excel_file_path = None
        self.excel_sheets = []
        self.current_sheet = None
        self.errors = []
        self.all_items = []
        self.reference_data = None
        self.reference_names = []
        self.reference_codes = []
        self.orig_codes = []
        self.reference_table_name = None
        self.reference_date = None
        self.reference_oktmo = None
        self.reference_names_map = {}  # Словарь {table_name: display_name}
        
        # Для расширенной проверки
        self.advanced_mode = False
        self.custom_name_column = 'name'
        self.custom_code_columns = []
        
        self.setWindowTitle("Проверка текстов из Excel")
        self.resize(1400, 900)
        self.setup_ui()
        self._load_oktmo_list()
        
    def setup_ui(self):
        """Создание интерфейса"""
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(12)
        
        # Панель настроек
        settings_group = QGroupBox("Параметры проверки")
        settings_layout = QVBoxLayout()
        settings_group.setLayout(settings_layout)
        
        # Excel файл
        excel_row = QHBoxLayout()
        excel_row.addWidget(QLabel("Excel файл:"))
        self.excel_path_label = QLabel("Файл не выбран")
        self.excel_path_label.setStyleSheet("color: #666; font-style: italic;")
        excel_row.addWidget(self.excel_path_label, 1)
        self.load_excel_btn = QPushButton("📂 Загрузить Excel")
        self.load_excel_btn.clicked.connect(self._load_excel)
        excel_row.addWidget(self.load_excel_btn)
        settings_layout.addLayout(excel_row)
        
        # Выбор листа Excel
        sheet_row = QHBoxLayout()
        sheet_row.addWidget(QLabel("Лист:"))
        self.sheet_combo = QComboBox()
        self.sheet_combo.setMinimumWidth(200)
        self.sheet_combo.setEnabled(False)
        self.sheet_combo.currentTextChanged.connect(self._on_sheet_changed)
        sheet_row.addWidget(self.sheet_combo)
        sheet_row.addStretch()
        settings_layout.addLayout(sheet_row)
        
        # Параметры загрузки Excel
        excel_params_row = QHBoxLayout()
        excel_params_row.addWidget(QLabel("Строка заголовков:"))
        self.header_row_spinbox = QSpinBox()
        self.header_row_spinbox.setRange(1, 100)
        self.header_row_spinbox.setValue(1)
        self.header_row_spinbox.setMaximumWidth(80)
        self.header_row_spinbox.setToolTip("Номер строки в файле, где находятся названия столбцов")
        self.header_row_spinbox.valueChanged.connect(self._on_header_row_changed)
        excel_params_row.addWidget(self.header_row_spinbox)
        
        excel_params_row.addSpacing(20)
        excel_params_row.addWidget(QLabel("Начало данных:"))
        self.data_start_row_spinbox = QSpinBox()
        self.data_start_row_spinbox.setRange(1, 100)
        self.data_start_row_spinbox.setValue(2)
        self.data_start_row_spinbox.setMaximumWidth(80)
        self.data_start_row_spinbox.setToolTip("Номер строки в файле, где начинаются данные")
        excel_params_row.addWidget(self.data_start_row_spinbox)
        
        self.reload_preview_btn = QPushButton("🔄 Обновить")
        self.reload_preview_btn.clicked.connect(self._reload_excel_preview)
        self.reload_preview_btn.setEnabled(False)
        excel_params_row.addWidget(self.reload_preview_btn)
        
        self.auto_detect_btn = QPushButton("🔍 Авто")
        self.auto_detect_btn.setToolTip("Автоматически определить начало таблицы")
        self.auto_detect_btn.clicked.connect(self._auto_detect_table_start)
        self.auto_detect_btn.setEnabled(False)
        excel_params_row.addWidget(self.auto_detect_btn)
        
        excel_params_row.addSpacing(20)
        self.preview_btn = QPushButton("📄 Просмотр")
        self.preview_btn.setToolTip("Предпросмотр данных из Excel")
        self.preview_btn.clicked.connect(self._show_excel_preview)
        self.preview_btn.setEnabled(False)
        excel_params_row.addWidget(self.preview_btn)
        
        excel_params_row.addStretch()
        settings_layout.addLayout(excel_params_row)
        
        # Столбцы Excel
        columns_row = QHBoxLayout()
        columns_row.addWidget(QLabel("Столбец с текстом:"))
        self.text_column_combo = QComboBox()
        self.text_column_combo.setMinimumWidth(150)
        columns_row.addWidget(self.text_column_combo)
        
        columns_row.addSpacing(20)
        columns_row.addWidget(QLabel("Столбец с кодом:"))
        self.code_column_combo = QComboBox()
        self.code_column_combo.setMinimumWidth(150)
        self.code_column_combo.addItem("-- Без проверки кода --", None)
        columns_row.addWidget(self.code_column_combo)
        columns_row.addStretch()
        settings_layout.addLayout(columns_row)
        
        # Справочник
        reference_row = QHBoxLayout()
        reference_row.addWidget(QLabel("Справочник:"))
        self.reference_combo = QComboBox()
        self.reference_combo.addItem("Классификация доходов", "v_budgetclastypeinc_merged")
        self.reference_combo.addItem("Классификация расходов", "v_budgetclascosts_merged")
        self.reference_combo.addItem("Источники финансирования", "v_budgetclassources_merged")
        reference_row.addWidget(self.reference_combo, 1)
        
        reference_row.addSpacing(20)
        reference_row.addWidget(QLabel("Дата:"))
        self.date_edit = QDateEdit()
        self.date_edit.setCalendarPopup(True)
        self.date_edit.setDate(QDate.currentDate())
        self.date_edit.setDisplayFormat("dd.MM.yyyy")
        self.date_edit.dateChanged.connect(self._on_date_changed)
        reference_row.addWidget(self.date_edit)
        
        reference_row.addSpacing(20)
        reference_row.addWidget(QLabel("ОКТМО:"))
        self.oktmo_combo = QComboBox()
        self.oktmo_combo.setMinimumWidth(250)
        reference_row.addWidget(self.oktmo_combo)
        settings_layout.addLayout(reference_row)
        
        # Расширенная проверка
        advanced_row = QHBoxLayout()
        self.advanced_checkbox = QCheckBox("⚙️ Расширенная проверка (любой справочник)")
        self.advanced_checkbox.setToolTip("Включить расширенный режим для проверки в произвольных справочниках")
        self.advanced_checkbox.stateChanged.connect(self._on_advanced_mode_changed)
        advanced_row.addWidget(self.advanced_checkbox)
        advanced_row.addStretch()
        settings_layout.addLayout(advanced_row)
        
        # Расширенные настройки справочника (скрыты по умолчанию)
        self.advanced_settings_widget = QWidget()
        advanced_settings_layout = QVBoxLayout(self.advanced_settings_widget)
        advanced_settings_layout.setContentsMargins(20, 0, 0, 0)
        
        # Строка 1: Таблица справочника
        adv_table_row = QHBoxLayout()
        adv_table_row.addWidget(QLabel("Таблица справочника:"))
        self.custom_table_combo = QComboBox()
        self.custom_table_combo.setMinimumWidth(300)
        self.custom_table_combo.setEditable(True)
        self.custom_table_combo.currentTextChanged.connect(self._on_custom_table_changed)
        adv_table_row.addWidget(self.custom_table_combo, 1)
        
        refresh_tables_btn = QPushButton("🔄")
        refresh_tables_btn.setToolTip("Обновить список таблиц из БД")
        refresh_tables_btn.setMaximumWidth(40)
        refresh_tables_btn.clicked.connect(self._load_database_tables)
        adv_table_row.addWidget(refresh_tables_btn)
        advanced_settings_layout.addLayout(adv_table_row)
        
        # Строка 2: Колонка с наименованием
        adv_name_col_row = QHBoxLayout()
        adv_name_col_row.addWidget(QLabel("Колонка с наименованием:"))
        self.custom_name_column_combo = QComboBox()
        self.custom_name_column_combo.setMinimumWidth(300)
        adv_name_col_row.addWidget(self.custom_name_column_combo, 1)
        advanced_settings_layout.addLayout(adv_name_col_row)
        
        # Строка 3: Колонки для кода (список с чекбоксами)
        adv_code_cols_layout = QVBoxLayout()
        adv_code_cols_layout.addWidget(QLabel("Колонки для кода (отметьте нужные, будут склеены через пробел):"))
        self.custom_code_columns_list = QListWidget()
        self.custom_code_columns_list.setMaximumHeight(120)
        self.custom_code_columns_list.setSelectionMode(QListWidget.NoSelection)  # Отключаем выделение
        adv_code_cols_layout.addWidget(self.custom_code_columns_list)
        advanced_settings_layout.addLayout(adv_code_cols_layout)
        
        settings_layout.addWidget(self.advanced_settings_widget)
        self.advanced_settings_widget.setVisible(False)
        
        # Параметры проверки
        params_row = QHBoxLayout()
        params_row.addWidget(QLabel("Макс. расстояние:"))
        self.distance_spinbox = QSpinBox()
        self.distance_spinbox.setRange(1, 301)
        self.distance_spinbox.setValue(200)
        self.distance_spinbox.setMaximumWidth(80)
        self.distance_spinbox.valueChanged.connect(self._apply_filters)
        params_row.addWidget(self.distance_spinbox)
        
        params_row.addSpacing(20)
        self.show_all_checkbox = QCheckBox("Показать все тексты")
        self.show_all_checkbox.stateChanged.connect(self._apply_filters)
        params_row.addWidget(self.show_all_checkbox)
        
        params_row.addStretch()
        self.check_btn = QPushButton("🔍 Проверить")
        self.check_btn.clicked.connect(self._run_validation)
        self.check_btn.setEnabled(False)
        params_row.addWidget(self.check_btn)
        settings_layout.addLayout(params_row)
        
        layout.addWidget(settings_group)
        
        # Результаты
        results_group = QGroupBox("Результаты проверки")
        results_layout = QVBoxLayout()
        results_group.setLayout(results_layout)
        
        # Информация о загруженных данных и статистика
        info_stats_row = QHBoxLayout()
        self.excel_info_label = QLabel("Данные не загружены")
        self.excel_info_label.setStyleSheet("color: #666; font-size: 9pt;")
        info_stats_row.addWidget(self.excel_info_label)
        info_stats_row.addStretch()
        self.stats_label = QLabel("")
        self.stats_label.setStyleSheet("font-size: 9pt; font-weight: bold;")
        info_stats_row.addWidget(self.stats_label)
        results_layout.addLayout(info_stats_row)
        
        self.errors_table = QTableWidget()
        self.errors_table.setAlternatingRowColors(True)
        self.errors_table.setColumnCount(5)
        self.errors_table.setHorizontalHeaderLabels([
            "Текст в файле",
            "Эталон из справочника",
            "Distance",
            "Код из файла",
            "Код справочника"
        ])
        
        # Настройка таблицы
        header = self.errors_table.horizontalHeader()
        header.setStretchLastSection(False)
        for i in range(5):
            header.setSectionResizeMode(i, QHeaderView.Interactive)
        
        header.resizeSection(0, 350)  # Текст в проекте
        header.resizeSection(1, 350)  # Эталон
        header.resizeSection(2, 80)   # Distance
        header.resizeSection(3, 200)  # Код в проекте
        header.resizeSection(4, 200)  # Код справочника
        
        self.errors_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.errors_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.errors_table.setWordWrap(True)
        self.errors_table.cellDoubleClicked.connect(self._show_error_details)
        
        # Устанавливаем делегат для отображения HTML во всех текстовых столбцах
        from views.text_validation_widget import HtmlDelegate
        html_delegate = HtmlDelegate(self.errors_table)
        self.errors_table.setItemDelegateForColumn(0, html_delegate)  # Текст в файле
        self.errors_table.setItemDelegateForColumn(1, html_delegate)  # Эталон из справочника
        self.errors_table.setItemDelegateForColumn(3, html_delegate)  # Код из файла
        self.errors_table.setItemDelegateForColumn(4, html_delegate)  # Код справочника
        
        results_layout.addWidget(self.errors_table)
        
        layout.addWidget(results_group)
        
        # Кнопки
        buttons_row = QHBoxLayout()
        buttons_row.addStretch()
        
        self.export_btn = QPushButton("💾 Экспорт в Excel")
        self.export_btn.clicked.connect(self._export_results)
        self.export_btn.setEnabled(False)
        buttons_row.addWidget(self.export_btn)
        
        close_btn = QPushButton("Закрыть")
        close_btn.clicked.connect(self.accept)
        buttons_row.addWidget(close_btn)
        
        layout.addLayout(buttons_row)
        
        # Загружаем названия справочников перед загрузкой таблиц
        self._load_reference_names()  # Загружаем названия справочников
        self._load_database_tables()  # Используют reference_names_map
        
    def _load_reference_names(self):
        """Загрузка названий справочников из единой конфигурации"""
        # Импортируем конфигурацию справочников
        try:
            from models.references.references_config import REFERENCE_TYPES
            self.reference_names_map = {}
            
            for name, config in REFERENCE_TYPES.items():
                if config.get('table') and not config.get('is_separator'):
                    self.reference_names_map[config['table']] = name
            
            logger.info(f"Загружено {len(self.reference_names_map)} названий справочников из конфигурации")
        except Exception as e:
            logger.error(f"Ошибка загрузки названий справочников: {e}", exc_info=True)
            self.reference_names_map = {}
    
    def _on_advanced_mode_changed(self, state):
        """Переключение расширенного режима"""
        self.advanced_mode = (state == Qt.Checked)
        
        # Показываем/скрываем элементы
        self.advanced_settings_widget.setVisible(self.advanced_mode)
        
        # Скрываем/показываем стандартный комбобокс справочника
        self.reference_combo.setEnabled(not self.advanced_mode)
        
        if self.advanced_mode:
            self.reference_combo.setStyleSheet("QComboBox { color: gray; }")
        else:
            self.reference_combo.setStyleSheet("")
            
        logger.info(f"Расширенный режим: {'включен' if self.advanced_mode else 'выключен'}")
    
    def _load_database_tables(self):
        """Загрузка списка таблиц/представлений из БД"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # Получаем список таблиц и представлений
            cursor.execute("""
                SELECT name FROM sqlite_master 
                WHERE type IN ('table', 'view') 
                AND name NOT LIKE 'sqlite_%'AND name NOT LIKE 'app_config'
                ORDER BY name
            """)
            
            tables = [row[0] for row in cursor.fetchall()]
            conn.close()
            
            self.custom_table_combo.clear()
            
            # Добавляем таблицы с их русскими названиями
            for table in tables:
                # Получаем русское название из маппинга, если есть
                display_name = self.reference_names_map.get(table, table)
                # Формат: "Русское название (table_name)"
                display_text = f"{display_name} ({table})" if display_name != table else table
                self.custom_table_combo.addItem(display_text, table)            
            logger.info(f"Загружено таблиц из БД: {len(tables)}")
            
        except Exception as e:
            logger.error(f"Ошибка загрузки списка таблиц: {e}", exc_info=True)
            QMessageBox.warning(self, "Ошибка", f"Не удалось загрузить список таблиц:\n{str(e)}")
    
    def _on_custom_table_changed(self, table_name: str):
        """Обработчик изменения выбранной таблицы - загружает список колонок"""
        if not table_name or not table_name.strip():
            return
        
        # Извлекаем реальное имя таблицы из userData
        table_name = self.custom_table_combo.currentData()
        if not table_name:
            return
            
        try:
            conn = sqlite3.connect(self.db_path)
            
            # Получаем список колонок из таблицы
            query = f"SELECT * FROM {table_name} LIMIT 0"
            cursor = conn.cursor()
            cursor.execute(query)
            
            columns = [description[0] for description in cursor.description]
            conn.close()
            
            # Загружаем маппинг колонок из единой конфигурации
            try:
                from models.references.references_config import get_russian_name
                get_russian_name_func = get_russian_name
            except Exception as e:
                logger.warning(f"Не удалось загрузить маппинг колонок из конфигурации: {e}")
                # Если не удалось загрузить, возвращаем колонку как есть
                get_russian_name_func = lambda table, col: col
            
            def get_column_label(col):
                """Получить русское название колонки из конфигурации"""
                russian = get_russian_name_func(table_name, col)
                if russian and russian != col:
                    return f"{russian} ({col})"
                return col
            
            # Заполняем комбобокс для колонки наименования
            self.custom_name_column_combo.clear()
            for col in columns:
                display_text = get_column_label(col)
                self.custom_name_column_combo.addItem(display_text, col)
            
            # Находим и выбираем 'name' если есть
            if 'name' in columns:
                idx = columns.index('name')
                self.custom_name_column_combo.setCurrentIndex(idx)
            
            # Заполняем список с чекбоксами для колонок кода
            self.custom_code_columns_list.clear()
            
            # Список потенциальных колонок кода (для автоматической отметки)
            code_keywords = ['code']
            
            for col in columns:
                display_text = get_column_label(col)
                item = QListWidgetItem(display_text)
                item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
                item.setCheckState(Qt.Unchecked)
                item.setData(Qt.UserRole, col)  # Сохраняем оригинальное имя колонки
                
                # Автоматически отмечаем колонки с "code" в названии
                col_lower = col.lower()
                if any(keyword in col_lower for keyword in code_keywords):
                    item.setCheckState(Qt.Checked)
                
                self.custom_code_columns_list.addItem(item)
            
            logger.info(f"Загружены колонки таблицы {table_name}: {len(columns)} шт.")
            
        except Exception as e:
            logger.error(f"Ошибка загрузки колонок таблицы {table_name}: {e}", exc_info=True)
            QMessageBox.warning(self, "Ошибка", f"Не удалось загрузить колонки таблицы:\n{str(e)}")
    
    def _load_excel(self):
        """Загрузка Excel файла"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Выберите Excel файл", "", "Excel Files (*.xlsx *.xls)"
        )
        
        if not file_path:
            return
        
        self.excel_file_path = file_path
        
        # Загружаем список листов
        self._load_sheet_list()
        
        self.header_row_spinbox.setValue(1)
        self.data_start_row_spinbox.setValue(2)
        
        # Автоматически определяем начало таблицы
        self._auto_detect_table_start()
        
        self._reload_excel_preview()
    
    def _load_sheet_list(self):
        """Загрузка списка листов из Excel файла"""
        if not self.excel_file_path:
            return
        
        try:
            # Получаем список листов
            excel_file = pd.ExcelFile(self.excel_file_path)
            self.excel_sheets = excel_file.sheet_names
            
            # Заполняем комбобокс
            self.sheet_combo.clear()
            self.sheet_combo.addItems(self.excel_sheets)
            self.sheet_combo.setEnabled(True)
            
            # По умолчанию выбираем первый лист
            if self.excel_sheets:
                self.current_sheet = self.excel_sheets[0]
                self.sheet_combo.setCurrentIndex(0)
            
            logger.info(f"Найдено листов в Excel: {len(self.excel_sheets)}")
            
        except Exception as e:
            logger.error(f"Ошибка загрузки списка листов: {e}", exc_info=True)
            QMessageBox.warning(self, "Ошибка", f"Не удалось получить список листов:\n{str(e)}")
    
    def _on_sheet_changed(self, sheet_name):
        """Обработчик смены листа"""
        if not sheet_name or sheet_name == self.current_sheet:
            return
        
        self.current_sheet = sheet_name
        logger.info(f"Выбран лист: {sheet_name}")
        
        # Перезагружаем данные из нового листа
        if self.excel_file_path:
            self._reload_excel_preview()
            
            # Обновляем метку с именем файла и листом
            filename = self.excel_file_path.split('\\')[-1]
            self.excel_path_label.setText(f"{filename} [{sheet_name}]")
    
    def _on_header_row_changed(self, value):
        """При изменении строки заголовков - обновить минимум для начала данных"""
        # Данные не могут начинаться раньше заголовков
        if self.data_start_row_spinbox.value() < value:
            self.data_start_row_spinbox.setValue(value)
        self.data_start_row_spinbox.setMinimum(value)
    
    def _auto_detect_table_start(self):
        """Автоматическое определение начала таблицы"""
        if not self.excel_file_path:
            return
        
        try:
            # Читаем первые 30 строк файла для анализа
            df_sample = pd.read_excel(
                self.excel_file_path,
                sheet_name=self.current_sheet if self.current_sheet else 0,
                header=None,
                nrows=30
            )
            
            if df_sample.empty:
                return
            
            # Ищем строку с максимальным количеством непустых значений
            # Это вероятно заголовки
            max_filled = 0
            header_row_candidate = 1
            
            for idx, row in df_sample.iterrows():
                # Количество непустых ячеек
                filled_count = row.notna().sum()
                
                # Если больше половины столбцов заполнено и это максимум
                if filled_count > max_filled and filled_count >= len(df_sample.columns) * 0.5:
                    max_filled = filled_count
                    header_row_candidate = idx + 1  # +1 т.к. индексы с 0
            
            # Находим первую строку с данными после заголовков
            data_start_candidate = header_row_candidate + 1
            
            # Пропускаем пустые строки
            for idx in range(header_row_candidate, min(len(df_sample), header_row_candidate + 10)):
                row = df_sample.iloc[idx]
                filled_count = row.notna().sum()
                
                # Если строка имеет данные (хотя бы 20% заполнено)
                if filled_count >= len(df_sample.columns) * 0.2:
                    data_start_candidate = idx + 1
                    break
            
            # Устанавливаем значения
            self.header_row_spinbox.setValue(header_row_candidate)
            self.data_start_row_spinbox.setValue(data_start_candidate)
            
            self.auto_detect_btn.setEnabled(True)
            
            logger.info(f"Автоопределение: заголовки=строка {header_row_candidate}, данные=строка {data_start_candidate}")
            
        except Exception as e:
            logger.error(f"Ошибка автоопределения: {e}", exc_info=True)
            # В случае ошибки оставляем значения по умолчанию
            pass
    
    def _reload_excel_preview(self):
        """Перезагрузка Excel с учетом параметров"""
        if not self.excel_file_path:
            return
            
        try:
            header_row = self.header_row_spinbox.value()
            data_start_row = self.data_start_row_spinbox.value()
            
            # Пропускаем строки до заголовков
            skiprows = header_row - 1 if header_row > 1 else None
            
            # Читаем Excel как текст для сохранения ведущих нулей в кодах
            # dtype=str сохраняет "007", "0123", "01.05" как текст
            df = pd.read_excel(
                self.excel_file_path,
                sheet_name=self.current_sheet if self.current_sheet else 0,
                skiprows=skiprows, 
                header=0,
                dtype=str,
                keep_default_na=False
            )
            
            # Если данные начинаются не сразу после заголовков, удаляем промежуточные строки
            rows_between = data_start_row - header_row - 1
            if rows_between > 0:
                df = df.iloc[rows_between:].reset_index(drop=True)
            
            if df.empty:
                QMessageBox.warning(self, "Ошибка", "Excel файл пустой или все строки пропущены")
                return
                
            self.excel_data = df
            # Формируем текст метки с именем файла и листом
            filename = self.excel_file_path.split('\\')[-1]
            if self.current_sheet:
                label_text = f"{filename} [{self.current_sheet}]"
            else:
                label_text = filename
            self.excel_path_label.setText(label_text)
            self.excel_path_label.setStyleSheet("color: #28a745; font-weight: bold;")
            
            # Заполняем комбобоксы (сохраняем оригинальные имена столбцов в userData)
            self.text_column_combo.clear()
            for col in df.columns:
                self.text_column_combo.addItem(str(col), col)  # Текст для отображения, оригинальное имя в data
            
            self.code_column_combo.clear()
            self.code_column_combo.addItem("-- Без проверки кода --", None)
            for col in df.columns:
                self.code_column_combo.addItem(str(col), col)
            
            # Активируем кнопки
            self.check_btn.setEnabled(True)
            self.reload_preview_btn.setEnabled(True)
            self.auto_detect_btn.setEnabled(True)
            self.preview_btn.setEnabled(True)
            
            # Обновляем информацию
            self._update_excel_info()
            
            logger.info(f"Загружен Excel: {self.excel_file_path}, header={header_row}, data_start={data_start_row}, строк: {len(df)}")
            
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось загрузить файл:\n{str(e)}")
            logger.error(f"Ошибка загрузки Excel: {e}", exc_info=True)
            
    def _show_excel_preview(self):
        """Показать предпросмотр данных Excel в отдельном окне"""
        if self.excel_data is None:
            QMessageBox.warning(self, "Ошибка", "Данные не загружены")
            return
        
        dialog = ExcelPreviewDialog(
            self.excel_data,
            self.header_row_spinbox.value(),
            self.data_start_row_spinbox.value(),
            self.current_sheet,
            self
        )
        dialog.exec_()
    
    def _update_excel_info(self):
        """Обновление информации о загруженных данных"""
        if self.excel_data is None:
            self.excel_info_label.setText("Данные не загружены")
            self.excel_info_label.setStyleSheet("color: #666; font-size: 9pt;")
            return
        
        header_row = self.header_row_spinbox.value()
        data_start_row = self.data_start_row_spinbox.value()
        total_rows = len(self.excel_data)
        
        info_parts = [f"Всего строк данных: <b>{total_rows}</b>"]
        
        if self.current_sheet:
            info_parts.append(f"Лист: <b>{self.current_sheet}</b>")
        
        info_parts.extend([
            f"Заголовки: строка <b>{header_row}</b>",
            f"Данные: с строки <b>{data_start_row}</b>",
            f"Столбцов: <b>{len(self.excel_data.columns)}</b>"
        ])
        
        info_text = " | ".join(info_parts)
        
        rows_between = data_start_row - header_row - 1
        if rows_between > 0:
            info_text += f" | <span style='color: #ff8c00;'>Пропущено после заголовков: {rows_between}</span>"
        
        self.excel_info_label.setText(info_text)
        self.excel_info_label.setStyleSheet("color: #333; font-size: 9pt;")
    
    def _load_oktmo_list(self):
        """Загрузка списка ОКТМО из базы данных (8 разрядов)"""
        try:
            filter_date = self.date_edit.date().toString("yyyy-MM-dd")
            
            # Сохраняем текущий выбор
            selected_code = self.oktmo_combo.currentData() if self.oktmo_combo.count() > 0 else None
            
            # Очищаем и добавляем пункт "Все"
            self.oktmo_combo.clear()
            self.oktmo_combo.addItem("00000000 — Все муниципалитеты", "00000000")
            
            # Загружаем ОКТМО через DatabaseManager (8 разрядов, как в проектах)
            pairs = self.db_manager.load_oktmo_for_municipality(filter_date, code_length=8)
            pairs.sort(key=lambda p: (p[0] or "").lower())
            
            # Добавляем в комбобокс
            for code, name in pairs:
                display = f"{code} — {name}" if name else code
                self.oktmo_combo.addItem(display, code)
            
            # Восстанавливаем выбор
            if selected_code:
                idx = self.oktmo_combo.findData(selected_code)
                if idx >= 0:
                    self.oktmo_combo.setCurrentIndex(idx)
                    
            logger.info(f"Загружено ОКТМО: {len(pairs)} (дата: {filter_date})")
            
        except Exception as e:
            logger.error(f"Ошибка загрузки ОКТМО: {e}", exc_info=True)
    
    def _on_date_changed(self, qdate):
        """Обработчик изменения даты - обновляем список ОКТМО"""
        self._load_oktmo_list()
        
    def _run_validation(self):
        """Запуск проверки"""
        if self.excel_data is None:
            QMessageBox.warning(self, "Ошибка", "Сначала загрузите Excel файл")
            return
            
        # Используем currentData() для получения оригинальных имен столбцов
        text_column = self.text_column_combo.currentData()
        code_column = self.code_column_combo.currentData() if self.code_column_combo.currentIndex() > 0 else None
        
        # Определяем справочник в зависимости от режима
        if self.advanced_mode:
            # Используем currentData() для получения реального имени таблицы
            reference_table = self.custom_table_combo.currentData()
            if not reference_table:
                QMessageBox.warning(self, "Ошибка", "Укажите таблицу справочника")
                return
            
            # Сохраняем настройки колонок из новых виджетов (берем оригинальные имена из userData)
            self.custom_name_column = self.custom_name_column_combo.currentData() or 'name'
            
            # Собираем отмеченные колонки для кода из списка с чекбоксами
            selected_code_columns = []
            for i in range(self.custom_code_columns_list.count()):
                item = self.custom_code_columns_list.item(i)
                if item.checkState() == Qt.Checked:
                    # Берем оригинальное имя колонки из UserRole
                    col_name = item.data(Qt.UserRole)
                    if col_name:
                        selected_code_columns.append(col_name)
            self.custom_code_columns = selected_code_columns
            
            logger.info(f"Расширенный режим: таблица={reference_table}, name_col={self.custom_name_column}, code_cols={self.custom_code_columns}")
        else:
            reference_table = self.reference_combo.currentData()
            
        check_date = self.date_edit.date().toString("yyyy-MM-dd")
        oktmo = self.oktmo_combo.currentData()
        max_distance = self.distance_spinbox.value()
        
        # Извлекаем данные (используя оригинальные имена столбцов)
        texts = self.excel_data[text_column].astype(str).str.strip().tolist()
        codes = self.excel_data[code_column].astype(str).str.strip().tolist() if code_column else None
        
        # Сохраняем коды для отображения
        self.orig_codes = codes if codes else []
        logger.info(f"Загружено кодов из файла: {len(self.orig_codes)} (колонка: {code_column})")
        
        # Загружаем справочник
        try:
            conn = sqlite3.connect(self.db_path)
            
            # Формируем фильтр по ОКТМО
            if oktmo and oktmo != "00000000":
                ppocode_filter = ["00000000", oktmo]
            else:
                ppocode_filter = "00000000"
            
            # В расширенном режиме загружаем напрямую из таблицы
            if self.advanced_mode:
                # Простой SELECT без фильтров (пользователь может выбрать любую таблицу)
                query = f"SELECT * FROM {reference_table}"
                reference_df = pd.read_sql_query(query, conn)
            else:
                reference_df = get_filtered_view(
                    conn, reference_table,
                    filter_date=check_date,
                    filter_ppocode=ppocode_filter,
                    deduplicate=True
                )
            conn.close()
            
            if reference_df.empty:
                QMessageBox.warning(self, "Ошибка", "Справочник пуст")
                return
            
            # Проверяем наличие колонки с наименованием
            name_col = self.custom_name_column if self.advanced_mode else 'name'
            if name_col not in reference_df.columns:
                QMessageBox.critical(self, "Ошибка", 
                    f"Колонка '{name_col}' не найдена в справочнике.\n"
                    f"Доступные колонки: {', '.join(reference_df.columns.tolist())}")
                return
                
            self.reference_data = reference_df
            self.reference_names = reference_df[name_col].astype(str).str.strip().tolist()
            self.reference_table_name = reference_table
            self.reference_date = check_date
            self.reference_oktmo = oktmo
            
            if codes:
                self.reference_codes = self._extract_reference_codes(reference_df, reference_table)
            else:
                self.reference_codes = None
                
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось загрузить справочник:\n{str(e)}")
            logger.error(f"Ошибка загрузки справочника: {e}", exc_info=True)
            return
            
        # Проверка
        progress = QProgressDialog("Проверка текстов...", "Отмена", 0, len(texts), self)
        progress.setWindowModality(Qt.WindowModal)
        progress.setMinimumDuration(0)
        
        try:
            def progress_callback(current, total):
                progress.setValue(current)
                return not progress.wasCanceled()
                
            all_results = find_errors(
                texts, self.reference_names,
                orig_codes=codes,
                reference_codes=self.reference_codes,
                max_distance=max_distance,
                progress_callback=progress_callback
            )
            
            progress.close()
            
            # Сохраняем все результаты
            self.all_items = sorted(all_results, key=lambda x: x.original_index)
            self.errors = [e for e in all_results if e.distance > 0]
            
            self._apply_filters()
            
            logger.info(f"Проверка завершена. Всего: {len(all_results)}, ошибок: {len(self.errors)}")
            
        except Exception as e:
            progress.close()
            QMessageBox.critical(self, "Ошибка", f"Ошибка при проверке:\n{str(e)}")
            logger.error(f"Ошибка проверки: {e}", exc_info=True)
            
    def _extract_reference_codes(self, df: pd.DataFrame, table_name: str) -> list:
        """Извлечение кодов из справочника"""
        try:
            # В расширенном режиме используем пользовательские колонки
            if self.advanced_mode and self.custom_code_columns:
                codes = []
                for _, row in df.iterrows():
                    parts = []
                    for col in self.custom_code_columns:
                        if col in df.columns and col in row and pd.notna(row[col]):
                            parts.append(str(row[col]).strip())
                    codes.append(''.join(parts) if parts else '')  # Склеиваем БЕЗ пробела
                logger.info(f"Расширенный режим: экстрагировано {len(codes)} кодов из колонок: {self.custom_code_columns}")
                return codes
            
            # Стандартная логика для известных справочников
            if 'inc' in table_name.lower():
                codes = []
                for _, row in df.iterrows():
                    parts = []
                    for col in ['inctypecode', 'incsubtypecode', 'analyticalgroupcode']:
                        if col in row and pd.notna(row[col]):
                            parts.append(str(row[col]).strip())
                    codes.append(' '.join(parts))
                return codes
                
            elif 'cost' in table_name.lower():
                codes = []
                for _, row in df.iterrows():
                    parts = []
                    for col in ['grbscode', 'rzpr', 'kcsr', 'kvr']:
                        if col in row and pd.notna(row[col]):
                            parts.append(str(row[col]).strip())
                    codes.append(' '.join(parts))
                return codes
                
            elif 'source' in table_name.lower():
                if 'code' in df.columns:
                    return df['code'].astype(str).str.strip().tolist()
                    
            return []
            
        except Exception as e:
            logger.error(f"Ошибка извлечения кодов: {e}", exc_info=True)
            return []
            
    def _apply_filters(self):
        """Применение фильтров расстояния и показа всех"""
        if not self.all_items:
            return
            
        show_all = self.show_all_checkbox.isChecked()
        max_distance = self.distance_spinbox.value()
        
        # Выбираем источник
        source = self.all_items if show_all else self.errors
        
        # Фильтруем по расстоянию
        actual_max = max_distance + 1 if max_distance == self.distance_spinbox.maximum() else max_distance
        filtered = [e for e in source if e.distance <= actual_max]
        
        self._display_results(filtered)
        
    def _display_results(self, items):
        """Отображение результатов"""
        from utils.text_validation import create_highlighted_text, normalize_classification_code
        
        show_all = self.show_all_checkbox.isChecked()
        
        logger.info(f"Обновление таблицы: {'все элементы' if show_all else 'только ошибки'} ({len(items)} записей)")
        logger.info(f"Коды в orig_codes: {len(self.orig_codes) if hasattr(self, 'orig_codes') else 'нет'}, в reference_codes: {len(self.reference_codes) if hasattr(self, 'reference_codes') and self.reference_codes else 'нет'}")
        
        if not items:
            self.errors_table.setRowCount(0)
            logger.info("Нет элементов для отображения")
            return
        
        self.errors_table.setRowCount(len(items))
        
        # Статистика
        total = len(self.excel_data) if self.excel_data is not None else 0
        errors_count = len(self.errors)
        correct_count = len(self.all_items) - errors_count
        
        if show_all:
            self.stats_label.setText(
                f"Всего: {total} | "
                f"<span style='color: #dc3545;'>Ошибок: {errors_count}</span> | "
                f"<span style='color: #28a745;'>Корректных: {correct_count}</span> | "
                f"Отображено: {len(items)}"
            )
        else:
            self.stats_label.setText(
                f"Всего: {total} | Найдено ошибок: <span style='color: #dc3545;'>{len(items)}</span>"
            )
        
        # Заполнение таблицы
        for row_idx, error in enumerate(items):
            try:
                # Текст в проекте - HTML с подсветкой
                html_text = self._richtext_to_html(
                    error.original_text,
                    error.diff_indices,
                    error.corrections
                )
                item = QTableWidgetItem()
                item.setData(Qt.DisplayRole, html_text)
                item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
                self.errors_table.setItem(row_idx, 0, item)
                
                # Эталон из справочника (оборачиваем в HTML для единообразного отображения)
                ref_text = error.reference_text if error.reference_text else ""
                ref_item = QTableWidgetItem()
                ref_item.setData(Qt.DisplayRole, ref_text)  # Обычный текст, без HTML
                ref_item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
                self.errors_table.setItem(row_idx, 1, ref_item)
                
                # Distance
                self.errors_table.setItem(row_idx, 2, QTableWidgetItem(str(error.distance)))
                
                # Код в проекте (с подсветкой, если есть ошибка в коде)
                orig_code = ""
                if hasattr(self, 'orig_codes') and self.orig_codes and error.original_index < len(self.orig_codes):
                    orig_code = self.orig_codes[error.original_index]
                
                if error.code_error:
                    # Нормализуем код и добавляем HTML подсветку
                    normalized_code = normalize_classification_code(orig_code)
                    code_html = self._richtext_to_html(
                        normalized_code,
                        error.code_error.diff_indices,
                        error.code_error.corrections
                    )
                    item = QTableWidgetItem()
                    item.setData(Qt.DisplayRole, code_html)
                    item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
                    self.errors_table.setItem(row_idx, 3, item)
                else:
                    # Нормализуем код даже если нет ошибки
                    normalized_code = normalize_classification_code(orig_code) if orig_code else ""
                    self.errors_table.setItem(row_idx, 3, QTableWidgetItem(normalized_code))
                
                # Код справочника (также нормализуем и оборачиваем в простой HTML)
                ref_code_item = QTableWidgetItem()
                if error.code_error and error.code_error.ref_code:
                    ref_code = normalize_classification_code(error.code_error.ref_code)
                    ref_code_item.setData(Qt.DisplayRole, ref_code)
                elif error.reference_index is not None and hasattr(self, 'reference_codes') and self.reference_codes and error.reference_index < len(self.reference_codes):
                    ref_code = normalize_classification_code(self.reference_codes[error.reference_index])
                    ref_code_item.setData(Qt.DisplayRole, ref_code)
                else:
                    ref_code_item.setData(Qt.DisplayRole, "")
                ref_code_item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
                self.errors_table.setItem(row_idx, 4, ref_code_item)
            
            except Exception as e:
                logger.error(f"Ошибка при обработке строки {row_idx}: {e}", exc_info=True)
                # Заполняем строку пустыми значениями
                for col in range(5):
                    if self.errors_table.item(row_idx, col) is None:
                        self.errors_table.setItem(row_idx, col, QTableWidgetItem(""))
        
        # Подгоняем высоту строк
        self.errors_table.resizeRowsToContents()
        self.export_btn.setEnabled(len(items) > 0)
    
    def _richtext_to_html(self, orig_text, diff_idx, corrections, invert_colors=False):
        """Преобразование подсвеченного текста в HTML (делегирует в общую функцию)"""
        return richtext_to_html(orig_text, diff_idx, corrections, invert_colors)
    
    def _show_error_details(self, row, column):
        """Показ детальной информации об ошибке текста при двойном клике с альтернативными вариантами"""
        if row < 0:
            return
        
        show_all = self.show_all_checkbox.isChecked()
        max_distance = self.distance_spinbox.value()
        actual_max = max_distance + 1 if max_distance == self.distance_spinbox.maximum() else max_distance
        
        source = self.all_items if show_all else self.errors
        filtered = [e for e in source if e.distance <= actual_max]
        
        if row >= len(filtered):
            return
            
        error = filtered[row]
        
        # Получаем коды
        from utils.text_validation import normalize_classification_code
        
        orig_code = ""
        if hasattr(self, 'orig_codes') and self.orig_codes and error.original_index < len(self.orig_codes):
            orig_code = normalize_classification_code(self.orig_codes[error.original_index])
        
        ref_code = ""
        if error.code_error and error.code_error.ref_code:
            ref_code = normalize_classification_code(error.code_error.ref_code)
        elif error.reference_index is not None and hasattr(self, 'reference_codes') and self.reference_codes and error.reference_index < len(self.reference_codes):
            ref_code = normalize_classification_code(self.reference_codes[error.reference_index])
        
        # Информация о коде
        if error.code_error and error.code_error.has_error:
            code_status = "<span style='color: red;'>❌ Несоответствие</span>"
        elif orig_code and ref_code:
            code_status = "<span style='color: green;'>✓ Совпадает</span>"
        else:
            code_status = "<span style='color: gray;'>— Не проверялся</span>"
        
        # Получаем startdate для текущего эталона
        ref_startdate = ""
        if error.reference_index is not None and self.reference_data is not None and error.reference_index < len(self.reference_data):
            ref_data_dict = self.reference_data.to_dict('records') if hasattr(self.reference_data, 'to_dict') else []
            if error.reference_index < len(ref_data_dict):
                ref_startdate = ref_data_dict[error.reference_index].get('startdate', '')
                if ref_startdate:
                    try:
                        from datetime import datetime
                        date_obj = datetime.strptime(ref_startdate.split()[0], '%Y-%m-%d')
                        ref_startdate = date_obj.strftime('%d.%m.%Y')
                    except:
                        pass
        
        # Создаем расширенный диалог
        from PyQt5.QtWidgets import QSplitter, QTabWidget
        from PyQt5.QtCore import QTimer
        
        dialog = QDialog(self)
        dialog.setWindowTitle("🔍 Детали ошибки текста с альтернативами")
        dialog.setMinimumSize(1200, 650)
        
        main_layout = QVBoxLayout(dialog)
        main_layout.setSpacing(8)
        main_layout.setContentsMargins(10, 10, 10, 10)
        
        # === КОМПАКТНАЯ КАРТОЧКА С ОСНОВНОЙ ИНФОРМАЦИЕЙ ===
        info_card = QWidget()
        info_card.setObjectName("infoCard")
        info_card.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Minimum)
        info_card.setMaximumHeight(30)
        info_layout = QHBoxLayout(info_card)
        info_layout.setSpacing(8)
        info_layout.setContentsMargins(6, 2, 6, 2)
        
        # Distance для текста
        text_distance_color = "#28a745" if error.distance < 10 else ("#ffc107" if error.distance < 50 else "#dc3545")
        text_status = QLabel(f"<b>📝 Текст:</b> <span style='color: {text_distance_color}; font-weight: bold;'>D:{error.distance}</span>")
        info_layout.addWidget(text_status)
        
        # Статус кода
        code_status_label = QLabel(f"<b>🔢 Код:</b> {code_status}")
        info_layout.addWidget(code_status_label)
        
        # Дата
        if ref_startdate:
            date_label = QLabel(f"<b>📅 Дата:</b> {ref_startdate}")
            info_layout.addWidget(date_label)
        
        info_layout.addStretch()
        info_card.setStyleSheet("""
            QWidget#infoCard {
                background-color: #f8f9fa;
                border: 1px solid #dee2e6;
                border-radius: 4px;
                padding: 4px;
            }
        """)
        main_layout.addWidget(info_card)
        
        # === ОСНОВНОЙ SPLITTER ===
        splitter = QSplitter(Qt.Horizontal)
        
        # === ЛЕВАЯ ПАНЕЛЬ: Сравнение текстов ===
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(0, 0, 0, 0)
        
        left_label = QLabel("<b>🔄 Сравнение текстов:</b>")
        left_label.setObjectName("sectionHeader")
        left_layout.addWidget(left_label)
        
        details = f"""
        <div style='padding: 10px;'>
        <table cellpadding='8' style='border: 1px solid #dee2e6; border-radius: 4px; width: 100%;'>
        <tr style='background-color: #fff3cd;'>
            <td style='width: 120px;'><b>📄 В Excel:</b></td>
            <td>{error.original_text}</td>
        </tr>
        <tr style='background-color: #d1ecf1;'>
            <td><b>✅ Эталон:</b></td>
            <td>{error.reference_text if error.reference_text else '—'}</td>
        </tr>
        <tr>
            <td><b>🔢 Код (Excel):</b></td>
            <td><code>{orig_code if orig_code else '—'}</code></td>
        </tr>
        <tr>
            <td><b>🔢 Код (справочник):</b></td>
            <td><code>{ref_code if ref_code else '—'}</code></td>
        </tr>
        </table>
        <p style='margin-top: 10px; color: #6c757d; font-size: 9pt;'>
        💡 <i>Красным выделены различия, зеленым — исправления</i>
        </p>
        </div>
        """
        
        text_browser = QTextBrowser()
        text_browser.setHtml(details)
        text_browser.setOpenExternalLinks(False)
        left_layout.addWidget(text_browser)
        
        splitter.addWidget(left_widget)
        
        # === ПРАВАЯ ПАНЕЛЬ: ВКЛАДКИ С АЛЬТЕРНАТИВАМИ ===
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setContentsMargins(0, 0, 0, 0)
        
        right_label = QLabel("<b>🔎 Альтернативные варианты:</b>")
        right_label.setObjectName("sectionHeader")
        right_layout.addWidget(right_label)
        
        # Вкладки для двух режимов
        tabs = QTabWidget()
        tabs.setObjectName("alternativesTabs")
        
        # ВКЛАДКА 1: Актуальные (дедублированные)
        tab1 = QWidget()
        tab1_layout = QVBoxLayout(tab1)
        alternatives_table_dedup = QTableWidget()
        alternatives_table_dedup.setColumnCount(2)
        alternatives_table_dedup.setHorizontalHeaderLabels(["D / Код / Дата / ОКТМО", "Текст"])
        alternatives_table_dedup.horizontalHeader().setStretchLastSection(False)
        alternatives_table_dedup.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        alternatives_table_dedup.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        alternatives_table_dedup.setSelectionBehavior(QTableWidget.SelectRows)
        alternatives_table_dedup.setSelectionMode(QTableWidget.SingleSelection)
        alternatives_table_dedup.setEditTriggers(QTableWidget.NoEditTriggers)
        alternatives_table_dedup.setWordWrap(True)
        alternatives_table_dedup.setAlternatingRowColors(True)
        alternatives_table_dedup.setObjectName("alternativesTable")
        
        # Устанавливаем делегаты
        from views.text_validation_widget import HtmlDelegate, WordWrapDelegate
        alternatives_table_dedup.setItemDelegateForColumn(0, WordWrapDelegate(alternatives_table_dedup))
        alternatives_table_dedup.setItemDelegateForColumn(1, HtmlDelegate(alternatives_table_dedup))
        
        # Находим альтернативные варианты с дедупликацией
        alternatives = self._find_alternative_variants(error.original_text, error.original_index, deduplicate=True)
        alternatives_table_dedup.setRowCount(len(alternatives))
        
        # Импортируем функцию для поиска различий
        from utils.text_validation.text_comparator import find_differences
        from utils.text_validation import normalize_classification_code
        
        for idx, (distance, text, code, ref_idx, startdate, ppocode, pponame) in enumerate(alternatives):
            # Объединяем Distance, код, startdate и ОКТМО в одну ячейку
            # Нормализуем код для отображения
            normalized_code = normalize_classification_code(code) if code else ""
            distance_code_text = f"D:{distance}\n{normalized_code}"
            if startdate:
                try:
                    from datetime import datetime
                    date_obj = datetime.strptime(startdate.split()[0], '%Y-%m-%d')
                    formatted_date = date_obj.strftime('%d.%m.%Y')
                    distance_code_text += f"\n{formatted_date}"
                except:
                    distance_code_text += f"\n{startdate}"
            if ppocode:
                distance_code_text += f"\n{ppocode}"
            distance_code_item = QTableWidgetItem(distance_code_text)
            
            # Генерируем HTML с подсветкой различий
            diff_indices, corrections = find_differences(text, error.original_text)
            html_text = self._richtext_to_html(text, diff_indices, corrections, invert_colors=True)
            
            text_item = QTableWidgetItem()
            text_item.setData(Qt.DisplayRole, html_text)
            text_item.setData(Qt.UserRole, text)
            
            # Проверяем совпадение кодов
            normalized_orig_code = normalize_classification_code(orig_code) if orig_code else ""
            normalized_alt_code = normalize_classification_code(code) if code else ""
            codes_match = normalized_orig_code and normalized_alt_code and normalized_orig_code == normalized_alt_code
            
            # Цветовое кодирование
            if codes_match:
                distance_code_item.setBackground(QColor("#28a745"))
                distance_code_item.setForeground(Qt.white)
            elif distance < 10:
                distance_code_item.setBackground(Qt.green)
                distance_code_item.setForeground(Qt.white)
            elif distance < 50:
                distance_code_item.setBackground(Qt.yellow)
            else:
                distance_code_item.setBackground(Qt.red)
                distance_code_item.setForeground(Qt.white)
            
            distance_code_item.setTextAlignment(Qt.AlignTop | Qt.AlignLeft)
            text_item.setTextAlignment(Qt.AlignTop | Qt.AlignLeft)
            
            alternatives_table_dedup.setItem(idx, 0, distance_code_item)
            alternatives_table_dedup.setItem(idx, 1, text_item)
            
            distance_code_item.setData(Qt.UserRole, ref_idx)
            distance_code_item.setData(Qt.UserRole + 1, code)
            distance_code_item.setData(Qt.UserRole + 2, distance)
            distance_code_item.setData(Qt.UserRole + 3, ppocode)
            distance_code_item.setData(Qt.UserRole + 4, pponame)
        
        alternatives_table_dedup.resizeRowsToContents()
        
        alternatives_table_dedup.cellDoubleClicked.connect(
            lambda r, c: self._show_alternative_details(alternatives_table_dedup, r)
        )
        
        tab1_layout.addWidget(alternatives_table_dedup)
        tabs.addTab(tab1, f"✨ Актуальные ({len(alternatives)})")
        
        # ВКЛАДКА 2: Все варианты
        tab2 = QWidget()
        tab2_layout = QVBoxLayout(tab2)
        
        alternatives_table_all = QTableWidget()
        alternatives_table_all.setColumnCount(2)
        alternatives_table_all.setHorizontalHeaderLabels(["D / Код / Дата / ОКТМО", "Текст"])
        alternatives_table_all.horizontalHeader().setStretchLastSection(False)
        alternatives_table_all.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        alternatives_table_all.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        alternatives_table_all.setSelectionBehavior(QTableWidget.SelectRows)
        alternatives_table_all.setSelectionMode(QTableWidget.SingleSelection)
        alternatives_table_all.setEditTriggers(QTableWidget.NoEditTriggers)
        alternatives_table_all.setWordWrap(True)
        alternatives_table_all.setAlternatingRowColors(True)
        alternatives_table_all.setObjectName("alternativesTable")
        
        alternatives_table_all.setItemDelegateForColumn(0, WordWrapDelegate(alternatives_table_all))
        alternatives_table_all.setItemDelegateForColumn(1, HtmlDelegate(alternatives_table_all))
        
        alternatives_all = self._find_alternative_variants(error.original_text, error.original_index, deduplicate=False)
        alternatives_table_all.setRowCount(len(alternatives_all))
        
        for idx, (distance, text, code, ref_idx, startdate, ppocode, pponame) in enumerate(alternatives_all):
            # Нормализуем код для отображения
            normalized_code = normalize_classification_code(code) if code else ""
            distance_code_text = f"D:{distance}\n{normalized_code}"
            if startdate:
                try:
                    from datetime import datetime
                    date_obj = datetime.strptime(startdate.split()[0], '%Y-%m-%d')
                    formatted_date = date_obj.strftime('%d.%m.%Y')
                    distance_code_text += f"\n{formatted_date}"
                except:
                    distance_code_text += f"\n{startdate}"
            if ppocode:
                distance_code_text += f"\n{ppocode}"
            distance_code_item = QTableWidgetItem(distance_code_text)
            
            diff_indices, corrections = find_differences(text, error.original_text)
            html_text = self._richtext_to_html(text, diff_indices, corrections, invert_colors=True)
            
            text_item = QTableWidgetItem()
            text_item.setData(Qt.DisplayRole, html_text)
            text_item.setData(Qt.UserRole, text)
            
            normalized_orig_code = normalize_classification_code(orig_code) if orig_code else ""
            normalized_alt_code = normalize_classification_code(code) if code else ""
            codes_match = normalized_orig_code and normalized_alt_code and normalized_orig_code == normalized_alt_code
            
            if codes_match:
                distance_code_item.setBackground(QColor("#28a745"))
                distance_code_item.setForeground(Qt.white)
            elif distance < 10:
                distance_code_item.setBackground(Qt.green)
                distance_code_item.setForeground(Qt.white)
            elif distance < 50:
                distance_code_item.setBackground(Qt.yellow)
            else:
                distance_code_item.setBackground(Qt.red)
                distance_code_item.setForeground(Qt.white)
            
            distance_code_item.setTextAlignment(Qt.AlignTop | Qt.AlignLeft)
            text_item.setTextAlignment(Qt.AlignTop | Qt.AlignLeft)
            
            alternatives_table_all.setItem(idx, 0, distance_code_item)
            alternatives_table_all.setItem(idx, 1, text_item)
            
            distance_code_item.setData(Qt.UserRole, ref_idx)
            distance_code_item.setData(Qt.UserRole + 1, code)
            distance_code_item.setData(Qt.UserRole + 2, distance)
            distance_code_item.setData(Qt.UserRole + 3, ppocode)
            distance_code_item.setData(Qt.UserRole + 4, pponame)
        
        alternatives_table_all.resizeRowsToContents()
        
        alternatives_table_all.cellDoubleClicked.connect(
            lambda r, c: self._show_alternative_details(alternatives_table_all, r)
        )
        
        tab2_layout.addWidget(alternatives_table_all)
        tabs.addTab(tab2, f"📚 Все варианты ({len(alternatives_all)})")
        
        right_layout.addWidget(tabs)
        splitter.addWidget(right_widget)
        
        # Устанавливаем пропорции
        splitter.setStretchFactor(0, 55)
        splitter.setStretchFactor(1, 45)
        
        main_layout.addWidget(splitter)
        
        # === ИНФОРМАЦИОННАЯ ПАНЕЛЬ С ПОДСКАЗКАМИ ===
        info_footer = QLabel()
        info_footer.setObjectName("infoFooter")
        info_footer_text = "<b>💡 Подсказка:</b> <b>⭐ Актуальные</b> - топ-5 из текущих версий справочника (без дублей по коду) | <b>📚 Все варианты</b> - топ-5 из всех версий (с историческими)"
        info_footer.setText(info_footer_text)
        info_footer.setWordWrap(True)
        info_footer.setMaximumHeight(60)
        main_layout.addWidget(info_footer)
        
        # Кнопка закрытия
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        close_btn = QPushButton("✖ Закрыть")
        close_btn.setObjectName("buttonClose")
        close_btn.clicked.connect(dialog.accept)
        button_layout.addWidget(close_btn)
        main_layout.addLayout(button_layout)
        
        # Отложенный вызов для правильного расчета высоты строк
        QTimer.singleShot(100, lambda: (alternatives_table_dedup.resizeRowsToContents(), alternatives_table_all.resizeRowsToContents()))
        
        dialog.exec_()
    
    def _find_alternative_variants(self, original_text: str, original_index: int, deduplicate: bool = True):
        """Поиск альтернативных вариантов из справочника (делегирует в общую функцию)"""
        if not self.reference_table_name:
            return []
        
        ppocode_filter = self.reference_oktmo
        if ppocode_filter and ppocode_filter != "00000000":
            ppocode_filter = ["00000000", ppocode_filter]
        else:
            ppocode_filter = "00000000"
        
        # Создаем функцию-экстрактор кодов, которая использует наш метод
        def code_extractor(df):
            return self._extract_reference_codes(df, self.reference_table_name)
        
        return find_alternative_variants(
            original_text=original_text,
            db_path=self.db_path,
            table_name=self.reference_table_name,
            filter_date=self.reference_date,
            ppocode_filter=ppocode_filter,
            code_extractor=code_extractor,
            deduplicate=deduplicate,
            top_n=5
        )
    
    def _show_alternative_details(self, table: QTableWidget, row: int):
        """Показ детальной информации об альтернативном варианте"""
        if row < 0 or row >= table.rowCount():
            return
        
        distance_code_item = table.item(row, 0)
        text_item = table.item(row, 1)
        
        if not all([distance_code_item, text_item]):
            return
        
        # Извлекаем данные
        ref_idx = distance_code_item.data(Qt.UserRole)
        distance = distance_code_item.data(Qt.UserRole + 2)
        code = distance_code_item.data(Qt.UserRole + 1)
        text = text_item.data(Qt.UserRole)
        ppocode = distance_code_item.data(Qt.UserRole + 3)
        pponame = distance_code_item.data(Qt.UserRole + 4)
        
        # Формируем детальное описание
        # Нормализуем код для отображения
        from utils.text_validation import normalize_classification_code
        normalized_code = normalize_classification_code(code) if code else '—'
        
        details = f"""<h3>Детали альтернативного варианта</h3>
        <table cellpadding='5' style='border: 1px solid #ccc;'>
        <tr style='background-color: #e6f7ff;'><td><b>Текст из справочника:</b></td><td>{text}</td></tr>
        <tr><td><b>Код классификации:</b></td><td>{normalized_code}</td></tr>
        <tr><td><b>Distance от текста в Excel:</b></td><td><b>{distance}</b></td></tr>
        <tr><td><b>ОКТМО:</b></td><td>{ppocode if ppocode else '—'} {pponame if pponame else ''}</td></tr>
        </table>
        <p style='margin-top: 10px;'><i>ℹ️ Это один из возможных вариантов из справочника</i></p>
        """
        
        # Создаем диалог
        detail_dialog = QDialog(self)
        detail_dialog.setWindowTitle("📋 Детали альтернативного варианта")
        detail_dialog.resize(700, 400)
        
        layout = QVBoxLayout(detail_dialog)
        
        browser = QTextBrowser()
        browser.setHtml(details)
        layout.addWidget(browser)
        
        close_btn = QPushButton("Закрыть")
        close_btn.clicked.connect(detail_dialog.accept)
        layout.addWidget(close_btn, alignment=Qt.AlignRight)
        
        detail_dialog.exec_()
    
    def _export_results(self):
        """Экспорт в Excel"""
        show_all = self.show_all_checkbox.isChecked()
        max_distance = self.distance_spinbox.value()
        actual_max = max_distance + 1 if max_distance == self.distance_spinbox.maximum() else max_distance
        
        source = self.all_items if show_all else self.errors
        items = [e for e in source if e.distance <= actual_max]
        
        if not items:
            QMessageBox.warning(self, "Предупреждение", "Нет данных для экспорта")
            return
            
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Сохранить",
            f"validation_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            "Excel Files (*.xlsx)"
        )
        
        if not file_path:
            return
            
        try:
            from utils.text_validation import normalize_classification_code
            
            data = []
            for e in items:
                # Получаем код из Excel файла
                orig_code = ""
                if hasattr(self, 'orig_codes') and self.orig_codes and e.original_index < len(self.orig_codes):
                    orig_code = normalize_classification_code(self.orig_codes[e.original_index])
                
                # Получаем код из справочника
                ref_code = ""
                if e.code_error and e.code_error.ref_code:
                    ref_code = normalize_classification_code(e.code_error.ref_code)
                elif e.reference_index is not None and hasattr(self, 'reference_codes') and self.reference_codes and e.reference_index < len(self.reference_codes):
                    ref_code = normalize_classification_code(self.reference_codes[e.reference_index])
                
                data.append({
                    '№': e.original_index + 1,
                    'Текст из Excel': e.original_text,
                    'Код из Excel': orig_code,
                    'Расстояние': e.distance,
                    'Эталонный текст': e.reference_text,
                    'Код из справочника': ref_code,
                    'Статус': 'OK' if e.distance == 0 else 'Ошибка'
                })
                
            df = pd.DataFrame(data)
            df.to_excel(file_path, index=False, engine='openpyxl')
            
            QMessageBox.information(self, "Успех", f"Сохранено:\n{file_path}")
            logger.info(f"Экспорт: {file_path}")
            
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось экспортировать:\n{str(e)}")
            logger.error(f"Ошибка экспорта: {e}", exc_info=True)
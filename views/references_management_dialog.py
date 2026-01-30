"""
Диалог для управления справочниками (загрузка из Excel, просмотр, редактирование)
"""
from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QTabWidget, QWidget,
    QTableWidget, QTableWidgetItem, QPushButton, QFileDialog,
    QMessageBox, QLabel, QHeaderView, QLineEdit,
    QSplitter, QListWidget, QListWidgetItem, QFormLayout,
    QTextEdit, QScrollArea, QSizePolicy, QMenu, QSpinBox, QDateEdit, QCheckBox
)
from PyQt5.QtCore import Qt, QDate
from PyQt5.QtGui import QFont, QFontMetrics
from pathlib import Path
from datetime import datetime
import pandas as pd
import sqlite3
from logger import logger

from models.database import DatabaseManager
from models.base_models import YearRef, MunicipalityRef, FormTypeMeta, PeriodRef
from views.budget_references_update_dialog import REFERENCE_NAMES
from views.reference_detail_dialog import ReferenceDetailDialog
from views.column_visibility_dialog import ColumnVisibilityDialog
from controllers.reference_controller import ReferenceController
from styles.styles import set_tab_bar_min_width
from utils.db_utils import get_filtered_view


class ReferencesManagementDialog(QDialog):
    """Диалог для управления справочниками"""
    
    # Список справочников с их методами загрузки
    REFERENCE_TYPES = {
        'Коды доходов': {
            'table': 'v_budgetclastypeinc_merged',
            'load_method': 'load_income_sources_reference',
            'load_type': 'доходы',
            'is_view': True,
            'columns': ['inctypecode', 'incsubtypecode', 'analyticalgroupcode', 'name', 'level'],
            'display_columns': ['concatenated_code AS код', 'name AS наименование', 'level AS уровень'],
            'search_columns': ['inctypecode', 'incsubtypecode', 'analyticalgroupcode', 'name', 'level'],
        },
        'Коды источников': {
            'table': 'source_reference_records',
            'load_method': 'load_income_sources_reference',  # Специальный метод для доходов/источников
            'load_type': 'источники',  # Тип для ReferenceController
            'columns': ['code', 'name', 'level', 'doc'],
            'display_columns': ['code AS код', 'name AS наименование', 'level AS уровень', 'doc AS документ']
        },
        # Справочники конфигурации
        'Годы': {
            'table': 'ref_years',
            'load_method': None,
            'columns': ['year', 'is_active'],
            'is_config': True,
            'load_func': '_load_years',
            'save_func': '_save_years'
        },
        'Муниципальные образования': {
            'table': 'ref_municipalities',
            'load_method': None,
            'columns': ['code', 'name', 'is_active'],
            'is_config': True,
            'load_func': '_load_municipalities',
            'save_func': '_save_municipalities'
        },
        'Типы форм': {
            'table': 'ref_form_types',
            'load_method': None,
            'columns': ['id', 'code', 'name', 'periodicity', 'is_active'],
            'is_config': True,
            'load_func': '_load_forms',
            'save_func': '_save_forms'
        },
        'Периоды': {
            'table': 'ref_periods',
            'load_method': None,
            'columns': ['id', 'code', 'name', 'sort_order', 'form_type_code', 'is_active'],
            'is_config': True,
            'load_func': '_load_periods',
            'save_func': '_save_periods'
        },
        # Справочники из бюджетной системы (онлайн справочники)
        '─── Онлайн справочники ───': {
            'table': None,
            'is_separator': True
        },
        'ОКТМО': {
            'table': 'oktmo',
            'load_method': None,
            'columns': None,  # None означает загрузить все колонки
            'is_online': True
        },
        'Классификаторы доходов бюджета ФУ': {
            'table': 'budgetclastypeinc',
            'load_method': None,
            'columns': None,  # None означает загрузить все колонки
            'is_online': True,
            'has_npa': True  # Указываем, что есть связь с npa
        },
        'Классификаторы доходов бюджета МО': {
            'table': 'budgetclassubtypincmo',
            'load_method': None,
            'columns': None,
            'is_online': True,
            'has_npa': True
        },
        'Администраторы бюджета ФУ': {
            'table': 'budgetclasgabs',
            'load_method': None,
            'columns': None,
            'is_online': True,
            'has_npa': True
        },
        'Администраторы бюджета МО': {
            'table': 'budgetclasgabsmo',
            'load_method': None,
            'columns': None,
            'is_online': True,
            'has_npa': True
        },
        'Распорядители бюджета ФУ': {
            'table': 'budgetclasgrbs',
            'load_method': None,
            'columns': None,
            'is_online': True,
            'has_npa': True
        },
        'Распорядители бюджета МО': {
            'table': 'budgetclasgrbsmo',
            'load_method': None,
            'columns': None,
            'is_online': True,
            'has_npa': True
        },
        'Классификаторы расходов бюджета ФУ': {
            'table': 'budgetclascosts',
            'load_method': None,
            'columns': None,
            'is_online': True,
            'has_npa': True
        },
        'Классификаторы расходов бюджета МО': {
            'table': 'budgetclascostsmo',
            'load_method': None,
            'columns': None,
            'is_online': True,
            'has_npa': True
        },
        'Источники финансирования дефицита ФУ': {
            'table': 'budgetclasgaiffb',
            'load_method': None,
            'columns': None,
            'is_online': True,
            'has_npa': True
        },
        'Источники финансирования дефицита МО': {
            'table': 'budgetclasgaifmo',
            'load_method': None,
            'columns': None,
            'is_online': True,
            'has_npa': True
        },
        'Классификаторы источников финансирования ФУ': {
            'table': 'budgetclassources',
            'load_method': None,
            'columns': None,
            'is_online': True,
            'has_npa': True
        },
        'Классификаторы источников финансирования МО': {
            'table': 'budgetclassourcesmo',
            'load_method': None,
            'columns': None,
            'is_online': True,
            'has_npa': True
        },
        # Объединенные представления (ФУ + МО)
        '─── Объединенные представления ───': {
            'table': None,
            'is_separator': True
        },
        'Классификаторы доходов (объединенные)': {
            'table': 'v_budgetclastypeinc_merged',
            'load_method': None,
            'columns': None,
            'is_view': True,
            'has_npa': True
        },
        'Классификаторы расходов (объединенные)': {
            'table': 'v_budgetclascosts_merged',
            'load_method': None,
            'columns': None,
            'is_view': True,
            'has_npa': True
        },
        'Распорядители бюджета (объединенные)': {
            'table': 'v_budgetclasgrbs_merged',
            'load_method': None,
            'columns': None,
            'is_view': True,
            'has_npa': True
        },
        'Администраторы бюджета (объединенные)': {
            'table': 'v_budgetclasgabs_merged',
            'load_method': None,
            'columns': None,
            'is_view': True,
            'has_npa': True
        },
        'Источники финансирования дефицита (объединенные)': {
            'table': 'v_budgetclasgaiffb_merged',
            'load_method': None,
            'columns': None,
            'is_view': True,
            'has_npa': True
        },
        'Классификаторы источников финансирования (объединенные)': {
            'table': 'v_budgetclassources_merged',
            'load_method': None,
            'columns': None,
            'is_view': True,
            'has_npa': True
        }
    }
    
    def __init__(self, db_manager: DatabaseManager, parent=None):
        super().__init__(parent)
        self.db_manager = db_manager
        self.setWindowTitle("Справочники")
        self.setMinimumSize(1200, 700)
        # Добавляем кнопку максимизации
        self.setWindowFlags(self.windowFlags() | Qt.WindowMaximizeButtonHint)
        
        # Стили применяются глобально через QApplication
        
        # Получаем reference_controller из родительского окна, если доступен
        self.reference_controller = None
        if parent and hasattr(parent, 'controller') and hasattr(parent.controller, 'reference_controller'):
            self.reference_controller = parent.controller.reference_controller
        
        self.current_reference_type = None
        
        # Переменные для пагинации
        self.current_page = 1
        self.page_size = 1000  # Количество записей на странице
        self.total_records = 0
        self.total_pages = 0
        
        # Переменная для хранения поискового запроса
        self.search_text = ""
        
        # Переменная для хранения выбранной даты фильтрации
        self.filter_date = None  # По умолчанию фильтрация отключена
        
        self.init_ui()
        # Дата по умолчанию из конфига (из метаданных ревизии или текущая)
        ref_date = self.db_manager.load_config("reference_filter_date")
        if ref_date and hasattr(self, 'date_filter') and hasattr(self, 'date_filter_checkbox'):
            try:
                from PyQt5.QtCore import QDate
                parts = ref_date.split('-')
                if len(parts) == 3:
                    y, m, d = int(parts[0]), int(parts[1]), int(parts[2])
                    self.date_filter.setDate(QDate(y, m, d))
                    self.date_filter_checkbox.setChecked(True)
                    self.date_filter.setEnabled(True)
            except Exception:
                pass
        self._update_filter_date()
        self.load_reference_list()
    
    def init_ui(self):
        """Инициализация UI"""
        layout = QVBoxLayout(self)
        
        # Создаем сплиттер для разделения на левую и правую части
        splitter = QSplitter(Qt.Horizontal)
        
        # Левая панель - список справочников
        left_panel = self.create_left_panel()
        splitter.addWidget(left_panel)
        
        # Правая панель - данные справочника
        right_panel = self.create_right_panel()
        splitter.addWidget(right_panel)
        
        # Устанавливаем пропорции (левая панель уже, правая шире)
        splitter.setSizes([250, 950])
        
        layout.addWidget(splitter)
        
        # Кнопки внизу
        buttons_layout = QHBoxLayout()
        buttons_layout.addStretch()
        
        close_btn = QPushButton("✕ Закрыть")
        close_btn.clicked.connect(self.accept)
        close_btn.setDefault(True)
        close_btn.setSizePolicy(QSizePolicy.Minimum, QSizePolicy.Fixed)
        buttons_layout.addWidget(close_btn)
        
        layout.addLayout(buttons_layout)
    
    def create_left_panel(self) -> QWidget:
        """Создание левой панели со списком справочников"""
        panel = QWidget()
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(10)
        
        # Заголовок
        title_label = QLabel("📚 Справочники")
        title_font = QFont("Arial", 12, QFont.Bold)
        title_label.setFont(title_font)
        title_label.setWordWrap(True)
        layout.addWidget(title_label)
        
        # Поле выбора даты для фильтрации
        date_layout = QHBoxLayout()
        date_layout.setSpacing(5)
        
        self.date_filter_checkbox = QCheckBox("Фильтровать по дате:")
        self.date_filter_checkbox.setChecked(False)  # По умолчанию фильтрация отключена
        self.date_filter_checkbox.stateChanged.connect(self.on_date_filter_toggled)
        date_layout.addWidget(self.date_filter_checkbox)
        
        self.date_filter = QDateEdit()
        self.date_filter.setCalendarPopup(True)
        self.date_filter.setDate(QDate.currentDate())
        self.date_filter.setDisplayFormat("dd.MM.yyyy")
        self.date_filter.setEnabled(False)  # По умолчанию отключено
        self.date_filter.dateChanged.connect(self.on_date_filter_changed)
        date_layout.addWidget(self.date_filter)
        
        layout.addLayout(date_layout)
        
        # Список справочников
        self.references_list = QListWidget()
        self.references_list.itemClicked.connect(self.on_reference_item_clicked)
        layout.addWidget(self.references_list)
        
        return panel
    
    def create_right_panel(self) -> QWidget:
        """Создание правой панели с данными справочника"""
        panel = QWidget()
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(10)
        
        # Заголовок с названием выбранного справочника
        self.reference_title_label = QLabel("Выберите справочник")
        title_font = QFont("Arial", 12, QFont.Bold)
        self.reference_title_label.setFont(title_font)
        self.reference_title_label.setWordWrap(True)
        layout.addWidget(self.reference_title_label)
        
        # Вкладки для просмотра и загрузки
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
        
        # Настраиваем вкладки для предотвращения обрезания текста (после добавления всех вкладок)
        self.tabs.tabBar().setUsesScrollButtons(True)
        self.tabs.tabBar().setElideMode(Qt.ElideNone)
        
        # Вычисляем минимальную ширину на основе самого длинного текста вкладки с учетом жирного шрифта
        # и применяем через CSS стили для всех вкладок
        max_bold_width = 0
        for i in range(self.tabs.count()):
            tab_text = self.tabs.tabText(i)
            if tab_text:
                # Получаем метрики шрифта для жирного текста (выделенные вкладки)
                bold_font = self.tabs.tabBar().font()
                bold_font.setBold(True)
                bold_font_metrics = QFontMetrics(bold_font)
                
                # Вычисляем ширину текста для жирного шрифта
                try:
                    bold_width = bold_font_metrics.horizontalAdvance(tab_text)
                except AttributeError:
                    # Для старых версий PyQt5 используем width()
                    bold_width = bold_font_metrics.width(tab_text)
                
                # Находим максимальную ширину
                if bold_width > max_bold_width:
                    max_bold_width = bold_width
        
        # Устанавливаем минимальную ширину для всех вкладок через CSS
        # Добавляем padding (20px с каждой стороны для выделенных вкладок)
        if max_bold_width > 0:
            min_width = max_bold_width + 40
            # Применяем стиль к QTabBar
            set_tab_bar_min_width(self.tabs.tabBar(), min_width)
        
        layout.addWidget(self.tabs)
        
        return panel
    
    def setup_view_tab(self):
        """Настройка вкладки просмотра"""
        layout = QVBoxLayout(self.view_tab)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(10)
        
        # Кнопки управления
        buttons_layout = QHBoxLayout()
        buttons_layout.setSpacing(10)
        
        refresh_btn = QPushButton("⟳ Обновить")
        refresh_btn.clicked.connect(self.load_current_reference)
        refresh_btn.setObjectName("compactButton")
        refresh_btn.setSizePolicy(QSizePolicy.Minimum, QSizePolicy.Preferred)
        buttons_layout.addWidget(refresh_btn)
        
        # Кнопка сохранения для справочников конфигурации (будет показываться только для них)
        self.save_config_btn = QPushButton("💾 Сохранить")
        self.save_config_btn.clicked.connect(self.save_config_reference)
        self.save_config_btn.setVisible(False)
        self.save_config_btn.setObjectName("compactButton")
        self.save_config_btn.setSizePolicy(QSizePolicy.Minimum, QSizePolicy.Preferred)
        buttons_layout.addWidget(self.save_config_btn)
        
        buttons_layout.addStretch()
        layout.addLayout(buttons_layout)
          
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Введите текст для поиска...")
        self.search_input.textChanged.connect(self.on_search_text_changed)
        buttons_layout.addWidget(self.search_input)
        
        # Таблица
        self.view_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.view_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.view_table.horizontalHeader().setStretchLastSection(True)
        # Включаем сортировку по столбцам
        self.view_table.setSortingEnabled(True)
        # Добавляем обработчик двойного клика для открытия детальной информации
        self.view_table.itemDoubleClicked.connect(self.on_row_double_clicked)
        # Контекстное меню для заголовков таблицы
        header = self.view_table.horizontalHeader()
        header.setContextMenuPolicy(Qt.CustomContextMenu)
        header.customContextMenuRequested.connect(self.show_table_header_context_menu)
        layout.addWidget(self.view_table)
        
        # Панель пагинации (контейнерный виджет)
        self.pagination_widget = QWidget()
        self.pagination_widget.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        pagination_layout = QHBoxLayout(self.pagination_widget)
        pagination_layout.setSpacing(3)
        pagination_layout.setContentsMargins(5, 5, 5, 5)
        
        # Растягиваем панель по всей ширине
        pagination_layout.addStretch()
        
        # Кнопка "Первая страница" - только иконка
        self.first_page_btn = QPushButton("◀◀")
        self.first_page_btn.setToolTip("Первая страница")
        self.first_page_btn.setObjectName("paginationButton")
        self.first_page_btn.clicked.connect(self.go_to_first_page)
        self.first_page_btn.setEnabled(False)
        pagination_layout.addWidget(self.first_page_btn)
        
        # Кнопка "Предыдущая страница" - только иконка
        self.prev_page_btn = QPushButton("◀")
        self.prev_page_btn.setToolTip("Предыдущая страница")
        self.prev_page_btn.setObjectName("paginationButton")
        self.prev_page_btn.clicked.connect(self.go_to_prev_page)
        self.prev_page_btn.setEnabled(False)
        pagination_layout.addWidget(self.prev_page_btn)
        
        # Информация о странице - компактная версия
        self.page_info_label = QLabel("1 / 1")
        self.page_info_label.setMinimumWidth(50)
        pagination_layout.addWidget(self.page_info_label)
        
        # Поле ввода номера страницы - компактная версия
        page_input_label = QLabel("Стр:")
        page_input_label.setMinimumWidth(30)
        pagination_layout.addWidget(page_input_label)
        self.page_input = QSpinBox()
        self.page_input.setMinimum(1)
        self.page_input.setMaximum(1)
        self.page_input.setValue(1)
        self.page_input.setMinimumWidth(60)
        self.page_input.setMaximumWidth(60)
        self.page_input.valueChanged.connect(self.on_page_input_changed)
        pagination_layout.addWidget(self.page_input)
        
        # Размер страницы - компактная версия
        page_size_label = QLabel("На стр:")
        page_size_label.setMinimumWidth(45)
        pagination_layout.addWidget(page_size_label)
        self.page_size_input = QSpinBox()
        self.page_size_input.setMinimum(100)
        self.page_size_input.setMaximum(10000)
        self.page_size_input.setSingleStep(100)
        self.page_size_input.setValue(1000)
        self.page_size_input.setMinimumWidth(80)
        self.page_size_input.setMaximumWidth(80)
        self.page_size_input.valueChanged.connect(self.on_page_size_changed)
        pagination_layout.addWidget(self.page_size_input)
        
        # Кнопка "Следующая страница" - только иконка
        self.next_page_btn = QPushButton("▶")
        self.next_page_btn.setToolTip("Следующая страница")
        self.next_page_btn.setObjectName("paginationButton")
        self.next_page_btn.clicked.connect(self.go_to_next_page)
        self.next_page_btn.setEnabled(False)
        pagination_layout.addWidget(self.next_page_btn)
        
        # Кнопка "Последняя страница" - только иконка
        self.last_page_btn = QPushButton("▶▶")
        self.last_page_btn.setToolTip("Последняя страница")
        self.last_page_btn.setObjectName("paginationButton")
        self.last_page_btn.clicked.connect(self.go_to_last_page)
        self.last_page_btn.setEnabled(False)
        pagination_layout.addWidget(self.last_page_btn)
        
        pagination_layout.addStretch()
        layout.addWidget(self.pagination_widget)
        
        # Изначально скрываем пагинацию
        self.pagination_widget.setVisible(False)
        
        # Статус
        self.status_label = QLabel("Выберите справочник для просмотра")
        self.status_label.setWordWrap(True)
        layout.addWidget(self.status_label)
    
    def setup_load_tab(self):
        """Настройка вкладки загрузки"""
        layout = QVBoxLayout(self.load_tab)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(15)
        
        # Название справочника (для доходов и источников)
        name_layout = QHBoxLayout()
        name_layout.setSpacing(10)
        
        name_label = QLabel("Название справочника:")
        name_label.setMinimumWidth(150)
        name_layout.addWidget(name_label)
        
        self.ref_name_edit = QLineEdit()
        name_layout.addWidget(self.ref_name_edit)
        
        layout.addLayout(name_layout)
        
        # Выбор файла
        file_layout = QHBoxLayout()
        file_layout.setSpacing(10)
        
        file_label = QLabel("Файл Excel:")
        file_label.setMinimumWidth(150)
        file_layout.addWidget(file_label)
        
        self.file_path_edit = QLineEdit()
        self.file_path_edit.setReadOnly(True)
        file_layout.addWidget(self.file_path_edit)
        
        browse_btn = QPushButton("📁 Обзор...")
        browse_btn.clicked.connect(self.browse_excel_file)
        browse_btn.setSizePolicy(QSizePolicy.Minimum, QSizePolicy.Preferred)
        file_layout.addWidget(browse_btn)
        
        layout.addLayout(file_layout)
        
        # Информация о необходимых колонках (для доходов и источников)
        self.info_label = QLabel()
        self.info_label.setWordWrap(True)
        self.info_label.setVisible(False)
        layout.addWidget(self.info_label)
        
        # Кнопка загрузки
        self.load_btn = QPushButton("⬆ Загрузить справочник")
        self.load_btn.clicked.connect(self.load_reference_from_excel)
        self.load_btn.setEnabled(False)
        self.load_btn.setSizePolicy(QSizePolicy.Minimum, QSizePolicy.Preferred)
        layout.addWidget(self.load_btn)
        
        # Статус загрузки
        self.load_status_label = QLabel("Выберите файл Excel для загрузки")
        self.load_status_label.setWordWrap(True)
        layout.addWidget(self.load_status_label)
        
        layout.addStretch()
    
    def load_reference_list(self):
        """Загрузка списка справочников"""
        self.references_list.clear()
        
        # Разделяем справочники на группы для лучшей организации
        regular_refs = []
        config_refs = []
        online_refs = []
        view_refs = []
        
        for ref_name, ref_data in self.REFERENCE_TYPES.items():
            if ref_data.get('is_separator', False):
                continue  # Пропускаем разделители
            elif ref_data.get('is_config', False):
                config_refs.append(ref_name)
            elif ref_data.get('is_online', False):
                online_refs.append(ref_name)
            elif ref_data.get('is_view', False):
                view_refs.append(ref_name)
            else:
                regular_refs.append(ref_name)
        
        # Добавляем обычные справочники
        for ref_name in regular_refs:
            item = QListWidgetItem(ref_name)
            self.references_list.addItem(item)
        
        # Добавляем разделитель для онлайн справочников, если они есть
        if online_refs:
            separator = QListWidgetItem("─── Онлайн справочники ───")
            separator.setFlags(Qt.NoItemFlags)  # Делаем разделитель невыбираемым
            separator.setForeground(Qt.gray)
            self.references_list.addItem(separator)
            
            # Добавляем онлайн справочники
            for ref_name in online_refs:
                item = QListWidgetItem(ref_name)
                self.references_list.addItem(item)
        
        # Добавляем разделитель для объединенных представлений, если они есть
        if view_refs:
            separator = QListWidgetItem("─── Объединенные представления ───")
            separator.setFlags(Qt.NoItemFlags)  # Делаем разделитель невыбираемым
            separator.setForeground(Qt.gray)
            self.references_list.addItem(separator)
            
            # Добавляем представления
            for ref_name in view_refs:
                item = QListWidgetItem(ref_name)
                self.references_list.addItem(item)
        
        # Добавляем разделитель, если есть справочники конфигурации
        if config_refs:
            separator = QListWidgetItem("─── Справочники конфигурации ───")
            separator.setFlags(Qt.NoItemFlags)  # Делаем разделитель невыбираемым
            separator.setForeground(Qt.gray)
            self.references_list.addItem(separator)
            
            # Добавляем справочники конфигурации
            for ref_name in config_refs:
                item = QListWidgetItem(ref_name)
                self.references_list.addItem(item)
        
        # Выбираем первый справочник по умолчанию
        if self.references_list.count() > 0:
            # Пропускаем разделители
            for i in range(self.references_list.count()):
                item = self.references_list.item(i)
                if item and item.flags() != Qt.NoItemFlags:
                    self.references_list.setCurrentRow(i)
                    self.on_reference_item_clicked(item)
                    break
    
    def on_reference_item_clicked(self, item: QListWidgetItem):
        """Обработчик клика по справочнику в списке"""
        if not item:
            return
        
        # Пропускаем разделители
        if item.flags() == Qt.NoItemFlags:
            return
        
        # Сохраняем видимость столбцов текущей таблицы до смены справочника (иначе сохранится под ключом новой таблицы)
        self._save_column_visibility()
        
        reference_name = item.text()
        self.current_reference_type = self.REFERENCE_TYPES.get(reference_name)
        
        if not self.current_reference_type:
            return
        
        # Обновляем заголовок
        if hasattr(self, 'reference_title_label'):
            self.reference_title_label.setText(f"📋 Справочник: {reference_name}")
        
        # Показываем/скрываем кнопку сохранения и вкладку загрузки для справочников конфигурации
        is_config = self.current_reference_type.get('is_config', False)
        if self.save_config_btn:
            self.save_config_btn.setVisible(is_config)
            # Для справочников конфигурации делаем таблицу редактируемой
            if is_config:
                self.view_table.setEditTriggers(QTableWidget.AllEditTriggers)
            else:
                self.view_table.setEditTriggers(QTableWidget.NoEditTriggers)
        
        # Настраиваем вкладку загрузки
        if self.tabs:
            load_tab_index = self.tabs.indexOf(self.load_tab)
            if load_tab_index >= 0:
                is_online = self.current_reference_type.get('is_online', False)
                is_view = self.current_reference_type.get('is_view', False)
                load_method = self.current_reference_type.get('load_method')
                
                # Показываем вкладку загрузки для справочников с методом загрузки
                # (включая доходы и источники)
                # Представления и онлайн справочники - только для просмотра
                if is_config or is_online or is_view:
                    self.tabs.setTabVisible(load_tab_index, False)
                elif load_method:
                    self.tabs.setTabVisible(load_tab_index, True)
                    # Обновляем информацию о необходимых колонках для доходов и источников
                    load_type = self.current_reference_type.get('load_type')
                    if load_type == 'доходы':
                        self.info_label.setText(
                            "Необходимые колонки:\n"
                            "• код_классификации_ДБ\n"
                            "• уровень_кода"
                        )
                        self.info_label.setVisible(True)
                        # Устанавливаем название по умолчанию
                        if not self.ref_name_edit.text():
                            self.ref_name_edit.setText('Классификация доходов бюджетов')
                    elif load_type == 'источники':
                        self.info_label.setText(
                            "Необходимые колонки:\n"
                            "• код_классификации_ИФДБ\n"
                            "• уровень_кода"
                        )
                        self.info_label.setVisible(True)
                        # Устанавливаем название по умолчанию
                        if not self.ref_name_edit.text():
                            self.ref_name_edit.setText('Классификация источников финансирования дефицитов')
                    else:
                        self.info_label.setVisible(False)
                else:
                    self.tabs.setTabVisible(load_tab_index, False)
                    # Очищаем поля при скрытии вкладки
                    self.file_path_edit.clear()
                    self.ref_name_edit.clear()
                    self.info_label.setVisible(False)
        
        # Сбрасываем пагинацию и поиск при смене справочника
        self.current_page = 1
        self.page_size = 1000
        self.page_size_input.setValue(1000)
        if hasattr(self, 'search_input'):
            self.search_input.clear()
        # Обновляем дату фильтрации из поля выбора
        self._update_filter_date()
        self.load_current_reference(skip_save_visibility=True)
    
    def _update_filter_date(self):
        """Обновляет filter_date из поля выбора даты"""
        if hasattr(self, 'date_filter_checkbox') and self.date_filter_checkbox.isChecked():
            if hasattr(self, 'date_filter'):
                qdate = self.date_filter.date()
                self.filter_date = qdate.toString("yyyy-MM-dd")
            else:
                self.filter_date = datetime.now().strftime('%Y-%m-%d')
        else:
            self.filter_date = None  # Фильтрация отключена
    
    def on_date_filter_toggled(self, state: int):
        """Обработчик включения/выключения фильтрации по дате"""
        # Включаем/выключаем поле выбора даты
        if hasattr(self, 'date_filter'):
            self.date_filter.setEnabled(state == Qt.Checked)
        
        # Обновляем дату фильтрации
        self._update_filter_date()
        # Сбрасываем на первую страницу при изменении фильтрации
        self.current_page = 1
        # Перезагружаем данные
        self.load_current_reference()
    
    def on_date_filter_changed(self, date: QDate):
        """Обработчик изменения даты фильтрации"""
        self._update_filter_date()
        # Сбрасываем на первую страницу при изменении даты
        self.current_page = 1
        # Перезагружаем данные с учетом новой даты
        self.load_current_reference()
    
    # --- Вспомогательные методы для работы с данными ---
    
    def _execute_query(self, conn, query: str, params: list = None) -> pd.DataFrame:
        """Выполняет SQL запрос с параметрами или без"""
        if params:
            return pd.read_sql_query(query, conn, params=params)
        return pd.read_sql_query(query, conn)
    
    def _save_column_visibility(self) -> dict:
        """Сохраняет текущую видимость столбцов и возвращает словарь видимости"""
        column_visibility = {}
        if self.view_table.columnCount() > 0 and self.view_table.rowCount() > 0:
            for col in range(self.view_table.columnCount()):
                header_item = self.view_table.horizontalHeaderItem(col)
                if header_item:
                    column_name = header_item.text()
                    column_visibility[column_name] = not self.view_table.isColumnHidden(col)
            
            # Сохраняем видимость столбцов в конфигурацию
            if self.current_reference_type and column_visibility:
                table_name = self.current_reference_type.get('table')
                if table_name:
                    config_key = f"references_table_columns:{table_name}"
                    self.db_manager.save_config(config_key, column_visibility)
        return column_visibility
    
    def _restore_column_visibility(self, column_visibility: dict = None):
        """Восстанавливает видимость столбцов из сохраненной конфигурации для текущей таблицы
        
        Args:
            column_visibility: Опциональный словарь текущей видимости (не используется, оставлен для совместимости)
        """
        saved_column_visibility = {}
        if self.current_reference_type:
            table_name = self.current_reference_type.get('table')
            if table_name:
                config_key = f"references_table_columns:{table_name}"
                saved_column_visibility = self.db_manager.load_config(config_key, {})
        
        # Используем только сохраненную конфигурацию для текущей таблицы
        # Не используем column_visibility из предыдущей таблицы, чтобы избежать применения настроек к другим таблицам
        if saved_column_visibility:
            for col in range(self.view_table.columnCount()):
                header_item = self.view_table.horizontalHeaderItem(col)
                if header_item:
                    column_name = header_item.text()
                    if column_name in saved_column_visibility:
                        self.view_table.setColumnHidden(col, not saved_column_visibility[column_name])
    
    def _fill_table_from_dataframe(self, df: pd.DataFrame, available_columns: list, 
                                    column_visibility: dict, offset: int, total_records: int,
                                    use_date_filter: bool = False):
        """Заполняет таблицу из DataFrame с восстановлением видимости столбцов и обновлением статуса"""
        # Заполняем таблицу
        self.view_table.setRowCount(len(df))
        self.view_table.setColumnCount(len(available_columns))
        self.view_table.setHorizontalHeaderLabels(available_columns)
        
        for row_idx, (_, row) in enumerate(df.iterrows()):
            for col_idx, col_name in enumerate(available_columns):
                value = row.get(col_name, '')
                item = QTableWidgetItem(str(value) if pd.notna(value) else '')
                self.view_table.setItem(row_idx, col_idx, item)
        
        # Настраиваем ширину столбцов
        self.adjust_reference_columns_width()
        
        # Восстанавливаем видимость столбцов
        self._restore_column_visibility(column_visibility)
        
        # Включаем сортировку
        self.view_table.setSortingEnabled(True)
        
        # Показываем пагинацию
        self._show_pagination()
        self._update_pagination_info()
        self._update_pagination_buttons()
        
        # Обновляем статус
        self._update_status_label(offset, len(df), total_records, use_date_filter)
    
    def _update_status_label(self, offset: int, df_length: int, total_records: int, 
                              use_date_filter: bool = False):
        """Обновляет статусную строку с информацией о записях"""
        start_record = offset + 1
        end_record = min(offset + df_length, total_records)
        date_info = f" (дата: {self.filter_date})" if use_date_filter and self.filter_date else ""
        
        if self.search_text:
            self.status_label.setText(
                f"Найдено записей: {start_record}-{end_record} из {total_records}{date_info} "
                f"(Страница {self.current_page} из {self.total_pages})"
            )
        else:
            self.status_label.setText(
                f"Показано записей: {start_record}-{end_record} из {total_records}{date_info} "
                f"(Страница {self.current_page} из {self.total_pages})"
            )
    
    def _build_query_with_npa_join(self, cursor, table_name: str, available_columns: list,
                                    search_where: str, search_params: list, limit: int, offset: int) -> tuple:
        """Строит SQL запрос для онлайн справочника с учетом NPA JOIN
        
        Returns:
            tuple: (query, params) - SQL запрос и параметры
        """
        # Проверяем существование таблицы npa
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='npa'")
        if cursor.fetchone():
            # Формируем SELECT с JOIN
            main_cols = [f"{table_name}.{col}" for col in available_columns]
            # Добавляем колонки из npa
            cursor.execute("PRAGMA table_info(npa)")
            npa_columns = [row[1] for row in cursor.fetchall()]
            npa_display_cols = [col for col in npa_columns if col not in ['id']]
            npa_cols = [f"npa.{col} AS npa_{col}" for col in npa_display_cols]
            
            all_cols = main_cols + npa_cols
            query = f'''SELECT {", ".join(all_cols)} 
                      FROM {table_name} 
                      LEFT JOIN npa ON {table_name}.npa_id = npa.id'''
            if search_where:
                query += f" {search_where}"
            query += f" LIMIT {limit} OFFSET {offset}"
            return query, search_params
        else:
            # Если таблицы npa нет, просто исключаем npa_id
            query = f'SELECT {", ".join(available_columns)} FROM {table_name}'
            if search_where:
                query += f" {search_where}"
            query += f" LIMIT {limit} OFFSET {offset}"
            return query, search_params
    
    def _build_select_query(self, table_name: str, columns: list, search_where: str, 
                            search_params: list, limit: int, offset: int) -> tuple:
        """Строит простой SELECT запрос
        
        Returns:
            tuple: (query, params) - SQL запрос и параметры
        """
        query = f'SELECT {", ".join(columns)} FROM {table_name}'
        if search_where:
            query += f" {search_where}"
        query += f" LIMIT {limit} OFFSET {offset}"
        return query, search_params
    
    def load_current_reference(self, page: int = None, skip_save_visibility: bool = False):
        """Загрузка текущего справочника в таблицу с поддержкой пагинации.
        
        Args:
            page: Номер страницы (опционально).
            skip_save_visibility: Если True, не сохранять видимость столбцов перед загрузкой
                (используется при переключении справочника — сохранение уже сделано в on_reference_item_clicked).
        """
        if not self.current_reference_type:
            return
        
        # Если указана страница, используем её, иначе текущую
        if page is not None:
            self.current_page = page
        
        # Сохраняем видимость столбцов текущей таблицы перед очисткой (при обновлении/пагинации; при переключении — уже сохранено)
        if not skip_save_visibility:
            self._save_column_visibility()
        
        # Очищаем таблицу перед загрузкой нового справочника
        self.view_table.clear()
        self.view_table.setRowCount(0)
        self.view_table.setColumnCount(0)
        # Отключаем сортировку при очистке, чтобы избежать ошибок
        self.view_table.setSortingEnabled(False)
        
        # Загрузка справочников конфигурации (без пагинации)
        if self.current_reference_type.get('is_config', False):
            load_func_name = self.current_reference_type.get('load_func')
            if load_func_name:
                load_func = getattr(self, load_func_name, None)
                if load_func:
                    load_func()
            # Скрываем пагинацию для справочников конфигурации
            self._hide_pagination()
            return
        
        # Загрузка обычных справочников из БД
        try:
            table_name = self.current_reference_type['table']
            columns = self.current_reference_type.get('columns', [])
            display_columns = self.current_reference_type.get('display_columns', [])
            conn = sqlite3.connect(self.db_manager.db_path)
            cursor = conn.cursor()
            # Проверяем существование таблицы или представления (без учёта регистра)
            is_view = self.current_reference_type.get('is_view', False)
            if is_view:
                cursor.execute("SELECT name FROM sqlite_master WHERE type='view' AND LOWER(name)=LOWER(?)", (table_name,))
            else:
                cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND LOWER(name)=LOWER(?)", (table_name,))
            row = cursor.fetchone()
            if not row:
                entity_type = "представление" if is_view else "таблица"
                self.status_label.setText(f"{entity_type.capitalize()} {table_name} не найдена в БД")
                conn.close()
                return
            resolved_table_name = row[0]
            # Получаем список колонок таблицы по точному имени из БД
            cursor.execute(f"PRAGMA table_info({resolved_table_name})")
            existing_columns = [row[1] for row in cursor.fetchall()]
            
            # Проверяем, есть ли поля startdate и enddate для фильтрации по дате
            has_date_fields = 'startdate' in existing_columns and 'enddate' in existing_columns
            # Для справочников с датами (в т.ч. v_budgetclastypeinc_merged) обязательно фильтруем по дате: по умолчанию из конфига или текущая
            if has_date_fields and table_name == 'v_budgetclastypeinc_merged' and self.filter_date is None:
                self.filter_date = self.db_manager.load_config("reference_filter_date") or datetime.now().strftime('%Y-%m-%d')
                if hasattr(self, 'date_filter') and hasattr(self, 'date_filter_checkbox'):
                    try:
                        parts = self.filter_date.split('-')
                        if len(parts) == 3:
                            from PyQt5.QtCore import QDate
                            self.date_filter.setDate(QDate(int(parts[0]), int(parts[1]), int(parts[2])))
                            self.date_filter_checkbox.setChecked(True)
                            self.date_filter.setEnabled(True)
                    except Exception:
                        pass
            use_date_filter = has_date_fields and self.filter_date is not None
            # Обновляем дату фильтрации из поля выбора
            self._update_filter_date()
            
            # Для таблиц и VIEW с полями дат используем get_filtered_view для фильтрации и дедупликации
            if use_date_filter:
                # Определяем, нужен ли JOIN к npa
                has_npa = self.current_reference_type.get('has_npa', False)
                has_npa_column = 'npa_id' in existing_columns
                join_npa = has_npa and has_npa_column
                
                # Загружаем отфильтрованные данные через get_filtered_view (работает и для таблиц, и для VIEW)
                df_filtered = get_filtered_view(conn, table_name, self.filter_date, join_npa=join_npa)
                # Для справочников с display_columns оставляем только нужные столбцы и заголовки (код, наименование, уровень)
                display_columns = self.current_reference_type.get('display_columns')
                if display_columns and len(df_filtered) > 0:
                    col_map = {}
                    for expr in display_columns:
                        if ' AS ' in expr:
                            left, _, right = expr.partition(' AS ')
                            orig = left.strip()
                            col_map[orig] = right.strip()
                    if col_map:
                        keep = [c for c in col_map if c in df_filtered.columns]
                        if keep:
                            df_filtered = df_filtered[keep].copy()
                            df_filtered.columns = [col_map.get(c, c) for c in keep]
                available_columns = list(df_filtered.columns)
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
                self.total_pages = max(1, (self.total_records + self.page_size - 1) // self.page_size)
                
                # Корректируем текущую страницу
                if self.current_page > self.total_pages:
                    self.current_page = self.total_pages
                if self.current_page < 1:
                    self.current_page = 1
                
                offset = (self.current_page - 1) * self.page_size
                limit = self.page_size
                df = df_filtered.iloc[offset:offset + limit].copy()
                if not available_columns:
                    available_columns = list(df.columns)
                
                conn.close()
                
                # Заполняем таблицу используя вспомогательный метод
                # Не передаем column_visibility, чтобы не применять настройки из предыдущей таблицы
                self._fill_table_from_dataframe(df, available_columns, {}, 
                                                offset, self.total_records, use_date_filter=True)
                
                return
            
            # Формируем WHERE условие для поиска и фильтрации по дате
            def build_search_where(base_columns, add_date_filter=False, table_prefix=None):
                """Строит WHERE условие для поиска по указанным колонкам и фильтрации по дате
                
                Args:
                    base_columns: Список колонок для поиска
                    add_date_filter: Добавить фильтр по дате
                    table_prefix: Префикс таблицы для колонок (например, 'budgetclascostsmo') для избежания неоднозначности при JOIN
                """
                conditions = []
                params = []
                
                # Префикс для колонок (если указан)
                prefix = f"{table_prefix}." if table_prefix else ""
                
                # Добавляем фильтрацию по дате
                if add_date_filter and self.filter_date:
                    conditions.append(f"({prefix}startdate IS NULL OR date({prefix}startdate) <= date(?))")
                    conditions.append(f"({prefix}enddate IS NULL OR {prefix}enddate = '' OR date({prefix}enddate) >= date(?))")
                    params.extend([self.filter_date, self.filter_date])
                
                # Добавляем поиск
                if self.search_text:
                    text_columns = [col for col in base_columns 
                                  if col not in ['id', 'guid', 'created_at', 'loaddate', 'startdate', 'enddate']]
                    if text_columns:
                        search_pattern = f"%{self.search_text}%"
                        search_conditions = " OR ".join([f"{prefix}{col} LIKE ?" for col in text_columns])
                        conditions.append(f"({search_conditions})")
                        params.extend([search_pattern] * len(text_columns))
                
                if conditions:
                    return "WHERE " + " AND ".join(conditions), params
                return "", []
            
            # Сначала получаем общее количество записей для пагинации (с учетом поиска и фильтрации по дате)
            # Используем точное имя таблицы из БД для запросов
            effective_table = resolved_table_name
            has_npa = self.current_reference_type.get('has_npa', False)
            has_npa_column = 'npa_id' in existing_columns
            use_table_prefix = has_npa and has_npa_column
            
            if use_table_prefix:
                count_query = f'SELECT COUNT(*) FROM {effective_table} LEFT JOIN npa ON {effective_table}.npa_id = npa.id'
                search_where, search_params = build_search_where(existing_columns, add_date_filter=use_date_filter, table_prefix=effective_table)
            else:
                count_query = f'SELECT COUNT(*) FROM {effective_table}'
                search_where, search_params = build_search_where(existing_columns, add_date_filter=use_date_filter)
            
            if search_where:
                count_query += f" {search_where}"
                cursor.execute(count_query, search_params)
            else:
                cursor.execute(count_query)
            
            self.total_records = cursor.fetchone()[0]
            self.total_pages = max(1, (self.total_records + self.page_size - 1) // self.page_size)
            
            # Корректируем текущую страницу, если она выходит за пределы
            if self.current_page > self.total_pages:
                self.current_page = self.total_pages
            if self.current_page < 1:
                self.current_page = 1
            
            # Вычисляем OFFSET и LIMIT для пагинации
            offset = (self.current_page - 1) * self.page_size
            limit = self.page_size
            
            # Формируем запрос с учетом поиска и фильтрации по дате (таблица с display_columns и опционально search_columns)
            if display_columns:
                base_cols = self.current_reference_type.get('search_columns') or ['code', 'name', 'level', 'doc']
                search_where, search_params = build_search_where(base_cols, add_date_filter=use_date_filter)
                query = f'SELECT {", ".join(display_columns)} FROM {effective_table}'
                if search_where:
                    query += f" {search_where}"
                query += f" LIMIT {limit} OFFSET {offset}"
                df = self._execute_query(conn, query, search_params)
                available_columns = list(df.columns)
            else:
                
                # Определяем доступные колонки и строим запрос
                # has_npa и has_npa_column уже определены выше для count_query
                
                if self.current_reference_type.get('is_online', False):
                    excluded_cols = ['id', 'guid', 'created_at', 'loaddate']
                    if has_npa and has_npa_column:
                        excluded_cols.append('npa_id')
                    available_columns = [col for col in existing_columns if col not in excluded_cols]
                    
                    if has_npa and has_npa_column:
                        search_where, search_params = build_search_where(available_columns, add_date_filter=use_date_filter, table_prefix=effective_table)
                        query, params = self._build_query_with_npa_join(
                            cursor, effective_table, available_columns, search_where, search_params, limit, offset
                        )
                    else:
                        search_where, search_params = build_search_where(available_columns, add_date_filter=use_date_filter)
                        query, params = self._build_select_query(
                            effective_table, available_columns, search_where, search_params, limit, offset
                        )
                    df = self._execute_query(conn, query, params)
                elif columns:
                    available_columns = [col for col in columns if col in existing_columns]
                    if not available_columns:
                        available_columns = existing_columns[:10]
                    search_where, search_params = build_search_where(available_columns, add_date_filter=use_date_filter)
                    query, params = self._build_select_query(
                        effective_table, available_columns, search_where, search_params, limit, offset
                    )
                    df = self._execute_query(conn, query, params)
                else:
                    available_columns = [col for col in existing_columns 
                                       if col not in ['id', 'guid', 'npa_id', 'created_at', 'loaddate']]
                    if not available_columns:
                        available_columns = existing_columns[:10]
                    search_where, search_params = build_search_where(available_columns)
                    query, params = self._build_select_query(
                        effective_table, available_columns, search_where, search_params, limit, offset
                    )
                    df = self._execute_query(conn, query, params)
                
                available_columns = list(df.columns)  # Обновляем список колонок после запроса
            
            conn.close()
            
            # Заполняем таблицу используя вспомогательный метод
            # Не передаем column_visibility, чтобы не применять настройки из предыдущей таблицы
            self._fill_table_from_dataframe(df, available_columns, {}, 
                                            offset, self.total_records, use_date_filter)
            
        except Exception as e:
            logger.error(f"Ошибка загрузки справочника: {e}", exc_info=True)
            QMessageBox.warning(self, "Ошибка", f"Не удалось загрузить справочник:\n{str(e)}")
            self.status_label.setText("Ошибка загрузки")
            self._hide_pagination()
    
    def adjust_reference_columns_width(self):
        """Настройка ширины столбцов справочника
        
        Логика:
        1. Устанавливается фиксированный размер для поля name (наименование)
        2. Если есть поле code (код), то для оставшихся столбцов определяем 
           максимальное количество символов и задаем ширину, только если эти 
           поля не содержат много текста (максимальная длина <= 50 символов)
        """
        if self.view_table.columnCount() == 0:
            return
        
        header = self.view_table.horizontalHeader()
        
        # Фиксированная ширина для столбца name (в пикселях)
        NAME_COLUMN_WIDTH = 280
        
        # Фиксированная ширина для столбца code (в пикселях)
        CODE_COLUMN_WIDTH = 120
        
        # Порог для определения "много текста" (в символах)
        MAX_TEXT_LENGTH_THRESHOLD = 50
        
        # Коэффициент для расчета ширины столбца на основе количества символов
        # Примерно 7-8 пикселей на символ в зависимости от шрифта
        CHAR_WIDTH_MULTIPLIER = 8
        
        # Минимальная и максимальная ширина для столбцов с коротким текстом
        MIN_COLUMN_WIDTH = 75
        MAX_COLUMN_WIDTH = 350
        
        # Находим индексы столбцов name и code
        name_col_idx = None
        code_col_idx = None
        
        for col in range(self.view_table.columnCount()):
            header_text = self.view_table.horizontalHeaderItem(col)
            if header_text:
                col_name = header_text.text().lower()
                # Проверяем различные варианты названий
                if col_name in ['name', 'наименование', 'название']:
                    name_col_idx = col
                elif col_name in ['code', 'код']:
                    code_col_idx = col
        
        # Устанавливаем фиксированную ширину для столбца name
        if name_col_idx is not None:
            header.setSectionResizeMode(name_col_idx, QHeaderView.Fixed)
            self.view_table.setColumnWidth(name_col_idx, NAME_COLUMN_WIDTH)
        
        # Устанавливаем фиксированную ширину для столбца code
        if code_col_idx is not None:
            header.setSectionResizeMode(code_col_idx, QHeaderView.Fixed)
            self.view_table.setColumnWidth(code_col_idx, CODE_COLUMN_WIDTH)
        
        # Для остальных столбцов определяем ширину на основе максимального количества символов
        # только если есть столбец code (значит это справочник с кодом)
        if code_col_idx is not None:
            for col in range(self.view_table.columnCount()):
                # Пропускаем столбцы name и code
                if col == name_col_idx or col == code_col_idx:
                    continue
                
                # Находим максимальную длину текста в столбце
                max_length = 0
                for row in range(self.view_table.rowCount()):
                    item = self.view_table.item(row, col)
                    if item:
                        text = item.text()
                        if text:
                            max_length = max(max_length, len(text))
                
                # Также учитываем длину заголовка
                header_item = self.view_table.horizontalHeaderItem(col)
                if header_item:
                    header_length = len(header_item.text())
                    max_length = max(max_length, header_length)
                
                # Если максимальная длина не превышает порог, устанавливаем ширину
                if max_length <= MAX_TEXT_LENGTH_THRESHOLD:
                    # Вычисляем ширину на основе максимальной длины
                    calculated_width = max_length * CHAR_WIDTH_MULTIPLIER
                    # Добавляем небольшой отступ
                    calculated_width += 20
                    # Ограничиваем минимальной и максимальной шириной
                    column_width = max(MIN_COLUMN_WIDTH, min(calculated_width, MAX_COLUMN_WIDTH))
                    
                    header.setSectionResizeMode(col, QHeaderView.Fixed)
                    self.view_table.setColumnWidth(col, column_width)
                else:
                    # Для столбцов с длинным текстом используем режим ResizeToContents
                    # или оставляем Interactive для ручной настройки
                    header.setSectionResizeMode(col, QHeaderView.ResizeToContents)
        else:
            # Если нет столбца code, для всех столбцов (кроме name) используем ResizeToContents
            for col in range(self.view_table.columnCount()):
                if col != name_col_idx:
                    header.setSectionResizeMode(col, QHeaderView.ResizeToContents)
    
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
            
            # Автоматически заполняем название справочника, если оно пустое
            if not self.ref_name_edit.text():
                file_name = Path(file_path).stem
                self.ref_name_edit.setText(file_name)
    
    def load_reference_from_excel(self):
        """Загрузка справочника из Excel"""
        if not self.current_reference_type:
            QMessageBox.warning(self, "Ошибка", "Выберите справочник")
            return
        
        file_path = self.file_path_edit.text()
        if not file_path:
            QMessageBox.warning(self, "Ошибка", "Выберите файл Excel")
            return
        
        # Проверяем, поддерживается ли загрузка для этого справочника
        load_method_name = self.current_reference_type.get('load_method')
        if not load_method_name:
            QMessageBox.warning(
                self,
                "Ошибка",
                "Для этого справочника загрузка из Excel не поддерживается."
            )
            return
        
        try:
            self.load_btn.setEnabled(False)
            self.load_status_label.setText("Загрузка...")
            
            # Специальная обработка для доходов и источников через ReferenceController
            if load_method_name == 'load_income_sources_reference':
                if not self.reference_controller:
                    QMessageBox.warning(
                        self,
                        "Ошибка",
                        "Не удалось получить доступ к контроллеру справочников."
                    )
                    return
                
                ref_type = self.current_reference_type.get('load_type')
                
                # Получаем название справочника из поля ввода или используем значение по умолчанию
                ref_name = self.ref_name_edit.text().strip()
                if not ref_name:
                    if ref_type == 'доходы':
                        ref_name = 'Классификация доходов бюджетов'
                    elif ref_type == 'источники':
                        ref_name = 'Классификация источников финансирования дефицитов'
                    else:
                        ref_name = Path(file_path).stem
                
                success = self.reference_controller.load_reference_file(file_path, ref_type, ref_name)
                
                if success:
                    # Подсчитываем количество записей
                    if ref_type == 'доходы':
                        df = self.db_manager.load_income_reference_df()
                    else:
                        df = self.db_manager.load_sources_reference_df()
                    count = len(df) if df is not None and not df.empty else 0
                    
                    self.load_status_label.setText(f"Загружено записей: {count}")
                    QMessageBox.information(
                        self,
                        "Успех",
                        f"Справочник успешно загружен.\nЗагружено записей: {count}"
                    )
                    
                    # Обновляем таблицу просмотра
                    self.load_current_reference()
                else:
                    self.load_status_label.setText("Ошибка загрузки")
                    QMessageBox.warning(self, "Ошибка", "Не удалось загрузить справочник.")
            else:
                # Обычная загрузка через методы db_manager
                load_method = getattr(self.db_manager, load_method_name)
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
            QMessageBox.critical(self, "Ошибка", f"Ошибка загрузки справочника:\n{str(e)}")
            self.load_status_label.setText("Ошибка загрузки")
        finally:
            self.load_btn.setEnabled(True)
    
    # --- Методы для справочников конфигурации ---
    
    def _load_config_reference(self, data_list, columns, field_extractors):
        """Универсальный метод загрузки конфигурационных справочников
        
        Args:
            data_list: Список объектов для загрузки
            columns: Список названий колонок
            field_extractors: Список функций для извлечения значений из объекта
        """
        self.view_table.setColumnCount(len(columns))
        self.view_table.setHorizontalHeaderLabels(columns)
        self.view_table.setRowCount(len(data_list) + 1)  # +1 пустая строка для добавления
        
        for row_idx, item in enumerate(data_list):
            for col_idx, extractor in enumerate(field_extractors):
                value = extractor(item)
                self.view_table.setItem(row_idx, col_idx, QTableWidgetItem(str(value) if value is not None else ""))
        
        # Настраиваем ширину столбцов
        self.adjust_reference_columns_width()
        self.status_label.setText(f"Загружено записей: {len(data_list)}")
    
    def _load_years(self):
        """Загрузка справочника годов"""
        years = self.db_manager.load_years()
        columns = ['Год', 'Активен (1/0)']
        extractors = [
            lambda y: y.year,
            lambda y: "1" if y.is_active else "0"
        ]
        self._load_config_reference(years, columns, extractors)
    
    def _load_municipalities(self):
        """Загрузка справочника муниципальных образований"""
        municip = self.db_manager.load_municipalities()
        columns = ['Код', 'Наименование', 'Активен (1/0)']
        extractors = [
            lambda m: m.code or "",
            lambda m: m.name or "",
            lambda m: "1" if m.is_active else "0"
        ]
        self._load_config_reference(municip, columns, extractors)
    
    def _load_forms(self):
        """Загрузка справочника типов форм"""
        forms = self.db_manager.load_form_types_meta()
        columns = ['ID', 'Код формы', 'Наименование', 'Периодичность', 'Активен (1/0)']
        extractors = [
            lambda f: f.id if f.id is not None else "",
            lambda f: f.code or "",
            lambda f: f.name or "",
            lambda f: f.periodicity or "",
            lambda f: "1" if f.is_active else "0"
        ]
        self._load_config_reference(forms, columns, extractors)
    
    def _load_periods(self):
        """Загрузка справочника периодов"""
        periods = self.db_manager.load_periods()
        columns = ['ID', 'Код периода', 'Наименование', 'Порядок сортировки', 'Код формы (опц.)', 'Активен (1/0)']
        extractors = [
            lambda p: p.id if p.id is not None else "",
            lambda p: p.code or "",
            lambda p: p.name or "",
            lambda p: p.sort_order,
            lambda p: p.form_type_code or "",
            lambda p: "1" if p.is_active else "0"
        ]
        self._load_config_reference(periods, columns, extractors)
    
    def show_table_header_context_menu(self, position):
        """Показать контекстное меню для заголовков таблицы"""
        try:
            header = self.view_table.horizontalHeader()
            column = header.logicalIndexAt(position.x())
            
            menu = QMenu(self)
            action = menu.addAction("Выбрать столбцы...")
            action.triggered.connect(self.show_column_visibility_dialog)
            
            menu.exec_(header.mapToGlobal(position))
        except Exception as e:
            logger.error(f"Ошибка при показе контекстного меню заголовков таблицы: {e}", exc_info=True)
    
    def show_column_visibility_dialog(self):
        """Показать диалог выбора столбцов для таблицы"""
        try:
            dialog = ColumnVisibilityDialog(self.view_table, self)
            if dialog.exec_() == QDialog.Accepted:
                # Сохраняем видимость столбцов в конфигурацию после применения изменений
                if self.current_reference_type:
                    table_name = self.current_reference_type.get('table')
                    if table_name:
                        column_visibility = {}
                        for col in range(self.view_table.columnCount()):
                            header_item = self.view_table.horizontalHeaderItem(col)
                            if header_item:
                                column_name = header_item.text()
                                column_visibility[column_name] = not self.view_table.isColumnHidden(col)
                        
                        config_key = f"references_table_columns:{table_name}"
                        self.db_manager.save_config(config_key, column_visibility)
        except Exception as e:
            logger.error(f"Ошибка при открытии диалога выбора столбцов: {e}", exc_info=True)
    
    def on_row_double_clicked(self, item: QTableWidgetItem):
        """Обработчик двойного клика по строке таблицы"""
        if not item:
            return
        
        row = item.row()
        if row < 0:
            return
        
        # Получаем данные строки
        row_data = {}
        column_headers = []
        
        for col in range(self.view_table.columnCount()):
            header = self.view_table.horizontalHeaderItem(col)
            header_text = header.text() if header else f"Колонка {col}"
            column_headers.append(header_text)
            
            cell_item = self.view_table.item(row, col)
            cell_value = cell_item.text() if cell_item else ""
            row_data[header_text] = cell_value
        
        # Открываем диалог с детальной информацией
        dialog = ReferenceDetailDialog(row_data, column_headers, self)
        dialog.exec_()
    
    def save_config_reference(self):
        """Сохранение справочника конфигурации"""
        if not self.current_reference_type or not self.current_reference_type.get('is_config', False):
            return
        
        try:
            save_func_name = self.current_reference_type.get('save_func')
            if save_func_name and hasattr(self, save_func_name):
                save_func = getattr(self, save_func_name)
                save_func()
                QMessageBox.information(self, "Сохранено", "Справочник успешно сохранен.")
                # Перезагружаем данные
                self.load_current_reference()
        except Exception as e:
            logger.error(f"Ошибка сохранения справочника конфигурации: {e}", exc_info=True)
            QMessageBox.critical(self, "Ошибка", f"Ошибка сохранения справочника:\n{str(e)}")
    
    def _save_config_reference(self, model_class, required_col_idx, field_setters, save_method):
        """Универсальный метод сохранения конфигурационных справочников
        
        Args:
            model_class: Класс модели (YearRef, MunicipalityRef, etc.)
            required_col_idx: Индекс обязательной колонки для проверки
            field_setters: Список функций для установки значений полей модели
            save_method: Метод db_manager для сохранения (save_years_bulk, etc.)
        """
        items_list = []
        for row in range(self.view_table.rowCount()):
            required_item = self.view_table.item(row, required_col_idx)
            if not required_item or not required_item.text().strip():
                continue
            
            item = model_class()
            for setter in field_setters:
                setter(item, row)
            items_list.append(item)
        
        save_method(items_list)
    
    def _save_years(self):
        """Сохранение справочника годов"""
        def set_year_fields(item, row):
            year_item = self.view_table.item(row, 0)
            active_item = self.view_table.item(row, 1)
            try:
                item.year = int(year_item.text())
            except ValueError:
                raise ValueError(f"Неверное значение года в строке {row + 1}")
            item.is_active = (active_item.text().strip() == "1") if active_item else True
        
        self._save_config_reference(YearRef, 0, [set_year_fields], self.db_manager.save_years_bulk)
    
    def _save_municipalities(self):
        """Сохранение справочника муниципальных образований"""
        def set_municip_fields(item, row):
            code_item = self.view_table.item(row, 0)
            name_item = self.view_table.item(row, 1)
            active_item = self.view_table.item(row, 2)
            item.code = (code_item.text() if code_item else "").strip()
            item.name = name_item.text().strip()
            item.is_active = (active_item.text().strip() == "1") if active_item else True
        
        self._save_config_reference(MunicipalityRef, 1, [set_municip_fields], self.db_manager.save_municipalities_bulk)
    
    def _save_forms(self):
        """Сохранение справочника типов форм"""
        def set_form_fields(item, row):
            id_item = self.view_table.item(row, 0)
            code_item = self.view_table.item(row, 1)
            name_item = self.view_table.item(row, 2)
            periodicity_item = self.view_table.item(row, 3)
            active_item = self.view_table.item(row, 4)
            if id_item and id_item.text().strip():
                try:
                    item.id = int(id_item.text().strip())
                except ValueError:
                    item.id = None
            item.code = code_item.text().strip()
            item.name = (name_item.text() if name_item else item.code).strip()
            item.periodicity = (periodicity_item.text() if periodicity_item else "").strip()
            item.is_active = (active_item.text().strip() == "1") if active_item else True
        
        self._save_config_reference(FormTypeMeta, 1, [set_form_fields], self.db_manager.save_form_types_bulk)
    
    def _save_periods(self):
        """Сохранение справочника периодов"""
        def set_period_fields(item, row):
            id_item = self.view_table.item(row, 0)
            code_item = self.view_table.item(row, 1)
            name_item = self.view_table.item(row, 2)
            sort_item = self.view_table.item(row, 3)
            form_code_item = self.view_table.item(row, 4)
            active_item = self.view_table.item(row, 5)
            if id_item and id_item.text().strip():
                try:
                    item.id = int(id_item.text().strip())
                except ValueError:
                    item.id = None
            item.code = code_item.text().strip()
            item.name = (name_item.text() if name_item else item.code).strip()
            try:
                item.sort_order = int(sort_item.text()) if sort_item and sort_item.text().strip() else 0
            except ValueError:
                item.sort_order = 0
            item.form_type_code = (form_code_item.text() if form_code_item else "").strip()
            item.is_active = (active_item.text().strip() == "1") if active_item else True
        
        self._save_config_reference(PeriodRef, 1, [set_period_fields], self.db_manager.save_periods_bulk)
    
    # --- Методы для поиска ---
    
    def on_search_text_changed(self, text: str):
        """Обработчик изменения текста поиска - перезагружает данные с учетом поиска"""
        self.search_text = text.strip()
        # Сбрасываем на первую страницу при новом поиске
        self.current_page = 1
        # Перезагружаем данные с учетом поиска
        self.load_current_reference()
    
    def clear_search(self):
        """Очистка поля поиска"""
        self.search_input.clear()
    
    # --- Методы для пагинации ---
    
    def _update_pagination_info(self):
        """Обновление информации о пагинации"""
        self.page_info_label.setText(f"{self.current_page} / {self.total_pages}")
        self.page_input.setMaximum(self.total_pages)
        self.page_input.setValue(self.current_page)
    
    def _update_pagination_buttons(self):
        """Обновление состояния кнопок пагинации"""
        self.first_page_btn.setEnabled(self.current_page > 1)
        self.prev_page_btn.setEnabled(self.current_page > 1)
        self.next_page_btn.setEnabled(self.current_page < self.total_pages)
        self.last_page_btn.setEnabled(self.current_page < self.total_pages)
    
    def _show_pagination(self):
        """Показать элементы пагинации"""
        self.pagination_widget.setVisible(True)
    
    def _hide_pagination(self):
        """Скрыть элементы пагинации"""
        self.pagination_widget.setVisible(False)
    
    def _go_to_page(self, page: int):
        """Переход на указанную страницу"""
        if 1 <= page <= self.total_pages and page != self.current_page:
            self.current_page = page
            self.load_current_reference()
    
    def go_to_first_page(self):
        """Переход на первую страницу"""
        self._go_to_page(1)
    
    def go_to_prev_page(self):
        """Переход на предыдущую страницу"""
        self._go_to_page(self.current_page - 1)
    
    def go_to_next_page(self):
        """Переход на следующую страницу"""
        self._go_to_page(self.current_page + 1)
    
    def go_to_last_page(self):
        """Переход на последнюю страницу"""
        self._go_to_page(self.total_pages)
    
    def on_page_input_changed(self, value: int):
        """Обработчик изменения номера страницы в поле ввода"""
        self._go_to_page(value)
    
    def on_page_size_changed(self, value: int):
        """Обработчик изменения размера страницы"""
        old_page_size = self.page_size
        self.page_size = value
        # Пересчитываем текущую страницу, чтобы остаться примерно в той же позиции
        if self.total_records > 0 and old_page_size > 0:
            old_offset = (self.current_page - 1) * old_page_size
            self.current_page = max(1, (old_offset // self.page_size) + 1)
        else:
            self.current_page = 1
        self.load_current_reference()

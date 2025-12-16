"""Панель вкладок"""
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QTabWidget,
                             QComboBox, QLabel, QCheckBox, QPushButton, QToolButton,
                             QTextEdit, QTableWidget, QHeaderView, QMenu)
from PyQt5.QtCore import Qt, QSize
from PyQt5.QtGui import QFont
from PyQt5.QtWidgets import QStyle
from views.excel_viewer import ExcelViewer
from views.widgets import WordWrapItemDelegate


class TabsPanel:
    """Класс для управления панелью вкладок"""
    
    def __init__(self, main_window):
        """
        Args:
            main_window: Ссылка на главное окно для доступа к контроллеру и обработчикам
        """
        self.main_window = main_window
        self.controller = main_window.controller
    
    def create_tabs_panel(self) -> QWidget:
        """Создание панели с вкладками"""
        tabs = QTabWidget()
        self.tabs_panel = tabs
        self.main_window.tabs_panel = tabs
        tabs.setTabsClosable(False)  # Отключаем стандартное закрытие вкладок
        
        # Добавляем контекстное меню для вкладок
        tabs.setContextMenuPolicy(Qt.CustomContextMenu)
        tabs.customContextMenuRequested.connect(self.main_window.show_tab_context_menu)
        
        # Вкладка с древовидными данными
        self.tree_tab = QWidget()
        
        tree_layout = QVBoxLayout(self.tree_tab)
        
        # Панель управления древом
        tree_control_layout = QHBoxLayout()
        # Кнопки управления деревом (максимально компактные)
        self.expand_all_btn = QToolButton()
        self.expand_all_btn.setIcon(self.main_window.style().standardIcon(QStyle.SP_ArrowDown))
        self.expand_all_btn.setToolTip("Развернуть все узлы дерева")
        self.expand_all_btn.setIconSize(QSize(14, 14))
        self.expand_all_btn.setToolButtonStyle(Qt.ToolButtonIconOnly)
        self.expand_all_btn.setAutoRaise(True)
        self.expand_all_btn.setFixedSize(22, 22)
        self.expand_all_btn.clicked.connect(self.main_window.expand_all_tree)
        tree_control_layout.addWidget(self.expand_all_btn)
        
        self.collapse_all_btn = QToolButton()
        self.collapse_all_btn.setIcon(self.main_window.style().standardIcon(QStyle.SP_ArrowUp))
        self.collapse_all_btn.setToolTip("Свернуть все узлы дерева")
        self.collapse_all_btn.setIconSize(QSize(14, 14))
        self.collapse_all_btn.setToolButtonStyle(Qt.ToolButtonIconOnly)
        self.collapse_all_btn.setAutoRaise(True)
        self.collapse_all_btn.setFixedSize(22, 22)
        self.collapse_all_btn.clicked.connect(self.main_window.collapse_all_tree)
        tree_control_layout.addWidget(self.collapse_all_btn)
        
        tree_control_layout.addStretch()
        
        # Выбор раздела
        tree_control_layout.addWidget(QLabel("Раздел:"))
        self.section_combo = QComboBox()
        self.section_combo.addItems(["Доходы", "Расходы", "Источники финансирования", "Консолидируемые расчеты"])
        self.section_combo.currentTextChanged.connect(self.main_window.on_section_changed)
        tree_control_layout.addWidget(self.section_combo)
        
        # Выбор типа данных
        tree_control_layout.addWidget(QLabel("Тип данных:"))
        self.data_type_combo = QComboBox()
        self.data_type_combo.addItems(["Утвержденный", "Исполненный", "Оба"])
        self.data_type_combo.currentTextChanged.connect(self.main_window.on_data_type_changed)
        tree_control_layout.addWidget(self.data_type_combo)
        
        # Чекбокс для скрытия нулевых столбцов
        self.hide_zero_columns_checkbox = QCheckBox("Скрыть нулевые столбцы")
        self.hide_zero_columns_checkbox.setToolTip("Скрыть столбцы, где в итоговой строке оба значения (утвержденный и исполненный) равны 0")
        self.hide_zero_columns_checkbox.stateChanged.connect(self.main_window.on_hide_zero_columns_changed)
        tree_control_layout.addWidget(self.hide_zero_columns_checkbox)
        
        # Панель инструментов для ревизии (активна только при выбранной ревизии)
        self.revision_toolbar = QHBoxLayout()
        self.revision_toolbar.setSpacing(5)
        
        # Кнопка пересчета
        self.recalculate_btn = QPushButton("Пересчитать")
        self.recalculate_btn.setIcon(self.main_window.style().standardIcon(QStyle.SP_BrowserReload))
        self.recalculate_btn.setToolTip("Пересчитать агрегированные суммы (F9)")
        self.recalculate_btn.setEnabled(False)
        self.recalculate_btn.clicked.connect(self.main_window.calculate_sums)
        self.revision_toolbar.addWidget(self.recalculate_btn)
        
        # Кнопка экспорта пересчитанной таблицы
        self.export_calculated_btn = QPushButton("Экспорт пересчитанной")
        self.export_calculated_btn.setIcon(self.main_window.style().standardIcon(QStyle.SP_DialogSaveButton))
        self.export_calculated_btn.setToolTip("Экспортировать форму с пересчитанными значениями")
        self.export_calculated_btn.setEnabled(False)
        self.export_calculated_btn.clicked.connect(self.main_window.export_calculated_table)
        self.revision_toolbar.addWidget(self.export_calculated_btn)
        
        # Кнопка показа ошибок расчетов
        self.show_errors_btn = QPushButton("Ошибки расчетов")
        self.show_errors_btn.setIcon(self.main_window.style().standardIcon(QStyle.SP_MessageBoxWarning))
        self.show_errors_btn.setToolTip("Показать ошибки расчетов")
        self.show_errors_btn.setEnabled(False)
        self.show_errors_btn.clicked.connect(self.main_window.show_calculation_errors)
        self.revision_toolbar.addWidget(self.show_errors_btn)
        
        # Кнопка открытия файла
        self.open_file_btn = QPushButton("Открыть файл")
        self.open_file_btn.setIcon(self.main_window.style().standardIcon(QStyle.SP_DirOpenIcon))
        self.open_file_btn.setToolTip("Открыть файл (doc, docx, xls, xlsx)")
        self.open_file_btn.setEnabled(True)
        self.open_file_btn.clicked.connect(self.main_window.open_file_dialog)
        self.revision_toolbar.addWidget(self.open_file_btn)
        
        # Кнопка открытия последнего экспортированного файла
        self.open_last_file_btn = QPushButton("Открыть последний")
        self.open_last_file_btn.setIcon(self.main_window.style().standardIcon(QStyle.SP_FileDialogStart))
        self.open_last_file_btn.setToolTip("Открыть последний экспортированный файл")
        self.open_last_file_btn.setEnabled(False)
        self.open_last_file_btn.clicked.connect(self.main_window.open_last_exported_file)
        self.revision_toolbar.addWidget(self.open_last_file_btn)
        
        # Меню документов
        self.documents_menu_btn = QPushButton("Документы ▼")
        self.documents_menu_btn.setIcon(self.main_window.style().standardIcon(QStyle.SP_FileDialogNewFolder))
        self.documents_menu_btn.setToolTip("Формирование документов")
        self.documents_menu_btn.setEnabled(False)
        self.documents_menu_btn.setMenu(QMenu(self.main_window))
        documents_menu = self.documents_menu_btn.menu()
        
        from PyQt5.QtWidgets import QAction
        generate_conclusion_action = QAction("Сформировать заключение...", self.main_window)
        generate_conclusion_action.setIcon(self.main_window.style().standardIcon(QStyle.SP_FileDialogNewFolder))
        generate_conclusion_action.triggered.connect(self.main_window.show_document_dialog)
        documents_menu.addAction(generate_conclusion_action)
        
        generate_letters_action = QAction("Сформировать письма...", self.main_window)
        generate_letters_action.setIcon(self.main_window.style().standardIcon(QStyle.SP_FileDialogNewFolder))
        generate_letters_action.triggered.connect(self.main_window.show_document_dialog)
        documents_menu.addAction(generate_letters_action)
        
        documents_menu.addSeparator()
        
        parse_solution_action = QAction("Обработать решение о бюджете...", self.main_window)
        parse_solution_action.setIcon(self.main_window.style().standardIcon(QStyle.SP_DialogOpenButton))
        parse_solution_action.triggered.connect(self.main_window.parse_solution_document)
        documents_menu.addAction(parse_solution_action)
        
        self.revision_toolbar.addWidget(self.documents_menu_btn)
        
        tree_control_layout.addLayout(self.revision_toolbar)
        tree_layout.addLayout(tree_control_layout)
        
        # Древовидный виджет (используем стандартный заголовок QTreeWidget)
        from PyQt5.QtWidgets import QTreeWidget
        self.data_tree = QTreeWidget()
        # Настраиваем заголовки дерева
        self.data_tree.setIndentation(10)
        # Отключаем единую высоту строк, чтобы высота подстраивалась под содержимое
        self.data_tree.setUniformRowHeights(False)
        # Включаем множественный выбор (Shift и Ctrl)
        self.data_tree.setSelectionMode(QTreeWidget.ExtendedSelection)
        # Устанавливаем делегат для переноса текста в ячейках
        self.data_tree.setItemDelegate(WordWrapItemDelegate())
        # Конфигурация заголовков будет выполнена позже (в main_window.configure_tree_headers)
        self.data_tree.setContextMenuPolicy(Qt.CustomContextMenu)
        self.data_tree.customContextMenuRequested.connect(self.main_window.show_tree_context_menu)
        self.data_tree.itemExpanded.connect(self.main_window.on_tree_item_expanded)
        self.data_tree.itemCollapsed.connect(self.main_window.on_tree_item_collapsed)
        # Обработчики выделения
        self.data_tree.itemSelectionChanged.connect(self.main_window.on_tree_selection_changed)
        self.data_tree.itemClicked.connect(self.main_window.on_tree_item_clicked)

        # Контекстное меню по заголовкам дерева (управление столбцами)
        header = self.data_tree.header()
        header.setContextMenuPolicy(Qt.CustomContextMenu)
        header.customContextMenuRequested.connect(self.main_window.show_tree_header_context_menu)

        tree_layout.addWidget(self.data_tree)
        
        # Сохраняем ссылки на виджеты в главном окне
        self.main_window.tree_tab = self.tree_tab
        self.main_window.data_tree = self.data_tree
        self.main_window.section_combo = self.section_combo
        self.main_window.data_type_combo = self.data_type_combo
        self.main_window.hide_zero_columns_checkbox = self.hide_zero_columns_checkbox
        self.main_window.expand_all_btn = self.expand_all_btn
        self.main_window.collapse_all_btn = self.collapse_all_btn
        self.main_window.revision_toolbar = self.revision_toolbar
        self.main_window.recalculate_btn = self.recalculate_btn
        self.main_window.export_calculated_btn = self.export_calculated_btn
        self.main_window.show_errors_btn = self.show_errors_btn
        self.main_window.open_file_btn = self.open_file_btn
        self.main_window.open_last_file_btn = self.open_last_file_btn
        self.main_window.documents_menu_btn = self.documents_menu_btn
        
        tabs.addTab(self.tree_tab, "Древовидные данные")
        
        # Вкладка с метаданными
        self.metadata_tab = QWidget()
        metadata_layout = QVBoxLayout(self.metadata_tab)
        
        self.metadata_text = QTextEdit()
        self.metadata_text.setReadOnly(True)
        metadata_layout.addWidget(self.metadata_text)
        
        self.main_window.metadata_tab = self.metadata_tab
        self.main_window.metadata_text = self.metadata_text
        
        tabs.addTab(self.metadata_tab, "Метаданные")
        
        # Вкладка с ошибками
        self.errors_tab = QWidget()
        errors_layout = QVBoxLayout(self.errors_tab)
        errors_layout.setContentsMargins(5, 5, 5, 5)
        
        # Заголовок и фильтры
        header_layout = QHBoxLayout()
        
        info_label = QLabel("Ошибки расчетов (несоответствия между оригинальными и расчетными значениями):")
        info_label.setFont(QFont("Arial", 10, QFont.Bold))
        header_layout.addWidget(info_label)
        
        header_layout.addStretch()
        
        # Фильтр по разделу
        header_layout.addWidget(QLabel("Раздел:"))
        self.errors_section_filter = QComboBox()
        self.errors_section_filter.addItems(["Все", "Доходы", "Расходы", "Источники финансирования", "Консолидируемые расчеты"])
        self.errors_section_filter.currentTextChanged.connect(
            lambda: self.main_window.errors_manager._update_errors_table()
        )
        header_layout.addWidget(self.errors_section_filter)
        
        errors_layout.addLayout(header_layout)
        
        # Таблица ошибок
        self.errors_table = QTableWidget()
        self.errors_table.setColumnCount(9)
        self.errors_table.setHorizontalHeaderLabels([
            "Раздел",
            "Наименование",
            "Код строки",
            "Уровень",
            "Тип",
            "Колонка",
            "Оригинальное",
            "Расчетное",
            "Разница"
        ])
        
        # Настройка таблицы
        header = self.errors_table.horizontalHeader()
        # Отключаем растягивание последнего столбца
        header.setStretchLastSection(False)
        # Используем Interactive режим для каждого столбца отдельно, чтобы можно было вручную изменять ширину
        for i in range(9):
            header.setSectionResizeMode(i, QHeaderView.Interactive)
        # Устанавливаем начальные размеры столбцов
        header.resizeSection(0, 120)  # Раздел
        header.resizeSection(1, 300)  # Наименование
        header.resizeSection(2, 100)  # Код строки
        header.resizeSection(3, 60)   # Уровень
        header.resizeSection(4, 120)  # Тип
        header.resizeSection(5, 100)  # Колонка
        header.resizeSection(6, 120)  # Оригинальное
        header.resizeSection(7, 120)  # Расчетное
        header.resizeSection(8, 120)  # Разница
        
        self.errors_table.setAlternatingRowColors(True)
        self.errors_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.errors_table.setEditTriggers(QTableWidget.NoEditTriggers)
        
        errors_layout.addWidget(self.errors_table)
        
        # Кнопки и статистика
        buttons_layout = QHBoxLayout()
        buttons_layout.addStretch()
        
        self.errors_export_btn = QPushButton("Экспорт...")
        self.errors_export_btn.setIcon(self.main_window.style().standardIcon(QStyle.SP_DialogSaveButton))
        self.errors_export_btn.clicked.connect(self.main_window.errors_manager._export_errors)
        buttons_layout.addWidget(self.errors_export_btn)
        
        errors_layout.addLayout(buttons_layout)
        
        # Статистика
        self.errors_stats_label = QLabel("Ошибок не найдено")
        self.errors_stats_label.setFont(QFont("Arial", 9))
        errors_layout.addWidget(self.errors_stats_label)
        
        # Сохраняем ссылки на виджеты ошибок в главном окне
        self.main_window.errors_tab = self.errors_tab
        self.main_window.errors_table = self.errors_table
        self.main_window.errors_section_filter = self.errors_section_filter
        self.main_window.errors_stats_label = self.errors_stats_label
        self.main_window.errors_export_btn = self.errors_export_btn
        
        tabs.addTab(self.errors_tab, "Ошибки")
        
        # Вкладка с просмотром Excel
        self.excel_viewer = ExcelViewer()
        self.main_window.excel_viewer = self.excel_viewer
        tabs.addTab(self.excel_viewer, "Просмотр формы")
        
        return tabs

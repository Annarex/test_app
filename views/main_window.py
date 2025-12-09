from PyQt5.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                             QPushButton, QLabel, QSplitter, 
                             QTabWidget, QTableWidget, QTableWidgetItem, 
                             QHeaderView, QMessageBox, QFileDialog, QProgressBar,
                             QToolBar, QStatusBar, QAction, QTextEdit,
                             QComboBox, QTreeWidget, QTreeWidgetItem, QMenu, 
                             QInputDialog, QDialog, QDialogButtonBox, QFormLayout,
                             QLineEdit, QCheckBox, QApplication, QStyle)
from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtGui import QFont, QColor, QBrush
import os
from pathlib import Path
import pandas as pd

from controllers.main_controller import MainController
from logger import logger
from models.form_0503317 import Form0503317Constants
from views.project_dialog import ProjectDialog
from views.reference_dialog import ReferenceDialog
from views.excel_viewer import ExcelViewer
from views.reference_viewer import ReferenceViewer
from views.dictionaries_dialog import DictionariesDialog
from views.form_load_dialog import FormLoadDialog

class MainWindow(QMainWindow):
    """Главное окно приложения"""
    
    def __init__(self):
        super().__init__()
        self.controller = MainController()
        self.current_section = "Доходы"
        self.current_data_type = "Оба"
        self.main_splitter = None
        self.projects_panel_index = 0
        self.projects_inner_panel = None
        self.projects_toggle_button = None
        self.projects_panel_last_size = 260
        self.reference_window = None
        self.tree_headers = []
        self.tree_header_tooltips = []
        self.tree_column_mapping = {}
        self._updating_header_height = False  # Флаг для предотвращения бесконечного цикла
        self.init_ui()
        self.connect_signals()
        self.controller.load_initial_data()
    
    def init_ui(self):
        """Инициализация интерфейса"""
        self.setWindowTitle("Система обработки бюджетных форм")
        self.setGeometry(100, 100, 1600, 900)
        
        # Создаем меню-бар
        self.create_menu_bar()
        
        # Создаем центральный виджет
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # Основной layout
        main_layout = QHBoxLayout(central_widget)
        
        # Создаем сплиттер
        splitter = QSplitter(Qt.Horizontal)
        self.main_splitter = splitter
        
        # Левая панель - список проектов
        self.projects_panel = self.create_projects_panel()
        splitter.addWidget(self.projects_panel)
        self.projects_panel_index = splitter.indexOf(self.projects_panel)
        
        # Центральная панель - вкладки с данными
        self.tabs_panel = self.create_tabs_panel()
        splitter.addWidget(self.tabs_panel)
        
        # Устанавливаем пропорции
        splitter.setSizes([300, 1300])
        
        main_layout.addWidget(splitter)
        
        # Создаем тулбар
        # self.create_toolbar()
        
        # Создаем статусбар
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("Готов к работе")
        
        # Прогресс-бар
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.status_bar.addPermanentWidget(self.progress_bar)
        
        # Создаем док-виджеты
        self.create_dock_widgets()
    
    def create_menu_bar(self):
        """Создание меню-бара"""
        menubar = self.menuBar()
        
        # ========== Меню "Файл" ==========
        file_menu = menubar.addMenu("&Файл")
        
        new_project_action = QAction("&Новый проект...", self)
        new_project_action.setIcon(self.style().standardIcon(QStyle.SP_FileIcon))
        new_project_action.setShortcut("Ctrl+N")
        new_project_action.setStatusTip("Создать новый проект")
        new_project_action.triggered.connect(self.show_new_project_dialog)
        file_menu.addAction(new_project_action)
        
        load_form_action = QAction("&Загрузить форму...", self)
        load_form_action.setIcon(self.style().standardIcon(QStyle.SP_DirOpenIcon))
        load_form_action.setShortcut("Ctrl+O")
        load_form_action.setStatusTip("Загрузить файл формы")
        load_form_action.triggered.connect(self.load_form_file)
        file_menu.addAction(load_form_action)
        
        file_menu.addSeparator()
        
        export_action = QAction("&Экспорт проверки...", self)
        export_action.setIcon(self.style().standardIcon(QStyle.SP_DialogSaveButton))
        export_action.setShortcut("Ctrl+E")
        export_action.setStatusTip("Экспортировать форму с проверкой")
        export_action.triggered.connect(self.export_validation)
        file_menu.addAction(export_action)
        
        file_menu.addSeparator()
        
        exit_action = QAction("&Выход", self)
        exit_action.setIcon(self.style().standardIcon(QStyle.SP_DialogCloseButton))
        exit_action.setShortcut("Ctrl+Q")
        exit_action.setStatusTip("Выход из приложения")
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)
        
        # ========== Меню "Проект" ==========
        project_menu = menubar.addMenu("&Проект")
        
        edit_project_action = QAction("&Редактировать проект...", self)
        edit_project_action.setIcon(self.style().standardIcon(QStyle.SP_FileDialogDetailedView))
        edit_project_action.setShortcut("Ctrl+P")
        edit_project_action.setStatusTip("Редактировать текущий проект")
        edit_project_action.triggered.connect(self.edit_current_project)
        project_menu.addAction(edit_project_action)
        
        delete_project_action = QAction("&Удалить проект", self)
        delete_project_action.setIcon(self.style().standardIcon(QStyle.SP_TrashIcon))
        delete_project_action.setShortcut("Ctrl+Delete")
        delete_project_action.setStatusTip("Удалить текущий проект")
        delete_project_action.triggered.connect(self.delete_current_project)
        project_menu.addAction(delete_project_action)
        
        project_menu.addSeparator()
        
        refresh_projects_action = QAction("&Обновить список", self)
        refresh_projects_action.setIcon(self.style().standardIcon(QStyle.SP_BrowserReload))
        refresh_projects_action.setShortcut("F5")
        refresh_projects_action.setStatusTip("Обновить список проектов")
        refresh_projects_action.triggered.connect(lambda: self.controller.projects_updated.emit(self.controller.project_controller.load_projects()))
        project_menu.addAction(refresh_projects_action)
        
        # ========== Меню "Данные" ==========
        data_menu = menubar.addMenu("&Данные")
        
        calculate_action = QAction("&Пересчитать суммы", self)
        calculate_action.setIcon(self.style().standardIcon(QStyle.SP_BrowserReload))
        calculate_action.setShortcut("F9")
        calculate_action.setStatusTip("Пересчитать агрегированные суммы")
        calculate_action.triggered.connect(self.calculate_sums)
        data_menu.addAction(calculate_action)
        
        data_menu.addSeparator()
        
        hide_zeros_action = QAction("&Скрыть нулевые столбцы", self)
        hide_zeros_action.setIcon(self.style().standardIcon(QStyle.SP_TrashIcon))
        hide_zeros_action.setShortcut("Ctrl+H")
        hide_zeros_action.setStatusTip("Скрыть столбцы с нулевыми значениями")
        hide_zeros_action.triggered.connect(self.hide_zero_columns_global)
        data_menu.addAction(hide_zeros_action)
        
        # ========== Меню "Справочники" ==========
        reference_menu = menubar.addMenu("&Справочники")
        
        load_income_ref_action = QAction("&Загрузить справочник доходов...", self)
        load_income_ref_action.setIcon(self.style().standardIcon(QStyle.SP_DialogOpenButton))
        load_income_ref_action.setStatusTip("Загрузить справочник доходов")
        load_income_ref_action.triggered.connect(lambda: self.show_reference_dialog("доходы"))
        reference_menu.addAction(load_income_ref_action)
        
        load_sources_ref_action = QAction("&Загрузить справочник источников...", self)
        load_sources_ref_action.setIcon(self.style().standardIcon(QStyle.SP_DialogOpenButton))
        load_sources_ref_action.setStatusTip("Загрузить справочник источников финансирования")
        load_sources_ref_action.triggered.connect(lambda: self.show_reference_dialog("источники"))
        reference_menu.addAction(load_sources_ref_action)
        
        reference_menu.addSeparator()
        
        show_references_action = QAction("&Просмотр справочников", self)
        show_references_action.setIcon(self.style().standardIcon(QStyle.SP_FileDialogInfoView))
        show_references_action.setShortcut("Ctrl+R")
        show_references_action.setStatusTip("Открыть окно просмотра справочников")
        show_references_action.triggered.connect(self.show_reference_viewer)
        reference_menu.addAction(show_references_action)
        
        reference_menu.addSeparator()
        
        config_dicts_action = QAction("&Справочники конфигурации...", self)
        config_dicts_action.setIcon(self.style().standardIcon(QStyle.SP_FileDialogListView))
        config_dicts_action.setShortcut("Ctrl+D")
        config_dicts_action.setStatusTip("Редактировать справочники конфигурации (годы, МО, типы форм, периоды)")
        config_dicts_action.triggered.connect(self.show_config_dictionaries)
        reference_menu.addAction(config_dicts_action)
        
        # ========== Меню "Вид" ==========
        view_menu = menubar.addMenu("&Вид")
        
        toggle_projects_panel_action = QAction("&Панель проектов", self)
        toggle_projects_panel_action.setCheckable(True)
        toggle_projects_panel_action.setChecked(True)
        toggle_projects_panel_action.setShortcut("Ctrl+1")
        toggle_projects_panel_action.setStatusTip("Показать/скрыть панель проектов")
        toggle_projects_panel_action.triggered.connect(self.toggle_projects_panel)
        view_menu.addAction(toggle_projects_panel_action)
        
        view_menu.addSeparator()
        
        fullscreen_action = QAction("&Полноэкранный режим", self)
        fullscreen_action.setShortcut("F11")
        fullscreen_action.setCheckable(True)
        fullscreen_action.setStatusTip("Переключить полноэкранный режим")
        fullscreen_action.triggered.connect(self.toggle_fullscreen)
        view_menu.addAction(fullscreen_action)
        
        # ========== Меню "Справка" ==========
        help_menu = menubar.addMenu("&Справка")
        
        about_action = QAction("&О программе", self)
        about_action.setIcon(self.style().standardIcon(QStyle.SP_MessageBoxInformation))
        about_action.setStatusTip("Информация о программе")
        about_action.triggered.connect(self.show_about)
        help_menu.addAction(about_action)
        
        help_menu.addSeparator()
        
        shortcuts_action = QAction("&Горячие клавиши", self)
        shortcuts_action.setIcon(self.style().standardIcon(QStyle.SP_FileDialogInfoView))
        shortcuts_action.setStatusTip("Список горячих клавиш")
        shortcuts_action.triggered.connect(self.show_shortcuts)
        help_menu.addAction(shortcuts_action)
    
    def create_toolbar(self):
        """Создание панели инструментов"""
        toolbar = QToolBar("Основные инструменты")
        self.addToolBar(toolbar)
        
        # Действия
        new_project_action = QAction("Новый проект", self)
        new_project_action.setIcon(self.style().standardIcon(QStyle.SP_FileIcon))
        new_project_action.triggered.connect(self.show_new_project_dialog)
        toolbar.addAction(new_project_action)
        
        load_form_action = QAction("Загрузить форму", self)
        load_form_action.setIcon(self.style().standardIcon(QStyle.SP_DirOpenIcon))
        load_form_action.triggered.connect(self.load_form_file)
        toolbar.addAction(load_form_action)
        
        calculate_action = QAction("Пересчитать", self)
        calculate_action.setIcon(self.style().standardIcon(QStyle.SP_BrowserReload))
        calculate_action.triggered.connect(self.calculate_sums)
        toolbar.addAction(calculate_action)
        
        export_action = QAction("Экспорт проверки", self)
        export_action.setIcon(self.style().standardIcon(QStyle.SP_DialogSaveButton))
        export_action.triggered.connect(self.export_validation)
        toolbar.addAction(export_action)
        
        toolbar.addSeparator()
        
        # Отдельные действия для справочников доходов и источников
        load_income_ref_action = QAction("Справочник доходов", self)
        load_income_ref_action.setIcon(self.style().standardIcon(QStyle.SP_DialogOpenButton))
        load_income_ref_action.triggered.connect(lambda: self.show_reference_dialog("доходы"))
        toolbar.addAction(load_income_ref_action)

        load_sources_ref_action = QAction("Справочник источников", self)
        load_sources_ref_action.setIcon(self.style().standardIcon(QStyle.SP_DialogOpenButton))
        load_sources_ref_action.triggered.connect(lambda: self.show_reference_dialog("источники"))
        toolbar.addAction(load_sources_ref_action)

        # Кнопка для сворачивания нулевых столбцов (таблица + дерево)
        hide_zeros_action = QAction("Нулевые столбцы", self)
        hide_zeros_action.setIcon(self.style().standardIcon(QStyle.SP_TrashIcon))
        hide_zeros_action.triggered.connect(self.hide_zero_columns_global)
        toolbar.addAction(hide_zeros_action)
        
        show_references_action = QAction("Просмотр справочников", self)
        show_references_action.setIcon(self.style().standardIcon(QStyle.SP_FileDialogInfoView))
        show_references_action.triggered.connect(self.show_reference_viewer)
        toolbar.addAction(show_references_action)

        # Редактор конфигурационных справочников (годы, МО, типы форм, периоды)
        config_dicts_action = QAction("Справочники конфигурации", self)
        config_dicts_action.setIcon(self.style().standardIcon(QStyle.SP_FileDialogListView))
        config_dicts_action.triggered.connect(self.show_config_dictionaries)
        toolbar.addAction(config_dicts_action)

        # Кнопки управления панелью проектов размещены непосредственно на самой панели
    
    def create_dock_widgets(self):
        """Инициализация структур для просмотра справочников (отдельное окно)"""
        # Метод оставлен для совместимости, но структуры инициализируются в show_reference_viewer
        pass
    
    def create_projects_panel(self) -> QWidget:
        """Создание панели проектов"""
        # Основная панель с содержимым
        inner_panel = QWidget()
        layout = QVBoxLayout(inner_panel)
        layout.setContentsMargins(6, 6, 2, 6)
        
        # Заголовок
        title_label = QLabel("Проекты")
        title_label.setFont(QFont("Arial", 12, QFont.Bold))
        layout.addWidget(title_label)
        
        # Кнопки управления проектами
        buttons_layout = QHBoxLayout()
        
        new_project_btn = QPushButton("Новый")
        new_project_btn.clicked.connect(self.show_new_project_dialog)
        buttons_layout.addWidget(new_project_btn)
        
        refresh_btn = QPushButton("Обновить")
        refresh_btn.clicked.connect(self.refresh_projects)
        buttons_layout.addWidget(refresh_btn)
        
        layout.addLayout(buttons_layout)
        
        # Дерево проектов: Год -> Проект -> Форма -> Ревизия
        from PyQt5.QtWidgets import QTreeWidget
        self.projects_tree = QTreeWidget()
        self.projects_tree.setIndentation(10)
        self.projects_tree.setHeaderHidden(True)
        self.projects_tree.itemDoubleClicked.connect(self.on_project_tree_double_clicked)
        self.projects_tree.setContextMenuPolicy(Qt.CustomContextMenu)
        self.projects_tree.customContextMenuRequested.connect(self.show_project_context_menu)
        layout.addWidget(self.projects_tree)
        
        # Информация о проекте
        self.project_info_label = QLabel("Выберите проект")
        self.project_info_label.setWordWrap(True)
        layout.addWidget(self.project_info_label)
        
        # Контейнер, в котором слева основная панель, справа узкая кнопка-свертка
        container = QWidget()
        container_layout = QHBoxLayout(container)
        container_layout.setContentsMargins(0, 0, 0, 0)
        container_layout.setSpacing(0)

        container_layout.addWidget(inner_panel)

        # Узкая вертикальная кнопка на правом краю панели
        toggle_button = QPushButton("◀")
        toggle_button.setFixedWidth(14)
        toggle_button.setFlat(True)
        toggle_button.setFocusPolicy(Qt.NoFocus)
        toggle_button.setToolTip("Свернуть/развернуть панель проектов")
        toggle_button.clicked.connect(self.on_projects_side_button_clicked)
        container_layout.addWidget(toggle_button)

        self.projects_inner_panel = inner_panel
        self.projects_toggle_button = toggle_button

        return container
    
    def create_tabs_panel(self) -> QWidget:
        """Создание панели с вкладками"""
        tabs = QTabWidget()
        
        # Вкладка с древовидными данными
        self.tree_tab = QWidget()
        tree_layout = QVBoxLayout(self.tree_tab)
        
        # Панель управления древом
        tree_control_layout = QHBoxLayout()
        
        # Выбор раздела
        tree_control_layout.addWidget(QLabel("Раздел:"))
        self.section_combo = QComboBox()
        self.section_combo.addItems(["Доходы", "Расходы", "Источники финансирования", "Консолидируемые расчеты"])
        self.section_combo.currentTextChanged.connect(self.on_section_changed)
        tree_control_layout.addWidget(self.section_combo)
        
        # Выбор типа данных
        tree_control_layout.addWidget(QLabel("Тип данных:"))
        self.data_type_combo = QComboBox()
        self.data_type_combo.addItems(["Утвержденный", "Исполненный", "Оба"])
        self.data_type_combo.currentTextChanged.connect(self.on_data_type_changed)
        tree_control_layout.addWidget(self.data_type_combo)
        
        tree_control_layout.addStretch()
        tree_layout.addLayout(tree_control_layout)
        
        # Древовидный виджет (используем стандартный заголовок QTreeWidget)
        self.data_tree = QTreeWidget()
        # Настраиваем заголовки дерева
        self.data_tree.setIndentation(10)
        self.configure_tree_headers(self.current_section)
        self.data_tree.setContextMenuPolicy(Qt.CustomContextMenu)
        self.data_tree.customContextMenuRequested.connect(self.show_tree_context_menu)
        self.data_tree.itemExpanded.connect(self.on_tree_item_expanded)
        self.data_tree.itemCollapsed.connect(self.on_tree_item_collapsed)

        # Контекстное меню по заголовкам дерева (управление столбцами)
        header = self.data_tree.header()
        header.setContextMenuPolicy(Qt.CustomContextMenu)
        header.customContextMenuRequested.connect(self.show_tree_header_context_menu)

        tree_layout.addWidget(self.data_tree)
        
        tabs.addTab(self.tree_tab, "Древовидные данные")
        
        # Вкладка с табличными данными
        self.table_tab = QWidget()
        table_layout = QVBoxLayout(self.table_tab)
        
        # Таблица для отображения данных
        self.data_table = QTableWidget()
        self.data_table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.data_table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.data_table.customContextMenuRequested.connect(self.show_table_context_menu)
        table_layout.addWidget(self.data_table)
        
        tabs.addTab(self.table_tab, "Табличные данные")
        
        # Вкладка с метаданными
        self.metadata_tab = QWidget()
        metadata_layout = QVBoxLayout(self.metadata_tab)
        
        self.metadata_text = QTextEdit()
        self.metadata_text.setReadOnly(True)
        metadata_layout.addWidget(self.metadata_text)
        
        tabs.addTab(self.metadata_tab, "Метаданные")
        
        # Вкладка с ошибками
        self.errors_tab = QWidget()
        errors_layout = QVBoxLayout(self.errors_tab)
        
        self.errors_table = QTableWidget()
        errors_layout.addWidget(self.errors_table)
        
        tabs.addTab(self.errors_tab, "Ошибки")
        
        # Вкладка с просмотром Excel
        self.excel_viewer = ExcelViewer()
        tabs.addTab(self.excel_viewer, "Просмотр формы")
        
        return tabs
    
    def connect_signals(self):
        """Подключение сигналов"""
        self.controller.projects_updated.connect(self.update_projects_list)
        self.controller.project_loaded.connect(self.on_project_loaded)
        self.controller.calculation_completed.connect(self.on_calculation_completed)
        self.controller.export_completed.connect(self.on_export_completed)
        self.controller.error_occurred.connect(self.on_error_occurred)
    
    def update_projects_list(self, _projects):
        """Обновление дерева проектов по новой архитектуре MainController.build_project_tree"""
        from PyQt5.QtWidgets import QTreeWidgetItem

        self.projects_tree.clear()

        # Получаем структурированные данные от контроллера
        tree_data = self.controller.build_project_tree()

        for year_entry in tree_data:
            year_label = f"Год {year_entry['year']}"
            year_item = QTreeWidgetItem([year_label])
            self.projects_tree.addTopLevelItem(year_item)

            for proj in year_entry["projects"]:
                proj_item = QTreeWidgetItem([proj["name"]])
                # Сохраняем ID проекта на уровне узла проекта
                proj_item.setData(0, Qt.UserRole, proj["id"])
                year_item.addChild(proj_item)

                # Формы/периоды/ревизии (показываем даже пустые, с заглушками)
                if proj.get("forms"):
                    for form in proj["forms"]:
                        form_label = f"{form['form_name']} ({form['form_code']})"
                        form_item = QTreeWidgetItem([form_label])
                        proj_item.addChild(form_item)

                        periods = form.get("periods") or []
                        if not periods:
                            form_item.addChild(QTreeWidgetItem(["Нет периодов"]))
                            continue

                        for period in periods:
                            period_label = period.get("period_name") or period.get("period_code") or "—"
                            period_item = QTreeWidgetItem([period_label])
                            form_item.addChild(period_item)

                            revisions = period.get("revisions") or []
                            if revisions:
                                for rev in revisions:
                                    status_icon = "✅" if rev["status"] == "calculated" else "📝"
                                    rev_text = f"{status_icon} рев. {rev['revision']}"
                                    rev_item = QTreeWidgetItem([rev_text])
                                    rev_item.setData(0, Qt.UserRole, rev.get("project_id"))
                                    revision_id = rev.get("revision_id")
                                    rev_item.setData(0, Qt.UserRole + 1, revision_id)
                                    if revision_id:
                                        logger.debug(
                                            f"Сохранена ревизия в дереве: "
                                            f"revision_id={revision_id}, project_id={rev.get('project_id')}, revision={rev.get('revision')}"
                                        )
                                    period_item.addChild(rev_item)
                            else:
                                period_item.addChild(QTreeWidgetItem(["Нет ревизий"]))
                else:
                    # Совсем нет форм — заглушка
                    placeholder = QTreeWidgetItem(["Нет ревизий"])
                    proj_item.addChild(placeholder)

        # Разворачиваем верхние уровни (год, проект, форма, период)
        # Ревизии остаются свернутыми по умолчанию
        for i in range(self.projects_tree.topLevelItemCount()):
            year_item = self.projects_tree.topLevelItem(i)
            year_item.setExpanded(True)
            for j in range(year_item.childCount()):
                proj_item = year_item.child(j)
                proj_item.setExpanded(True)
                for k in range(proj_item.childCount()):
                    form_item = proj_item.child(k)
                    form_item.setExpanded(True)
                    for m in range(form_item.childCount()):
                        period_item = form_item.child(m)
                        period_item.setExpanded(True)

    def on_project_tree_double_clicked(self, item, column):
        """Обработка двойного клика по дереву проектов"""
        # Поднимаемся по дереву, чтобы найти project_id/revision_id даже при клике на заглушки
        def _resolve_ids(it):
            proj_id = None
            rev_id = None
            cur = it
            while cur:
                if proj_id is None:
                    proj_id = cur.data(0, Qt.UserRole)
                if rev_id is None:
                    rev_id = cur.data(0, Qt.UserRole + 1)
                if proj_id is not None and rev_id is not None:
                    break
                cur = cur.parent()
            return proj_id, rev_id

        project_id, revision_id = _resolve_ids(item)
        
        if not project_id:
            return
        
        # Определяем, является ли узел ревизией (ревизия имеет revision_id и является дочерним элементом периода)
        is_revision = False
        if revision_id is not None and revision_id != 0:
            # Проверяем структуру дерева: ревизия является дочерним элементом периода
            parent = item.parent()
            if parent and item.childCount() == 0:
                # Период является дочерним элементом формы
                grandparent = parent.parent() if parent else None
                if grandparent:
                    grandparent_text = grandparent.text(0).lower()
                    if "форма" in grandparent_text or "(" in grandparent_text:
                        is_revision = True
        
        if is_revision:
            # Подтягиваем параметры формы из ревизии для последующей загрузки файлов
            self.controller.set_form_params_from_revision(revision_id)
            # Загружаем конкретную ревизию
            logger.info(f"Загрузка ревизии {revision_id} для проекта {project_id}")
            self.controller.load_revision(revision_id, project_id)
        else:
            # Клик по проекту/форме/периоду/заглушке — выбираем проект, чтобы можно было загрузить новую форму
            if project_id:
                logger.debug(f"Выбор проекта {project_id}")
                self.controller.project_controller.load_project(project_id)
            else:
                logger.warning("Проект не определён для выбранного узла")

    def show_project_context_menu(self, position):
        """Контекстное меню для дерева проектов"""
        item = self.projects_tree.itemAt(position)
        if not item:
            return
        project_id = item.data(0, Qt.UserRole)
        revision_id = item.data(0, Qt.UserRole + 1)

        # Если нет ID проекта — контекстное меню не показываем
        if not project_id:
            return

        # Определяем, является ли узел ревизией
        # Структура дерева: Год -> Проект -> Форма -> Период -> Ревизия
        # Ревизия - это узел, который является дочерним элементом периода
        # и не имеет дочерних элементов
        is_revision = False
        
        # Проверяем структуру дерева: ревизия является дочерним элементом периода
        parent = item.parent()
        if parent and item.childCount() == 0:
            # Период является дочерним элементом формы
            grandparent = parent.parent() if parent else None
            if grandparent:
                # Проверяем, что дедушка - это форма (содержит "форма" или "(")
                grandparent_text = grandparent.text(0).lower()
                if "форма" in grandparent_text or "(" in grandparent_text:
                    # Родитель - период, значит текущий узел - ревизия
                    is_revision = True

        menu = QMenu()
        edit_action = None
        edit_rev_action = None
        delete_rev_action = None
        delete_project_action = None

        # Если это узел ревизии
        if is_revision:
            # Для ревизии нужен revision_id для редактирования/удаления
            if revision_id is not None:
                edit_rev_action = menu.addAction("Редактировать ревизию")
                delete_rev_action = menu.addAction("Удалить ревизию")
            # Если revision_id не установлен (виртуальная ревизия из старой модели),
            # действия редактирования/удаления недоступны
        else:
            # Для узла проекта (не ревизии) показываем действия проекта
            edit_action = menu.addAction("Редактировать проект")
            delete_project_action = menu.addAction("Удалить проект")

        action = menu.exec_(self.projects_tree.mapToGlobal(position))

        if action == edit_action:
            self.edit_project(project_id)
        elif edit_rev_action is not None and action == edit_rev_action and revision_id:
            self.edit_revision(revision_id, project_id)
        elif delete_rev_action is not None and action == delete_rev_action and revision_id:
            reply = QMessageBox.question(
                self,
                "Подтверждение",
                "Вы уверены, что хотите удалить выбранную ревизию?",
                QMessageBox.Yes | QMessageBox.No,
            )
            if reply == QMessageBox.Yes:
                self.controller.delete_form_revision(revision_id)
                # После удаления ревизии обновляем дерево
                self.update_projects_list(None)
        elif action == delete_project_action:
            reply = QMessageBox.question(
                self,
                "Подтверждение",
                "Вы уверены, что хотите удалить проект (все ревизии)?",
                QMessageBox.Yes | QMessageBox.No,
            )
            if reply == QMessageBox.Yes:
                self.controller.delete_project(project_id)

    def edit_project(self, project_id: int):
        """Редактирование проекта через диалог"""
        try:
            # Загружаем проект в контроллер (установит current_project)
            self.controller.load_project(project_id)

            from views.project_dialog import ProjectDialog

            dlg = ProjectDialog(self)
            # Заполняем диалог текущим проектом
            if self.controller.current_project:
                dlg.set_project(self.controller.current_project)

            if dlg.exec_():
                project_data = dlg.get_project_data()
                if self.controller.update_project(project_data):
                    self.status_bar.showMessage(
                        f"Проект '{self.controller.current_project.name}' обновлён"
                    )
                    # Обновляем дерево проектов
                    self.update_projects_list(None)
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Ошибка редактирования проекта: {e}")
    
    def edit_revision(self, revision_id: int, project_id: int):
        """Редактирование ревизии через диалог"""
        try:
            from views.revision_dialog import RevisionDialog

            dlg = RevisionDialog(self.controller.db_manager, self)
            # Загружаем данные ревизии
            revision = self.controller.db_manager.get_form_revision_by_id(revision_id)
            if not revision:
                QMessageBox.warning(self, "Ошибка", "Ревизия не найдена")
                return
            
            dlg.set_revision(revision, project_id)

            if dlg.exec_():
                revision_data = dlg.get_revision_data()
                if self.controller.update_form_revision(revision_id, revision_data):
                    self.status_bar.showMessage("Ревизия обновлена")
                    # Дерево проектов обновится автоматически через сигнал projects_updated
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Ошибка редактирования ревизии: {e}")
    
    def on_project_loaded(self, project):
        """Обработка загруженного проекта"""
        try:
            # Убеждаемся, что прогресс-бар скрыт
            self.progress_bar.setVisible(False)

            # --------------------------------------------------
            # Определяем текущую ревизию и связанную мета‑информацию
            # --------------------------------------------------
            rev_id = getattr(self.controller, "current_revision_id", None)
            form_text = "—"
            revision_text = "—"
            status_text = "—"
            period_text = "—"

            excel_path = None

            if rev_id:
                try:
                    db = self.controller.db_manager
                    revision = db.get_form_revision_by_id(rev_id)
                    if revision:
                        # Ревизия и статус
                        revision_text = revision.revision or "—"
                        from models.base_models import ProjectStatus  # локальный импорт, чтобы избежать циклов
                        if isinstance(revision.status, ProjectStatus):
                            status_text = revision.status.value
                        else:
                            # На случай строкового статуса
                            status_text = str(revision.status or "—")

                        # Путь к файлу для Excel‑просмотра
                        excel_path = revision.file_path or None

                        # Находим связанную форму и её тип / период
                        project_forms = db.load_project_forms(project.id)
                        pf = next((p for p in project_forms if p.id == revision.project_form_id), None)
                        if pf:
                            # Тип формы
                            form_types_meta = {ft.id: ft for ft in db.load_form_types_meta()}
                            ft_meta = form_types_meta.get(pf.form_type_id)
                            if ft_meta:
                                # Показываем и код, и читаемое имя, если есть
                                if ft_meta.name:
                                    form_text = f"{ft_meta.name} ({ft_meta.code})"
                                else:
                                    form_text = ft_meta.code
                            # Период
                            if pf.period_id:
                                periods = db.load_periods()
                                period_ref = next((p for p in periods if p.id == pf.period_id), None)
                                if period_ref:
                                    period_text = period_ref.name or period_ref.code or period_text
                    else:
                        # Если ревизия по ID не найдена — fallback на старые поля проекта
                        revision_text = project.revision or "—"
                        status_text = getattr(project.status, "value", str(project.status)) if project.status else "—"
                        form_text = getattr(project.form_type, "value", str(project.form_type)) if project.form_type else "—"
                except Exception as e:
                    logger.error(f"Ошибка получения информации о ревизии: {e}", exc_info=True)
                    # Fallback на старые поля проекта
                    revision_text = project.revision or "—"
                    status_text = getattr(project.status, "value", str(project.status)) if project.status else "—"
                    form_text = getattr(project.form_type, "value", str(project.form_type)) if project.form_type else "—"
            else:
                # Проект без выбранной ревизии (старые проекты или только что созданные)
                form_text = getattr(project.form_type, "value", str(project.form_type)) if project.form_type else "—"
                revision_text = project.revision or "—"
                status_text = getattr(project.status, "value", str(project.status)) if project.status else "—"

            # МО — берём из справочника по municipality_id проекта
            municipality_text = "—"
            try:
                if hasattr(project, "municipality_id") and project.municipality_id:
                    db = self.controller.db_manager
                    municip_list = db.load_municipalities()
                    municip_ref = next((m for m in municip_list if m.id == project.municipality_id), None)
                    if municip_ref:
                        municipality_text = municip_ref.name or municipality_text
            except Exception as e:
                logger.warning(f"Ошибка получения МО для проекта {project.id}: {e}", exc_info=True)

            # Обновляем информацию о проекте
            info_text = (
                f"<b>Проект:</b> {project.name}<br>"
                f"<b>Форма:</b> {form_text}<br>"
                f"<b>Ревизия:</b> {revision_text}<br>"
                f"<b>МО:</b> {municipality_text}<br>"
                f"<b>Период:</b> {period_text}<br>"
                f"<b>Статус:</b> {status_text}<br>"
                f"<b>Создан:</b> {project.created_at.strftime('%d.%m.%Y %H:%M')}"
            )
            self.project_info_label.setText(info_text)

            # Загружаем данные в древовидное представление
            self.load_project_data_to_tree(project)

            # Загружаем метаданные
            self.load_metadata(project)

            # Загружаем файл в просмотрщик Excel:
            # Используем исходный файл ревизии (form_revisions.file_path), а не экспортированный
            # Экспортированный файл сохраняется отдельно и не должен заменять исходный
            if excel_path and os.path.exists(excel_path):
                # excel_path уже содержит путь к исходному файлу ревизии из revision_record.file_path
                self.excel_viewer.load_excel_file(excel_path)
            # Если файл не найден, просто не загружаем его

            self.status_bar.showMessage(f"Проект '{project.name}' загружен")
        except Exception as e:
            error_msg = f"Ошибка при загрузке проекта: {e}"
            logger.error(error_msg, exc_info=True)
            self.status_bar.setVisible(True)
            self.status_bar.showMessage(error_msg)
            self.progress_bar.setVisible(False)
    
    def load_project_data_to_tree(self, project):
        """Загрузка данных проекта в древовидное представление"""
        try:
            if not project:
                self.status_bar.showMessage("Проект не выбран")
                return
            
            if not project.data:
                self.status_bar.showMessage("В проекте нет данных для отображения")
                self.data_tree.clear()
                self.data_table.clear()
                self.data_table.setRowCount(0)
                self.data_table.setColumnCount(0)
                return
            
            # Очищаем дерево
            self.data_tree.clear()
            
            # Загружаем данные текущего раздела
            section_map = {
                "Доходы": "доходы_data",
                "Расходы": "расходы_data", 
                "Источники финансирования": "источники_финансирования_data",
                "Консолидируемые расчеты": "консолидируемые_расчеты_data"
            }

            # Настраиваем заголовки дерева под выбранный раздел
            self.configure_tree_headers(self.current_section)
            
            section_key = section_map.get(self.current_section)
            if section_key and section_key in project.data:
                data = project.data[section_key]
                if data and len(data) > 0:
                    # Для раздела "Расходы" обновляем строку с кодом 450 расчетными значениями
                    # из результат_исполнения_data для подсветки ошибок в расчетах
                    if self.current_section == "Расходы" and project.data.get('результат_исполнения_data'):
                        результат_data = project.data['результат_исполнения_data']
                        # Ищем строку с кодом 450
                        for row in data:
                            if str(row.get('код_строки', '')).strip() == '450':
                                # Добавляем расчетные значения для проверки несоответствий
                                for col in Form0503317Constants.BUDGET_COLUMNS:
                                    row[f'расчетный_утвержденный_{col}'] = результат_data.get('утвержденный', {}).get(col, 0)
                                    row[f'расчетный_исполненный_{col}'] = результат_data.get('исполненный', {}).get(col, 0)
                                break
                    
                    self.build_tree_from_data(data)
                    self.load_project_data_to_table(section_key, data)
                    # Обновляем высоту заголовка после загрузки данных
                    QTimer.singleShot(100, self._update_tree_header_height)
                    self.status_bar.showMessage(f"Загружено {len(data)} записей в разделе '{self.current_section}'")
                else:
                    # Если данных нет, очищаем таблицу и показываем сообщение
                    self.data_table.clear()
                    self.data_table.setRowCount(0)
                    self.data_table.setColumnCount(0)
                    self.status_bar.showMessage(f"В разделе '{self.current_section}' нет данных для отображения")
            else:
                # Если данных нет, очищаем таблицу и показываем сообщение
                self.data_table.clear()
                self.data_table.setRowCount(0)
                self.data_table.setColumnCount(0)
                self.status_bar.showMessage(f"Раздел '{self.current_section}' не найден в данных проекта")
        except Exception as e:
            error_msg = f"Ошибка загрузки данных в дерево: {e}"
            logger.error(error_msg, exc_info=True)
            self.status_bar.showMessage(error_msg)

    def load_project_data_to_table(self, section_key: str, data):
        """Загрузка данных проекта в табличное представление (все столбцы)"""
        self.data_table.clear()

        if not data:
            self.data_table.setRowCount(0)
            self.data_table.setColumnCount(0)
            return

        # Общие колонки
        base_headers = ["Наименование", "Код строки", "Код классификации", "Уровень"]

        if section_key == "консолидируемые_расчеты_data":
            # Используем CONSOLIDATED_COLUMNS
            cons_cols = Form0503317Constants.CONSOLIDATED_COLUMNS
            headers = base_headers + cons_cols
            self.data_table.setColumnCount(len(headers))
            self.data_table.setHorizontalHeaderLabels(headers)

            self.data_table.setRowCount(len(data))
            error_color = QColor("#FF6B6B")

            for row_idx, item in enumerate(data):
                self.data_table.setItem(row_idx, 0, QTableWidgetItem(str(item.get("наименование_показателя", ""))))
                self.data_table.setItem(row_idx, 1, QTableWidgetItem(str(item.get("код_строки", ""))))
                self.data_table.setItem(row_idx, 2, QTableWidgetItem(str(item.get("код_классификации", ""))))
                self.data_table.setItem(row_idx, 3, QTableWidgetItem(str(item.get("уровень", 0))))

                # Оригинальные значения поступлений (вложенный словарь или плоские поля)
                поступления = item.get("поступления", {}) or {}

                for col_idx, col_name in enumerate(cons_cols, start=len(base_headers)):
                    original_value = (
                        поступления.get(col_name, 0)
                        if isinstance(поступления, dict) else item.get(f"поступления_{col_name}", 0)
                    )
                    calculated_value = item.get(f"расчетный_поступления_{col_name}")
                    if calculated_value is None:
                        calculated_value = original_value

                    cell = QTableWidgetItem()

                    # Отображаем расхождения так же, как в дереве: значение и расчет в скобках
                    # Для консолидированных расчетов проверяем на всех уровнях (как в дереве)
                    level = item.get("уровень", 0)
                    is_total_column = (col_name == 'ИТОГО')
                    should_check = (level < 6) or is_total_column
                    
                    if should_check and self._is_value_different(original_value, calculated_value):
                        if isinstance(original_value, (int, float)) and isinstance(calculated_value, (int, float)):
                            display_value = f"{original_value:,.2f} ({calculated_value:,.2f})"
                        else:
                            display_value = f"{original_value} ({calculated_value})"
                        cell.setText(display_value)
                        cell.setForeground(QBrush(error_color))
                    else:
                        cell.setText(self.format_budget_value(original_value))

                    self.data_table.setItem(row_idx, col_idx, cell)

            self.hide_zero_columns_in_table(section_key, data)

            # Ограничиваем ширину столбцов
            header = self.data_table.horizontalHeader()
            max_width = max(80, self.width() // 8 if self.width() > 0 else 200)
            for i in range(self.data_table.columnCount()):
                header.setSectionResizeMode(i, QHeaderView.Interactive)
                header.resizeSection(i, min(header.sectionSize(i), max_width))

            # Высота заголовка в зависимости от количества строк в названиях колонок
            font_metrics = header.fontMetrics()
            max_lines = 1
            for text in headers:
                lines = text.count("\n") + 1
                if lines > max_lines:
                    max_lines = lines
            line_height = font_metrics.lineSpacing()
            header.setFixedHeight(line_height * max_lines + 6)
            return

        # Для доходов, расходов и источников используем BUDGET_COLUMNS
        budget_cols = Form0503317Constants.BUDGET_COLUMNS
        approved_headers = [f"Утв: {col}" for col in budget_cols]
        executed_headers = [f"Исп: {col}" for col in budget_cols]

        headers = base_headers + approved_headers + executed_headers
        self.data_table.setColumnCount(len(headers))
        self.data_table.setHorizontalHeaderLabels(headers)

        self.data_table.setRowCount(len(data))
        for row_idx, item in enumerate(data):
            self.data_table.setItem(row_idx, 0, QTableWidgetItem(str(item.get("наименование_показателя", ""))))
            self.data_table.setItem(row_idx, 1, QTableWidgetItem(str(item.get("код_строки", ""))))
            self.data_table.setItem(row_idx, 2, QTableWidgetItem(str(item.get("код_классификации", ""))))
            self.data_table.setItem(row_idx, 3, QTableWidgetItem(str(item.get("уровень", 0))))

            approved = item.get("утвержденный", {}) or {}
            executed = item.get("исполненный", {}) or {}

            # Утвержденные
            for i, col_name in enumerate(budget_cols):
                value = approved.get(col_name, 0)
                text = "" if value in (None, "x") else f"{value:,.2f}" if isinstance(value, (int, float)) else str(value)
                self.data_table.setItem(row_idx, len(base_headers) + i, QTableWidgetItem(text))

            # Исполненные
            offset = len(base_headers) + len(budget_cols)
            for i, col_name in enumerate(budget_cols):
                value = executed.get(col_name, 0)
                text = "" if value in (None, "x") else f"{value:,.2f}" if isinstance(value, (int, float)) else str(value)
                self.data_table.setItem(row_idx, offset + i, QTableWidgetItem(text))

        self.hide_zero_columns_in_table(section_key, data)

        # Ограничиваем ширину столбцов
        header = self.data_table.horizontalHeader()
        max_width = max(80, self.width() // 8 if self.width() > 0 else 200)
        for i in range(self.data_table.columnCount()):
            header.setSectionResizeMode(i, QHeaderView.Interactive)
            header.resizeSection(i, min(header.sectionSize(i), max_width))

        # Высота заголовка в зависимости от количества строк в названиях колонок
        font_metrics = header.fontMetrics()
        max_lines = 1
        for text in headers:
            lines = text.count("\n") + 1
            if lines > max_lines:
                max_lines = lines
        line_height = font_metrics.lineSpacing()
        header.setFixedHeight(line_height * max_lines + 6)

        # Применяем такое же скрытие нулевых столбцов к дереву
        self.hide_zero_columns_in_tree(section_key, data)

    def hide_zero_columns_in_table(self, section_key: str, data):
        """
        Автоматическое скрытие столбцов, в которых итоговое значение равно 0.
        Для доходов/расходов/источников ищем строку '...всего', для
        консолидируемых расчетов — строку с 'итого' или кодом 899.
        """
        base_offset = 4  # Наименование, код строки, код классификации, уровень

        if section_key == "консолидируемые_расчеты_data":
            cons_cols = Form0503317Constants.CONSOLIDATED_COLUMNS
            # Ищем итоговую строку
            total_item = None
            for item in data:
                name = str(item.get("наименование_показателя", "")).lower()
                code = str(item.get("код_строки", "")).lower()
                if "итого" in name or code == "899":
                    total_item = item
                    break
            if not total_item:
                return

            поступления = total_item.get("поступления", {}) or {}
            for i, col_name in enumerate(cons_cols):
                value = поступления.get(col_name, 0)
                if isinstance(value, (int, float)) and abs(value) < 1e-9:
                    col_index = base_offset + i
                    if 0 <= col_index < self.data_table.columnCount():
                        self.data_table.horizontalHeader().setSectionHidden(col_index, True)
            return

        # Доходы, расходы, источники
        budget_cols = Form0503317Constants.BUDGET_COLUMNS
        total_item = None
        for item in data:
            name = str(item.get("наименование_показателя", "")).lower()
            if "всего" in name:
                total_item = item
                break
        if not total_item:
            return

        approved = total_item.get("утвержденный", {}) or {}
        executed = total_item.get("исполненный", {}) or {}

        for i, col_name in enumerate(budget_cols):
            a_val = approved.get(col_name, 0) or 0
            e_val = executed.get(col_name, 0) or 0
            if isinstance(a_val, (int, float)) and isinstance(e_val, (int, float)):
                if abs(a_val) < 1e-9 and abs(e_val) < 1e-9:
                    # Скрываем и утвержденный, и исполненный столбец для этой колонки
                    approved_col_index = base_offset + i
                    executed_col_index = base_offset + len(budget_cols) + i
                    if 0 <= approved_col_index < self.data_table.columnCount():
                        self.data_table.horizontalHeader().setSectionHidden(approved_col_index, True)
                    if 0 <= executed_col_index < self.data_table.columnCount():
                        self.data_table.horizontalHeader().setSectionHidden(executed_col_index, True)

    def configure_tree_headers(self, section_name: str):
        """Конфигурация заголовков дерева под выбранный раздел"""
        base_headers = ["Наименование", "Код строки", "Код классификации", "Уровень"]
        display_headers = base_headers[:]
        tooltip_headers = base_headers[:]
        mapping = {
            "type": "base",
            "base_count": len(base_headers)
        }

        if section_name in ["Доходы", "Расходы", "Источники финансирования"]:
            budget_cols = Form0503317Constants.BUDGET_COLUMNS
            mapping.update({
                "type": "budget",
                "budget_columns": budget_cols,
                "approved_start": len(display_headers),
                "executed_start": len(display_headers) + len(budget_cols)
            })

            for col in budget_cols:
                display_headers.append(f"Утв:\n{col}")
                tooltip_headers.append(f"Утвержденный — {col}")
            for col in budget_cols:
                display_headers.append(f"Исп:\n{col}")
                tooltip_headers.append(f"Исполненный — {col}")

        elif section_name == "Консолидируемые расчеты":
            cons_cols = Form0503317Constants.CONSOLIDATED_COLUMNS
            mapping.update({
                "type": "consolidated",
                "value_start": len(display_headers),
                "columns": cons_cols
            })
            for col in cons_cols:
                display_headers.append(col)
                tooltip_headers.append(col)

        self.tree_headers = display_headers
        self.tree_header_tooltips = tooltip_headers
        self.tree_column_mapping = mapping

        self.data_tree.setColumnCount(len(display_headers))
        self.data_tree.setHeaderLabels(display_headers)
        header = self.data_tree.header()
        header.setDefaultAlignment(Qt.AlignCenter)

        # Ограничиваем ширину колонок
        max_width = max(80, self.width() // 8 if self.width() > 0 else 200)
        for idx in range(len(display_headers)):
            header.setSectionResizeMode(idx, QHeaderView.Interactive)
            header.resizeSection(idx, min(header.sectionSize(idx), max_width))

        # Вычисляем высоту заголовка с учетом автоматического переноса текста
        # Используем QTimer для обновления после того, как заголовки будут установлены
        QTimer.singleShot(50, self._update_tree_header_height)
        
        # Подключаем обработчик изменения размера столбцов для обновления высоты заголовка
        try:
            if hasattr(header, 'sectionResized'):
                try:
                    header.sectionResized.disconnect(self._on_tree_header_section_resized)
                except:
                    pass
                header.sectionResized.connect(self._on_tree_header_section_resized)
        except Exception as e:
            logger.warning(f"Ошибка подключения обработчика sectionResized: {e}", exc_info=True)

        # Для консолидируемых расчетов колонку "Код классификации" не показываем
        if section_name == "Консолидируемые расчеты" and len(display_headers) > 2:
            self.data_tree.setColumnHidden(2, True)

    def _update_tree_header_height(self):
        """Обновляет высоту заголовка дерева с учетом автоматического переноса текста"""
        try:
            header = self.data_tree.header()
            font_metrics = header.fontMetrics()
            max_lines = 1
            
            # Получаем заголовки из headerItem
            header_item = self.data_tree.headerItem()
            if header_item:
                # Проходим по всем заголовкам и вычисляем максимальное количество строк
                for idx in range(self.data_tree.columnCount()):
                    if self.data_tree.isColumnHidden(idx):
                        continue
                    
                    # Получаем текст из headerItem
                    text = header_item.text(idx) if idx < self.data_tree.columnCount() else ""
                    if not text and idx < len(self.tree_headers):
                        text = self.tree_headers[idx]
                    
                    if text:
                        # Получаем ширину столбца
                        width = max(header.sectionSize(idx), 50)
                        
                        # Стандартный расчет по количеству явных переносов строк
                        lines = str(text).count("\n") + 1
                        max_lines = max(max_lines, lines)
            else:
                # Если нет headerItem, используем tree_headers
                for text in self.tree_headers:
                    if text:
                        lines = str(text).count("\n") + 1
                        max_lines = max(max_lines, lines)
            
            line_height = font_metrics.lineSpacing()
            new_height = line_height * max_lines + 6
            header.setFixedHeight(new_height)
        except Exception as e:
            logger.warning(f"Ошибка обновления высоты заголовка дерева: {e}", exc_info=True)
            # В случае ошибки используем минимальную высоту
            try:
                header = self.data_tree.header()
                font_metrics = header.fontMetrics()
                header.setFixedHeight(font_metrics.lineSpacing() + 6)
            except:
                pass
    
    def _on_tree_header_section_resized(self, logicalIndex, oldSize, newSize):
        """Обработчик изменения размера столбца заголовка дерева"""
        # Обновляем высоту заголовка при изменении размера столбца
        try:
            QTimer.singleShot(100, self._update_tree_header_height)
        except Exception as e:
            logger.warning(f"Ошибка в _on_tree_header_section_resized: {e}", exc_info=True)

    def hide_zero_columns_in_tree(self, section_key: str, data):
        """
        Скрытие столбцов дерева, в которых итоговое значение равно 0.
        Логика аналогична табличному представлению.
        """
        if not data:
            return

        if section_key == "консолидируемые_расчеты_data":
            cons_cols = Form0503317Constants.CONSOLIDATED_COLUMNS
            mapping = self.tree_column_mapping or {}
            if mapping.get("type") != "consolidated":
                return

            # Ищем итоговую строку
            total_item = None
            for item in data:
                name = str(item.get("наименование_показателя", "")).lower()
                code = str(item.get("код_строки", "")).lower()
                if "итого" in name or code == "899":
                    total_item = item
                    break
            if not total_item:
                return

            value_start = mapping.get("value_start", 4)
            totals = total_item.get("поступления", {}) or {}

            for i, col_name in enumerate(cons_cols):
                val = totals.get(col_name, 0)
                if isinstance(val, (int, float)) and abs(val) < 1e-9:
                    col_index = value_start + i
                    if 0 <= col_index < self.data_tree.columnCount():
                        self.data_tree.setColumnHidden(col_index, True)
            return

        # Доходы, расходы, источники
        budget_cols = Form0503317Constants.BUDGET_COLUMNS
        mapping = self.tree_column_mapping or {}
        if mapping.get("type") != "budget":
            return

        total_item = None
        for item in data:
            name = str(item.get("наименование_показателя", "")).lower()
            if "всего" in name:
                total_item = item
                break
        if not total_item:
            return

        approved = total_item.get("утвержденный", {}) or {}
        executed = total_item.get("исполненный", {}) or {}

        approved_start = mapping.get("approved_start", 4)
        executed_start = mapping.get("executed_start", approved_start + len(budget_cols))

        for i, col_name in enumerate(budget_cols):
            a_val = approved.get(col_name, 0) or 0
            e_val = executed.get(col_name, 0) or 0
            if isinstance(a_val, (int, float)) and isinstance(e_val, (int, float)):
                if abs(a_val) < 1e-9 and abs(e_val) < 1e-9:
                    appr_idx = approved_start + i
                    exec_idx = executed_start + i
                    if 0 <= appr_idx < self.data_tree.columnCount():
                        self.data_tree.setColumnHidden(appr_idx, True)
                    if 0 <= exec_idx < self.data_tree.columnCount():
                        self.data_tree.setColumnHidden(exec_idx, True)
        header_item = self.data_tree.headerItem()
        if header_item:
            for idx, tip in enumerate(self.tree_header_tooltips):
                if idx < self.data_tree.columnCount():
                    header_item.setToolTip(idx, tip)
                    # Убеждаемся, что текст заголовка установлен
                    if idx < len(self.tree_headers):
                        current_text = header_item.text(idx)
                        if not current_text or current_text != self.tree_headers[idx]:
                            header_item.setText(idx, self.tree_headers[idx])

        # Применяем отображение колонок в зависимости от выбранного типа данных
        self.apply_tree_data_type_visibility()

    def apply_tree_data_type_visibility(self):
        """Скрывает столбцы дерева в зависимости от выбранного типа данных"""
        if not self.tree_column_mapping:
            return

        column_total = len(self.tree_headers)
        for col in range(column_total):
            self.data_tree.setColumnHidden(col, False)

        if self.tree_column_mapping.get("type") != "budget":
            return

        approved_start = self.tree_column_mapping.get("approved_start", 0)
        executed_start = self.tree_column_mapping.get("executed_start", 0)
        budget_cols = self.tree_column_mapping.get("budget_columns", [])

        approved_range = range(approved_start, approved_start + len(budget_cols))
        executed_range = range(executed_start, executed_start + len(budget_cols))

        show_approved = self.current_data_type in ("Утвержденный", "Оба")
        show_executed = self.current_data_type in ("Исполненный", "Оба")

        for idx in approved_range:
            self.data_tree.setColumnHidden(idx, not show_approved)
        for idx in executed_range:
            self.data_tree.setColumnHidden(idx, not show_executed)

    def format_budget_value(self, value):
        """Форматирование значения бюджета для отображения"""
        if value in (None, "", "0", 0):
            return ""
        if value == 'x':
            return 'x'
        try:
            return f"{float(value):,.2f}"
        except (ValueError, TypeError):
            return str(value)
    
    def build_tree_from_data(self, data):
        """Построение дерева из данных"""
        try:
            if not data:
                self.status_bar.showMessage("Нет данных для построения дерева")
                return
            
            if not isinstance(data, list) or len(data) == 0:
                self.status_bar.showMessage("Данные пусты или имеют неверный формат")
                return
            
            # Цвета для уровней
            level_colors = {
                0: "#E6E6FA", 1: "#68e368", 2: "#98FB98", 3: "#FFFF99", 
                4: "#FFB366", 5: "#FF9999", 6: "#FFCCCC"
            }
            
            # Строим дерево, учитывая последовательность уровней:
            # каждая строка является дочерней для ближайшей предыдущей строки
            # с меньшим уровнем (обычно level-1).
            parents_stack = []  # список кортежей (level, QTreeWidgetItem)
            items_created = 0
            items_failed = 0

            for item in data:
                try:
                    if not isinstance(item, dict):
                        items_failed += 1
                        continue
                    
                    level = item.get('уровень', 0)
                    tree_item = self.create_tree_item(item, level_colors)
                
                    # Убираем из стека все уровни, которые не могут быть родителями
                    while parents_stack and parents_stack[-1][0] >= level:
                        parents_stack.pop()

                    if parents_stack:
                        # Текущий элемент становится ребёнком последнего подходящего родителя
                        parents_stack[-1][1].addChild(tree_item)
                    else:
                        # Если родителя нет, это корневой элемент
                        self.data_tree.addTopLevelItem(tree_item)

                    # Запоминаем текущий элемент как последний для своего уровня
                    parents_stack.append((level, tree_item))
                    items_created += 1
                except Exception as e:
                    items_failed += 1
                    logger.warning(f"Ошибка создания элемента дерева: {e}", exc_info=True)
                    continue
            
            # Разворачиваем уровень 0
            for i in range(self.data_tree.topLevelItemCount()):
                try:
                    self.data_tree.topLevelItem(i).setExpanded(True)
                except:
                    pass
            
            if items_created > 0:
                msg = f"Построено дерево: {items_created} элементов"
                if items_failed > 0:
                    msg += f", ошибок: {items_failed}"
                self.status_bar.showMessage(msg)
            else:
                self.status_bar.showMessage("Не удалось построить дерево: все элементы содержат ошибки")
        except Exception as e:
            error_msg = f"Ошибка построения дерева: {e}"
            logger.error(error_msg, exc_info=True)
            self.status_bar.showMessage(error_msg)
    
    def create_tree_item(self, item, level_colors):
        """Создание элемента дерева"""
        try:
            level = item.get('уровень', 0)

            column_count = self.data_tree.columnCount()
            if column_count == 0:
                # Если колонок нет, создаем хотя бы одну
                self.data_tree.setColumnCount(1)
                column_count = 1
            
            tree_item = QTreeWidgetItem([""] * column_count)
            
            # Основные данные
            name = str(item.get('наименование_показателя', ''))
            code_line = str(item.get('код_строки', ''))
            class_code = str(item.get('код_классификации_форматированный', item.get('код_классификации', '')))

            if column_count > 0:
                tree_item.setText(0, name)
            if column_count > 1:
                tree_item.setText(1, code_line)
            if column_count > 2:
                tree_item.setText(2, class_code)
            if column_count > 3:
                tree_item.setText(3, str(level))

            mapping = self.tree_column_mapping or {}
            column_type = mapping.get("type", "base")

            if column_type == "budget":
                budget_cols = mapping.get("budget_columns", [])
                approved_start = mapping.get("approved_start", 4)
                executed_start = mapping.get("executed_start", approved_start + len(budget_cols))
                approved_data = item.get('утвержденный', {}) or {}
                executed_data = item.get('исполненный', {}) or {}
                
                # Цвет для выделения несоответствий (красный)
                error_color = QColor("#FF6B6B")

                for idx, col in enumerate(budget_cols):
                    try:
                        # Утвержденные значения
                        original_approved = approved_data.get(col, 0) or 0
                        calculated_approved = item.get(f'расчетный_утвержденный_{col}', original_approved)
                        
                        # Проверяем несоответствие (только для уровней < 6)
                        if level < 6 and self._is_value_different(original_approved, calculated_approved):
                            # Показываем значение с расчетным в скобках
                            if isinstance(original_approved, (int, float)) and isinstance(calculated_approved, (int, float)):
                                approved_value = f"{original_approved:,.2f} ({calculated_approved:,.2f})"
                            else:
                                approved_value = f"{original_approved} ({calculated_approved})"
                            # Выделяем красным цветом
                            if approved_start + idx < column_count:
                                tree_item.setText(approved_start + idx, approved_value)
                                tree_item.setForeground(approved_start + idx, QBrush(error_color))
                        else:
                            approved_value = self.format_budget_value(original_approved)
                            if approved_start + idx < column_count:
                                tree_item.setText(approved_start + idx, approved_value)
                        
                        # Исполненные значения
                        original_executed = executed_data.get(col, 0) or 0
                        calculated_executed = item.get(f'расчетный_исполненный_{col}', original_executed)
                        
                        # Проверяем несоответствие (только для уровней < 6)
                        if level < 6 and self._is_value_different(original_executed, calculated_executed):
                            # Показываем значение с расчетным в скобках
                            if isinstance(original_executed, (int, float)) and isinstance(calculated_executed, (int, float)):
                                executed_value = f"{original_executed:,.2f} ({calculated_executed:,.2f})"
                            else:
                                executed_value = f"{original_executed} ({calculated_executed})"
                            # Выделяем красным цветом
                            if executed_start + idx < column_count:
                                tree_item.setText(executed_start + idx, executed_value)
                                tree_item.setForeground(executed_start + idx, QBrush(error_color))
                        else:
                            executed_value = self.format_budget_value(original_executed)
                            if executed_start + idx < column_count:
                                tree_item.setText(executed_start + idx, executed_value)
                    except Exception as e:
                        logger.warning(f"Ошибка обработки несоответствий для колонки {col}: {e}", exc_info=True)
                        pass

            elif column_type == "consolidated":
                value_start = mapping.get("value_start", 4)
                cons_cols = mapping.get("columns", [])
                
                # Получаем данные поступлений (может быть вложенным словарем или плоскими полями)
                cons_data = item.get('поступления', {}) or {}
                
                # Цвет для выделения несоответствий (красный)
                error_color = QColor("#FF6B6B")
                
                for idx, col in enumerate(cons_cols):
                    try:
                        # Оригинальное значение - проверяем и вложенный словарь, и плоские поля
                        if isinstance(cons_data, dict) and col in cons_data:
                            original_value = cons_data.get(col, 0) or 0
                        else:
                            # Если нет вложенного словаря, проверяем плоские поля
                            original_value = item.get(f'поступления_{col}', 0) or 0
                        
                        # Расчетное значение - проверяем плоские поля (после to_dict('records'))
                        calculated_value = item.get(f'расчетный_поступления_{col}')
                        if calculated_value is None:
                            # Fallback на оригинальное значение, если расчетного нет
                            calculated_value = original_value
                        
                        # Проверяем несоответствие (аналогично бюджетным разделам — до 5 уровня),
                        # а для столбца "ИТОГО" проверяем на всех уровнях, так как это итоговая сумма
                        is_total_column = (col == 'ИТОГО')
                        should_check = (level < 6) or is_total_column
                        
                        if should_check and self._is_value_different(original_value, calculated_value):
                            # Показываем значение с расчетным в скобках
                            if isinstance(original_value, (int, float)) and isinstance(calculated_value, (int, float)):
                                display_value = f"{original_value:,.2f} ({calculated_value:,.2f})"
                            else:
                                display_value = f"{original_value} ({calculated_value})"
                            # Выделяем красным цветом
                            if value_start + idx < column_count:
                                tree_item.setText(value_start + idx, display_value)
                                tree_item.setForeground(value_start + idx, QBrush(error_color))
                        else:
                            # Обычное отображение без несоответствий
                            if value_start + idx < column_count:
                                tree_item.setText(value_start + idx, self.format_budget_value(original_value))
                    except Exception as e:
                        logger.warning(f"Ошибка обработки несоответствий для консолидируемых расчетов, колонка {col}: {e}", exc_info=True)
                        pass
            
            # Устанавливаем цвет фона
            try:
                if level in level_colors:
                    color = QColor(level_colors[level])
                    for i in range(min(tree_item.columnCount(), column_count)):
                        tree_item.setBackground(i, QBrush(color))
            except:
                pass
            
            # Устанавливаем подсказки (колонка -> заголовок)
            try:
                for idx, tip in enumerate(self.tree_header_tooltips):
                    if idx < tree_item.columnCount() and idx < len(self.tree_header_tooltips):
                        current_text = tree_item.text(idx)
                        if current_text:
                            tree_item.setToolTip(idx, f"{tip}: {current_text}")
                        else:
                            tree_item.setToolTip(idx, tip)
            except:
                pass

            # Сохраняем исходные данные
            try:
                tree_item.setData(0, Qt.UserRole, item)
            except:
                pass
            
            return tree_item
        except Exception as e:
            logger.error(f"Ошибка создания элемента дерева: {e}", exc_info=True)
            # Возвращаем пустой элемент в случае ошибки
            column_count = max(self.data_tree.columnCount(), 1)
            tree_item = QTreeWidgetItem([""] * column_count)
            return tree_item
    
    def _is_value_different(self, original: float, calculated: float) -> bool:
        """Проверка различия значений (аналогично методу в Form0503317)"""
        try:
            original_val = float(original) if original not in (None, "", "x") else 0.0
            calculated_val = float(calculated) if calculated not in (None, "", "x") else 0.0
            return abs(original_val - calculated_val) > 0.00001
        except (ValueError, TypeError):
            return False
    
    def load_metadata(self, project):
        """Загрузка метаданных для выбранной ревизии"""
        # Метаданные должны быть только у ревизии, а не у проекта
        # Проверяем, что загружена ревизия (current_revision_id установлен)
        rev_id = getattr(self.controller, "current_revision_id", None)
        if not rev_id:
            # Если ревизия не загружена, метаданные не отображаем
            self.metadata_text.setHtml("")
            return
        
        # Метаданные берём из данных проекта (которые загружаются из ревизии)
        if not project or not project.data:
            self.metadata_text.setHtml("")
            return
        
        meta_info = project.data.get('meta_info', {})
        if not meta_info:
            self.metadata_text.setHtml("")
            return
        
        metadata_text = ""
        for key, value in meta_info.items():
            metadata_text += f"<b>{key}:</b> {value}<br>"
        self.metadata_text.setHtml(metadata_text)
    
    def on_section_changed(self, section_name):
        """Обработка смены раздела"""
        self.current_section = section_name
        if self.controller.current_project:
            self.load_project_data_to_tree(self.controller.current_project)
    
    def on_data_type_changed(self, data_type):
        """Обработка смены типа данных"""
        self.current_data_type = data_type
        self.apply_tree_data_type_visibility()
        if self.controller.current_project:
            self.load_project_data_to_tree(self.controller.current_project)
    
    def expand_all_tree(self):
        """Развернуть все узлы дерева"""
        self.data_tree.expandAll()
    
    def collapse_all_tree(self):
        """Свернуть все узлы дерева"""
        self.data_tree.collapseAll()
    
    def on_tree_item_expanded(self, item):
        """Обработка разворачивания узла дерева"""
        pass
    
    def on_tree_item_collapsed(self, item):
        """Обработка сворачивания узла дерева"""
        pass
    
    def show_tree_context_menu(self, position):
        """Контекстное меню для дерева"""
        item = self.data_tree.itemAt(position)
        if not item:
            return
        
        menu = QMenu()
        copy_action = menu.addAction("Копировать значение")
        
        action = menu.exec_(self.data_tree.mapToGlobal(position))
        
        if action == copy_action:
            self.copy_tree_item_value(item)

    def show_tree_header_context_menu(self, position):
        """Контекстное меню для заголовков дерева (скрытие/отображение столбцов)"""
        header = self.data_tree.header()
        col = header.logicalIndexAt(position)
        if col < 0:
            return

        menu = QMenu(self)
        hide_action = menu.addAction("Скрыть столбец")
        show_all_action = menu.addAction("Показать все столбцы")
        chosen = menu.exec_(header.mapToGlobal(position))

        if chosen == hide_action:
            # Не скрываем первый столбец с названием
            if col > 0:
                self.data_tree.setColumnHidden(col, True)
        elif chosen == show_all_action:
            for i in range(self.data_tree.columnCount()):
                self.data_tree.setColumnHidden(i, False)
    
    def show_table_context_menu(self, position):
        """Контекстное меню для таблицы"""
        menu = QMenu()
        
        hide_column_action = menu.addAction("Скрыть столбец")
        show_all_columns_action = menu.addAction("Показать все столбцы")
        menu.addSeparator()
        hide_zero_columns_action = menu.addAction("Скрыть нулевые столбцы")
        menu.addSeparator()
        copy_action = menu.addAction("Копировать значение")
        
        action = menu.exec_(self.data_table.mapToGlobal(position))
        
        if action == hide_column_action:
            self.hide_current_column()
        elif action == show_all_columns_action:
            self.show_all_columns()
        elif action == hide_zero_columns_action:
            # Повторно применяем логику скрытия нулевых столбцов для текущих данных
            if self.controller.current_project and self.controller.current_project.data:
                section_map = {
                    "Доходы": "доходы_data",
                    "Расходы": "расходы_data", 
                    "Источники финансирования": "источники_финансирования_data",
                    "Консолидируемые расчеты": "консолидируемые_расчеты_data"
                }
                section_key = section_map.get(self.current_section)
                if section_key and section_key in self.controller.current_project.data:
                    data = self.controller.current_project.data[section_key]
                    # Сначала показываем все, затем снова скрываем нулевые
                    self.show_all_columns()
                    self.hide_zero_columns_in_table(section_key, data)
        elif action == copy_action:
            self.copy_table_cell_value()
    
    def hide_current_column(self):
        """Скрыть текущий столбец"""
        current_column = self.data_table.currentColumn()
        if current_column >= 0:
            self.data_table.horizontalHeader().setSectionHidden(current_column, True)
    
    def show_all_columns(self):
        """Показать все столбцы"""
        for i in range(self.data_table.columnCount()):
            self.data_table.horizontalHeader().setSectionHidden(i, False)
        # Показать все столбцы и в дереве (кроме скрытого кода для консолидированных)
        header = self.data_tree.header()
        for i in range(self.data_tree.columnCount()):
            self.data_tree.setColumnHidden(i, False)
        if self.current_section == "Консолидируемые расчеты" and self.data_tree.columnCount() > 2:
            self.data_tree.setColumnHidden(2, True)

    def hide_zero_columns_global(self):
        """Сворачивает все столбцы с нулевыми значениями в итоговой строке"""
        if not (self.controller.current_project and self.controller.current_project.data):
            return

        section_map = {
            "Доходы": "доходы_data",
            "Расходы": "расходы_data", 
            "Источники финансирования": "источники_финансирования_data",
            "Консолидируемые расчеты": "консолидируемые_расчеты_data"
        }
        section_key = section_map.get(self.current_section)
        if not section_key or section_key not in self.controller.current_project.data:
            return

        data = self.controller.current_project.data[section_key]

        # Сначала показываем все столбцы
        self.show_all_columns()
        for i in range(self.data_tree.columnCount()):
            self.data_tree.setColumnHidden(i, False)
        if self.current_section == "Консолидируемые расчеты" and self.data_tree.columnCount() > 2:
            self.data_tree.setColumnHidden(2, True)

        # Применяем скрытие нулевых столбцов
        self.hide_zero_columns_in_table(section_key, data)
        self.hide_zero_columns_in_tree(section_key, data)
    
    def copy_tree_item_value(self, item):
        """Копировать значение из дерева"""
        if item:
            text = item.text(0)  # Копируем значение из первого столбца
            clipboard = QApplication.clipboard()
            clipboard.setText(text)
    
    def copy_table_cell_value(self):
        """Копировать значение из таблицы"""
        current_item = self.data_table.currentItem()
        if current_item:
            clipboard = QApplication.clipboard()
            clipboard.setText(current_item.text())
    
    def show_new_project_dialog(self):
        """Показать диалог создания проекта"""
        dialog = ProjectDialog(self)
        if dialog.exec_():
            project_data = dialog.get_project_data()
            project = self.controller.create_project(project_data)
            if project:
                QMessageBox.information(self, "Успех", f"Проект '{project.name}' создан")
    
    def show_reference_dialog(self, ref_type: str = None):
        """Показать диалог загрузки справочника"""
        dialog = ReferenceDialog(self, ref_type)
        if dialog.exec_():
            ref_data = dialog.get_reference_data()
            success = self.controller.load_reference_file(
                ref_data['file_path'],
                ref_data['reference_type'],
                ref_data['name']
            )
            if success:
                QMessageBox.information(self, "Успех", "Справочник загружен")
    
    def show_reference_viewer(self):
        """Показать просмотрщик справочников в отдельном окне"""
        from PyQt5.QtWidgets import QMainWindow

        if self.reference_window is None:
            self.reference_window = QMainWindow(self)
            self.reference_window.setWindowTitle("Справочники")
            self.reference_window.resize(900, 600)

            self.reference_viewer = ReferenceViewer()
            self.reference_window.setCentralWidget(self.reference_viewer)

        # Загружаем актуальные справочники и показываем окно
        self.reference_viewer.load_references(self.controller.references)
        self.reference_window.show()
        self.reference_window.raise_()
        self.reference_window.activateWindow()

    def show_config_dictionaries(self):
        """Показать диалог редактирования справочников конфигурации"""
        dlg = DictionariesDialog(self.controller.db_manager, self)
        dlg.exec_()

    def on_projects_side_button_clicked(self):
        """Обработчик клика по боковой кнопке панели проектов"""
        if not self.projects_inner_panel:
            return
        # Инвертируем состояние по видимости внутренней панели
        self.toggle_projects_panel(not self.projects_inner_panel.isVisible())

    def toggle_projects_panel(self, checked: bool = None):
        """Показать/скрыть панель проектов"""
        if not self.main_splitter or not self.projects_inner_panel:
            return
        
        # Если checked не указан, инвертируем текущее состояние
        if checked is None:
            checked = not self.projects_inner_panel.isVisible()
        else:
            # Обновляем состояние меню
            for action in self.menuBar().actions():
                if action.text() == "&Вид":
                    for sub_action in action.menu().actions():
                        if sub_action.text() == "&Панель проектов":
                            sub_action.setChecked(checked)
                            break
                    break

        if not checked:
            # Запоминаем текущую ширину панели перед схлопыванием
            sizes = self.main_splitter.sizes()
            if sizes and sizes[0] > 0:
                self.projects_panel_last_size = sizes[0]

            # Скрываем содержимое, оставляя узкую кнопку
            self.projects_inner_panel.setVisible(False)
            if self.projects_toggle_button:
                self.projects_toggle_button.setText("▶")

            handle_width = self.projects_toggle_button.width() if self.projects_toggle_button else 20
            self.main_splitter.setSizes([handle_width, max(400, self.width() - handle_width)])
        else:
            # Показываем внутреннюю панель
            self.projects_inner_panel.setVisible(True)
            if self.projects_toggle_button:
                self.projects_toggle_button.setText("◀")

            total_width = max(self.width(), self.projects_panel_last_size + 400)
            self.main_splitter.setSizes(
                [self.projects_panel_last_size, total_width - self.projects_panel_last_size]
            )
    
    def load_form_file(self):
        """Загрузка файла формы"""
        # Если проект не выбран, пытаемся выбрать из текущего выделения в дереве
        if not self.controller.current_project:
            item = self.projects_tree.currentItem()
            if item:
                proj_id = item.data(0, Qt.UserRole) or (item.parent().data(0, Qt.UserRole) if item.parent() else None)
                if proj_id:
                    self.controller.project_controller.load_project(proj_id)
        if not self.controller.current_project:
            QMessageBox.warning(self, "Ошибка", "Сначала выберите или создайте проект")
            return
        
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Выберите файл формы",
            "",
            "Excel files (*.xlsx *.xls);;All files (*.*)"
        )
        
        if file_path:
            # Перед загрузкой файла спрашиваем тип формы, период и ревизию
            defaults = self.controller.get_pending_form_params() if hasattr(self.controller, "get_pending_form_params") else {}
            params_dialog = FormLoadDialog(self.controller.db_manager, self, defaults=defaults)
            if params_dialog.exec_() != QDialog.Accepted:
                return

            form_params = params_dialog.get_form_params()

            # Сохраняем выбранные пользователем параметры формы в контроллере
            if self.controller.current_project:
                form_code = form_params["form_code"]
                revision = form_params["revision"]
                period_code = form_params["period_code"]

                self.controller.set_current_form_params(
                    form_code=form_code,
                    revision=revision,
                    period_code=period_code,
                )

            self.progress_bar.setVisible(True)
            self.progress_bar.setRange(0, 0)
            self.status_bar.showMessage("Загрузка файла формы...")

            QTimer.singleShot(100, lambda: self._process_form_file(file_path))
    
    def _process_form_file(self, file_path):
        """Обработка файла формы"""
        try:
            success = self.controller.load_form_file(file_path)
            if success:
                # Перезагружаем данные проекта после загрузки формы
                if self.controller.current_project:
                    self.load_project_data_to_tree(self.controller.current_project)
                QMessageBox.information(self, "Успех", "Форма загружена и распарсена")
                self.status_bar.showMessage("Форма успешно загружена")
            else:
                QMessageBox.warning(self, "Ошибка", "Не удалось загрузить форму")
                self.status_bar.showMessage("Ошибка загрузки формы")
        except Exception as e:
            error_msg = f"Ошибка обработки файла формы: {e}"
            logger.error(error_msg, exc_info=True)
            QMessageBox.critical(self, "Ошибка", error_msg)
            self.status_bar.showMessage(error_msg)
        finally:
            self.progress_bar.setVisible(False)
    
    def calculate_sums(self):
        """Расчет агрегированных сумм"""
        if not self.controller.current_project:
            QMessageBox.warning(self, "Ошибка", "Сначала выберите проект")
            return
        
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, 0)
        
        QTimer.singleShot(100, self.controller.calculate_sums)
        QTimer.singleShot(1000, self._do_refresh_projects)
        
    
    def on_calculation_completed(self, results):
        """Обработка завершения расчета"""
        self.progress_bar.setVisible(False)
        QMessageBox.information(self, "Успех", "Расчет завершен")
        
        # Обновляем отображение данных
        if self.controller.current_project:
            self.load_project_data_to_tree(self.controller.current_project)
    
    def export_validation(self):
        """Экспорт формы с проверкой"""
        if not self.controller.current_project:
            QMessageBox.warning(self, "Ошибка", "Сначала выберите проект")
            return
        
        output_path, _ = QFileDialog.getSaveFileName(
            self,
            "Сохранить проверенную форму",
            f"{self.controller.current_project.name}_проверка.xlsx",
            "Excel files (*.xlsx)"
        )
        
        if output_path:
            self.progress_bar.setVisible(True)
            self.progress_bar.setRange(0, 0)
            
            QTimer.singleShot(100, lambda: self._process_export(output_path))
    
    def _process_export(self, output_path):
        """Обработка экспорта"""
        success = self.controller.export_validation(output_path)
        self.progress_bar.setVisible(False)
        
        if success:
            QMessageBox.information(self, "Успех", f"Форма экспортирована: {output_path}")
        else:
            QMessageBox.warning(self, "Ошибка", "Не удалось экспортировать форму")
    
    def on_export_completed(self, file_path):
        """Обработка завершения экспорта"""
        self.status_bar.showMessage(f"Форма экспортирована: {file_path}")
    
    def on_error_occurred(self, error_message):
        """Обработка ошибки"""
        self.progress_bar.setVisible(False)
        QMessageBox.critical(self, "Ошибка", error_message)
        self.status_bar.showMessage(f"Ошибка: {error_message}")
    
    def refresh_projects(self):
        """Обновление списка проектов"""
        # Показываем прогресс-бар во время обновления
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, 0)  # Неопределенный прогресс
        self.status_bar.showMessage("Обновление списка проектов...")
        
        # Обновляем данные с небольшой задержкой, чтобы UI успел обновиться
        # Это позволяет показать прогресс-бар до начала загрузки
        QTimer.singleShot(10, self._do_refresh_projects)
    
    def _do_refresh_projects(self):
        """Выполнение обновления списка проектов"""
        try:
            # Обновляем только список проектов, не перезагружая текущий проект
            # Это предотвращает зависание из-за пересчета уровней
            projects = self.controller.project_controller.load_projects()
            self.controller.projects_updated.emit(projects)
            
            # Обновляем справочники отдельно, чтобы не блокировать UI
            self.controller.refresh_references()
            
            self.status_bar.showMessage("Список проектов обновлен")
        except Exception as e:
            self.status_bar.showMessage(f"Ошибка обновления: {str(e)}")
            QMessageBox.critical(self, "Ошибка", f"Ошибка обновления списка проектов: {str(e)}")
        finally:
            self.progress_bar.setVisible(False)
    
    def edit_current_project(self):
        """Редактировать текущий проект"""
        if not self.controller.current_project or not self.controller.current_project.id:
            QMessageBox.warning(self, "Предупреждение", "Проект не выбран")
            return
        self.edit_project(self.controller.current_project.id)
    
    def delete_current_project(self):
        """Удалить текущий проект"""
        if not self.controller.current_project or not self.controller.current_project.id:
            QMessageBox.warning(self, "Предупреждение", "Проект не выбран")
            return
        
        reply = QMessageBox.question(
            self, 
            "Подтверждение удаления",
            f"Вы уверены, что хотите удалить проект '{self.controller.current_project.name}'?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            self.controller.delete_project(self.controller.current_project.id)
            QMessageBox.information(self, "Успех", "Проект удален")
    
    def toggle_fullscreen(self, checked: bool):
        """Переключить полноэкранный режим"""
        if checked:
            self.showFullScreen()
        else:
            self.showNormal()
    
    def show_about(self):
        """Показать информацию о программе"""
        QMessageBox.about(
            self,
            "О программе",
            "<h2>Система обработки бюджетных форм</h2>"
            "<p>Версия 1.0</p>"
            "<p>Приложение для обработки и анализа бюджетных форм, "
            "включая формы 0503317 и другие.</p>"
            "<p><b>Основные возможности:</b></p>"
            "<ul>"
            "<li>Загрузка и парсинг бюджетных форм</li>"
            "<li>Расчет агрегированных сумм</li>"
            "<li>Валидация данных</li>"
            "<li>Работа со справочниками</li>"
            "<li>Экспорт с проверкой</li>"
            "</ul>"
        )
    
    def show_shortcuts(self):
        """Показать список горячих клавиш"""
        shortcuts_text = """
        <h2>Горячие клавиши</h2>
        <table border="1" cellpadding="5">
        <tr><th>Действие</th><th>Клавиша</th></tr>
        <tr><td>Новый проект</td><td><b>Ctrl+N</b></td></tr>
        <tr><td>Загрузить форму</td><td><b>Ctrl+O</b></td></tr>
        <tr><td>Экспорт проверки</td><td><b>Ctrl+E</b></td></tr>
        <tr><td>Выход</td><td><b>Ctrl+Q</b></td></tr>
        <tr><td>Редактировать проект</td><td><b>Ctrl+P</b></td></tr>
        <tr><td>Удалить проект</td><td><b>Ctrl+Delete</b></td></tr>
        <tr><td>Обновить список</td><td><b>F5</b></td></tr>
        <tr><td>Пересчитать суммы</td><td><b>F9</b></td></tr>
        <tr><td>Скрыть нулевые столбцы</td><td><b>Ctrl+H</b></td></tr>
        <tr><td>Просмотр справочников</td><td><b>Ctrl+R</b></td></tr>
        <tr><td>Справочники конфигурации</td><td><b>Ctrl+D</b></td></tr>
        <tr><td>Панель проектов</td><td><b>Ctrl+1</b></td></tr>
        <tr><td>Полноэкранный режим</td><td><b>F11</b></td></tr>
        </table>
        """
        msg = QMessageBox(self)
        msg.setWindowTitle("Горячие клавиши")
        msg.setText(shortcuts_text)
        msg.setTextFormat(Qt.RichText)
        msg.exec_()
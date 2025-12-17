from PyQt5.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                             QPushButton, QLabel, QSplitter, 
                             QTabWidget, QTableWidget, QTableWidgetItem, 
                             QHeaderView, QMessageBox, QFileDialog, QProgressBar,
                             QToolBar, QStatusBar, QAction, QTextEdit,
                             QComboBox, QTreeWidget, QTreeWidgetItem, QMenu, 
                             QInputDialog, QDialog, QDialogButtonBox, QFormLayout,
                             QLineEdit, QCheckBox, QApplication, QStyle, QToolButton,
                             QSpinBox, QWidgetAction)
from PyQt5.QtCore import Qt, QTimer, QSize, QRect
from PyQt5.QtGui import (QFont, QColor, QBrush, QTextDocument, QTextOption, 
                        QTextCharFormat, QTextCursor, QPainter)
import os
import subprocess
import platform
import re
from pathlib import Path
import pandas as pd

from controllers.main_controller import MainController
from logger import logger
from models.constants.form_0503317_constants import Form0503317Constants
from views.project_dialog import ProjectDialog
from views.reference_dialog import ReferenceDialog
from views.excel_viewer import ExcelViewer
from views.reference_viewer import ReferenceViewer
from views.dictionaries_dialog import DictionariesDialog
from views.references_management_dialog import ReferencesManagementDialog
from views.form_load_dialog import FormLoadDialog
from views.document_dialog import DocumentDialog
from views.widgets import WrapHeaderView, WordWrapItemDelegate, DetachedTabWindow
from views.menu import MenuBar, ToolBar
from views.panels import ProjectsPanel, TabsPanel
from views.tree import TreeBuilder, TreeConfig, TreeHandlers
from views.errors import ErrorsManager
from views.metadata import MetadataPanel


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
        self.last_exported_file = None  # Путь к последнему экспортированного файла
        self.errors_tab_fullscreen = False  # Флаг полноэкранного режима вкладки ошибок
        # Окна для открепленных вкладок
        self.detached_windows = {}  # {tab_name: QMainWindow}
        self.tabs_panel = None  # Будет установлен в create_tabs_panel
        # Настройки шрифтов
        self.font_size = 10  # Размер шрифта для данных
        self.header_font_size = 10  # Размер шрифта для заголовков
        # Отслеживание выделения
        self.selection_start_column = None  # Столбец, с которого началось выделение
        
        # Инициализируем компоненты интерфейса
        self.projects_panel_obj = ProjectsPanel(self)
        self.tabs_panel_obj = TabsPanel(self)
        self.tree_builder = TreeBuilder(self)
        self.tree_config = TreeConfig(self)
        self.tree_handlers = TreeHandlers(self)
        self.errors_manager = ErrorsManager(self)
        self.metadata_panel = MetadataPanel(self)
        
        # Инициализируем менеджеры и контроллеры
        from views.managers.tab_manager import TabManager
        from views.controllers.documents_ui_controller import DocumentsUIController
        self.tab_manager = TabManager(self)
        self.documents_ui = DocumentsUIController(self)
        
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
        self.projects_panel = self.projects_panel_obj.create_projects_panel()
        splitter.addWidget(self.projects_panel)
        self.projects_panel_index = splitter.indexOf(self.projects_panel)
        
        # Центральная панель - вкладки с данными
        self.tabs_panel = self.tabs_panel_obj.create_tabs_panel()
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
        # self.create_dock_widgets()
    
    def create_menu_bar(self):
        """Создание меню-бара"""
        menu_bar = MenuBar(self)
        menu_bar.create_menu_bar()
    
    def create_toolbar(self):
        """Создание тулбара"""
        toolbar = ToolBar(self)
        toolbar.create_toolbar()
    
    def _create_menu_bar_old(self):
        """Старый метод создания меню-бара (для справки)"""
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
        
        # Открытие файлов
        open_file_action = QAction("&Открыть файл...", self)
        open_file_action.setIcon(self.style().standardIcon(QStyle.SP_DirOpenIcon))
        open_file_action.setShortcut("Ctrl+Shift+O")
        open_file_action.setStatusTip("Открыть файл (doc, docx, xls, xlsx)")
        open_file_action.triggered.connect(self.open_file_dialog)
        file_menu.addAction(open_file_action)
        
        # Открыть последний экспортированный файл
        self.open_last_file_action = QAction("Открыть последний экспортированный файл", self)
        self.open_last_file_action.setIcon(self.style().standardIcon(QStyle.SP_FileDialogStart))
        self.open_last_file_action.setStatusTip("Открыть последний экспортированный файл")
        self.open_last_file_action.setEnabled(False)
        self.open_last_file_action.triggered.connect(self.open_last_exported_file)
        file_menu.addAction(self.open_last_file_action)
        
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
        
        # Управление размером шрифта данных
        font_size_widget = QWidget()
        font_size_layout = QHBoxLayout(font_size_widget)
        font_size_layout.setContentsMargins(10, 5, 10, 5)
        font_size_label = QLabel("Размер шрифта данных:")
        font_size_layout.addWidget(font_size_label)
        self.font_size_spinbox = QSpinBox()
        self.font_size_spinbox.setMinimum(6)
        self.font_size_spinbox.setMaximum(20)
        self.font_size_spinbox.setValue(self.font_size)
        self.font_size_spinbox.setSuffix(" пт")
        self.font_size_spinbox.valueChanged.connect(self.on_font_size_changed)
        font_size_layout.addWidget(self.font_size_spinbox)
        font_size_action = QWidgetAction(self)
        font_size_action.setDefaultWidget(font_size_widget)
        view_menu.addAction(font_size_action)
        
        # Управление размером шрифта заголовков
        header_font_size_widget = QWidget()
        header_font_size_layout = QHBoxLayout(header_font_size_widget)
        header_font_size_layout.setContentsMargins(10, 5, 10, 5)
        header_font_size_label = QLabel("Размер шрифта заголовков:")
        header_font_size_layout.addWidget(header_font_size_label)
        self.header_font_size_spinbox = QSpinBox()
        self.header_font_size_spinbox.setMinimum(6)
        self.header_font_size_spinbox.setMaximum(20)
        self.header_font_size_spinbox.setValue(self.header_font_size)
        self.header_font_size_spinbox.setSuffix(" пт")
        self.header_font_size_spinbox.valueChanged.connect(self.on_header_font_size_changed)
        header_font_size_layout.addWidget(self.header_font_size_spinbox)
        header_font_size_action = QWidgetAction(self)
        header_font_size_action.setDefaultWidget(header_font_size_widget)
        view_menu.addAction(header_font_size_action)
        
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
        # Действие "Скрыть нулевые столбцы" перенесено в интерфейс формы (чекбокс)
        # Удалено из тулбара, т.к. теперь доступно в интерфейсе формы
        
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
    
    # Метод create_projects_panel перенесен в views.panels.projects_panel.ProjectsPanel
    
    # Метод create_tabs_panel перенесен в views.panels.tabs_panel.TabsPanel
    
    def connect_signals(self):
        """Подключение сигналов"""
        self.controller.projects_updated.connect(self.projects_panel_obj.update_projects_list)
        self.controller.project_loaded.connect(self.on_project_loaded)
        self.controller.calculation_completed.connect(self.on_calculation_completed)
        self.controller.export_completed.connect(self.on_export_completed)
        self.controller.error_occurred.connect(self.on_error_occurred)
    
    # Метод update_projects_list перенесен в views.panels.projects_panel.ProjectsPanel
    
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
                self.projects_panel_obj.update_projects_list(None)
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
                    self.projects_panel_obj.update_projects_list(None)
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

            # Получаем информацию о проекте/ревизии из контроллера
            project_info = self.controller.get_project_info(project)
            rev_id = getattr(self.controller, "current_revision_id", None)

            # Обновляем информацию о проекте
            info_text = (
                f"<b>Проект:</b> {project.name}<br>"
                f"<b>Форма:</b> {project_info['form_text']}<br>"
                f"<b>Ревизия:</b> {project_info['revision_text']}<br>"
                f"<b>МО:</b> {project_info['municipality_text']}<br>"
                f"<b>Период:</b> {project_info['period_text']}<br>"
                f"<b>Статус:</b> {project_info['status_text']}<br>"
                f"<b>Создан:</b> {project.created_at.strftime('%d.%m.%Y %H:%M')}"
            )
            self.project_info_label.setText(info_text)

            # Обновляем состояние кнопок ревизии
            self.update_revision_buttons_state(rev_id is not None)

            # Загружаем данные в древовидное представление
            self.tree_builder.load_project_data_to_tree(project)

            # Загружаем метаданные
            self.metadata_panel.load_metadata(project)
            
            # Обновляем вкладку ошибок
            self.errors_manager.load_errors_to_tab(project.data)

            # Загружаем файл в просмотрщик Excel:
            # Используем исходный файл ревизии (form_revisions.file_path), а не экспортированный
            # Экспортированный файл сохраняется отдельно и не должен заменять исходный
            excel_path = project_info.get('excel_path')
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
    
    # Методы _get_tree_widgets, _get_errors_widgets, _get_metadata_widgets перенесены в соответствующие модули
    # Метод load_project_data_to_tree перенесен в views.tree.tree_builder.TreeBuilder
    
    # Методы работы с ошибками перенесены в views.errors.errors_manager.ErrorsManager
    
    # Методы работы с ошибками (_check_budget_errors, _check_consolidated_errors, 
    # _update_errors_table, _export_errors, _format_error_value, _calculate_error_difference)
    # перенесены в views.errors.errors_manager.ErrorsManager
    
    def configure_tree_headers(self, section_name: str):
        """Конфигурация заголовков дерева под выбранный раздел (делегирует к tree_config)"""
        self.tree_config.configure_tree_headers(section_name)
    
    def _configure_tree_headers_for_widget(self, tree_widget, section_name, display_headers=None, mapping=None):
        """Настройка заголовков для конкретного виджета дерева (делегирует к tree_config)"""
        self.tree_config._configure_tree_headers_for_widget(tree_widget, section_name, display_headers, mapping)
    
    def _update_tree_header_height(self, tree_widget=None):
        """Обновляет высоту заголовка дерева (делегирует к tree_config)"""
        self.tree_config._update_tree_header_height(tree_widget)

    def hide_zero_columns_in_tree(self, section_key: str, data):
        """Скрытие столбцов дерева (делегирует к tree_config)"""
        self.tree_config.hide_zero_columns_in_tree(section_key, data)

    def apply_tree_data_type_visibility(self):
        """Скрывает столбцы дерева в зависимости от выбранного типа данных (делегирует к tree_config)"""
        self.tree_config.apply_tree_data_type_visibility()

    def format_budget_value(self, value):
        """Форматирование значения бюджета для отображения (делегирует к tree_builder)"""
        return self.tree_builder.format_budget_value(value)
    
    def build_tree_from_data(self, data, tree_widget=None):
        """Построение дерева из данных (делегирует к tree_builder)"""
        self.tree_builder.build_tree_from_data(data, tree_widget)
    
    def create_tree_item(self, item, level_colors, tree_widget=None):
        """Создание элемента дерева (делегирует к tree_builder)"""
        return self.tree_builder.create_tree_item(item, level_colors, tree_widget)
    
    def on_section_changed(self, section_name):
        """Обработка смены раздела"""
        self.current_section = section_name
        # Сбрасываем столбец выделения при смене раздела
        self.selection_start_column = None
        if self.controller.current_project:
            self.tree_builder.load_project_data_to_tree(self.controller.current_project)
            # Применяем скрытие нулевых столбцов, если чекбокс включен
            if hasattr(self, 'hide_zero_columns_checkbox') and self.hide_zero_columns_checkbox.isChecked():
                QTimer.singleShot(200, lambda: self.apply_hide_zero_columns())
    
    def on_data_type_changed(self, data_type):
        """Обработка смены типа данных"""
        self.current_data_type = data_type
        self.tree_config.apply_tree_data_type_visibility()
        # Применяем скрытие нулевых столбцов, если чекбокс включен
        if hasattr(self, 'hide_zero_columns_checkbox') and self.hide_zero_columns_checkbox.isChecked():
            self.apply_hide_zero_columns()
        if self.controller.current_project:
            self.tree_builder.load_project_data_to_tree(self.controller.current_project)
    
    def on_hide_zero_columns_changed(self, state):
        """Обработка изменения состояния чекбокса 'Скрыть нулевые столбцы'"""
        if state == Qt.Checked:
            self.apply_hide_zero_columns()
        else:
            self.show_all_columns()
    
    def apply_hide_zero_columns(self):
        """Применить скрытие нулевых столбцов"""
        if not (self.controller.current_project and self.controller.current_project.data):
            logger.debug("apply_hide_zero_columns: нет проекта или данных")
            return

        section_map = {
            "Доходы": "доходы_data",
            "Расходы": "расходы_data",
            "Источники финансирования": "источники_финансирования_data",
            "Консолидируемые расчеты": "консолидируемые_расчеты_data"
        }
        section_key = section_map.get(self.current_section)
        if not section_key or section_key not in self.controller.current_project.data:
            logger.debug(f"apply_hide_zero_columns: раздел {self.current_section} не найден")
            return

        data = self.controller.current_project.data[section_key]
        if not data:
            logger.debug(f"apply_hide_zero_columns: нет данных для раздела {section_key}")
            return

        logger.debug(f"apply_hide_zero_columns: применяю скрытие для раздела {section_key}, записей: {len(data)}")

        # Сначала показываем все столбцы
        self.tree_handlers.show_all_columns()
        
        # Применяем отображение колонок в зависимости от выбранного типа данных
        self.tree_config.apply_tree_data_type_visibility()

        # Затем применяем скрытие нулевых столбцов (после применения видимости по типу данных)
        self.tree_config.hide_zero_columns_in_tree(section_key, data)
    
    def expand_all_tree(self):
        """Развернуть все узлы дерева"""
        for tree_widget in self.tree_builder._get_tree_widgets():
            tree_widget.expandAll()
    
    def collapse_all_tree(self):
        """Свернуть все узлы дерева"""
        for tree_widget in self.tree_builder._get_tree_widgets():
            tree_widget.collapseAll()
    
    def on_tree_item_expanded(self, item):
        """Обработка разворачивания узла дерева (делегирует к tree_handlers)"""
        self.tree_handlers.on_tree_item_expanded(item)
    
    def on_tree_item_collapsed(self, item):
        """Обработка сворачивания узла дерева (делегирует к tree_handlers)"""
        self.tree_handlers.on_tree_item_collapsed(item)
    
    def show_tree_context_menu(self, position):
        """Контекстное меню для дерева (делегирует к tree_handlers)"""
        self.tree_handlers.show_tree_context_menu(position)

    def show_tree_header_context_menu(self, position):
        """Контекстное меню для заголовков дерева (делегирует к tree_handlers)"""
        self.tree_handlers.show_tree_header_context_menu(position)
    
    def show_all_columns(self):
        """Показать все столбцы в дереве и вернуть им нормальные ширины/заголовки"""
        self.tree_handlers.show_all_columns()

    def copy_tree_item_value(self, item):
        """Копировать значение из дерева (делегирует к tree_handlers)"""
        self.tree_handlers.copy_tree_item_value(item)
    
    
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
        from PyQt5.QtWidgets import QMainWindow, QToolBar

        if self.reference_window is None:
            self.reference_window = QMainWindow(self)
            self.reference_window.setWindowTitle("Справочники")
            self.reference_window.resize(900, 600)
            # Включаем стандартные кнопки окна (включая максимизацию)
            self.reference_window.setWindowFlags(self.reference_window.windowFlags() | Qt.WindowMaximizeButtonHint)
            self.reference_window.is_fullscreen = False

            self.reference_viewer = ReferenceViewer()
            self.reference_window.setCentralWidget(self.reference_viewer)
            
            # Обработка F11 для полноэкранного режима
            # Создаем обработчик событий для окна справочников
            def key_press_handler(event):
                if event.key() == Qt.Key_F11:
                    self._toggle_reference_fullscreen()
                else:
                    QMainWindow.keyPressEvent(self.reference_window, event)
            
            self.reference_window.keyPressEvent = key_press_handler

        # Устанавливаем callback для обновления данных из контроллера
        def refresh_callback():
            # Обновляем справочники в контроллере
            self.controller.refresh_references()
            # Обновляем данные в окне справочников
            self.reference_viewer.load_references(self.controller.references)
        
        self.reference_viewer.refresh_callback = refresh_callback
        
        # Загружаем актуальные справочники и показываем окно
        self.reference_viewer.load_references(self.controller.references)
        self.reference_window.show()
        self.reference_window.raise_()
        self.reference_window.activateWindow()
    
    def _toggle_reference_fullscreen(self):
        """Переключение полноэкранного режима для окна справочников"""
        if self.reference_window is None:
            return
        
        if self.reference_window.is_fullscreen:
            self.reference_window.showNormal()
            self.reference_window.is_fullscreen = False
        else:
            self.reference_window.showFullScreen()
            self.reference_window.is_fullscreen = True
    

    def show_config_dictionaries(self):
        """Показать диалог редактирования справочников конфигурации"""
        dlg = DictionariesDialog(self.controller.db_manager, self)
        dlg.exec_()
    
    def show_references_management(self):
        """Показать диалог управления справочниками (коды доходов, расходов и т.д.)"""
        dlg = ReferencesManagementDialog(self.controller.db_manager, self)
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
                    self.tree_builder.load_project_data_to_tree(self.controller.current_project)
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
    
    def update_revision_buttons_state(self, has_revision: bool):
        """Обновление состояния кнопок ревизии в зависимости от наличия выбранной ревизии"""
        if hasattr(self, 'recalculate_btn'):
            self.recalculate_btn.setEnabled(has_revision)
        if hasattr(self, 'export_calculated_btn'):
            self.export_calculated_btn.setEnabled(has_revision)
        if hasattr(self, 'show_errors_btn'):
            self.show_errors_btn.setEnabled(has_revision)
        if hasattr(self, 'documents_menu_btn'):
            self.documents_menu_btn.setEnabled(has_revision)
    
    def calculate_sums(self):
        """Расчет агрегированных сумм"""
        if not self.controller.current_project or not self.controller.current_revision_id:
            QMessageBox.warning(self, "Ошибка", "Сначала выберите проект и загрузите ревизию формы")
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
            self.tree_builder.load_project_data_to_tree(self.controller.current_project)
            # Обновляем вкладку ошибок
            self.errors_manager.load_errors_to_tab(self.controller.current_project.data)

    def export_validation(self):
        """Экспорт формы с проверкой (обертка для экспорта пересчитанной таблицы)"""
        self.export_calculated_table()
    
    def export_calculated_table(self):
        """Экспорт пересчитанной таблицы"""
        if not self.controller.current_project or not self.controller.current_revision_id:
            QMessageBox.warning(self, "Ошибка", "Сначала выберите проект и загрузите ревизию формы")
            return
        
        # Получаем информацию о ревизии для имени файла
        rev_id = self.controller.current_revision_id
        revision = self.controller.db_manager.get_form_revision_by_id(rev_id)
        revision_text = revision.revision if revision else "unknown"
        
        output_path, _ = QFileDialog.getSaveFileName(
            self,
            "Сохранить пересчитанную форму",
            f"{self.controller.current_project.name}_рев{revision_text}_пересчет.xlsx",
            "Excel files (*.xlsx)"
        )
        
        if output_path:
            self.progress_bar.setVisible(True)
            self.progress_bar.setRange(0, 0)
            
            QTimer.singleShot(100, lambda: self._process_export(output_path))
    
    def _process_export(self, output_path):
        """Обработка экспорта"""
        # Перед экспортом убеждаемся, что просмотрщик Excel не держит файл открытым
        if hasattr(self, "excel_viewer") and self.excel_viewer is not None:
            try:
                self.excel_viewer.close_workbook()
            except Exception:
                # Даже если что-то пошло не так, продолжаем — экспорт сам сообщит об ошибке
                pass

        success = self.controller.export_validation(output_path)
        self.progress_bar.setVisible(False)
        
        if success:
            # Сохраняем путь к последнему экспортированному файлу
            self.last_exported_file = output_path
            self.open_last_file_action.setEnabled(True)
            if hasattr(self, 'open_last_file_btn'):
                self.open_last_file_btn.setEnabled(True)
            
            # Предлагаем открыть файл
            reply = QMessageBox.question(
                self,
                "Успех",
                f"Форма экспортирована: {output_path}\n\nОткрыть файл?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.Yes
            )
            if reply == QMessageBox.Yes:
                self.open_file(output_path)
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
    
    def on_font_size_changed(self, size: int):
        """Обработка изменения размера шрифта данных"""
        self.font_size = size
        self.apply_font_sizes()
    
    def on_header_font_size_changed(self, size: int):
        """Обработка изменения размера шрифта заголовков"""
        self.header_font_size = size
        self.apply_font_sizes()
    
    def apply_font_sizes(self):
        """Применение размеров шрифтов ко всем деревьям"""
        # Получаем все виджеты дерева
        tree_widgets = self.tree_builder._get_tree_widgets()
        
        for tree_widget in tree_widgets:
            if tree_widget:
                # Применяем размер шрифта к дереву данных
                font = tree_widget.font()
                font.setPointSize(self.font_size)
                tree_widget.setFont(font)
                
                # Применяем размер шрифта к заголовкам
                header = tree_widget.header()
                if header:
                    header_font = header.font()
                    header_font.setPointSize(self.header_font_size)
                    header.setFont(header_font)
                    
                    # Обновляем высоту заголовка с учетом нового размера шрифта
                    self.tree_config._update_tree_header_height(tree_widget)
                
                # Обновляем делегат, если он использует шрифт
                delegate = tree_widget.itemDelegate()
                if delegate:
                    # Делегат будет использовать шрифт из option, который берется из виджета
                    tree_widget.update()
        
        # Обновляем отображение
        QApplication.processEvents()
    
    def toggle_fullscreen(self, checked: bool):
        """Переключить полноэкранный режим"""
        # Проверяем, активна ли вкладка ошибок
        if self.tabs_panel and self.tabs_panel.currentWidget() == self.errors_tab:
            self._toggle_errors_tab_fullscreen()
        else:
            if checked:
                self.showFullScreen()
            else:
                self.showNormal()
    
    def _toggle_errors_tab_fullscreen(self):
        """Переключение полноэкранного режима для вкладки ошибок"""
        if self.errors_tab_fullscreen:
            # Выходим из полноэкранного режима
            self.errors_tab_fullscreen = False
            self.showNormal()
        else:
            # Входим в полноэкранный режим
            self.errors_tab_fullscreen = True
            self.showFullScreen()
    
    def keyPressEvent(self, event):
        """Обработка нажатий клавиш"""
        if event.key() == Qt.Key_F11:
            # Если активна вкладка ошибок, переключаем её полноэкранный режим
            if self.tabs_panel and self.tabs_panel.currentWidget() == self.errors_tab:
                self._toggle_errors_tab_fullscreen()
            else:
                # Иначе переключаем полноэкранный режим главного окна
                self.toggle_fullscreen(not self.isFullScreen())
        else:
            super().keyPressEvent(event)
    
    def on_tree_item_clicked(self, item, column):
        """Обработчик клика по элементу дерева (делегирует к tree_handlers)"""
        self.tree_handlers.on_tree_item_clicked(item, column)
        # Также обновляем сумму сразу после клика
        QTimer.singleShot(10, self.on_tree_selection_changed)
    
    def on_tree_selection_changed(self):
        """Обработчик изменения выделения (делегирует к tree_handlers)"""
        self.tree_handlers.on_tree_selection_changed()
    
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
    
    def show_calculation_errors(self):
        """Показать вкладку с ошибками расчетов"""
        if not self.controller.current_project or not self.controller.current_revision_id:
            QMessageBox.warning(self, "Ошибка", "Сначала выберите проект и загрузите ревизию формы")
            return
        
        if not self.controller.current_project.data:
            QMessageBox.warning(self, "Предупреждение", "Нет данных для анализа ошибок")
            return
        
        # Загружаем ошибки из текущих данных проекта
        self.errors_manager.load_errors_to_tab(self.controller.current_project.data)
        
        # Переключаемся на вкладку ошибок
        tabs = self.tabs_panel
        if tabs:
            for i in range(tabs.count()):
                if tabs.tabText(i) == "Ошибки":
                    tabs.setCurrentIndex(i)
                    break
    
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
        <tr><td>Ошибки расчетов</td><td><b>Ctrl+Shift+E</b></td></tr>
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
    
    def show_document_dialog(self):
        """Показать диалог формирования документов (делегирует к documents_ui)"""
        self.documents_ui.show_document_dialog()
    
    def parse_solution_document(self):
        """Обработка решения о бюджете (делегирует к documents_ui)"""
        self.documents_ui.parse_solution_document()
    
    def open_file(self, file_path: str):
        """Открыть файл в системе"""
        if not file_path or not os.path.exists(file_path):
            QMessageBox.warning(self, "Ошибка", f"Файл не найден: {file_path}")
            return
        
        try:
            system = platform.system()
            if system == "Windows":
                os.startfile(file_path)
            elif system == "Darwin":  # macOS
                subprocess.run(["open", file_path])
            else:  # Linux
                subprocess.run(["xdg-open", file_path])
            self.status_bar.showMessage(f"Файл открыт: {file_path}")
        except Exception as e:
            logger.error(f"Ошибка открытия файла: {e}", exc_info=True)
            QMessageBox.critical(self, "Ошибка", f"Не удалось открыть файл:\n{str(e)}")
    
    def open_file_dialog(self):
        """Диалог выбора файла для открытия"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Выберите файл для открытия",
            "",
            "Все поддерживаемые файлы (*.doc *.docx *.xls *.xlsx);;"
            "Word Documents (*.doc *.docx);;"
            "Excel Files (*.xls *.xlsx);;"
            "All Files (*.*)"
        )
        
        if file_path:
            self.open_file(file_path)
    
    def open_last_exported_file(self):
        """
        Открыть последний экспортированный (пересчитанный) Excel,
        связанный с выбранной ревизией.
        """
        # 1) Пробуем открыть файл, привязанный к выбранной ревизии
        if self.controller.current_project and self.controller.current_revision_id:
            rev_id = self.controller.current_revision_id
            try:
                revision = self.controller.db_manager.get_form_revision_by_id(rev_id)
            except Exception as e:
                logger.error(
                    f"Ошибка получения ревизии {rev_id} при открытии последнего файла: {e}",
                    exc_info=True,
                )
                revision = None

            file_path = getattr(revision, "file_path", None) if revision else None
            if file_path and os.path.exists(file_path):
                self.open_file(file_path)
                return

        # 2) Fallback: используем last_exported_file (как раньше), если он есть
        if self.last_exported_file and os.path.exists(self.last_exported_file):
            self.open_file(self.last_exported_file)
            return

        # 3) Если ничего не нашли — показываем понятное сообщение
        QMessageBox.warning(
            self,
            "Ошибка",
            "Не удалось найти экспортированный файл.\n"
            "Убедитесь, что для выбранной ревизии выполнен экспорт с проверкой."
        )
    
    def show_tab_context_menu(self, position):
        """Контекстное меню для вкладок (делегирует к tab_manager)"""
        self.tab_manager.show_tab_context_menu(position)
    
    def detach_tab(self, tab_index, tab_name):
        """Открепление вкладки в отдельное окно (делегирует к tab_manager)"""
        self.tab_manager.detach_tab(tab_index, tab_name)
    
    def attach_tab(self, tab_name, tab_widget=None):
        """Возврат вкладки в главное окно (делегирует к tab_manager)"""
        self.tab_manager.attach_tab(tab_name, tab_widget)
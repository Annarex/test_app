"""Меню-бар приложения"""
from PyQt5.QtWidgets import (QAction, QWidget, QHBoxLayout, QLabel, 
                             QSpinBox, QWidgetAction)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont
from PyQt5.QtWidgets import QStyle


class MenuBar:
    """Класс для создания меню-бара"""
    
    def __init__(self, main_window):
        """
        Args:
            main_window: Ссылка на главное окно для доступа к обработчикам
        """
        self.main_window = main_window
    
    def create_menu_bar(self):
        """Создание меню-бара"""
        menubar = self.main_window.menuBar()
        
        # ========== Меню "Файл" ==========
        file_menu = menubar.addMenu("&Файл")
        
        new_project_action = QAction("&Новый проект...", self.main_window)
        new_project_action.setIcon(self.main_window.style().standardIcon(QStyle.SP_FileIcon))
        new_project_action.setShortcut("Ctrl+N")
        new_project_action.setStatusTip("Создать новый проект")
        new_project_action.triggered.connect(self.main_window.show_new_project_dialog)
        file_menu.addAction(new_project_action)
        
        load_form_action = QAction("&Загрузить форму...", self.main_window)
        load_form_action.setIcon(self.main_window.style().standardIcon(QStyle.SP_DirOpenIcon))
        load_form_action.setShortcut("Ctrl+O")
        load_form_action.setStatusTip("Загрузить файл формы")
        load_form_action.triggered.connect(self.main_window.load_form_file)
        file_menu.addAction(load_form_action)
        
        file_menu.addSeparator()
        
        export_action = QAction("&Экспорт проверки...", self.main_window)
        export_action.setIcon(self.main_window.style().standardIcon(QStyle.SP_DialogSaveButton))
        export_action.setShortcut("Ctrl+E")
        export_action.setStatusTip("Экспортировать форму с проверкой")
        export_action.triggered.connect(self.main_window.export_validation)
        file_menu.addAction(export_action)
        
        file_menu.addSeparator()
        
        # Открытие файлов
        open_file_action = QAction("&Открыть файл...", self.main_window)
        open_file_action.setIcon(self.main_window.style().standardIcon(QStyle.SP_DirOpenIcon))
        open_file_action.setShortcut("Ctrl+Shift+O")
        open_file_action.setStatusTip("Открыть файл (doc, docx, xls, xlsx)")
        open_file_action.triggered.connect(self.main_window.open_file_dialog)
        file_menu.addAction(open_file_action)
        
        # Открыть последний экспортированный файл
        self.main_window.open_last_file_action = QAction("Открыть последний экспортированный файл", self.main_window)
        self.main_window.open_last_file_action.setIcon(self.main_window.style().standardIcon(QStyle.SP_FileDialogStart))
        self.main_window.open_last_file_action.setStatusTip("Открыть последний экспортированный файл")
        self.main_window.open_last_file_action.setEnabled(False)
        self.main_window.open_last_file_action.triggered.connect(self.main_window.open_last_exported_file)
        file_menu.addAction(self.main_window.open_last_file_action)
        
        file_menu.addSeparator()
        
        exit_action = QAction("&Выход", self.main_window)
        exit_action.setIcon(self.main_window.style().standardIcon(QStyle.SP_DialogCloseButton))
        exit_action.setShortcut("Ctrl+Q")
        exit_action.setStatusTip("Выход из приложения")
        exit_action.triggered.connect(self.main_window.close)
        file_menu.addAction(exit_action)
        
        # ========== Меню "Проект" ==========
        project_menu = menubar.addMenu("&Проект")
        
        edit_project_action = QAction("&Редактировать проект...", self.main_window)
        edit_project_action.setIcon(self.main_window.style().standardIcon(QStyle.SP_FileDialogDetailedView))
        edit_project_action.setShortcut("Ctrl+P")
        edit_project_action.setStatusTip("Редактировать текущий проект")
        edit_project_action.triggered.connect(self.main_window.edit_current_project)
        project_menu.addAction(edit_project_action)
        
        delete_project_action = QAction("&Удалить проект", self.main_window)
        delete_project_action.setIcon(self.main_window.style().standardIcon(QStyle.SP_TrashIcon))
        delete_project_action.setShortcut("Ctrl+Delete")
        delete_project_action.setStatusTip("Удалить текущий проект")
        delete_project_action.triggered.connect(self.main_window.delete_current_project)
        project_menu.addAction(delete_project_action)
        
        project_menu.addSeparator()
        
        refresh_projects_action = QAction("&Обновить список", self.main_window)
        refresh_projects_action.setIcon(self.main_window.style().standardIcon(QStyle.SP_BrowserReload))
        refresh_projects_action.setShortcut("F5")
        refresh_projects_action.setStatusTip("Обновить список проектов")
        refresh_projects_action.triggered.connect(
            lambda: self.main_window.controller.projects_updated.emit(
                self.main_window.controller.project_controller.load_projects()
            )
        )
        project_menu.addAction(refresh_projects_action)
        
        # ========== Меню "Данные" ==========
        data_menu = menubar.addMenu("&Данные")
        # (действия для раздела данных сейчас управляются непосредственно формой)
        
        # ========== Меню "Справочники" ==========
        reference_menu = menubar.addMenu("&Справочники")
        
        load_income_ref_action = QAction("&Загрузить справочник доходов...", self.main_window)
        load_income_ref_action.setIcon(self.main_window.style().standardIcon(QStyle.SP_DialogOpenButton))
        load_income_ref_action.setStatusTip("Загрузить справочник доходов")
        load_income_ref_action.triggered.connect(lambda: self.main_window.show_reference_dialog("доходы"))
        reference_menu.addAction(load_income_ref_action)
        
        load_sources_ref_action = QAction("&Загрузить справочник источников...", self.main_window)
        load_sources_ref_action.setIcon(self.main_window.style().standardIcon(QStyle.SP_DialogOpenButton))
        load_sources_ref_action.setStatusTip("Загрузить справочник источников финансирования")
        load_sources_ref_action.triggered.connect(lambda: self.main_window.show_reference_dialog("источники"))
        reference_menu.addAction(load_sources_ref_action)
        
        reference_menu.addSeparator()
        
        show_references_action = QAction("&Просмотр справочников", self.main_window)
        show_references_action.setIcon(self.main_window.style().standardIcon(QStyle.SP_FileDialogInfoView))
        show_references_action.setShortcut("Ctrl+R")
        show_references_action.setStatusTip("Открыть окно просмотра справочников")
        show_references_action.triggered.connect(self.main_window.show_reference_viewer)
        reference_menu.addAction(show_references_action)
        
        reference_menu.addSeparator()
        
        config_dicts_action = QAction("&Справочники конфигурации...", self.main_window)
        config_dicts_action.setIcon(self.main_window.style().standardIcon(QStyle.SP_FileDialogListView))
        config_dicts_action.setShortcut("Ctrl+D")
        config_dicts_action.setStatusTip("Редактировать справочники конфигурации (годы, МО, типы форм, периоды)")
        config_dicts_action.triggered.connect(self.main_window.show_config_dictionaries)
        reference_menu.addAction(config_dicts_action)
        
        manage_refs_action = QAction("&Управление справочниками...", self.main_window)
        manage_refs_action.setIcon(self.main_window.style().standardIcon(QStyle.SP_FileDialogListView))
        manage_refs_action.setStatusTip("Управление справочниками (коды доходов, расходов, ГРБС и т.д.)")
        manage_refs_action.triggered.connect(self.main_window.show_references_management)
        reference_menu.addAction(manage_refs_action)
        
        # ========== Меню "Вид" ==========
        view_menu = menubar.addMenu("&Вид")
        
        toggle_projects_panel_action = QAction("&Панель проектов", self.main_window)
        toggle_projects_panel_action.setCheckable(True)
        toggle_projects_panel_action.setChecked(True)
        toggle_projects_panel_action.setShortcut("Ctrl+1")
        toggle_projects_panel_action.setStatusTip("Показать/скрыть панель проектов")
        toggle_projects_panel_action.triggered.connect(self.main_window.toggle_projects_panel)
        view_menu.addAction(toggle_projects_panel_action)
        
        view_menu.addSeparator()
        
        # Управление размером шрифта данных
        font_size_widget = QWidget()
        font_size_layout = QHBoxLayout(font_size_widget)
        font_size_layout.setContentsMargins(10, 5, 10, 5)
        font_size_label = QLabel("Размер шрифта данных:")
        font_size_layout.addWidget(font_size_label)
        self.main_window.font_size_spinbox = QSpinBox()
        self.main_window.font_size_spinbox.setMinimum(6)
        self.main_window.font_size_spinbox.setMaximum(20)
        self.main_window.font_size_spinbox.setValue(self.main_window.font_size)
        self.main_window.font_size_spinbox.setSuffix(" пт")
        self.main_window.font_size_spinbox.valueChanged.connect(self.main_window.on_font_size_changed)
        font_size_layout.addWidget(self.main_window.font_size_spinbox)
        font_size_action = QWidgetAction(self.main_window)
        font_size_action.setDefaultWidget(font_size_widget)
        view_menu.addAction(font_size_action)
        
        # Управление размером шрифта заголовков
        header_font_size_widget = QWidget()
        header_font_size_layout = QHBoxLayout(header_font_size_widget)
        header_font_size_layout.setContentsMargins(10, 5, 10, 5)
        header_font_size_label = QLabel("Размер шрифта заголовков:")
        header_font_size_layout.addWidget(header_font_size_label)
        self.main_window.header_font_size_spinbox = QSpinBox()
        self.main_window.header_font_size_spinbox.setMinimum(6)
        self.main_window.header_font_size_spinbox.setMaximum(20)
        self.main_window.header_font_size_spinbox.setValue(self.main_window.header_font_size)
        self.main_window.header_font_size_spinbox.setSuffix(" пт")
        self.main_window.header_font_size_spinbox.valueChanged.connect(self.main_window.on_header_font_size_changed)
        header_font_size_layout.addWidget(self.main_window.header_font_size_spinbox)
        header_font_size_action = QWidgetAction(self.main_window)
        header_font_size_action.setDefaultWidget(header_font_size_widget)
        view_menu.addAction(header_font_size_action)
        
        view_menu.addSeparator()
        
        fullscreen_action = QAction("&Полноэкранный режим", self.main_window)
        fullscreen_action.setShortcut("F11")
        fullscreen_action.setCheckable(True)
        fullscreen_action.setStatusTip("Переключить полноэкранный режим")
        fullscreen_action.triggered.connect(self.main_window.toggle_fullscreen)
        view_menu.addAction(fullscreen_action)
        
        # ========== Меню "Справка" ==========
        help_menu = menubar.addMenu("&Справка")
        
        about_action = QAction("&О программе", self.main_window)
        about_action.setIcon(self.main_window.style().standardIcon(QStyle.SP_MessageBoxInformation))
        about_action.setStatusTip("Информация о программе")
        about_action.triggered.connect(self.main_window.show_about)
        help_menu.addAction(about_action)
        
        help_menu.addSeparator()
        
        shortcuts_action = QAction("&Горячие клавиши", self.main_window)
        shortcuts_action.setIcon(self.main_window.style().standardIcon(QStyle.SP_FileDialogInfoView))
        shortcuts_action.setStatusTip("Список горячих клавиш")
        shortcuts_action.triggered.connect(self.main_window.show_shortcuts)
        help_menu.addAction(shortcuts_action)

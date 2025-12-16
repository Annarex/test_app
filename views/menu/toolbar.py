"""Тулбар приложения"""
from PyQt5.QtWidgets import QToolBar, QAction
from PyQt5.QtWidgets import QStyle


class ToolBar:
    """Класс для создания тулбара"""
    
    def __init__(self, main_window):
        """
        Args:
            main_window: Ссылка на главное окно для доступа к обработчикам
        """
        self.main_window = main_window
    
    def create_toolbar(self):
        """Создание панели инструментов"""
        toolbar = QToolBar("Основные инструменты")
        self.main_window.addToolBar(toolbar)
        
        # Действия
        new_project_action = QAction("Новый проект", self.main_window)
        new_project_action.setIcon(self.main_window.style().standardIcon(QStyle.SP_FileIcon))
        new_project_action.triggered.connect(self.main_window.show_new_project_dialog)
        toolbar.addAction(new_project_action)
        
        load_form_action = QAction("Загрузить форму", self.main_window)
        load_form_action.setIcon(self.main_window.style().standardIcon(QStyle.SP_DirOpenIcon))
        load_form_action.triggered.connect(self.main_window.load_form_file)
        toolbar.addAction(load_form_action)
        
        toolbar.addSeparator()
        
        # Отдельные действия для справочников доходов и источников
        load_income_ref_action = QAction("Справочник доходов", self.main_window)
        load_income_ref_action.setIcon(self.main_window.style().standardIcon(QStyle.SP_DialogOpenButton))
        load_income_ref_action.triggered.connect(lambda: self.main_window.show_reference_dialog("доходы"))
        toolbar.addAction(load_income_ref_action)
        
        load_sources_ref_action = QAction("Справочник источников", self.main_window)
        load_sources_ref_action.setIcon(self.main_window.style().standardIcon(QStyle.SP_DialogOpenButton))
        load_sources_ref_action.triggered.connect(lambda: self.main_window.show_reference_dialog("источники"))
        toolbar.addAction(load_sources_ref_action)
        
        show_references_action = QAction("Просмотр справочников", self.main_window)
        show_references_action.setIcon(self.main_window.style().standardIcon(QStyle.SP_FileDialogInfoView))
        show_references_action.triggered.connect(self.main_window.show_reference_viewer)
        toolbar.addAction(show_references_action)
        
        # Редактор конфигурационных справочников (годы, МО, типы форм, периоды)
        config_dicts_action = QAction("Справочники конфигурации", self.main_window)
        config_dicts_action.setIcon(self.main_window.style().standardIcon(QStyle.SP_FileDialogListView))
        config_dicts_action.triggered.connect(self.main_window.show_config_dictionaries)
        toolbar.addAction(config_dicts_action)
        
        # Кнопки управления панелью проектов размещены непосредственно на самой панели

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
        
        manage_refs_action = QAction("Управление справочниками", self.main_window)
        manage_refs_action.setIcon(self.main_window.style().standardIcon(QStyle.SP_FileDialogListView))
        manage_refs_action.triggered.connect(self.main_window.show_reference_viewer)
        toolbar.addAction(manage_refs_action)
        
        # Кнопки управления панелью проектов размещены непосредственно на самой панели

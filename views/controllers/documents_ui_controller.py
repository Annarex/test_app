"""UI-контроллер для работы с документами"""
from PyQt5.QtWidgets import QMessageBox, QFileDialog
from logger import logger
from views.document_dialog import DocumentDialog
from views.solution_load_dialog import SolutionLoadDialog


class DocumentsUIController:
    """UI-контроллер для работы с документами (заключения, письма, решения)"""
    
    def __init__(self, main_window):
        """
        Args:
            main_window: Ссылка на главное окно
        """
        self.main_window = main_window
    
    def show_document_dialog(self):
        """Показать диалог формирования документов"""
        if not self.main_window.controller.current_project or not self.main_window.controller.current_revision_id:
            QMessageBox.warning(
                self.main_window, 
                "Ошибка", 
                "Сначала выберите проект и загрузите ревизию формы"
            )
            return
        
        dialog = DocumentDialog(self.main_window)
        dialog.exec_()
    
    def parse_solution_document(self):
        """Обработка решения о бюджете"""
        if not self.main_window.controller.current_project:
            QMessageBox.warning(self.main_window, "Ошибка", "Сначала выберите проект")
            return
        
        # Используем новый диалог для загрузки решений
        dialog = SolutionLoadDialog(self.main_window)
        dialog.exec_()

"""UI-контроллер для работы с документами"""
from PyQt5.QtWidgets import QMessageBox, QFileDialog
from logger import logger
from views.document_dialog import DocumentDialog


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
        
        file_path, _ = QFileDialog.getOpenFileName(
            self.main_window,
            "Выберите файл решения о бюджете",
            "",
            "Word Documents (*.docx *.doc);;All Files (*)"
        )
        
        if file_path:
            try:
                result = self.main_window.controller.parse_solution_document(file_path)
                if result:
                    QMessageBox.information(
                        self.main_window,
                        "Успех",
                        f"Решение обработано:\n"
                        f"Доходов: {len(result.get('приложение1', []))}\n"
                        f"Расходов (общие): {len(result.get('приложение2', []))}\n"
                        f"Расходов (по ГРБС): {len(result.get('приложение3', []))}"
                    )
                else:
                    QMessageBox.warning(self.main_window, "Ошибка", "Не удалось обработать решение")
            except Exception as e:
                logger.error(f"Ошибка обработки решения: {e}", exc_info=True)
                QMessageBox.critical(
                    self.main_window, 
                    "Ошибка", 
                    f"Ошибка обработки решения:\n{str(e)}"
                )

"""Панель метаданных"""
from PyQt5.QtWidgets import QTextEdit


class MetadataPanel:
    """Класс для управления панелью метаданных"""
    
    def __init__(self, main_window):
        """
        Args:
            main_window: Ссылка на главное окно для доступа к контроллеру
        """
        self.main_window = main_window
        self.controller = main_window.controller
    
    def load_metadata(self, project):
        """Загрузка метаданных для выбранной ревизии"""
        # Метаданные должны быть только у ревизии, а не у проекта
        # Проверяем, что загружена ревизия (current_revision_id установлен)
        rev_id = getattr(self.controller, "current_revision_id", None)
        
        # Получаем все виджеты метаданных
        metadata_widgets = self._get_metadata_widgets()
        
        if not rev_id:
            # Если ревизия не загружена, метаданные не отображаем
            for metadata_widget in metadata_widgets:
                metadata_widget.setHtml("")
            return
        
        # Метаданные берём из данных проекта (которые загружаются из ревизии)
        if not project or not project.data:
            for metadata_widget in metadata_widgets:
                metadata_widget.setHtml("")
            return
        
        meta_info = project.data.get('meta_info', {})
        if not meta_info:
            for metadata_widget in metadata_widgets:
                metadata_widget.setHtml("")
            return
        
        metadata_text = ""
        for key, value in meta_info.items():
            metadata_text += f"<b>{key}:</b> {value}<br>"
        
        # Обновляем все виджеты метаданных
        for metadata_widget in metadata_widgets:
            metadata_widget.setHtml(metadata_text)
    
    def _get_metadata_widgets(self):
        """Получить все виджеты метаданных (в главном окне и открепленных)"""
        widgets = []
        # Виджет в главном окне
        if hasattr(self.main_window, 'metadata_text') and self.main_window.metadata_text:
            widgets.append(self.main_window.metadata_text)
        
        # Виджеты в открепленных окнах
        if hasattr(self.main_window, 'detached_windows') and "Метаданные" in self.main_window.detached_windows:
            detached_window = self.main_window.detached_windows["Метаданные"]
            tab_widget = detached_window.get_tab_widget()
            if tab_widget:
                from PyQt5.QtWidgets import QTextEdit
                for child in tab_widget.findChildren(QTextEdit):
                    if child not in widgets:
                        widgets.append(child)
        
        return widgets

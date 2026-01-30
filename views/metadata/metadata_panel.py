"""Панель метаданных"""
from datetime import datetime
from PyQt5.QtWidgets import QTextEdit


def get_reference_date_from_meta_info(meta_info: dict) -> str:
    """Извлекает дату актуальности из метаданных ревизии. Возвращает YYYY-MM-DD или текущую дату."""
    if not meta_info:
        return datetime.now().strftime("%Y-%m-%d")
    for key in ("дата", "date", "report_date", "Дата формирования", "Дата отчетности", "Дата", "period"):
        val = meta_info.get(key)
        if val is None:
            continue
        s = str(val).strip()
        if not s:
            continue
        for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%d/%m/%Y"):
            try:
                dt = datetime.strptime(s[:10], fmt)
                return dt.strftime("%Y-%m-%d")
            except (ValueError, IndexError):
                continue
        if len(s) >= 4 and s[:4].isdigit():
            try:
                y = int(s[:4])
                if 1990 <= y <= 2100:
                    return f"{y}-01-01"
            except ValueError:
                pass
    return datetime.now().strftime("%Y-%m-%d")


def get_reference_ppocode_from_meta_info(meta_info: dict):
    """Извлекает код ОКТМО (ppocode) из метаданных ревизии для фильтрации справочников. Возвращает строку или None."""
    if not meta_info:
        return None
    for key in ("код ОКТМО", "ОКТМО", "oktmo", "ppocode", "код_октмо"):
        val = meta_info.get(key)
        if val is None:
            continue
        s = str(val).strip()
        if s:
            return s
    return None


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
            # Если ревизия не загружена, метаданные не отображаем, дата актуальности — текущая
            for metadata_widget in metadata_widgets:
                metadata_widget.setHtml("")
            ref_date = datetime.now().strftime("%Y-%m-%d")
            self._update_reference_date_widget(ref_date)
            self._apply_reference_filters(ref_date, None)
            return
        
        # Метаданные берём из данных проекта (которые загружаются из ревизии)
        if not project or not project.data:
            for metadata_widget in metadata_widgets:
                metadata_widget.setHtml("")
            ref_date = datetime.now().strftime("%Y-%m-%d")
            self._update_reference_date_widget(ref_date)
            self._apply_reference_filters(ref_date, None)
            return
        
        meta_info = project.data.get('meta_info', {})
        if not meta_info:
            for metadata_widget in metadata_widgets:
                metadata_widget.setHtml("")
            ref_date = datetime.now().strftime("%Y-%m-%d")
            self._update_reference_date_widget(ref_date)
            self._apply_reference_filters(ref_date, None)
            return
        
        metadata_text = ""
        for key, value in meta_info.items():
            metadata_text += f"<b>{key}:</b> {value}<br>"
        
        # Обновляем все виджеты метаданных
        for metadata_widget in metadata_widgets:
            metadata_widget.setHtml(metadata_text)
        
        # Обновляем дату и ОКТМО (ppocode) для фильтрации справочников (панель дерева и конфиг)
        ref_date = get_reference_date_from_meta_info(meta_info)
        ref_ppocode = get_reference_ppocode_from_meta_info(meta_info)
        self._update_reference_date_widget(ref_date)
        self._apply_reference_filters(ref_date, ref_ppocode)
    
    def _apply_reference_filters(self, ref_date: str, ref_ppocode=None):
        """Сохраняет дату и ОКТМО (ppocode) в конфиг и перезагружает справочник доходов (v_budgetclastypeinc_merged)."""
        if self.controller.db_manager:
            self.controller.db_manager.save_config("reference_filter_date", ref_date)
            self.controller.db_manager.save_config("reference_filter_ppocode", ref_ppocode or "")
        if hasattr(self.controller, "refresh_references"):
            self.controller.refresh_references()

    def _update_reference_date_widget(self, date_str: str):
        """Обновляет виджет даты актуальности на панели дерева."""
        if hasattr(self.main_window, "reference_date_edit") and self.main_window.reference_date_edit:
            self.main_window.reference_date_edit.setText(date_str)
    
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

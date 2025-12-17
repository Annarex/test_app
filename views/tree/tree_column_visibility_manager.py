"""Менеджер видимости колонок дерева"""
from logger import logger
from models.constants.form_0503317_constants import Form0503317Constants


class TreeColumnVisibilityManager:
    """Класс для управления видимостью колонок дерева"""
    
    def __init__(self, main_window):
        """
        Args:
            main_window: Ссылка на главное окно для доступа к свойствам
        """
        self.main_window = main_window
    
    def hide_zero_columns(self, section_key: str, data, tree_widget):
        """
        Скрытие столбцов дерева, в которых итоговое значение равно 0.
        Логика аналогична табличному представлению.
        
        Args:
            section_key: Ключ раздела данных
            data: Данные раздела
            tree_widget: Виджет дерева
        """
        if not data:
            return

        if section_key == "консолидируемые_расчеты_data":
            self._hide_zero_columns_consolidated(data, tree_widget)
        else:
            self._hide_zero_columns_budget(section_key, data, tree_widget)
    
    def _hide_zero_columns_consolidated(self, data, tree_widget):
        """Скрытие нулевых колонок для консолидированных расчетов"""
        cons_cols = Form0503317Constants.CONSOLIDATED_COLUMNS
        mapping = getattr(self.main_window, 'tree_column_mapping', {})
        if mapping.get("type") != "consolidated":
            return

        # Ищем итоговую строку
        total_item = None
        for item in data:
            name = str(item.get("наименование_показателя", "")).strip().lower()
            code = str(item.get("код_строки", "")).strip().lower()
            # Для консолидированных: строка начинается с "всего" ИЛИ код 899
            if name.startswith("всего") or code == "899":
                total_item = item
                break
        if not total_item:
            return

        value_start = mapping.get("value_start", 4)
        totals = total_item.get("поступления", {}) or {}

        header = tree_widget.header()
        zero_cols = []
        for i, col_name in enumerate(cons_cols):
            val = totals.get(col_name, 0)
            if isinstance(val, (int, float)) and abs(val) < 1e-9:
                col_index = value_start + i
                if 0 <= col_index < tree_widget.columnCount():
                    zero_cols.append(col_index)

        # Сужаем «нулевые» колонки до минимальной ширины и очищаем заголовки
        header_item = tree_widget.headerItem()
        for col_index in zero_cols:
            header.resizeSection(col_index, 2)  # минимальная ширина
            if header_item:
                header_item.setText(col_index, "")
                header_item.setToolTip(col_index, "")
    
    def _hide_zero_columns_budget(self, section_key: str, data, tree_widget):
        """Скрытие нулевых колонок для бюджетных разделов"""
        budget_cols = Form0503317Constants.BUDGET_COLUMNS
        mapping = getattr(self.main_window, 'tree_column_mapping', {})
        if mapping.get("type") != "budget":
            return

        total_item = None
        for item in data:
            name = str(item.get("наименование_показателя", "")).strip().lower()
            # Ищем первую строку, где встречается слово "всего"
            if "всего" in name:
                total_item = item
                break
        
        if not total_item:
            logger.debug(f"Итоговая строка не найдена для раздела {section_key}")
            return

        approved = total_item.get("утвержденный", {}) or {}
        executed = total_item.get("исполненный", {}) or {}

        approved_start = mapping.get("approved_start", 4)
        executed_start = mapping.get("executed_start", approved_start + len(budget_cols))

        # Учитываем видимость столбцов по типу данных
        current_data_type = getattr(self.main_window, 'current_data_type', 'Оба')
        show_approved = current_data_type in ("Утвержденный", "Оба")
        show_executed = current_data_type in ("Исполненный", "Оба")

        header = tree_widget.header()
        zero_cols = set()
        for i, col_name in enumerate(budget_cols):
            a_val = approved.get(col_name, 0) or 0
            e_val = executed.get(col_name, 0) or 0
            if isinstance(a_val, (int, float)) and isinstance(e_val, (int, float)):
                if abs(a_val) < 1e-9 and abs(e_val) < 1e-9:
                    appr_idx = approved_start + i
                    exec_idx = executed_start + i
                    if show_approved and 0 <= appr_idx < tree_widget.columnCount():
                        zero_cols.add(appr_idx)
                    if show_executed and 0 <= exec_idx < tree_widget.columnCount():
                        zero_cols.add(exec_idx)

        # Сужаем «нулевые» колонки до минимальной ширины и очищаем заголовки
        header_item = tree_widget.headerItem()
        for col_index in zero_cols:
            header.resizeSection(col_index, 2)  # минимальная ширина
            if header_item:
                header_item.setText(col_index, "")
                header_item.setToolTip(col_index, "")
    
    def apply_data_type_visibility(self, data_type: str, tree_widget):
        """Скрывает столбцы дерева в зависимости от выбранного типа данных
        
        Args:
            data_type: Тип данных ("Утвержденный", "Исполненный", "Оба")
            tree_widget: Виджет дерева
        """
        mapping = getattr(self.main_window, 'tree_column_mapping', {})
        if not mapping:
            return

        if mapping.get("type") == "budget":
            budget_cols = mapping.get("budget_columns", [])
            approved_start = mapping.get("approved_start", 4)
            executed_start = mapping.get("executed_start", approved_start + len(budget_cols))
            
            show_approved = data_type in ("Утвержденный", "Оба")
            show_executed = data_type in ("Исполненный", "Оба")
            
            approved_range = range(approved_start, executed_start)
            executed_range = range(executed_start, executed_start + len(budget_cols))
            
            for idx in approved_range:
                tree_widget.setColumnHidden(idx, not show_approved)
            for idx in executed_range:
                tree_widget.setColumnHidden(idx, not show_executed)
    
    def show_all_columns(self, tree_widget):
        """Показать все столбцы в дереве и вернуть им нормальные ширины/заголовки
        
        Args:
            tree_widget: Виджет дерева
        """
        mapping = getattr(self.main_window, 'tree_column_mapping', {})
        if not mapping:
            return
        
        tree_headers = getattr(self.main_window, 'tree_headers', [])
        header = tree_widget.header()
        header_item = tree_widget.headerItem()
        
        # Показываем все скрытые колонки
        for idx in range(tree_widget.columnCount()):
            tree_widget.setColumnHidden(idx, False)
            
            # Восстанавливаем ширину колонок
            if mapping.get("type") == "budget":
                if idx == 0:
                    # Столбец "Наименование"
                    indentation = tree_widget.indentation()
                    indent_reserve = indentation * 6 + 50
                    header.resizeSection(idx, 400 + indent_reserve)
                elif idx == 1:
                    header.resizeSection(idx, 80)
                elif idx == 2:
                    header.resizeSection(idx, 200)
                elif idx == 3:
                    header.resizeSection(idx, 50)
                elif idx >= 4:
                    header.resizeSection(idx, 150)
            
            # Восстанавливаем заголовки
            if header_item and idx < len(tree_headers):
                header_item.setText(idx, tree_headers[idx])
                tooltips = getattr(self.main_window, 'tree_header_tooltips', [])
                if idx < len(tooltips):
                    header_item.setToolTip(idx, tooltips[idx])

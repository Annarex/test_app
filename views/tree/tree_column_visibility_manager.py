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
    
    def get_visible_indices_for_data_type(self, column_count: int, mapping: dict, data_type: str) -> set:
        """Возвращает множество индексов столбцов, видимых при выбранном типе данных.
        Используется при построении tree_visible_columns (дерево без скрытых столбцов)."""
        visible = set(range(column_count))
        if mapping.get("type") != "budget":
            return visible
        budget_cols = mapping.get("budget_columns", [])
        approved_start = mapping.get("approved_start", 4)
        executed_start = mapping.get("executed_start", approved_start + len(budget_cols))
        show_approved = data_type in ("Утвержденный", "Оба")
        show_executed = data_type in ("Исполненный", "Оба")
        if not show_approved:
            visible -= set(range(approved_start, executed_start))
        if not show_executed:
            visible -= set(range(executed_start, executed_start + len(budget_cols)))
        return visible
    
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
            
            from PyQt5.QtWidgets import QHeaderView
            header = tree_widget.header()
            for idx in approved_range:
                is_visible = show_approved
                # Для скрытых столбцов сначала устанавливаем режим Fixed и ширину в 0, затем скрываем
                if not is_visible:
                    header.setSectionResizeMode(idx, QHeaderView.Fixed)
                    header.resizeSection(idx, 0)
                    tree_widget.setColumnWidth(idx, 0)  # Также устанавливаем ширину на самом виджете
                    tree_widget.setColumnHidden(idx, True)
                else:
                    tree_widget.setColumnHidden(idx, False)
            for idx in executed_range:
                is_visible = show_executed
                # Для скрытых столбцов сначала устанавливаем режим Fixed и ширину в 0, затем скрываем
                if not is_visible:
                    header.setSectionResizeMode(idx, QHeaderView.Fixed)
                    header.resizeSection(idx, 0)
                    tree_widget.setColumnWidth(idx, 0)  # Также устанавливаем ширину на самом виджете
                    tree_widget.setColumnHidden(idx, True)
                else:
                    tree_widget.setColumnHidden(idx, False)
            
            # Обновляем геометрию заголовка и виджета после всех изменений
            header.updateGeometries()
            tree_widget.updateGeometry()
            tree_widget.viewport().update()  # Принудительно обновляем viewport
    
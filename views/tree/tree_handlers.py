"""Обработчики событий дерева"""
from PyQt5.QtWidgets import QMenu, QTreeWidgetItem, QApplication
from PyQt5.QtCore import Qt
from views.reference_detail_dialog import ReferenceDetailDialog
from views.column_visibility_dialog import ColumnVisibilityDialog
from logger import logger


class TreeHandlers:
    """Класс для обработчиков событий дерева"""
    
    def __init__(self, main_window):
        """
        Args:
            main_window: Ссылка на главное окно для доступа к методам и свойствам
        """
        self.main_window = main_window
    
    def on_tree_item_clicked(self, item, column):
        """Обработка клика по элементу дерева"""
        # Сохраняем столбец начала выделения для расширенного выделения
        self.main_window.selection_start_column = column
    
    def on_tree_item_double_clicked(self, item, column):
        """Обработка двойного клика по элементу дерева - открывает диалог с детальной информацией"""
        try:
            # Получаем исходные данные из UserRole
            item_data = item.data(0, Qt.UserRole)
            if not item_data or not isinstance(item_data, dict):
                # Если данных нет, формируем из текста столбцов дерева
                item_data = {}
                tree_headers = getattr(self.main_window, 'tree_headers', [])
                for col in range(self.main_window.data_tree.columnCount()):
                    header = tree_headers[col] if col < len(tree_headers) else f"Колонка {col}"
                    item_data[header] = item.text(col) if item else ""
            
            # Получаем заголовки столбцов дерева
            tree_headers = getattr(self.main_window, 'tree_headers', [])
            if not tree_headers:
                # Если заголовков нет, формируем из данных элемента
                column_headers = list(item_data.keys())
            else:
                column_headers = tree_headers.copy()
            
            # Формируем row_data из item_data, дополняя значениями из столбцов дерева
            row_data = {}
            
            # Сначала добавляем все данные из item_data
            for key, value in item_data.items():
                if value is not None:
                    # Форматируем значения
                    if isinstance(value, (int, float)):
                        row_data[key] = f"{value:,.2f}".replace(",", " ")
                    elif isinstance(value, dict):
                        # Для вложенных словарей (утвержденный, исполненный, поступления)
                        row_data[key] = self._format_dict_value(value)
                    elif isinstance(value, list):
                        row_data[key] = ", ".join(str(v) for v in value)
                    else:
                        row_data[key] = str(value)
            
            # Дополняем данными из столбцов дерева, если их нет в item_data
            for col in range(self.main_window.data_tree.columnCount()):
                header = column_headers[col] if col < len(column_headers) else f"Колонка {col}"
                if header not in row_data:
                    text = item.text(col) if item else ""
                    if text:
                        row_data[header] = text
            
            # Если row_data пуст, используем только данные из столбцов дерева
            if not row_data:
                for col in range(self.main_window.data_tree.columnCount()):
                    header = column_headers[col] if col < len(column_headers) else f"Колонка {col}"
                    row_data[header] = item.text(col) if item else ""
            
            # Открываем диалог с детальной информацией
            dialog = ReferenceDetailDialog(row_data, column_headers, self.main_window)
            dialog.exec_()
        except Exception as e:
            logger.error(f"Ошибка при открытии диалога детальной информации: {e}", exc_info=True)
    
    def _format_dict_value(self, value_dict):
        """Форматирует словарь для отображения в диалоге"""
        if not isinstance(value_dict, dict):
            return str(value_dict)
        
        formatted_items = []
        for key, val in value_dict.items():
            if val is not None and val != 0 and val != '':
                if isinstance(val, (int, float)):
                    formatted_items.append(f"{key}: {val:,.2f}".replace(",", " "))
                else:
                    formatted_items.append(f"{key}: {val}")
        
        return "; ".join(formatted_items) if formatted_items else "(пусто)"
    
    def on_tree_selection_changed(self):
        """Обработка изменения выделения в дереве - подсчитывает сумму по выбранному столбцу"""
        # Получаем выделенные элементы
        selected_items = self.main_window.data_tree.selectedItems()
        
        if not selected_items:
            # Если ничего не выбрано, очищаем статус
            self.main_window.status_bar.showMessage("Готов к работе")
            return
        
        # Определяем столбец: сначала используем сохраненный, если нет - текущий столбец
        column_index = getattr(self.main_window, 'selection_start_column', None)
        if column_index is None:
            # Пытаемся определить столбец из текущего элемента
            current_item = self.main_window.data_tree.currentItem()
            if current_item:
                # Используем столбец текущего элемента
                column_index = self.main_window.data_tree.currentColumn()
                if column_index < 0:
                    column_index = 0
            else:
                # Если не можем определить, используем первый столбец данных (после базовых)
                mapping = getattr(self.main_window, 'tree_column_mapping', {})
                column_type = mapping.get("type", "base")
                if column_type == "budget":
                    column_index = mapping.get("approved_start", 4)
                elif column_type == "consolidated":
                    column_index = mapping.get("value_start", 4)
                else:
                    column_index = 4  # По умолчанию
        
        # Определяем тип столбца и получаем данные
        mapping = getattr(self.main_window, 'tree_column_mapping', {})
        column_type = mapping.get("type", "base")
        
        total = 0.0
        count = 0
        column_name = ""
        
        # Определяем название столбца для отображения
        tree_headers = getattr(self.main_window, 'tree_headers', [])
        if column_index < len(tree_headers):
            column_name = tree_headers[column_index]
        
        # Обрабатываем выбранные элементы
        for tree_item in selected_items:
            # Получаем исходные данные из UserRole
            item_data = tree_item.data(0, Qt.UserRole)
            if not item_data or not isinstance(item_data, dict):
                continue
            
            value = None
            
            if column_type == "budget":
                # Бюджетные столбцы (утвержденный/исполненный)
                budget_cols = mapping.get("budget_columns", [])
                approved_start = mapping.get("approved_start", 4)
                executed_start = mapping.get("executed_start", approved_start + len(budget_cols))
                
                if approved_start <= column_index < executed_start:
                    # Столбец утвержденного
                    col_idx = column_index - approved_start
                    if col_idx < len(budget_cols):
                        col_name = budget_cols[col_idx]
                        approved_data = item_data.get('утвержденный', {}) or {}
                        value = approved_data.get(col_name, 0) or 0
                elif executed_start <= column_index < executed_start + len(budget_cols):
                    # Столбец исполненного
                    col_idx = column_index - executed_start
                    if col_idx < len(budget_cols):
                        col_name = budget_cols[col_idx]
                        executed_data = item_data.get('исполненный', {}) or {}
                        value = executed_data.get(col_name, 0) or 0
            
            elif column_type == "consolidated":
                # Консолидируемые расчеты
                value_start = mapping.get("value_start", 4)
                cons_cols = mapping.get("columns", [])
                
                if value_start <= column_index < value_start + len(cons_cols):
                    col_idx = column_index - value_start
                    if col_idx < len(cons_cols):
                        col_name = cons_cols[col_idx]
                        cons_data = item_data.get('поступления', {}) or {}
                        if isinstance(cons_data, dict) and col_name in cons_data:
                            value = cons_data.get(col_name, 0) or 0
                        else:
                            # Проверяем плоские поля
                            value = item_data.get(f'поступления_{col_name}', 0) or 0
            
            # Преобразуем значение в число и добавляем к сумме
            if value is not None:
                try:
                    if value == 'x' or value == '':
                        continue
                    num_value = float(value)
                    total += num_value
                    count += 1
                except (ValueError, TypeError):
                    continue
        
        # Форматируем и выводим результат
        if count > 0:
            formatted_total = f"{total:,.2f}".replace(",", " ")
            message = f"Выбрано строк: {count} | Сумма по столбцу '{column_name}': {formatted_total}"
            self.main_window.status_bar.showMessage(message)
        else:
            self.main_window.status_bar.showMessage("Готов к работе")
    
    def show_tree_context_menu(self, position):
        """Контекстное меню для дерева"""
        tree = self.main_window.data_tree
        menu = QMenu(self.main_window)
        item = tree.itemAt(position)
        copy_action = None
        if item:
            copy_action = menu.addAction("Копировать значение")
        
        action = menu.exec_(tree.mapToGlobal(position))
        
        if item and action == copy_action:
            self.copy_tree_item_value(item)
    
    def show_tree_header_context_menu(self, position):
        """Контекстное меню для заголовков дерева (скрытие/отображение столбцов)"""
        header = self.main_window.data_tree.header()
        tree = self.main_window.data_tree
        
        # position - это позиция относительно header
        col = header.logicalIndexAt(position.x())
        
        # Если не нашли столбец, пробуем найти визуальный индекс
        if col < 0:
            visual_col = header.visualIndexAt(position.x())
            if visual_col >= 0:
                col = header.logicalIndex(visual_col)
        
        # Если все еще не нашли, проверяем, что клик в пределах заголовка
        if col < 0:
            header_rect = header.rect()
            if not header_rect.contains(position):
                return
            col = 0  # Используем первый столбец как дефолтный
        
        menu = QMenu(self.main_window)
        select_columns_action = menu.addAction("Выбрать столбцы...")
        chosen = menu.exec_(header.mapToGlobal(position))

        if chosen == select_columns_action:
            self.show_column_visibility_dialog()

    
    def copy_tree_item_value(self, item):
        """Копировать значение из дерева"""
        if item:
            text = item.text(0)  # Копируем значение из первого столбца
            clipboard = QApplication.clipboard()
            clipboard.setText(text)
    
    def on_tree_item_expanded(self, item):
        """Обработка разворачивания узла дерева"""
        pass
    
    def on_tree_item_collapsed(self, item):
        """Обработка сворачивания узла дерева"""
        pass
    
    def show_column_visibility_dialog(self):
        """Показать диалог выбора столбцов для отображения"""
        try:
            tree_widget = self.main_window.data_tree
            if not tree_widget:
                return
            
            dialog = ColumnVisibilityDialog(tree_widget, self.main_window)
            dialog.exec_()
        except Exception as e:
            logger.error(f"Ошибка при открытии диалога выбора столбцов: {e}", exc_info=True)
    
    def _get_tree_widgets(self):
        """Получить все виджеты дерева (в главном окне и открепленных)"""
        widgets = []
        # Виджет в главном окне
        if hasattr(self.main_window, 'data_tree') and self.main_window.data_tree:
            widgets.append(self.main_window.data_tree)
        
        # Виджеты в открепленных окнах
        if hasattr(self.main_window, 'detached_windows') and "Древовидные данные" in self.main_window.detached_windows:
            detached_window = self.main_window.detached_windows["Древовидные данные"]
            tab_widget = detached_window.get_tab_widget()
            if tab_widget:
                from PyQt5.QtWidgets import QTreeWidget
                for child in tab_widget.findChildren(QTreeWidget):
                    if child not in widgets:
                        widgets.append(child)
        
        return widgets if widgets else []

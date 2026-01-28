"""Построение дерева из данных"""
from PyQt5.QtWidgets import QTreeWidget, QTreeWidgetItem, QHeaderView
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QColor, QBrush
from logger import logger
from models.constants.form_0503317_constants import Form0503317Constants


class TreeBuilder:
    """Класс для построения дерева из данных"""
    
    def __init__(self, main_window):
        """
        Args:
            main_window: Ссылка на главное окно для доступа к методам и свойствам
        """
        self.main_window = main_window
    
    def build_tree_from_data(self, data, tree_widget=None):
        """Построение дерева из данных"""
        try:
            if tree_widget is None:
                tree_widget = self.main_window.data_tree
            
            if not data:
                return
            
            if not isinstance(data, list) or len(data) == 0:
                return
            
            # Цвета для уровней
            level_colors = {
                0: "#E6E6FA", 1: "#68e368", 2: "#98FB98", 3: "#FFFF99", 
                4: "#FFB366", 5: "#FF9999", 6: "#FFCCCC"
            }
            
            # Строим дерево, учитывая последовательность уровней:
            # каждая строка является дочерней для ближайшей предыдущей строки
            # с меньшим уровнем (обычно level-1).
            parents_stack = []  # список кортежей (level, QTreeWidgetItem)
            items_created = 0
            items_failed = 0

            for item in data:
                try:
                    if not isinstance(item, dict):
                        items_failed += 1
                        continue
                    
                    level = item.get('уровень', 0)
                    tree_item = self.create_tree_item(item, level_colors, tree_widget)
                
                    # Убираем из стека все уровни, которые не могут быть родителями
                    while parents_stack and parents_stack[-1][0] >= level:
                        parents_stack.pop()

                    if parents_stack:
                        # Текущий элемент становится ребёнком последнего подходящего родителя
                        parents_stack[-1][1].addChild(tree_item)
                    else:
                        # Если родителя нет, это корневой элемент
                        tree_widget.addTopLevelItem(tree_item)

                    # Запоминаем текущий элемент как последний для своего уровня
                    parents_stack.append((level, tree_item))
                    items_created += 1
                except Exception as e:
                    items_failed += 1
                    logger.warning(f"Ошибка создания элемента дерева: {e}", exc_info=True)
                    continue
            
            # Разворачиваем уровень 0
            for i in range(tree_widget.topLevelItemCount()):
                try:
                    tree_widget.topLevelItem(i).setExpanded(True)
                except:
                    pass

            if items_created > 0 and tree_widget == self.main_window.data_tree:
                msg = f"Построено дерево: {items_created} элементов"
                if items_failed > 0:
                    msg += f", ошибок: {items_failed}"
                self.main_window.status_bar.showMessage(msg)
        except Exception as e:
            error_msg = f"Ошибка построения дерева: {e}"
            logger.error(error_msg, exc_info=True)
            if tree_widget == self.main_window.data_tree:
                self.main_window.status_bar.showMessage(error_msg)
    
    def create_tree_item(self, item, level_colors, tree_widget=None):
        """Создание элемента дерева"""
        try:
            if tree_widget is None:
                tree_widget = self.main_window.data_tree
            
            level = item.get('уровень', 0)

            column_count = tree_widget.columnCount()
            if column_count == 0:
                # Если колонок нет, создаем хотя бы одну
                tree_widget.setColumnCount(1)
                column_count = 1
            
            tree_item = QTreeWidgetItem([""] * column_count)

            # Маппинг «исходный индекс столбца -> индекс в дереве» (учёт скрытых столбцов)
            visible_columns = getattr(self.main_window, 'tree_visible_columns', None)
            if not visible_columns or len(visible_columns) != column_count:
                visible_columns = list(range(column_count))
            source_to_proxy = {src: idx for idx, src in enumerate(visible_columns)}

            # Основные данные
            name = str(item.get('наименование_показателя', ''))
            code_line = str(item.get('код_строки', ''))
            class_code = str(item.get('код_классификации_форматированный', item.get('код_классификации', '')))

            # 0: Наименование
            if 0 in source_to_proxy:
                col_idx = source_to_proxy[0]
                tree_item.setText(col_idx, name)
            # 1: Код строки
            if 1 in source_to_proxy:
                col_idx = source_to_proxy[1]
                tree_item.setText(col_idx, code_line)
            # 2: Код классификации
            if 2 in source_to_proxy:
                col_idx = source_to_proxy[2]
                tree_item.setText(col_idx, class_code)
            # 3: Уровень
            if 3 in source_to_proxy:
                col_idx = source_to_proxy[3]
                tree_item.setText(col_idx, str(level))

            # Получаем mapping из main_window
            mapping = getattr(self.main_window, 'tree_column_mapping', {})
            column_type = mapping.get("type", "base")

            if column_type == "budget":
                budget_cols = mapping.get("budget_columns", [])
                approved_start = mapping.get("approved_start", 4)
                executed_start = mapping.get("executed_start", approved_start + len(budget_cols))
                approved_data = item.get('утвержденный', {}) or {}
                executed_data = item.get('исполненный', {}) or {}
                
                # Цвет для выделения несоответствий (красный)
                error_color = QColor("#FF6B6B")

                for idx, col in enumerate(budget_cols):
                    try:
                        # Утвержденные значения
                        original_approved = approved_data.get(col, 0) or 0
                        calculated_approved = item.get(f'расчетный_утвержденный_{col}', original_approved)
                        
                        # Проверяем несоответствие (только для уровней < 6)
                        if level < 6 and self._is_value_different(original_approved, calculated_approved):
                            # Показываем значение с расчетным в скобках
                            if isinstance(original_approved, (int, float)) and isinstance(calculated_approved, (int, float)):
                                approved_value = f"{original_approved:,.2f} ({calculated_approved:,.2f})"
                            else:
                                approved_value = f"{original_approved} ({calculated_approved})"
                            # Выделяем красным цветом
                            src_col = approved_start + idx
                            proxy_col = source_to_proxy.get(src_col)
                            if proxy_col is not None:
                                tree_item.setText(proxy_col, approved_value)
                                tree_item.setForeground(proxy_col, QBrush(error_color))
                        else:
                            approved_value = self.format_budget_value(original_approved)
                            src_col = approved_start + idx
                            proxy_col = source_to_proxy.get(src_col)
                            if proxy_col is not None:
                                tree_item.setText(proxy_col, approved_value)
                        
                        # Исполненные значения
                        original_executed = executed_data.get(col, 0) or 0
                        calculated_executed = item.get(f'расчетный_исполненный_{col}', original_executed)
                        
                        # Проверяем несоответствие (только для уровней < 6)
                        if level < 6 and self._is_value_different(original_executed, calculated_executed):
                            # Показываем значение с расчетным в скобках
                            if isinstance(original_executed, (int, float)) and isinstance(calculated_executed, (int, float)):
                                executed_value = f"{original_executed:,.2f} ({calculated_executed:,.2f})"
                            else:
                                executed_value = f"{original_executed} ({calculated_executed})"
                            # Выделяем красным цветом
                            src_col = executed_start + idx
                            proxy_col = source_to_proxy.get(src_col)
                            if proxy_col is not None:
                                tree_item.setText(proxy_col, executed_value)
                                tree_item.setForeground(proxy_col, QBrush(error_color))
                        else:
                            executed_value = self.format_budget_value(original_executed)
                            src_col = executed_start + idx
                            proxy_col = source_to_proxy.get(src_col)
                            if proxy_col is not None:
                                tree_item.setText(proxy_col, executed_value)
                    except Exception as e:
                        logger.warning(f"Ошибка обработки несоответствий для колонки {col}: {e}", exc_info=True)
                        pass

            elif column_type == "consolidated":
                value_start = mapping.get("value_start", 4)
                cons_cols = mapping.get("columns", [])
                
                # Получаем данные поступлений (может быть вложенным словарем или плоскими полями)
                cons_data = item.get('поступления', {}) or {}
                
                # Цвет для выделения несоответствий (красный)
                error_color = QColor("#FF6B6B")
                
                for idx, col in enumerate(cons_cols):
                    try:
                        # Оригинальное значение - проверяем и вложенный словарь, и плоские поля
                        if isinstance(cons_data, dict) and col in cons_data:
                            original_value = cons_data.get(col, 0) or 0
                        else:
                            # Если нет вложенного словаря, проверяем плоские поля
                            original_value = item.get(f'поступления_{col}', 0) or 0
                        
                        # Расчетное значение - проверяем плоские поля (после to_dict('records'))
                        calculated_value = item.get(f'расчетный_поступления_{col}')
                        if calculated_value is None:
                            # Fallback на оригинальное значение, если расчетного нет
                            calculated_value = original_value
                        
                        # Проверяем несоответствие (аналогично бюджетным разделам — до 5 уровня),
                        # а для столбца "ИТОГО" проверяем на всех уровнях, так как это итоговая сумма
                        is_total_column = (col == 'ИТОГО')
                        should_check = (level < 6) or is_total_column
                        
                        if should_check and self._is_value_different(original_value, calculated_value):
                            # Показываем значение с расчетным в скобках
                            if isinstance(original_value, (int, float)) and isinstance(calculated_value, (int, float)):
                                display_value = f"{original_value:,.2f} ({calculated_value:,.2f})"
                            else:
                                display_value = f"{original_value} ({calculated_value})"
                            # Выделяем красным цветом
                            src_col = value_start + idx
                            proxy_col = source_to_proxy.get(src_col)
                            if proxy_col is not None:
                                tree_item.setText(proxy_col, display_value)
                                tree_item.setForeground(proxy_col, QBrush(error_color))
                        else:
                            # Обычное отображение без несоответствий
                            src_col = value_start + idx
                            proxy_col = source_to_proxy.get(src_col)
                            if proxy_col is not None:
                                tree_item.setText(proxy_col, self.format_budget_value(original_value))
                    except Exception as e:
                        logger.warning(f"Ошибка обработки несоответствий для консолидируемых расчетов, колонка {col}: {e}", exc_info=True)
                        pass
            
            # Устанавливаем цвет фона для всех видимых столбцов
            try:
                if level in level_colors:
                    color = QColor(level_colors[level])
                    brush = QBrush(color)
                    for proxy_col in range(column_count):
                        tree_item.setBackground(proxy_col, brush)
            except Exception as e:
                logger.warning(f"Ошибка установки цвета фона для уровня {level}: {e}", exc_info=True)
                pass
            
            # Устанавливаем подсказки (колонка -> заголовок) с учетом видимых столбцов
            try:
                tree_header_tooltips = getattr(self.main_window, 'tree_header_tooltips', [])
                for proxy_col, src_col in enumerate(visible_columns):
                    if 0 <= src_col < len(tree_header_tooltips):
                        tip = tree_header_tooltips[src_col]
                        current_text = tree_item.text(proxy_col)
                        if current_text:
                            tree_item.setToolTip(proxy_col, f"{tip}: {current_text}")
                        else:
                            tree_item.setToolTip(proxy_col, tip)
            except:
                pass

            # Сохраняем исходные данные
            try:
                tree_item.setData(0, Qt.UserRole, item)
            except:
                pass
            
            return tree_item
        except Exception as e:
            logger.error(f"Ошибка создания элемента дерева: {e}", exc_info=True)
            # Возвращаем пустой элемент в случае ошибки
            column_count = max(self.main_window.data_tree.columnCount(), 1)
            tree_item = QTreeWidgetItem([""] * column_count)
            return tree_item
    
    def _is_value_different(self, original: float, calculated: float) -> bool:
        """Проверка различия значений (аналогично методу в Form0503317)"""
        try:
            original_val = float(original) if original not in (None, "", "x") else 0.0
            calculated_val = float(calculated) if calculated not in (None, "", "x") else 0.0
            return abs(original_val - calculated_val) > 0.00001
        except (ValueError, TypeError):
            return False
    
    def format_budget_value(self, value):
        """Форматирование значения бюджета для отображения"""
        if value in (None, "", "0", 0):
            return ""
        if value == 'x':
            return 'x'
        try:
            return f"{float(value):,.2f}"
        except (ValueError, TypeError):
            return str(value)
    
    def load_project_data_to_tree(self, project):
        """Загрузка данных проекта в древовидное представление"""
        try:
            if not project:
                self.main_window.status_bar.showMessage("Проект не выбран")
                return
            
            if not project.data:
                self.main_window.status_bar.showMessage("В проекте нет данных для отображения")
                # Очищаем все деревья
                tree_widgets = self._get_tree_widgets()
                if tree_widgets:
                    for tree in tree_widgets:
                        if tree:
                            tree.clear()
                return
            
            # Получаем все виджеты дерева
            tree_widgets = self._get_tree_widgets()
            
            # Проверяем, что есть хотя бы одно дерево
            if not tree_widgets:
                logger.warning("Не найдены виджеты дерева для загрузки данных")
                self.main_window.status_bar.showMessage("Ошибка: виджеты дерева не инициализированы")
                return
            
            # Очищаем все деревья
            for tree in tree_widgets:
                if tree:
                    tree.clear()
            
            # Загружаем данные текущего раздела
            section_map = {
                "Доходы": "доходы_data",
                "Расходы": "расходы_data", 
                "Источники финансирования": "источники_финансирования_data",
                "Консолидируемые расчеты": "консолидируемые_расчеты_data"
            }

            # Настраиваем заголовки дерева под выбранный раздел
            if hasattr(self.main_window, 'tree_config'):
                self.main_window.tree_config.configure_tree_headers(self.main_window.current_section)
            elif hasattr(self.main_window, 'configure_tree_headers'):
                self.main_window.configure_tree_headers(self.main_window.current_section)
            
            section_key = section_map.get(self.main_window.current_section)
            if section_key and section_key in project.data:
                data = project.data[section_key]
                if data and len(data) > 0:
                    # Для раздела "Расходы" подсвечиваем строку 450, сравнивая
                    # план/исполнение с пересчитанным результатом исполнения бюджета
                    # (дефицит/профицит), который теперь берём из calculated_deficit_proficit.
                    if (
                        self.main_window.current_section == "Расходы"
                        and project.data.get('calculated_deficit_proficit')
                    ):
                        результат_data = project.data['calculated_deficit_proficit']
                        # Ищем строку с кодом 450
                        for row in data:
                            if str(row.get('код_строки', '')).strip() == '450':
                                # Добавляем расчетные значения для проверки несоответствий
                                for col in Form0503317Constants.BUDGET_COLUMNS:
                                    row[f'расчетный_утвержденный_{col}'] = результат_data.get(
                                        'утвержденный', {}
                                    ).get(col, 0)
                                    row[f'расчетный_исполненный_{col}'] = результат_data.get(
                                        'исполненный', {}
                                    ).get(col, 0)
                                break
                    
                    # Строим дерево для всех виджетов (в главном окне и открепленных)
                    for tree_widget in tree_widgets:
                        # Заголовки уже настроены в configure_tree_headers с учетом видимых столбцов
                        self.build_tree_from_data(data, tree_widget)
                    
                    # Обновляем высоту заголовка после загрузки данных
                    # Обновляем синхронно и через таймер для надежности
                    if hasattr(self.main_window, 'tree_config'):
                        self.main_window.tree_config._update_tree_header_height_for_all()
                        from PyQt5.QtCore import QTimer
                        QTimer.singleShot(100, lambda: self.main_window.tree_config._update_tree_header_height_for_all())
                    elif hasattr(self.main_window, '_update_tree_header_height_for_all'):
                        self.main_window._update_tree_header_height_for_all()
                        from PyQt5.QtCore import QTimer
                        QTimer.singleShot(100, lambda: self.main_window._update_tree_header_height_for_all())
                    
                    # Обновляем вкладку ошибок
                    if hasattr(self.main_window, 'errors_manager'):
                        self.main_window.errors_manager.load_errors_to_tab(project.data)
                    elif hasattr(self.main_window, 'load_errors_to_tab'):
                        self.main_window.load_errors_to_tab(project.data)
                    
                    # Применяем скрытие нулевых столбцов, если чекбокс включен
                    self.main_window.status_bar.showMessage(f"Загружено {len(data)} записей в разделе '{self.main_window.current_section}'")
                else:
                    self.main_window.status_bar.showMessage(f"В разделе '{self.main_window.current_section}' нет данных для отображения")
            else:
                self.main_window.status_bar.showMessage(f"Раздел '{self.main_window.current_section}' не найден в данных проекта")
        except Exception as e:
            error_msg = f"Ошибка загрузки данных в дерево: {e}"
            logger.error(error_msg, exc_info=True)
            self.main_window.status_bar.showMessage(error_msg)
    
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

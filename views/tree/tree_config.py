"""Конфигурация заголовков дерева"""
from PyQt5.QtWidgets import QHeaderView, QTreeWidget, QApplication
from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtGui import QTextDocument, QTextOption
from views.widgets import WrapHeaderView
from views.widgets import WordWrapItemDelegate
from models.form_0503317 import Form0503317Constants
from logger import logger


class TreeConfig:
    """Класс для конфигурации заголовков дерева"""
    
    def __init__(self, main_window):
        """
        Args:
            main_window: Ссылка на главное окно для доступа к свойствам
        """
        self.main_window = main_window
        self.tree_headers = []
        self.tree_header_tooltips = []
        self.tree_column_mapping = {}
    
    def configure_tree_headers(self, section_name: str):
        """Конфигурация заголовков дерева под выбранный раздел"""
        base_headers = ["Наименование", "Код строки", "Код классификации", "Уровень"]
        display_headers = base_headers[:]
        tooltip_headers = base_headers[:]
        mapping = {
            "type": "base",
            "base_count": len(base_headers)
        }

        if section_name in ["Доходы", "Расходы", "Источники финансирования"]:
            budget_cols = Form0503317Constants.BUDGET_COLUMNS
            mapping.update({
                "type": "budget",
                "budget_columns": budget_cols,
                "approved_start": len(display_headers),
                "executed_start": len(display_headers) + len(budget_cols)
            })

            for col in budget_cols:
                display_headers.append(f"У. {col}")
                tooltip_headers.append(f"Утвержденный — {col}")
            for col in budget_cols:
                display_headers.append(f"И. {col}")
                tooltip_headers.append(f"Исполненный — {col}")

        elif section_name == "Консолидируемые расчеты":
            cons_cols = Form0503317Constants.CONSOLIDATED_COLUMNS
            mapping.update({
                "type": "consolidated",
                "value_start": len(display_headers),
                "columns": cons_cols
            })
            for col in cons_cols:
                display_headers.append(col)
                tooltip_headers.append(col)

        # Сохраняем в main_window для обратной совместимости
        self.main_window.tree_headers = display_headers
        self.main_window.tree_header_tooltips = tooltip_headers
        self.main_window.tree_column_mapping = mapping
        
        # Также сохраняем в себе
        self.tree_headers = display_headers
        self.tree_header_tooltips = tooltip_headers
        self.tree_column_mapping = mapping

        # Настраиваем заголовки для всех деревьев
        for tree_widget in self._get_tree_widgets():
            self._configure_tree_headers_for_widget(tree_widget, section_name, display_headers, mapping)

        # Вычисляем высоту заголовка с учетом автоматического переноса текста
        # Обновляем высоту синхронно для всех деревьев
        self._update_tree_header_height_for_all()
        # Также обновляем через таймер на случай, если размеры столбцов еще не установлены
        QTimer.singleShot(100, lambda: self._update_tree_header_height_for_all())
    
    def _configure_tree_headers_for_widget(self, tree_widget, section_name, display_headers=None, mapping=None):
        """Настройка заголовков для конкретного виджета дерева"""
        if display_headers is None:
            display_headers = self.tree_headers or getattr(self.main_window, 'tree_headers', [])
        if mapping is None:
            mapping = self.tree_column_mapping or getattr(self.main_window, 'tree_column_mapping', {})
        
        # Устанавливаем делегат для переноса текста в ячейках
        tree_widget.setItemDelegate(WordWrapItemDelegate())
        # Отключаем единую высоту строк, чтобы высота подстраивалась под содержимое
        tree_widget.setUniformRowHeights(False)
        
        # Применяем текущий размер шрифта
        font = tree_widget.font()
        font.setPointSize(self.main_window.font_size)
        tree_widget.setFont(font)
        
        tree_widget.setColumnCount(len(display_headers))
        
        # Проверяем, есть ли уже кастомный заголовок, если нет - создаем новый
        header = tree_widget.header()
        if not isinstance(header, WrapHeaderView):
            # Создаем и устанавливаем кастомный заголовок с поддержкой переноса текста
            custom_header = WrapHeaderView(Qt.Horizontal, tree_widget)
            custom_header.setHeaderTexts({idx: text for idx, text in enumerate(display_headers)})
            tree_widget.setHeader(custom_header)
            header = tree_widget.header()
        
        # Устанавливаем заголовки ПОСЛЕ установки кастомного заголовка
        tree_widget.setHeaderLabels(display_headers)
        
        # Убеждаемся, что заголовок видим
        tree_widget.setHeaderHidden(False)
        
        # После setHeaderLabels нужно снова получить заголовок, так как он может быть пересоздан
        header = tree_widget.header()
        
        # Если заголовок не кастомный, создаем и устанавливаем его снова
        if not isinstance(header, WrapHeaderView):
            custom_header = WrapHeaderView(Qt.Horizontal, tree_widget)
            custom_header.setHeaderTexts({idx: text for idx, text in enumerate(display_headers)})
            tree_widget.setHeader(custom_header)
            header = tree_widget.header()
        
        # Обновляем тексты заголовков в кастомном заголовке
        if isinstance(header, WrapHeaderView):
            header.setHeaderTexts({idx: text for idx, text in enumerate(display_headers)})
        
        header.setDefaultAlignment(Qt.AlignCenter)
        
        # Применяем размер шрифта к заголовкам
        header_font = header.font()
        header_font.setPointSize(self.main_window.header_font_size)
        header.setFont(header_font)
        
        # Включаем перенос текста в заголовках
        header.setTextElideMode(Qt.ElideNone)
        
        # Убеждаемся, что заголовок видим
        tree_widget.setHeaderHidden(False)
        
        # Устанавливаем минимальную ширину столбцов
        for idx in range(len(display_headers)):
            header.setMinimumSectionSize(50)
        
        # Устанавливаем режимы изменения размера и ширину столбцов
        # Столбец 0 ("Наименование") - Interactive с фиксированной шириной с учетом отступов
        # Столбец 1 ("Код строки") - Fixed с шириной 80px
        # Столбец 2 ("Код классификации") - Interactive с фиксированной шириной 200px
        # Столбец 3 ("Уровень") - Fixed с шириной 50px
        # Остальные столбцы - Fixed с шириной 150px (текст будет переноситься)
        for idx in range(len(display_headers)):
            if idx == 0:
                # Столбец "Наименование" - Interactive режим с фиксированной шириной
                header.setSectionResizeMode(idx, QHeaderView.Interactive)
                # Получаем отступы дерева и добавляем запас
                indentation = tree_widget.indentation()
                # Добавляем запас на отступы (примерно 6 уровней * отступ + небольшой запас)
                indent_reserve = indentation * 6 + 50  # Запас на отступы и дополнительные элементы
                # Устанавливаем ширину 400 пикселей + запас на отступы
                header.resizeSection(idx, 400 + indent_reserve)
            elif idx == 1:
                # Столбец "Код строки" - Fixed режим с шириной 80px
                header.setSectionResizeMode(idx, QHeaderView.Fixed)
                header.resizeSection(idx, 80)
            elif idx == 2:
                # Столбец "Код классификации" - Interactive режим с фиксированной шириной
                header.setSectionResizeMode(idx, QHeaderView.Interactive)
                header.resizeSection(idx, 200)
            elif idx == 3:
                # Столбец "Уровень" - Fixed режим с шириной 50px
                header.setSectionResizeMode(idx, QHeaderView.Fixed)
                header.resizeSection(idx, 50)
            else:
                # Остальные столбцы - Fixed режим с шириной 150px
                header.setSectionResizeMode(idx, QHeaderView.Fixed)
                header.resizeSection(idx, 150)
        
        # Подключаем обработчик изменения размера столбцов для обновления высоты заголовка
        # и ограничения ширины столбца "Наименование"
        def on_section_resized(logical_index, old_size, new_size):
            # Ограничиваем ширину столбца "Наименование" (индекс 0) с учетом отступов
            if logical_index == 0:
                indentation = tree_widget.indentation()
                indent_reserve = indentation * 6 + 50  # Запас на отступы
                max_width = 400 + indent_reserve
                if new_size > max_width:
                    header.resizeSection(0, max_width)
            # Для столбцов с Fixed режимом восстанавливаем их фиксированные размеры
            elif logical_index == 1:  # Столбец "Код строки" - 80px
                if header.sectionResizeMode(logical_index) == QHeaderView.Fixed:
                    if new_size != 80:
                        header.resizeSection(logical_index, 80)
            elif logical_index == 3:  # Столбец "Уровень" - 50px
                if header.sectionResizeMode(logical_index) == QHeaderView.Fixed:
                    if new_size != 50:
                        header.resizeSection(logical_index, 50)
            elif logical_index != 2:  # Остальные столбцы (кроме 0 и 2) - 150px
                # Проверяем, что это столбец с Fixed режимом
                if header.sectionResizeMode(logical_index) == QHeaderView.Fixed:
                    if new_size != 150:
                        header.resizeSection(logical_index, 150)
            QTimer.singleShot(50, lambda tw=tree_widget: self._update_tree_header_height(tw))
        
        header.sectionResized.connect(on_section_resized)
        
        # Обновляем тексты заголовков в кастомном заголовке при изменении размера
        if isinstance(header, WrapHeaderView):
            header.setHeaderTexts({idx: text for idx, text in enumerate(display_headers)})
            header.update()  # Принудительно обновляем отрисовку

        # Для консолидируемых расчетов колонку "Код классификации" не показываем
        # Для остальных разделов - показываем
        if section_name == "Консолидируемые расчеты" and len(display_headers) > 2:
            tree_widget.setColumnHidden(2, True)
        else:
            # Убеждаемся, что столбец "Код классификации" видим для других разделов
            if len(display_headers) > 2:
                tree_widget.setColumnHidden(2, False)
        
        # Обновляем высоту заголовка сразу после настройки
        # Это предотвращает наезд заголовка на данные при смене раздела
        QApplication.processEvents()  # Обрабатываем события, чтобы заголовки были установлены
        self._update_tree_header_height(tree_widget)

    def _update_tree_header_height_for_all(self):
        """Обновляет высоту заголовка для всех деревьев"""
        for tree_widget in self._get_tree_widgets():
            self._update_tree_header_height(tree_widget)
    
    def _update_tree_header_height(self, tree_widget=None):
        """Обновляет высоту заголовка дерева с учетом автоматического переноса текста"""
        if tree_widget is None:
            tree_widget = self.main_window.data_tree
        try:
            header = tree_widget.header()
            font_metrics = header.fontMetrics()
            max_height = 0
            
            # Получаем заголовки из headerItem
            header_item = tree_widget.headerItem()
            tree_headers = self.tree_headers or getattr(self.main_window, 'tree_headers', [])
            
            if header_item:
                # Проходим по всем заголовкам и вычисляем максимальную высоту с учетом переноса
                for idx in range(tree_widget.columnCount()):
                    if tree_widget.isColumnHidden(idx):
                        continue
                    
                    # Получаем текст из headerItem
                    text = header_item.text(idx) if idx < tree_widget.columnCount() else ""
                    if not text and idx < len(tree_headers):
                        text = tree_headers[idx]
                    
                    if text:
                        # Получаем ширину столбца
                        width = max(header.sectionSize(idx), 50)
                        
                        # Создаем документ для расчета высоты с учетом переноса
                        doc = QTextDocument()
                        doc.setDefaultFont(header.font())
                        doc.setPlainText(str(text))
                        
                        # Настраиваем перенос текста
                        text_option = QTextOption()
                        text_option.setWrapMode(QTextOption.WrapAtWordBoundaryOrAnywhere)
                        doc.setDefaultTextOption(text_option)
                        
                        # Устанавливаем ширину документа (с учетом отступов)
                        padding = 4
                        doc.setTextWidth(width - 2 * padding)
                        
                        # Получаем высоту документа
                        doc_height = doc.size().height()
                        max_height = max(max_height, doc_height)
            else:
                # Если нет headerItem, используем tree_headers
                for idx, text in enumerate(tree_headers):
                    if text and not tree_widget.isColumnHidden(idx):
                        width = max(header.sectionSize(idx), 50)
                        
                        # Создаем документ для расчета высоты
                        doc = QTextDocument()
                        doc.setDefaultFont(header.font())
                        doc.setPlainText(str(text))
                        
                        text_option = QTextOption()
                        text_option.setWrapMode(QTextOption.WrapAtWordBoundaryOrAnywhere)
                        doc.setDefaultTextOption(text_option)
                        doc.setTextWidth(width)
                        
                        doc_height = doc.size().height()
                        max_height = max(max_height, doc_height)
            
            # Устанавливаем высоту заголовка с небольшим отступом
            if max_height > 0:
                header.setFixedHeight(int(max_height) + 8)
            else:
                # Если не удалось рассчитать, используем стандартную высоту
                line_height = font_metrics.lineSpacing()
                header.setFixedHeight(line_height + 6)
        except Exception as e:
            logger.warning(f"Ошибка обновления высоты заголовка дерева: {e}", exc_info=True)
            # В случае ошибки используем минимальную высоту
            try:
                header = self.main_window.data_tree.header()
                font_metrics = header.fontMetrics()
                header.setFixedHeight(font_metrics.lineSpacing() + 6)
            except:
                pass
    
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
                for child in tab_widget.findChildren(QTreeWidget):
                    if child not in widgets:
                        widgets.append(child)
        
        return widgets if widgets else []
    
    def hide_zero_columns_in_tree(self, section_key: str, data):
        """
        Скрытие столбцов дерева, в которых итоговое значение равно 0.
        Логика аналогична табличному представлению.
        """
        if not data:
            return

        if section_key == "консолидируемые_расчеты_data":
            cons_cols = Form0503317Constants.CONSOLIDATED_COLUMNS
            mapping = self.tree_column_mapping or getattr(self.main_window, 'tree_column_mapping', {})
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

            header = self.main_window.data_tree.header()
            zero_cols = []
            for i, col_name in enumerate(cons_cols):
                val = totals.get(col_name, 0)
                if isinstance(val, (int, float)) and abs(val) < 1e-9:
                    col_index = value_start + i
                    if 0 <= col_index < self.main_window.data_tree.columnCount():
                        zero_cols.append(col_index)

            # Сужаем «нулевые» колонки до минимальной ширины и очищаем заголовки
            header_item = self.main_window.data_tree.headerItem()
            for col_index in zero_cols:
                header.resizeSection(col_index, 2)  # минимальная ширина
                if header_item:
                    header_item.setText(col_index, "")
                    header_item.setToolTip(col_index, "")
            return

        # Доходы, расходы, источники
        budget_cols = Form0503317Constants.BUDGET_COLUMNS
        mapping = self.tree_column_mapping or getattr(self.main_window, 'tree_column_mapping', {})
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

        header = self.main_window.data_tree.header()
        zero_cols = set()
        for i, col_name in enumerate(budget_cols):
            a_val = approved.get(col_name, 0) or 0
            e_val = executed.get(col_name, 0) or 0
            if isinstance(a_val, (int, float)) and isinstance(e_val, (int, float)):
                if abs(a_val) < 1e-9 and abs(e_val) < 1e-9:
                    appr_idx = approved_start + i
                    exec_idx = executed_start + i
                    if show_approved and 0 <= appr_idx < self.main_window.data_tree.columnCount():
                        zero_cols.add(appr_idx)
                    if show_executed and 0 <= exec_idx < self.main_window.data_tree.columnCount():
                        zero_cols.add(exec_idx)

        # Сужаем «нулевые» колонки до минимальной ширины и очищаем заголовки
        header_item = self.main_window.data_tree.headerItem()
        for col_index in zero_cols:
            header.resizeSection(col_index, 2)  # минимальная ширина
            if header_item:
                header_item.setText(col_index, "")
                header_item.setToolTip(col_index, "")
    
    def apply_tree_data_type_visibility(self):
        """Скрывает столбцы дерева в зависимости от выбранного типа данных"""
        mapping = self.tree_column_mapping or getattr(self.main_window, 'tree_column_mapping', {})
        if not mapping:
            return

        column_total = len(self.tree_headers or getattr(self.main_window, 'tree_headers', []))
        
        if mapping.get("type") == "budget":
            budget_cols = mapping.get("budget_columns", [])
            approved_start = mapping.get("approved_start", 4)
            executed_start = mapping.get("executed_start", approved_start + len(budget_cols))
            
            current_data_type = getattr(self.main_window, 'current_data_type', 'Оба')
            show_approved = current_data_type in ("Утвержденный", "Оба")
            show_executed = current_data_type in ("Исполненный", "Оба")
            
            # Получаем все виджеты дерева
            tree_widgets = self._get_tree_widgets()
            
            for tree_widget in tree_widgets:
                approved_range = range(approved_start, executed_start)
                executed_range = range(executed_start, executed_start + len(budget_cols))
                
                for idx in approved_range:
                    tree_widget.setColumnHidden(idx, not show_approved)
                for idx in executed_range:
                    tree_widget.setColumnHidden(idx, not show_executed)

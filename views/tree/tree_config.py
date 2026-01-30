"""Конфигурация заголовков дерева"""
from PyQt5.QtWidgets import QHeaderView, QTreeWidget, QApplication
from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtGui import QTextDocument, QTextOption
from views.widgets import WrapHeaderView
from views.widgets import WordWrapItemDelegate
from models.constants.form_0503317_constants import Form0503317Constants
from logger import logger
from views.tree.tree_header_configurator import TreeHeaderConfigurator
from views.tree.tree_column_visibility_manager import TreeColumnVisibilityManager
from views.tree.tree_header_layout_helper import TreeHeaderLayoutHelper


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
        
        # Инициализация компонентов через композицию
        self.header_configurator = TreeHeaderConfigurator()
        self.visibility_manager = TreeColumnVisibilityManager(main_window)
        self.layout_helper = TreeHeaderLayoutHelper(main_window)
    
    def configure_tree_headers(self, section_name: str):
        """Конфигурация заголовков дерева под выбранный раздел"""
        # Используем конфигуратор заголовков
        config_result = self.header_configurator.configure_headers(section_name)
        display_headers = config_result["headers"]
        tooltip_headers = config_result["tooltips"]
        mapping = config_result["mapping"]

        # Сохраняем в main_window для обратной совместимости (полный список заголовков)
        self.main_window.tree_headers = display_headers
        self.main_window.tree_header_tooltips = tooltip_headers
        self.main_window.tree_column_mapping = mapping
        
        # Также сохраняем в себе
        self.tree_headers = display_headers
        self.tree_header_tooltips = tooltip_headers
        self.tree_column_mapping = mapping
        
        # Настраиваем заголовки для всех деревьев (показываем полный набор столбцов,
        # а видимость/скрытие управляется через setColumnHidden)
        for tree_widget in self._get_tree_widgets():
            self._configure_tree_headers_for_widget(tree_widget, section_name, display_headers, mapping)

        # Вычисляем высоту заголовка с учетом автоматического переноса текста
        self._update_tree_header_height_for_all()
        QTimer.singleShot(100, lambda: self._update_tree_header_height_for_all())
        
        # Восстанавливаем сохранённые настройки видимости столбцов (через setColumnHidden)
        QTimer.singleShot(150, lambda: self._restore_tree_column_visibility_for_all(section_name, display_headers))
    
    def _configure_tree_headers_for_widget(self, tree_widget, section_name, display_headers=None, mapping=None, tree_visible_columns=None):
        """Настройка заголовков для конкретного виджета дерева."""
        if display_headers is None:
            display_headers = self.tree_headers or getattr(self.main_window, 'tree_headers', [])
        if mapping is None:
            mapping = self.tree_column_mapping or getattr(self.main_window, 'tree_column_mapping', {})
        
        # Делегату передаём дерево, чтобы sizeHint сразу учитывал отступ по уровню (без запоздалого обновления высоты)
        tree_widget.setItemDelegate(WordWrapItemDelegate(tree_widget))
        tree_widget.setUniformRowHeights(False)
        
        font = tree_widget.font()
        font.setPointSize(self.main_window.font_size)
        tree_widget.setFont(font)
        
        tree_widget.setColumnCount(len(display_headers))
        
        header = tree_widget.header()
        if not isinstance(header, WrapHeaderView):
            custom_header = WrapHeaderView(Qt.Horizontal, tree_widget)
            custom_header.setHeaderTexts({idx: text for idx, text in enumerate(display_headers)})
            tree_widget.setHeader(custom_header)
            header = tree_widget.header()
        
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
        
        # Переустанавливаем контекстное меню для заголовка (на случай, если заголовок был пересоздан)
        header.setContextMenuPolicy(Qt.CustomContextMenu)
        # Отключаем старые соединения и подключаем заново
        try:
            header.customContextMenuRequested.disconnect()
        except:
            pass
        header.customContextMenuRequested.connect(self.main_window.show_tree_header_context_menu)
        
        header.setDefaultAlignment(Qt.AlignCenter)
        
        # Применяем размер шрифта к заголовкам
        header_font = header.font()
        header_font.setPointSize(self.main_window.header_font_size)
        header.setFont(header_font)
        
        # Включаем перенос текста в заголовках
        header.setTextElideMode(Qt.ElideNone)
        
        # Убеждаемся, что заголовок видим
        tree_widget.setHeaderHidden(False)
        
        for idx in range(len(display_headers)):
            header.setMinimumSectionSize(50)
        
        # Устанавливаем режимы и ширину столбцов по их индексу
        for idx in range(len(display_headers)):
            if idx == 0:
                header.setSectionResizeMode(idx, QHeaderView.Interactive)
                indentation = tree_widget.indentation() or 20
                indent_reserve = indentation * 3 + 24  # запас под отступы уровней без перебора
                header.resizeSection(idx, 400 + indent_reserve)
            elif idx == 1:
                header.setSectionResizeMode(idx, QHeaderView.Fixed)
                header.resizeSection(idx, 80)
            elif idx == 2:
                header.setSectionResizeMode(idx, QHeaderView.Interactive)
                header.resizeSection(idx, 200)
            elif idx == 3:
                header.setSectionResizeMode(idx, QHeaderView.Fixed)
                header.resizeSection(idx, 50)
            else:
                header.setSectionResizeMode(idx, QHeaderView.Fixed)
                header.resizeSection(idx, 150)
        
        def on_section_resized(logical_index, old_size, new_size):
            # Если столбец программно сворачивается до ширины 0 (скрытие),
            # не вмешиваемся в его ширину, чтобы избежать зацикливания resizeSection.
            if new_size == 0:
                return
            if logical_index == 0:
                indentation = tree_widget.indentation() or 20
                indent_reserve = indentation * 3 + 24
                max_width = 400 + indent_reserve
                if new_size > max_width:
                    header.resizeSection(logical_index, max_width)
                try:
                    tree_widget.doItemsLayout()
                except Exception:
                    pass
                QTimer.singleShot(0, lambda tw=tree_widget: tw.doItemsLayout() if tw else None)
            elif logical_index == 1 and header.sectionResizeMode(logical_index) == QHeaderView.Fixed and new_size != 80:
                header.resizeSection(logical_index, 80)
            elif logical_index == 3 and header.sectionResizeMode(logical_index) == QHeaderView.Fixed and new_size != 50:
                header.resizeSection(logical_index, 50)
            # Для остальных фиксированных столбцов, кроме 0,1,2,3, удерживаем ширину 150,
            # чтобы не вмешиваться в поведение специальных колонок (0,1,2,3).
            elif logical_index not in (0, 1, 2, 3) and header.sectionResizeMode(logical_index) == QHeaderView.Fixed and new_size != 150:
                header.resizeSection(logical_index, 150)
            QTimer.singleShot(50, lambda tw=tree_widget: self._update_tree_header_height(tw))
        
        header.sectionResized.connect(on_section_resized)
        
        if isinstance(header, WrapHeaderView):
            header.setHeaderTexts({idx: text for idx, text in enumerate(display_headers)})
            header.update()
        
        # Обновляем высоту заголовка сразу после настройки
        QApplication.processEvents()
        self._update_tree_header_height(tree_widget)

    def _update_tree_header_height_for_all(self):
        """Обновляет высоту заголовка для всех деревьев"""
        for tree_widget in self._get_tree_widgets():
            self._update_tree_header_height(tree_widget)
    
    def _update_tree_header_height(self, tree_widget=None):
        """Обновляет высоту заголовка дерева с учетом автоматического переноса текста"""
        if tree_widget is None:
            tree_widget = self.main_window.data_tree
        # Делегируем к layout_helper
        self.layout_helper.update_header_height(tree_widget)
    
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
    
    def _restore_tree_column_visibility(self, tree_widget, section_name: str, display_headers: list):
        """Восстанавливает сохраненные настройки видимости столбцов для дерева."""
        try:
            config_key = f"tree_columns:{section_name}"
            saved_column_visibility = self.main_window.controller.db_manager.load_config(config_key, {})
            
            if saved_column_visibility:
                header = tree_widget.header()
                for col in range(tree_widget.columnCount()):
                    if col < len(display_headers):
                        column_name = display_headers[col]
                        if column_name in saved_column_visibility:
                            is_visible = saved_column_visibility[column_name]
                            tree_widget.setColumnHidden(col, not is_visible)
                            if not is_visible:
                                header.setSectionResizeMode(col, QHeaderView.Fixed)
                                header.resizeSection(col, 0)
                                tree_widget.setColumnWidth(col, 0)
                                header.updateGeometries()
                                tree_widget.viewport().update()
        except Exception as e:
            logger.error(f"Ошибка при восстановлении настроек видимости столбцов дерева: {e}", exc_info=True)
    
    def _restore_tree_column_visibility_for_all(self, section_name: str, display_headers: list):
        """Восстанавливает сохраненные настройки видимости столбцов для всех деревьев."""
        for tree_widget in self._get_tree_widgets():
            self._restore_tree_column_visibility(tree_widget, section_name, display_headers)
    
    def apply_tree_data_type_visibility(self):
        """Скрывает столбцы дерева в зависимости от выбранного типа данных."""
        current_data_type = getattr(self.main_window, 'current_data_type', 'Оба')
        for tree_widget in self._get_tree_widgets():
            self.visibility_manager.apply_data_type_visibility(current_data_type, tree_widget)

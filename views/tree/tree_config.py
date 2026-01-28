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

        # Видимые столбцы: из конфига + фильтр по разделу + по типу данных (дерево без скрытых — только видимые колонки)
        config_key = f"tree_columns:{section_name}"
        saved_column_visibility = self.main_window.controller.db_manager.load_config(config_key, {})
        tree_visible_columns = [i for i in range(len(display_headers)) if saved_column_visibility.get(display_headers[i], True)]
        if section_name == "Консолидируемые расчеты" and len(display_headers) > 2 and 2 in tree_visible_columns:
            tree_visible_columns.remove(2)
        current_data_type = getattr(self.main_window, 'current_data_type', 'Оба')
        type_visible = self.visibility_manager.get_visible_indices_for_data_type(len(display_headers), mapping, current_data_type)
        tree_visible_columns = [i for i in tree_visible_columns if i in type_visible]
        if not tree_visible_columns:
            tree_visible_columns = list(range(len(display_headers)))
        self.main_window.tree_visible_columns = tree_visible_columns

        visible_headers = [display_headers[i] for i in tree_visible_columns]
        for tree_widget in self._get_tree_widgets():
            self._configure_tree_headers_for_widget(tree_widget, section_name, visible_headers, mapping, tree_visible_columns)

        # Вычисляем высоту заголовка с учетом автоматического переноса текста
        self._update_tree_header_height_for_all()
        QTimer.singleShot(100, lambda: self._update_tree_header_height_for_all())
    
    def _configure_tree_headers_for_widget(self, tree_widget, section_name, display_headers=None, mapping=None, tree_visible_columns=None):
        """Настройка заголовков для конкретного виджета дерева. display_headers — список видимых заголовков (подмножество полного)."""
        if display_headers is None:
            display_headers = self.tree_headers or getattr(self.main_window, 'tree_headers', [])
        if mapping is None:
            mapping = self.tree_column_mapping or getattr(self.main_window, 'tree_column_mapping', {})
        if tree_visible_columns is None:
            tree_visible_columns = getattr(self.main_window, 'tree_visible_columns', list(range(len(display_headers))))
        
        # Устанавливаем делегат для переноса текста в ячейках
        tree_widget.setItemDelegate(WordWrapItemDelegate())
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
        
        # Режимы и ширина по исходному индексу столбца (source_col)
        for idx in range(len(display_headers)):
            source_col = tree_visible_columns[idx] if idx < len(tree_visible_columns) else idx
            if source_col == 0:
                header.setSectionResizeMode(idx, QHeaderView.Interactive)
                indentation = tree_widget.indentation()
                indent_reserve = indentation * 6 + 50
                header.resizeSection(idx, 400 + indent_reserve)
            elif source_col == 1:
                header.setSectionResizeMode(idx, QHeaderView.Fixed)
                header.resizeSection(idx, 80)
            elif source_col == 2:
                header.setSectionResizeMode(idx, QHeaderView.Interactive)
                header.resizeSection(idx, 200)
            elif source_col == 3:
                header.setSectionResizeMode(idx, QHeaderView.Fixed)
                header.resizeSection(idx, 50)
            else:
                header.setSectionResizeMode(idx, QHeaderView.Fixed)
                header.resizeSection(idx, 150)
        
        def on_section_resized(logical_index, old_size, new_size):
            source_col = tree_visible_columns[logical_index] if logical_index < len(tree_visible_columns) else logical_index
            if source_col == 0:
                indentation = tree_widget.indentation()
                indent_reserve = indentation * 6 + 50
                max_width = 400 + indent_reserve
                if new_size > max_width:
                    header.resizeSection(logical_index, max_width)
            elif source_col == 1 and header.sectionResizeMode(logical_index) == QHeaderView.Fixed and new_size != 80:
                header.resizeSection(logical_index, 80)
            elif source_col == 3 and header.sectionResizeMode(logical_index) == QHeaderView.Fixed and new_size != 50:
                header.resizeSection(logical_index, 50)
            elif source_col not in (0, 2) and header.sectionResizeMode(logical_index) == QHeaderView.Fixed and new_size != 150:
                header.resizeSection(logical_index, 150)
            QTimer.singleShot(50, lambda tw=tree_widget: self._update_tree_header_height(tw))
        
        header.sectionResized.connect(on_section_resized)
        
        if isinstance(header, WrapHeaderView):
            header.setHeaderTexts({idx: text for idx, text in enumerate(display_headers)})
            header.update()
        
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
    
    def apply_tree_data_type_visibility(self):
        """Актуализирует видимость столбцов при смене типа данных.
        
        Фактическая видимость теперь задаётся через tree_visible_columns в configure_tree_headers,
        поэтому здесь достаточно при необходимости перенастроить заголовки для текущего раздела.
        Реальная перестройка дерева выполняется в load_project_data_to_tree.
        """
        # Ничего не делаем: configure_tree_headers будет вызван из load_project_data_to_tree
        return

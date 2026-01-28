"""
Диалог для выбора отображаемых столбцов в дереве данных и таблицах
"""
from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QScrollArea, QWidget, QCheckBox, QFrame, QTreeWidget, QTableWidget
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont
from logger import logger


class ColumnVisibilityDialog(QDialog):
    """Диалог для выбора отображаемых столбцов"""
    
    def __init__(self, widget, parent=None):
        """
        Args:
            widget: Виджет (QTreeWidget или QTableWidget), для которого настраивается видимость столбцов
            parent: Родительское окно
        """
        super().__init__(parent)
        self.widget = widget
        self.is_tree = isinstance(widget, QTreeWidget)
        self.is_table = isinstance(widget, QTableWidget)
        self.checkboxes = {}  # {column_index: checkbox}
        
        self.setWindowTitle("Выбор столбцов для отображения")
        self.setMinimumSize(400, 500)
        self.setModal(True)
        
        self.init_ui()
        self.load_current_state()
    
    def init_ui(self):
        """Инициализация интерфейса"""
        layout = QVBoxLayout(self)
        layout.setSpacing(15)
        layout.setContentsMargins(20, 20, 20, 20)
        
        # Заголовок
        header_frame = QFrame()
        header_frame.setObjectName("headerFrame")
        header_layout = QVBoxLayout(header_frame)
        header_layout.setContentsMargins(15, 10, 15, 10)
        
        title_label = QLabel("📋 Выбор столбцов для отображения")
        title_font = QFont()
        title_font.setPointSize(14)
        title_font.setBold(True)
        title_label.setFont(title_font)
        title_label.setWordWrap(True)
        header_layout.addWidget(title_label)
        
        widget_type = "таблице" if self.is_table else "дереве данных"
        info_label = QLabel(f"Выберите столбцы, которые должны быть видны в {widget_type}:")
        info_label.setWordWrap(True)
        info_label.setStyleSheet("border: none; background-color: transparent; color: #666666;")
        header_layout.addWidget(info_label)
        
        layout.addWidget(header_frame)
        
        # Область прокрутки для чекбоксов
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        
        # Виджет с чекбоксами
        content_widget = QWidget()
        content_layout = QVBoxLayout(content_widget)
        content_layout.setSpacing(8)
        content_layout.setContentsMargins(10, 10, 10, 10)
        
        headers = []
        
        if self.is_tree:
            # Для QTreeWidget используем полный список заголовков и сохранённые видимые колонки
            main_window = self.widget.window()
            tree_headers = getattr(main_window, 'tree_headers', [])
            headers = tree_headers or []
            column_count = len(headers)
        else:
            # Для QTableWidget работаем с фактическими колонками виджета
            column_count = self.widget.columnCount()
        
        if self.is_table:
            # Для QTableWidget
            for col in range(column_count):
                header_item = self.widget.horizontalHeaderItem(col)
                if header_item:
                    header_text = header_item.text()
                else:
                    # Пробуем получить из модели
                    model = self.widget.model()
                    if model:
                        header_text = model.headerData(col, Qt.Horizontal, Qt.DisplayRole)
                        if not header_text:
                            header_text = f"Столбец {col}"
                    else:
                        header_text = f"Столбец {col}"
                headers.append(header_text if header_text else f"Столбец {col}")
        
        # Создаем чекбоксы для каждого столбца
        if self.is_tree:
            main_window = self.widget.window()
            visible_columns = getattr(main_window, 'tree_visible_columns', list(range(len(headers))))
            for col in range(column_count):
                header_text = headers[col] if col < len(headers) else f"Столбец {col}"
                checkbox = QCheckBox(header_text)
                checkbox.setChecked(col in visible_columns)
                self.checkboxes[col] = checkbox
                content_layout.addWidget(checkbox)
        else:
            # Табличные виджеты: привязываемся к текущей видимости столбцов
            column_count = self.widget.columnCount()
            for col in range(column_count):
                if col < len(headers):
                    header_text = headers[col]
                else:
                    header_text = f"Столбец {col}"
                
                checkbox = QCheckBox(header_text)
                checkbox.setChecked(not self.widget.isColumnHidden(col))
                
                self.checkboxes[col] = checkbox
                content_layout.addWidget(checkbox)
        
        # Добавляем растягивающий элемент в конец
        content_layout.addStretch()
        
        scroll_area.setWidget(content_widget)
        layout.addWidget(scroll_area)
        
        # Кнопки управления
        buttons_frame = QFrame()
        buttons_layout = QHBoxLayout(buttons_frame)
        buttons_layout.setContentsMargins(0, 0, 0, 0)
        
        # Кнопки "Выбрать все" и "Снять все"
        select_all_btn = QPushButton("Выбрать все")
        select_all_btn.clicked.connect(self.select_all)
        buttons_layout.addWidget(select_all_btn)
        
        deselect_all_btn = QPushButton("Снять все")
        deselect_all_btn.clicked.connect(self.deselect_all)
        buttons_layout.addWidget(deselect_all_btn)
        
        buttons_layout.addStretch()
        
        # Кнопки "Применить" и "Отмена"
        apply_btn = QPushButton("✓ Применить")
        apply_btn.clicked.connect(self.apply_changes)
        apply_btn.setDefault(True)
        buttons_layout.addWidget(apply_btn)
        
        cancel_btn = QPushButton("✕ Отмена")
        cancel_btn.clicked.connect(self.reject)
        buttons_layout.addWidget(cancel_btn)
        
        layout.addWidget(buttons_frame)
    
    def load_current_state(self):
        """Загружает текущее состояние видимости столбцов"""
        if self.is_tree:
            # Для дерева состояние уже инициализировано из tree_visible_columns
            return
        for col, checkbox in self.checkboxes.items():
            checkbox.setChecked(not self.widget.isColumnHidden(col))
    
    def select_all(self):
        """Выбрать все столбцы"""
        for checkbox in self.checkboxes.values():
            checkbox.setChecked(True)
    
    def deselect_all(self):
        """Снять выбор со всех столбцов (кроме первого)"""
        for col, checkbox in self.checkboxes.items():
            # Первый столбец (с названием) всегда оставляем видимым
            if col == 0:
                checkbox.setChecked(True)
            else:
                checkbox.setChecked(False)
    
    def apply_changes(self):
        """Применить изменения видимости столбцов"""
        try:
            if self.is_tree:
                # Для дерева сохраняем конфиг и пересобираем дерево через TreeConfig/TreeBuilder
                main_window = self.widget.window()
                tree_headers = getattr(main_window, 'tree_headers', [])
                current_section = getattr(main_window, 'current_section', 'Доходы')

                # Флаги видимости по исходным индексам
                visibility_flags = {}
                visible_indices = []
                for col, checkbox in self.checkboxes.items():
                    is_visible = checkbox.isChecked()
                    visibility_flags[col] = is_visible
                    if is_visible:
                        visible_indices.append(col)

                # Обновляем список видимых колонок в главном окне
                main_window.tree_visible_columns = visible_indices

                # Сохраняем настройки в конфигурацию
                if tree_headers:
                    column_visibility = {}
                    for idx, name in enumerate(tree_headers):
                        is_visible = visibility_flags.get(idx, True)
                        column_visibility[name] = is_visible
                    config_key = f"tree_columns:{current_section}"
                    main_window.controller.db_manager.save_config(config_key, column_visibility)

                # Перестраиваем дерево с учётом новых настроек
                if main_window.controller.current_project:
                    main_window.tree_builder.load_project_data_to_tree(main_window.controller.current_project)
            else:
                # Для таблиц просто скрываем/показываем столбцы
                for col, checkbox in self.checkboxes.items():
                    is_visible = checkbox.isChecked()
                    self.widget.setColumnHidden(col, not is_visible)
            
            logger.info("Изменения видимости столбцов применены")
            self.accept()
        except Exception as e:
            logger.error(f"Ошибка при применении изменений видимости столбцов: {e}", exc_info=True)
            self.reject()

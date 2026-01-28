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
        
        # Получаем заголовки столбцов
        column_count = self.widget.columnCount()
        headers = []
        
        if self.is_tree:
            # Для QTreeWidget
            main_window = self.widget.window()
            tree_headers = getattr(main_window, 'tree_headers', [])
            
            if tree_headers:
                headers = tree_headers
            elif column_count > 0:
                header_item = self.widget.headerItem()
                if header_item:
                    for col in range(column_count):
                        header_text = header_item.text(col) if header_item.text(col) else f"Столбец {col}"
                        headers.append(header_text)
                else:
                    for col in range(column_count):
                        headers.append(f"Столбец {col}")
        elif self.is_table:
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
        for col in range(column_count):
            # Получаем название столбца
            if col < len(headers):
                header_text = headers[col]
            else:
                header_text = f"Столбец {col}"
            
            checkbox = QCheckBox(header_text)
            checkbox.setChecked(not self.widget.isColumnHidden(col))
            
            # Сохраняем связь между индексом столбца и чекбоксом
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
                # Для дерева нужно сохранять и восстанавливать ширины столбцов
                from PyQt5.QtWidgets import QHeaderView
                main_window = self.widget.window()
                header = self.widget.header()
                saved_widths = {}
                saved_resize_modes = {}
                
                def get_default_column_width_and_mode(col_idx, tree_widget, header):
                    """Получить дефолтную ширину и режим изменения размера для столбца"""
                    if col_idx == 0:
                        indentation = tree_widget.indentation()
                        indent_reserve = indentation * 6 + 50
                        return (400 + indent_reserve, QHeaderView.Interactive)
                    elif col_idx == 1:
                        return (80, QHeaderView.Fixed)
                    elif col_idx == 2:
                        return (200, QHeaderView.Interactive)
                    elif col_idx == 3:
                        return (50, QHeaderView.Fixed)
                    else:
                        return (150, QHeaderView.Fixed)
                
                for col, checkbox in self.checkboxes.items():
                    if col not in saved_widths:
                        was_hidden = self.widget.isColumnHidden(col)
                        section_size = header.sectionSize(col)
                        # Если столбец скрыт или его ширина равна 0, используем дефолтные значения
                        if was_hidden or section_size == 0:
                            default_width, default_mode = get_default_column_width_and_mode(col, self.widget, header)
                            saved_widths[col] = default_width
                            saved_resize_modes[col] = default_mode
                        else:
                            saved_widths[col] = section_size
                            saved_resize_modes[col] = header.sectionResizeMode(col)
                
                # Применяем видимость визуально
                for col, checkbox in self.checkboxes.items():
                    is_visible = checkbox.isChecked()

                    if not is_visible:
                        header.setSectionResizeMode(col, QHeaderView.Fixed)
                        header.resizeSection(col, 0)
                        self.widget.setColumnWidth(col, 0)
                        self.widget.setColumnHidden(col, True)
                    else:
                        self.widget.setColumnHidden(col, False)
                        # Восстанавливаем режим изменения размера и ширину столбца
                        if col in saved_resize_modes:
                            header.setSectionResizeMode(col, saved_resize_modes[col])
                        # Всегда восстанавливаем ширину, даже если она была 0 (используем дефолтную)
                        if col in saved_widths:
                            target_width = saved_widths[col]
                            if target_width > 0:
                                header.resizeSection(col, target_width)
                                self.widget.setColumnWidth(col, target_width)
                            else:
                                # Если ширина все еще 0, используем дефолтную
                                default_width, default_mode = get_default_column_width_and_mode(col, self.widget, header)
                                header.setSectionResizeMode(col, default_mode)
                                header.resizeSection(col, default_width)
                                self.widget.setColumnWidth(col, default_width)
                
                header.updateGeometries()
                self.widget.updateGeometry()
                self.widget.viewport().update()

                # Сохраняем настройки видимости столбцов дерева в конфигурацию
                try:
                    current_section = getattr(main_window, 'current_section', 'Доходы')
                    tree_headers = getattr(main_window, 'tree_headers', [])
                    column_visibility = {}
                    for col, checkbox in self.checkboxes.items():
                        # Определяем имя столбца
                        if tree_headers and col < len(tree_headers):
                            column_name = tree_headers[col]
                        else:
                            header_item = self.widget.headerItem()
                            if header_item and col < self.widget.columnCount():
                                column_name = header_item.text(col) or f"Колонка {col}"
                            else:
                                column_name = f"Колонка {col}"
                        column_visibility[column_name] = checkbox.isChecked()
                    if column_visibility:
                        config_key = f"tree_columns:{current_section}"
                        main_window.controller.db_manager.save_config(config_key, column_visibility)
                except Exception as e:
                    logger.error(f"Ошибка сохранения настроек видимости столбцов дерева: {e}", exc_info=True)
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

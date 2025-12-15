from PyQt5.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                             QPushButton, QLabel, QSplitter, 
                             QTabWidget, QTableWidget, QTableWidgetItem, 
                             QHeaderView, QMessageBox, QFileDialog, QProgressBar,
                             QToolBar, QStatusBar, QAction, QTextEdit,
                             QComboBox, QTreeWidget, QTreeWidgetItem, QMenu, 
                             QInputDialog, QDialog, QDialogButtonBox, QFormLayout,
                             QLineEdit, QCheckBox, QApplication, QStyle, QToolButton,
                             QStyledItemDelegate, QSpinBox, QWidgetAction)
from PyQt5.QtCore import Qt, QTimer, QSize, QRect
from PyQt5.QtCore import Qt, QTimer, QSize, QRect
from PyQt5.QtGui import (QFont, QColor, QBrush, QTextDocument, QTextOption, 
                        QTextCharFormat, QTextCursor, QPainter)
from PyQt5.QtWidgets import QStyleOptionHeader
import os
import subprocess
import platform
import re
from pathlib import Path
import pandas as pd

from controllers.main_controller import MainController
from logger import logger
from models.form_0503317 import Form0503317Constants
from views.project_dialog import ProjectDialog
from views.reference_dialog import ReferenceDialog
from views.excel_viewer import ExcelViewer
from views.reference_viewer import ReferenceViewer
from views.dictionaries_dialog import DictionariesDialog
from views.form_load_dialog import FormLoadDialog
from views.document_dialog import DocumentDialog


class WrapHeaderView(QHeaderView):
    """Кастомный заголовок с поддержкой переноса текста"""
    
    def __init__(self, orientation=Qt.Horizontal, parent=None):
        super().__init__(orientation, parent)
        self.setTextElideMode(Qt.ElideNone)
        self._header_texts = {}  # Кэш текстов заголовков
    
    def setHeaderTexts(self, texts):
        """Устанавливает тексты заголовков для кэширования"""
        self._header_texts = texts
    
    def paintSection(self, painter, rect, logicalIndex):
        """Переопределяем отрисовку секции заголовка с поддержкой переноса текста"""
        # Получаем текст заголовка из кэша или модели
        text = None
        if logicalIndex in self._header_texts:
            text = self._header_texts[logicalIndex]
        elif self.model():
            text = self.model().headerData(logicalIndex, self.orientation(), Qt.DisplayRole)
        
        if not text:
            # Используем стандартную отрисовку, если текста нет
            super().paintSection(painter, rect, logicalIndex)
            return
        
        text = str(text)
        
        # Рисуем фон заголовка вручную, используя стиль
        # Получаем опции стиля для отрисовки фона
        option = QStyleOptionHeader()
        option.initFrom(self)
        option.rect = rect
        option.section = logicalIndex
        
        # Определяем позицию секции (первая, средняя, последняя)
        if logicalIndex == 0:
            if self.count() > 1:
                option.position = QStyleOptionHeader.Beginning
            else:
                option.position = QStyleOptionHeader.OnlyOneSection
        elif logicalIndex == self.count() - 1:
            option.position = QStyleOptionHeader.End
        else:
            option.position = QStyleOptionHeader.Middle
        
        # Рисуем фон заголовка вручную через палитру
        # Получаем цвет фона из палитры
        palette = self.palette()
        bg_color = palette.color(palette.Button)
        painter.fillRect(rect, bg_color)
        
        # Рисуем границы заголовка
        border_color = palette.color(palette.Mid)
        painter.setPen(border_color)
        painter.drawRect(rect.adjusted(0, 0, -1, -1))
        
        # Создаем документ для переноса текста
        doc = QTextDocument()
        doc.setDefaultFont(self.font())
        doc.setPlainText(text)
        
        # Настраиваем перенос текста
        text_option = QTextOption()
        text_option.setWrapMode(QTextOption.WrapAtWordBoundaryOrAnywhere)
        text_option.setAlignment(Qt.AlignCenter)
        doc.setDefaultTextOption(text_option)
        
        # Устанавливаем ширину документа равной ширине секции (с небольшими отступами)
        padding = 4
        doc.setTextWidth(rect.width() - 2 * padding)
        
        # Рисуем текст с переносом
        painter.save()
        painter.translate(rect.left() + padding, rect.top() + (rect.height() - doc.size().height()) / 2)
        painter.setClipRect(QRect(0, 0, rect.width() - 2 * padding, rect.height()))
        doc.drawContents(painter)
        painter.restore()
    
    def sizeHint(self):
        """Возвращаем размер заголовка с учетом переноса текста"""
        size = super().sizeHint()
        
        # Вычисляем максимальную высоту с учетом переноса текста
        max_height = 0
        font_metrics = self.fontMetrics()
        
        for idx in range(self.count()):
            if self.isSectionHidden(idx):
                continue
            
            # Получаем текст из кэша или модели
            text = None
            if idx in self._header_texts:
                text = self._header_texts[idx]
            elif self.model():
                text = self.model().headerData(idx, self.orientation(), Qt.DisplayRole)
            
            if not text:
                continue
            
            text = str(text)
            width = self.sectionSize(idx)
            
            # Создаем документ для расчета высоты
            doc = QTextDocument()
            doc.setDefaultFont(self.font())
            doc.setPlainText(text)
            
            text_option = QTextOption()
            text_option.setWrapMode(QTextOption.WrapAtWordBoundaryOrAnywhere)
            doc.setDefaultTextOption(text_option)
            
            padding = 4
            doc.setTextWidth(width - 2 * padding)
            
            doc_height = doc.size().height()
            max_height = max(max_height, doc_height)
        
        if max_height > 0:
            size.setHeight(int(max_height) + 8)  # Добавляем отступы
        else:
            size.setHeight(font_metrics.lineSpacing() + 6)
        
        return size


class WordWrapItemDelegate(QStyledItemDelegate):
    """Делегат для переноса текста в ячейках дерева"""
    
    def paint(self, painter, option, index):
        if not index.isValid():
            return
        
        # Настраиваем опции отрисовки (нужно сделать до проверки текста)
        option = option.__class__(option)
        self.initStyleOption(option, index)
        
        # Получаем текст из модели
        text = index.data(Qt.DisplayRole) or ""
        
        # Рисуем фон даже если текст пустой (для окраски по уровням)
        # Получаем цвет фона
        background_brush = index.data(Qt.BackgroundRole)
        if background_brush:
            painter.fillRect(option.rect, background_brush)
        elif option.state & QStyle.State_Selected:
            # Для выделенных строк используем более светлый фон
            highlight_color = option.palette.highlight().color()
            # Делаем фон более прозрачным/светлым
            light_highlight = QColor(highlight_color)
            light_highlight.setAlpha(50)  # Полупрозрачный фон
            painter.fillRect(option.rect, light_highlight)
        else:
            painter.fillRect(option.rect, option.palette.base())
        
        # Если текст пустой, только рисуем фон и выходим
        if not text:
            return
        
        # Получаем номер столбца
        column = index.column()
        
        # Создаем документ для переноса текста
        doc = QTextDocument()
        doc.setDefaultFont(option.font)
        doc.setPlainText(str(text))
        
        # Настраиваем перенос текста
        # Для столбца "Код классификации" (индекс 2) отключаем перенос текста
        text_option = QTextOption()
        if column == 2:  # Код классификации - без переноса
            text_option.setWrapMode(QTextOption.NoWrap)
        else:
            text_option.setWrapMode(QTextOption.WrapAtWordBoundaryOrAnywhere)
        doc.setDefaultTextOption(text_option)
        
        # Для столбца "Код классификации" рисуем текст без переноса и без обрезания
        if column == 2:
            # Фон уже нарисован выше, только устанавливаем цвет текста
            text_color = index.data(Qt.ForegroundRole)
            # Сохраняем исходный шрифт
            original_font = painter.font()
            if text_color:
                painter.setPen(text_color)
            else:
                # Для выделенных строк используем обычный цвет текста, но жирный шрифт
                if option.state & QStyle.State_Selected:
                    painter.setPen(option.palette.text().color())
                    # Устанавливаем жирный шрифт
                    font = painter.font()
                    font.setBold(True)
                    painter.setFont(font)
                else:
                    painter.setPen(option.palette.text().color())
                    # Убеждаемся, что шрифт не жирный для невыделенных строк
                    font = painter.font()
                    font.setBold(False)
                    painter.setFont(font)
            
            # Рисуем текст без переноса
            text_rect = option.rect.adjusted(2, 0, -2, 0)  # Небольшой отступ слева и справа
            painter.drawText(text_rect, Qt.AlignLeft | Qt.AlignVCenter, str(text))
            # Восстанавливаем исходный шрифт
            painter.setFont(original_font)
            return
        
        # Для остальных столбцов используем документ с переносом
        # Устанавливаем ширину документа равной ширине ячейки
        # Для столбца "Наименование" используем ширину с учетом отступов дерева
        width = option.rect.width()
        right_padding = 0  # Отступ справа для столбца "Наименование"
        
        if column == 0:
            # Получаем отступы дерева из виджета
            widget = option.widget
            if widget and hasattr(widget, 'indentation'):
                indentation = widget.indentation()
                indent_reserve = indentation * 6 + 50  # Запас на отступы
                width = 400 + indent_reserve
                
                # Вычисляем уровень элемента (глубину вложенности относительно нулевого уровня)
                item_level = 0
                try:
                    model = index.model()
                    if model:
                        # Пытаемся получить уровень из данных элемента (столбец 3 - "Уровень")
                        level_index = model.index(index.row(), 3, index.parent())
                        if level_index.isValid():
                            level_text = model.data(level_index, Qt.DisplayRole)
                            if level_text:
                                try:
                                    item_level = int(str(level_text))
                                except (ValueError, TypeError):
                                    item_level = 0
                        
                        # Если не удалось получить из данных, вычисляем по глубине вложенности
                        if item_level == 0:
                            parent = index.parent()
                            while parent.isValid():
                                item_level += 1
                                parent = parent.parent()
                except Exception:
                    item_level = 0
                
                # Вычисляем внутренний отступ справа с учетом всех уровней от 0 до текущего
                # Сумма отступов всех уровней: indentation * (0 + 1 + 2 + ... + item_level)
                # Формула суммы арифметической прогрессии: n * (n + 1) / 2
                # Для уровней от 0 до item_level: indentation * item_level * (item_level + 1) / 2
                if item_level > 0:
                    right_padding = indentation * item_level * (item_level + 1) // 2
                else:
                    right_padding = 0
            else:
                width = 400  # Значение по умолчанию, если не удалось получить отступы
        
        doc.setTextWidth(width - right_padding)
        
        # Фон уже нарисован выше, устанавливаем цвет текста (для ошибок) через QTextCharFormat
        text_color = index.data(Qt.ForegroundRole)
        if text_color:
            # text_color может быть QBrush или QColor
            if isinstance(text_color, QBrush):
                color = text_color.color()
            else:
                color = text_color
            # Устанавливаем цвет текста через формат
            char_format = QTextCharFormat()
            char_format.setForeground(color)
            # Применяем формат ко всему документу через курсор
            cursor = QTextCursor(doc)
            cursor.select(QTextCursor.Document)
            cursor.setCharFormat(char_format)
        elif option.state & QStyle.State_Selected:
            # Для выделенных строк используем обычный цвет текста, но жирный шрифт
            color = option.palette.text().color()
            char_format = QTextCharFormat()
            char_format.setForeground(color)
            char_format.setFontWeight(QFont.Bold)  # Делаем текст жирным
            cursor = QTextCursor(doc)
            cursor.select(QTextCursor.Document)
            cursor.setCharFormat(char_format)
        else:
            color = option.palette.text().color()
            char_format = QTextCharFormat()
            char_format.setForeground(color)
            cursor = QTextCursor(doc)
            cursor.select(QTextCursor.Document)
            cursor.setCharFormat(char_format)
        
        # Рисуем текст с учетом внутреннего отступа справа
        text_rect = option.rect.adjusted(0, 0, -right_padding, 0)
        painter.save()
        painter.translate(text_rect.topLeft())
        doc.drawContents(painter)
        painter.restore()
    
    def sizeHint(self, option, index):
        if not index.isValid():
            return QSize()
        
        text = index.data(Qt.DisplayRole) or ""
        if not text:
            return QSize(0, option.fontMetrics.height())
        
        # Получаем ширину столбца из виджета
        widget = option.widget
        column_width = 200  # Значение по умолчанию
        column = index.column()
        
        right_padding = 0  # Отступ справа для столбца "Наименование"
        
        if widget and hasattr(widget, 'header'):
            header = widget.header()
            if column >= 0:
                column_width = max(header.sectionSize(column), 50)
                # Для столбца "Наименование" (индекс 0) используем ширину с учетом отступов
                if column == 0:
                    if hasattr(widget, 'indentation'):
                        indentation = widget.indentation()
                        indent_reserve = indentation * 6 + 50
                        column_width = 400 + indent_reserve
                        
                        # Вычисляем уровень элемента для внутреннего отступа справа
                        item_level = 0
                        try:
                            model = index.model()
                            if model:
                                # Пытаемся получить уровень из данных элемента (столбец 3 - "Уровень")
                                level_index = model.index(index.row(), 3, index.parent())
                                if level_index.isValid():
                                    level_text = model.data(level_index, Qt.DisplayRole)
                                    if level_text:
                                        try:
                                            item_level = int(str(level_text))
                                        except (ValueError, TypeError):
                                            item_level = 0
                                
                                # Если не удалось получить из данных, вычисляем по глубине вложенности
                                if item_level == 0:
                                    parent = index.parent()
                                    while parent.isValid():
                                        item_level += 1
                                        parent = parent.parent()
                        except Exception:
                            item_level = 0
                        
                        # Вычисляем внутренний отступ справа с учетом всех уровней от 0 до текущего
                        # Сумма отступов всех уровней: indentation * (0 + 1 + 2 + ... + item_level)
                        # Формула суммы арифметической прогрессии: n * (n + 1) / 2
                        if item_level > 0:
                            right_padding = indentation * item_level * (item_level + 1) // 2
                        else:
                            right_padding = 0
                    else:
                        column_width = 400
        
        # Если ширина из option доступна, используем её
        if option.rect.width() > 0:
            column_width = option.rect.width()
            # Для столбца "Наименование" учитываем отступы
            if column == 0:
                if widget and hasattr(widget, 'indentation'):
                    indentation = widget.indentation()
                    indent_reserve = indentation * 6 + 50
                    column_width = 400 + indent_reserve
                    
                    # Вычисляем уровень элемента для внутреннего отступа справа
                    item_level = 0
                    try:
                        model = index.model()
                        if model:
                            level_index = model.index(index.row(), 3, index.parent())
                            if level_index.isValid():
                                level_text = model.data(level_index, Qt.DisplayRole)
                                if level_text:
                                    try:
                                        item_level = int(str(level_text))
                                    except (ValueError, TypeError):
                                        item_level = 0
                            
                            if item_level == 0:
                                parent = index.parent()
                                while parent.isValid():
                                    item_level += 1
                                    parent = parent.parent()
                    except Exception:
                        item_level = 0
                    
                    # Вычисляем внутренний отступ справа с учетом всех уровней от 0 до текущего
                    # Сумма отступов всех уровней: indentation * (0 + 1 + 2 + ... + item_level)
                    if item_level > 0:
                        right_padding = indentation * item_level * (item_level + 1) // 2
                    else:
                        right_padding = 0
                else:
                    column_width = 400
        
        # Для столбца "Код классификации" (индекс 2) используем ширину текста без переноса
        if column == 2:
            # Возвращаем размер текста без переноса
            text_width = option.fontMetrics.horizontalAdvance(str(text))
            return QSize(text_width, option.fontMetrics.height())
        
        # Для остальных столбцов создаем документ для расчета размера с переносом
        doc = QTextDocument()
        doc.setDefaultFont(option.font)
        doc.setPlainText(str(text))
        
        # Настраиваем перенос текста
        text_option = QTextOption()
        text_option.setWrapMode(QTextOption.WrapAtWordBoundaryOrAnywhere)
        doc.setDefaultTextOption(text_option)
        
        # Устанавливаем ширину документа равной ширине столбца с учетом внутреннего отступа справа
        # Для столбца "Наименование" вычитаем отступ справа
        available_width = column_width - right_padding if column == 0 else column_width
        doc.setTextWidth(available_width)
        
        # Возвращаем размер с учетом переноса
        return QSize(int(doc.idealWidth()), int(doc.size().height()))


class DetachedTabWindow(QMainWindow):
    """Отдельное окно для открепленной вкладки"""
    
    def __init__(self, tab_widget, tab_name, parent=None):
        super().__init__(parent)
        self.tab_widget = tab_widget
        self.tab_name = tab_name
        self.main_window = parent
        
        self.setWindowTitle(tab_name)
        self.setGeometry(100, 100, 1200, 800)
        
        # Убеждаемся, что виджет видим и имеет правильный родитель
        if tab_widget.parent():
            # Удаляем виджет из старого layout, если он там был
            old_parent = tab_widget.parent()
            if isinstance(old_parent, QWidget):
                old_layout = old_parent.layout()
                if old_layout:
                    old_layout.removeWidget(tab_widget)
        
        # Устанавливаем виджет как центральный виджет напрямую
        # Это работает, если tab_widget уже содержит все необходимое
        self.setCentralWidget(tab_widget)
        
        # Убеждаемся, что виджет видим
        tab_widget.setVisible(True)
        tab_widget.show()

    def closeEvent(self, event):
        """Обработка закрытия окна - переопределяем метод closeEvent"""
        logger.debug(f"closeEvent вызван для окна '{self.tab_name}'")
        
        # Проверяем, не происходит ли уже возврат вкладки (чтобы избежать повторного вызова)
        if self.property("attaching"):
            logger.debug(f"Флаг 'attaching' установлен, пропускаем возврат вкладки")
            event.accept()
            return
        
        # При закрытии окна возвращаем вкладку в главное окно
        if self.main_window:
            try:
                # Используем tab_widget из centralWidget, если он доступен
                tab_widget = self.centralWidget() or self.tab_widget
                if tab_widget:
                    logger.debug(f"Вызов attach_tab из closeEvent для '{self.tab_name}'")
                    self.main_window.attach_tab(self.tab_name, tab_widget)
                else:
                    logger.warning(f"Не удалось получить виджет для возврата вкладки '{self.tab_name}'")
            except Exception as e:
                logger.error(f"Ошибка при возврате вкладки: {e}", exc_info=True)
        else:
            logger.warning(f"main_window не установлен для окна '{self.tab_name}'")
        
        event.accept()
    
    def get_tab_widget(self):
        """Получить виджет вкладки"""
        return self.tab_widget


class MainWindow(QMainWindow):
    """Главное окно приложения"""
    
    def __init__(self):
        super().__init__()
        self.controller = MainController()
        self.current_section = "Доходы"
        self.current_data_type = "Оба"
        self.main_splitter = None
        self.projects_panel_index = 0
        self.projects_inner_panel = None
        self.projects_toggle_button = None
        self.projects_panel_last_size = 260
        self.reference_window = None
        self.tree_headers = []
        self.tree_header_tooltips = []
        self.tree_column_mapping = {}
        self._updating_header_height = False  # Флаг для предотвращения бесконечного цикла
        self.last_exported_file = None  # Путь к последнему экспортированного файла
        self.errors_tab_fullscreen = False  # Флаг полноэкранного режима вкладки ошибок
        # Окна для открепленных вкладок
        self.detached_windows = {}  # {tab_name: QMainWindow}
        self.tabs_panel = None  # Будет установлен в create_tabs_panel
        # Настройки шрифтов
        self.font_size = 10  # Размер шрифта для данных
        self.header_font_size = 10  # Размер шрифта для заголовков
        # Отслеживание выделения
        self.selection_start_column = None  # Столбец, с которого началось выделение
        self.init_ui()
        self.connect_signals()
        self.controller.load_initial_data()
    
    def init_ui(self):
        """Инициализация интерфейса"""
        self.setWindowTitle("Система обработки бюджетных форм")
        self.setGeometry(100, 100, 1600, 900)
        
        # Создаем меню-бар
        self.create_menu_bar()
        
        # Создаем центральный виджет
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # Основной layout
        main_layout = QHBoxLayout(central_widget)
        
        # Создаем сплиттер
        splitter = QSplitter(Qt.Horizontal)
        self.main_splitter = splitter
        
        # Левая панель - список проектов
        self.projects_panel = self.create_projects_panel()
        splitter.addWidget(self.projects_panel)
        self.projects_panel_index = splitter.indexOf(self.projects_panel)
        
        # Центральная панель - вкладки с данными
        self.tabs_panel = self.create_tabs_panel()
        splitter.addWidget(self.tabs_panel)
        
        # Устанавливаем пропорции
        splitter.setSizes([300, 1300])
        
        main_layout.addWidget(splitter)
        
        # Создаем тулбар
        # self.create_toolbar()
        
        # Создаем статусбар
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("Готов к работе")
        
        # Прогресс-бар
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.status_bar.addPermanentWidget(self.progress_bar)
        
        # Создаем док-виджеты
        self.create_dock_widgets()
    
    def create_menu_bar(self):
        """Создание меню-бара"""
        menubar = self.menuBar()
        
        # ========== Меню "Файл" ==========
        file_menu = menubar.addMenu("&Файл")
        
        new_project_action = QAction("&Новый проект...", self)
        new_project_action.setIcon(self.style().standardIcon(QStyle.SP_FileIcon))
        new_project_action.setShortcut("Ctrl+N")
        new_project_action.setStatusTip("Создать новый проект")
        new_project_action.triggered.connect(self.show_new_project_dialog)
        file_menu.addAction(new_project_action)
        
        load_form_action = QAction("&Загрузить форму...", self)
        load_form_action.setIcon(self.style().standardIcon(QStyle.SP_DirOpenIcon))
        load_form_action.setShortcut("Ctrl+O")
        load_form_action.setStatusTip("Загрузить файл формы")
        load_form_action.triggered.connect(self.load_form_file)
        file_menu.addAction(load_form_action)
        
        file_menu.addSeparator()
        
        export_action = QAction("&Экспорт проверки...", self)
        export_action.setIcon(self.style().standardIcon(QStyle.SP_DialogSaveButton))
        export_action.setShortcut("Ctrl+E")
        export_action.setStatusTip("Экспортировать форму с проверкой")
        export_action.triggered.connect(self.export_validation)
        file_menu.addAction(export_action)
        
        file_menu.addSeparator()
        
        # Открытие файлов
        open_file_action = QAction("&Открыть файл...", self)
        open_file_action.setIcon(self.style().standardIcon(QStyle.SP_DirOpenIcon))
        open_file_action.setShortcut("Ctrl+Shift+O")
        open_file_action.setStatusTip("Открыть файл (doc, docx, xls, xlsx)")
        open_file_action.triggered.connect(self.open_file_dialog)
        file_menu.addAction(open_file_action)
        
        # Открыть последний экспортированный файл
        self.open_last_file_action = QAction("Открыть последний экспортированный файл", self)
        self.open_last_file_action.setIcon(self.style().standardIcon(QStyle.SP_FileDialogStart))
        self.open_last_file_action.setStatusTip("Открыть последний экспортированный файл")
        self.open_last_file_action.setEnabled(False)
        self.open_last_file_action.triggered.connect(self.open_last_exported_file)
        file_menu.addAction(self.open_last_file_action)
        
        file_menu.addSeparator()
        
        exit_action = QAction("&Выход", self)
        exit_action.setIcon(self.style().standardIcon(QStyle.SP_DialogCloseButton))
        exit_action.setShortcut("Ctrl+Q")
        exit_action.setStatusTip("Выход из приложения")
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)
        
        # ========== Меню "Проект" ==========
        project_menu = menubar.addMenu("&Проект")
        
        edit_project_action = QAction("&Редактировать проект...", self)
        edit_project_action.setIcon(self.style().standardIcon(QStyle.SP_FileDialogDetailedView))
        edit_project_action.setShortcut("Ctrl+P")
        edit_project_action.setStatusTip("Редактировать текущий проект")
        edit_project_action.triggered.connect(self.edit_current_project)
        project_menu.addAction(edit_project_action)
        
        delete_project_action = QAction("&Удалить проект", self)
        delete_project_action.setIcon(self.style().standardIcon(QStyle.SP_TrashIcon))
        delete_project_action.setShortcut("Ctrl+Delete")
        delete_project_action.setStatusTip("Удалить текущий проект")
        delete_project_action.triggered.connect(self.delete_current_project)
        project_menu.addAction(delete_project_action)
        
        project_menu.addSeparator()
        
        refresh_projects_action = QAction("&Обновить список", self)
        refresh_projects_action.setIcon(self.style().standardIcon(QStyle.SP_BrowserReload))
        refresh_projects_action.setShortcut("F5")
        refresh_projects_action.setStatusTip("Обновить список проектов")
        refresh_projects_action.triggered.connect(lambda: self.controller.projects_updated.emit(self.controller.project_controller.load_projects()))
        project_menu.addAction(refresh_projects_action)
        
        # ========== Меню "Данные" ==========
        data_menu = menubar.addMenu("&Данные")
        
        # Действия для работы с данными (не специфичные для ревизии)
        # Действие "Скрыть нулевые столбцы" перенесено в интерфейс формы (чекбокс)
        # Оставляем только для совместимости с горячей клавишей
        hide_zeros_action = QAction("&Скрыть нулевые столбцы", self)
        hide_zeros_action.setIcon(self.style().standardIcon(QStyle.SP_TrashIcon))
        hide_zeros_action.setShortcut("Ctrl+H")
        hide_zeros_action.setStatusTip("Скрыть столбцы с нулевыми значениями (используйте чекбокс в интерфейсе формы)")
        hide_zeros_action.triggered.connect(self.hide_zero_columns_global)
        data_menu.addAction(hide_zeros_action)
        
        # ========== Меню "Справочники" ==========
        reference_menu = menubar.addMenu("&Справочники")
        
        load_income_ref_action = QAction("&Загрузить справочник доходов...", self)
        load_income_ref_action.setIcon(self.style().standardIcon(QStyle.SP_DialogOpenButton))
        load_income_ref_action.setStatusTip("Загрузить справочник доходов")
        load_income_ref_action.triggered.connect(lambda: self.show_reference_dialog("доходы"))
        reference_menu.addAction(load_income_ref_action)
        
        load_sources_ref_action = QAction("&Загрузить справочник источников...", self)
        load_sources_ref_action.setIcon(self.style().standardIcon(QStyle.SP_DialogOpenButton))
        load_sources_ref_action.setStatusTip("Загрузить справочник источников финансирования")
        load_sources_ref_action.triggered.connect(lambda: self.show_reference_dialog("источники"))
        reference_menu.addAction(load_sources_ref_action)
        
        reference_menu.addSeparator()
        
        show_references_action = QAction("&Просмотр справочников", self)
        show_references_action.setIcon(self.style().standardIcon(QStyle.SP_FileDialogInfoView))
        show_references_action.setShortcut("Ctrl+R")
        show_references_action.setStatusTip("Открыть окно просмотра справочников")
        show_references_action.triggered.connect(self.show_reference_viewer)
        reference_menu.addAction(show_references_action)
        
        reference_menu.addSeparator()
        
        config_dicts_action = QAction("&Справочники конфигурации...", self)
        config_dicts_action.setIcon(self.style().standardIcon(QStyle.SP_FileDialogListView))
        config_dicts_action.setShortcut("Ctrl+D")
        config_dicts_action.setStatusTip("Редактировать справочники конфигурации (годы, МО, типы форм, периоды)")
        config_dicts_action.triggered.connect(self.show_config_dictionaries)
        reference_menu.addAction(config_dicts_action)
        
        # ========== Меню "Вид" ==========
        view_menu = menubar.addMenu("&Вид")
        
        toggle_projects_panel_action = QAction("&Панель проектов", self)
        toggle_projects_panel_action.setCheckable(True)
        toggle_projects_panel_action.setChecked(True)
        toggle_projects_panel_action.setShortcut("Ctrl+1")
        toggle_projects_panel_action.setStatusTip("Показать/скрыть панель проектов")
        toggle_projects_panel_action.triggered.connect(self.toggle_projects_panel)
        view_menu.addAction(toggle_projects_panel_action)
        
        view_menu.addSeparator()
        
        # Управление размером шрифта данных
        font_size_widget = QWidget()
        font_size_layout = QHBoxLayout(font_size_widget)
        font_size_layout.setContentsMargins(10, 5, 10, 5)
        font_size_label = QLabel("Размер шрифта данных:")
        font_size_layout.addWidget(font_size_label)
        self.font_size_spinbox = QSpinBox()
        self.font_size_spinbox.setMinimum(6)
        self.font_size_spinbox.setMaximum(20)
        self.font_size_spinbox.setValue(self.font_size)
        self.font_size_spinbox.setSuffix(" пт")
        self.font_size_spinbox.valueChanged.connect(self.on_font_size_changed)
        font_size_layout.addWidget(self.font_size_spinbox)
        font_size_action = QWidgetAction(self)
        font_size_action.setDefaultWidget(font_size_widget)
        view_menu.addAction(font_size_action)
        
        # Управление размером шрифта заголовков
        header_font_size_widget = QWidget()
        header_font_size_layout = QHBoxLayout(header_font_size_widget)
        header_font_size_layout.setContentsMargins(10, 5, 10, 5)
        header_font_size_label = QLabel("Размер шрифта заголовков:")
        header_font_size_layout.addWidget(header_font_size_label)
        self.header_font_size_spinbox = QSpinBox()
        self.header_font_size_spinbox.setMinimum(6)
        self.header_font_size_spinbox.setMaximum(20)
        self.header_font_size_spinbox.setValue(self.header_font_size)
        self.header_font_size_spinbox.setSuffix(" пт")
        self.header_font_size_spinbox.valueChanged.connect(self.on_header_font_size_changed)
        header_font_size_layout.addWidget(self.header_font_size_spinbox)
        header_font_size_action = QWidgetAction(self)
        header_font_size_action.setDefaultWidget(header_font_size_widget)
        view_menu.addAction(header_font_size_action)
        
        view_menu.addSeparator()
        
        fullscreen_action = QAction("&Полноэкранный режим", self)
        fullscreen_action.setShortcut("F11")
        fullscreen_action.setCheckable(True)
        fullscreen_action.setStatusTip("Переключить полноэкранный режим")
        fullscreen_action.triggered.connect(self.toggle_fullscreen)
        view_menu.addAction(fullscreen_action)
        
        # ========== Меню "Справка" ==========
        help_menu = menubar.addMenu("&Справка")
        
        about_action = QAction("&О программе", self)
        about_action.setIcon(self.style().standardIcon(QStyle.SP_MessageBoxInformation))
        about_action.setStatusTip("Информация о программе")
        about_action.triggered.connect(self.show_about)
        help_menu.addAction(about_action)
        
        help_menu.addSeparator()
        
        shortcuts_action = QAction("&Горячие клавиши", self)
        shortcuts_action.setIcon(self.style().standardIcon(QStyle.SP_FileDialogInfoView))
        shortcuts_action.setStatusTip("Список горячих клавиш")
        shortcuts_action.triggered.connect(self.show_shortcuts)
        help_menu.addAction(shortcuts_action)
    
    def create_toolbar(self):
        """Создание панели инструментов"""
        toolbar = QToolBar("Основные инструменты")
        self.addToolBar(toolbar)
        
        # Действия
        new_project_action = QAction("Новый проект", self)
        new_project_action.setIcon(self.style().standardIcon(QStyle.SP_FileIcon))
        new_project_action.triggered.connect(self.show_new_project_dialog)
        toolbar.addAction(new_project_action)
        
        load_form_action = QAction("Загрузить форму", self)
        load_form_action.setIcon(self.style().standardIcon(QStyle.SP_DirOpenIcon))
        load_form_action.triggered.connect(self.load_form_file)
        toolbar.addAction(load_form_action)
        
        toolbar.addSeparator()
        
        # Отдельные действия для справочников доходов и источников
        load_income_ref_action = QAction("Справочник доходов", self)
        load_income_ref_action.setIcon(self.style().standardIcon(QStyle.SP_DialogOpenButton))
        load_income_ref_action.triggered.connect(lambda: self.show_reference_dialog("доходы"))
        toolbar.addAction(load_income_ref_action)

        load_sources_ref_action = QAction("Справочник источников", self)
        load_sources_ref_action.setIcon(self.style().standardIcon(QStyle.SP_DialogOpenButton))
        load_sources_ref_action.triggered.connect(lambda: self.show_reference_dialog("источники"))
        toolbar.addAction(load_sources_ref_action)

        # Кнопка для сворачивания нулевых столбцов (таблица + дерево)
        # Действие "Скрыть нулевые столбцы" перенесено в интерфейс формы (чекбокс)
        # Удалено из тулбара, т.к. теперь доступно в интерфейсе формы
        
        show_references_action = QAction("Просмотр справочников", self)
        show_references_action.setIcon(self.style().standardIcon(QStyle.SP_FileDialogInfoView))
        show_references_action.triggered.connect(self.show_reference_viewer)
        toolbar.addAction(show_references_action)

        # Редактор конфигурационных справочников (годы, МО, типы форм, периоды)
        config_dicts_action = QAction("Справочники конфигурации", self)
        config_dicts_action.setIcon(self.style().standardIcon(QStyle.SP_FileDialogListView))
        config_dicts_action.triggered.connect(self.show_config_dictionaries)
        toolbar.addAction(config_dicts_action)

        # Кнопки управления панелью проектов размещены непосредственно на самой панели
    
    def create_dock_widgets(self):
        """Инициализация структур для просмотра справочников (отдельное окно)"""
        # Метод оставлен для совместимости, но структуры инициализируются в show_reference_viewer
        pass
    
    def create_projects_panel(self) -> QWidget:
        """Создание панели проектов"""
        # Основная панель с содержимым
        inner_panel = QWidget()
        layout = QVBoxLayout(inner_panel)
        layout.setContentsMargins(6, 6, 2, 6)
        
        # Заголовок
        title_label = QLabel("Проекты")
        title_label.setFont(QFont("Arial", 12, QFont.Bold))
        layout.addWidget(title_label)
        
        # Кнопки управления проектами
        buttons_layout = QHBoxLayout()
        
        new_project_btn = QPushButton("Новый")
        new_project_btn.clicked.connect(self.show_new_project_dialog)
        buttons_layout.addWidget(new_project_btn)
        
        refresh_btn = QPushButton("Обновить")
        refresh_btn.clicked.connect(self.refresh_projects)
        buttons_layout.addWidget(refresh_btn)
        
        layout.addLayout(buttons_layout)
        
        # Дерево проектов: Год -> Проект -> Форма -> Ревизия
        from PyQt5.QtWidgets import QTreeWidget
        self.projects_tree = QTreeWidget()
        self.projects_tree.setIndentation(10)
        self.projects_tree.setHeaderHidden(True)
        self.projects_tree.itemDoubleClicked.connect(self.on_project_tree_double_clicked)
        self.projects_tree.setContextMenuPolicy(Qt.CustomContextMenu)
        self.projects_tree.customContextMenuRequested.connect(self.show_project_context_menu)
        layout.addWidget(self.projects_tree)
        
        # Информация о проекте
        self.project_info_label = QLabel("Выберите проект")
        self.project_info_label.setWordWrap(True)
        layout.addWidget(self.project_info_label)
        
        # Контейнер, в котором слева основная панель, справа узкая кнопка-свертка
        container = QWidget()
        container_layout = QHBoxLayout(container)
        container_layout.setContentsMargins(0, 0, 0, 0)
        container_layout.setSpacing(0)

        container_layout.addWidget(inner_panel)

        # Узкая вертикальная кнопка на правом краю панели
        toggle_button = QPushButton("◀")
        toggle_button.setFixedWidth(14)
        toggle_button.setFlat(True)
        toggle_button.setFocusPolicy(Qt.NoFocus)
        toggle_button.setToolTip("Свернуть/развернуть панель проектов")
        toggle_button.clicked.connect(self.on_projects_side_button_clicked)
        container_layout.addWidget(toggle_button)

        self.projects_inner_panel = inner_panel
        self.projects_toggle_button = toggle_button

        return container
    
    def create_tabs_panel(self) -> QWidget:
        """Создание панели с вкладками"""
        tabs = QTabWidget()
        self.tabs_panel = tabs
        tabs.setTabsClosable(False)  # Отключаем стандартное закрытие вкладок
        
        # Добавляем контекстное меню для вкладок
        tabs.setContextMenuPolicy(Qt.CustomContextMenu)
        tabs.customContextMenuRequested.connect(self.show_tab_context_menu)
        
        # Вкладка с древовидными данными
        self.tree_tab = QWidget()
        
        tree_layout = QVBoxLayout(self.tree_tab)
        
        # Панель управления древом
        tree_control_layout = QHBoxLayout()
        # Кнопки управления деревом (максимально компактные)
        self.expand_all_btn = QToolButton()
        self.expand_all_btn.setIcon(self.style().standardIcon(QStyle.SP_ArrowDown))
        self.expand_all_btn.setToolTip("Развернуть все узлы дерева")
        self.expand_all_btn.setIconSize(QSize(14, 14))
        self.expand_all_btn.setToolButtonStyle(Qt.ToolButtonIconOnly)
        self.expand_all_btn.setAutoRaise(True)
        self.expand_all_btn.setFixedSize(22, 22)
        self.expand_all_btn.clicked.connect(self.expand_all_tree)
        tree_control_layout.addWidget(self.expand_all_btn)
        
        self.collapse_all_btn = QToolButton()
        self.collapse_all_btn.setIcon(self.style().standardIcon(QStyle.SP_ArrowUp))
        self.collapse_all_btn.setToolTip("Свернуть все узлы дерева")
        self.collapse_all_btn.setIconSize(QSize(14, 14))
        self.collapse_all_btn.setToolButtonStyle(Qt.ToolButtonIconOnly)
        self.collapse_all_btn.setAutoRaise(True)
        self.collapse_all_btn.setFixedSize(22, 22)
        self.collapse_all_btn.clicked.connect(self.collapse_all_tree)
        tree_control_layout.addWidget(self.collapse_all_btn)
        
        tree_control_layout.addStretch()
        
        # Выбор раздела
        tree_control_layout.addWidget(QLabel("Раздел:"))
        self.section_combo = QComboBox()
        self.section_combo.addItems(["Доходы", "Расходы", "Источники финансирования", "Консолидируемые расчеты"])
        self.section_combo.currentTextChanged.connect(self.on_section_changed)
        tree_control_layout.addWidget(self.section_combo)
        
        # Выбор типа данных
        tree_control_layout.addWidget(QLabel("Тип данных:"))
        self.data_type_combo = QComboBox()
        self.data_type_combo.addItems(["Утвержденный", "Исполненный", "Оба"])
        self.data_type_combo.currentTextChanged.connect(self.on_data_type_changed)
        tree_control_layout.addWidget(self.data_type_combo)
        
        # Чекбокс для скрытия нулевых столбцов
        self.hide_zero_columns_checkbox = QCheckBox("Скрыть нулевые столбцы")
        self.hide_zero_columns_checkbox.setToolTip("Скрыть столбцы, где в итоговой строке оба значения (утвержденный и исполненный) равны 0")
        self.hide_zero_columns_checkbox.stateChanged.connect(self.on_hide_zero_columns_changed)
        tree_control_layout.addWidget(self.hide_zero_columns_checkbox)
        
        # Панель инструментов для ревизии (активна только при выбранной ревизии)
        self.revision_toolbar = QHBoxLayout()
        self.revision_toolbar.setSpacing(5)
        
        # Кнопка пересчета
        self.recalculate_btn = QPushButton("Пересчитать")
        self.recalculate_btn.setIcon(self.style().standardIcon(QStyle.SP_BrowserReload))
        self.recalculate_btn.setToolTip("Пересчитать агрегированные суммы (F9)")
        self.recalculate_btn.setEnabled(False)
        self.recalculate_btn.clicked.connect(self.calculate_sums)
        self.revision_toolbar.addWidget(self.recalculate_btn)
        
        # Кнопка экспорта пересчитанной таблицы
        self.export_calculated_btn = QPushButton("Экспорт пересчитанной")
        self.export_calculated_btn.setIcon(self.style().standardIcon(QStyle.SP_DialogSaveButton))
        self.export_calculated_btn.setToolTip("Экспортировать форму с пересчитанными значениями")
        self.export_calculated_btn.setEnabled(False)
        self.export_calculated_btn.clicked.connect(self.export_calculated_table)
        self.revision_toolbar.addWidget(self.export_calculated_btn)
        
        # Кнопка показа ошибок расчетов
        self.show_errors_btn = QPushButton("Ошибки расчетов")
        self.show_errors_btn.setIcon(self.style().standardIcon(QStyle.SP_MessageBoxWarning))
        self.show_errors_btn.setToolTip("Показать ошибки расчетов")
        self.show_errors_btn.setEnabled(False)
        self.show_errors_btn.clicked.connect(self.show_calculation_errors)
        self.revision_toolbar.addWidget(self.show_errors_btn)
        
        # self.revision_toolbar.addSeparator()
        
        # Кнопка открытия файла
        self.open_file_btn = QPushButton("Открыть файл")
        self.open_file_btn.setIcon(self.style().standardIcon(QStyle.SP_DirOpenIcon))
        self.open_file_btn.setToolTip("Открыть файл (doc, docx, xls, xlsx)")
        self.open_file_btn.setEnabled(True)
        self.open_file_btn.clicked.connect(self.open_file_dialog)
        self.revision_toolbar.addWidget(self.open_file_btn)
        
        # Кнопка открытия последнего экспортированного файла
        self.open_last_file_btn = QPushButton("Открыть последний")
        self.open_last_file_btn.setIcon(self.style().standardIcon(QStyle.SP_FileDialogStart))
        self.open_last_file_btn.setToolTip("Открыть последний экспортированный файл")
        self.open_last_file_btn.setEnabled(False)
        self.open_last_file_btn.clicked.connect(self.open_last_exported_file)
        self.revision_toolbar.addWidget(self.open_last_file_btn)
        
        # self.revision_toolbar.addSeparator()
        
        # Меню документов
        self.documents_menu_btn = QPushButton("Документы ▼")
        self.documents_menu_btn.setIcon(self.style().standardIcon(QStyle.SP_FileDialogNewFolder))
        self.documents_menu_btn.setToolTip("Формирование документов")
        self.documents_menu_btn.setEnabled(False)
        self.documents_menu_btn.setMenu(QMenu(self))
        documents_menu = self.documents_menu_btn.menu()
        
        generate_conclusion_action = QAction("Сформировать заключение...", self)
        generate_conclusion_action.setIcon(self.style().standardIcon(QStyle.SP_FileDialogNewFolder))
        generate_conclusion_action.triggered.connect(self.show_document_dialog)
        documents_menu.addAction(generate_conclusion_action)
        
        generate_letters_action = QAction("Сформировать письма...", self)
        generate_letters_action.setIcon(self.style().standardIcon(QStyle.SP_FileDialogNewFolder))
        generate_letters_action.triggered.connect(self.show_document_dialog)
        documents_menu.addAction(generate_letters_action)
        
        documents_menu.addSeparator()
        
        parse_solution_action = QAction("Обработать решение о бюджете...", self)
        parse_solution_action.setIcon(self.style().standardIcon(QStyle.SP_DialogOpenButton))
        parse_solution_action.triggered.connect(self.parse_solution_document)
        documents_menu.addAction(parse_solution_action)
        
        self.revision_toolbar.addWidget(self.documents_menu_btn)
        
        tree_control_layout.addLayout(self.revision_toolbar)
        tree_layout.addLayout(tree_control_layout)
        
        # Древовидный виджет (используем стандартный заголовок QTreeWidget)
        self.data_tree = QTreeWidget()
        # Настраиваем заголовки дерева
        self.data_tree.setIndentation(10)
        # Отключаем единую высоту строк, чтобы высота подстраивалась под содержимое
        self.data_tree.setUniformRowHeights(False)
        # Включаем множественный выбор (Shift и Ctrl)
        self.data_tree.setSelectionMode(QTreeWidget.ExtendedSelection)
        # Устанавливаем делегат для переноса текста в ячейках
        self.data_tree.setItemDelegate(WordWrapItemDelegate())
        self.configure_tree_headers(self.current_section)
        self.data_tree.setContextMenuPolicy(Qt.CustomContextMenu)
        self.data_tree.customContextMenuRequested.connect(self.show_tree_context_menu)
        self.data_tree.itemExpanded.connect(self.on_tree_item_expanded)
        self.data_tree.itemCollapsed.connect(self.on_tree_item_collapsed)
        # Обработчики выделения
        self.data_tree.itemSelectionChanged.connect(self.on_tree_selection_changed)
        self.data_tree.itemClicked.connect(self.on_tree_item_clicked)

        # Контекстное меню по заголовкам дерева (управление столбцами)
        header = self.data_tree.header()
        header.setContextMenuPolicy(Qt.CustomContextMenu)
        header.customContextMenuRequested.connect(self.show_tree_header_context_menu)

        tree_layout.addWidget(self.data_tree)
        
        tabs.addTab(self.tree_tab, "Древовидные данные")
        
        # Вкладка с метаданными
        self.metadata_tab = QWidget()
        metadata_layout = QVBoxLayout(self.metadata_tab)
        
        self.metadata_text = QTextEdit()
        self.metadata_text.setReadOnly(True)
        metadata_layout.addWidget(self.metadata_text)
        
        tabs.addTab(self.metadata_tab, "Метаданные")
        
        # Вкладка с ошибками
        self.errors_tab = QWidget()
        errors_layout = QVBoxLayout(self.errors_tab)
        errors_layout.setContentsMargins(5, 5, 5, 5)
        
        # Заголовок и фильтры
        header_layout = QHBoxLayout()
        
        info_label = QLabel("Ошибки расчетов (несоответствия между оригинальными и расчетными значениями):")
        info_label.setFont(QFont("Arial", 10, QFont.Bold))
        header_layout.addWidget(info_label)
        
        header_layout.addStretch()
        
        # Фильтр по разделу
        header_layout.addWidget(QLabel("Раздел:"))
        self.errors_section_filter = QComboBox()
        self.errors_section_filter.addItems(["Все", "Доходы", "Расходы", "Источники финансирования", "Консолидируемые расчеты"])
        self.errors_section_filter.currentTextChanged.connect(lambda: self._update_errors_table())
        header_layout.addWidget(self.errors_section_filter)
        
        errors_layout.addLayout(header_layout)
        
        # Таблица ошибок
        self.errors_table = QTableWidget()
        self.errors_table.setColumnCount(9)
        self.errors_table.setHorizontalHeaderLabels([
            "Раздел",
            "Наименование",
            "Код строки",
            "Уровень",
            "Тип",
            "Колонка",
            "Оригинальное",
            "Расчетное",
            "Разница"
        ])
        
        # Настройка таблицы
        header = self.errors_table.horizontalHeader()
        # Отключаем растягивание последнего столбца
        header.setStretchLastSection(False)
        # Используем Interactive режим для каждого столбца отдельно, чтобы можно было вручную изменять ширину
        for i in range(9):
            header.setSectionResizeMode(i, QHeaderView.Interactive)
        # Устанавливаем начальные размеры столбцов
        header.resizeSection(0, 120)  # Раздел
        header.resizeSection(1, 300)  # Наименование
        header.resizeSection(2, 100)  # Код строки
        header.resizeSection(3, 60)   # Уровень
        header.resizeSection(4, 120)  # Тип
        header.resizeSection(5, 100)  # Колонка
        header.resizeSection(6, 120)  # Оригинальное
        header.resizeSection(7, 120)  # Расчетное
        header.resizeSection(8, 120)  # Разница
        
        self.errors_table.setAlternatingRowColors(True)
        self.errors_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.errors_table.setEditTriggers(QTableWidget.NoEditTriggers)
        
        errors_layout.addWidget(self.errors_table)
        
        # Кнопки и статистика
        buttons_layout = QHBoxLayout()
        buttons_layout.addStretch()
        
        self.errors_export_btn = QPushButton("Экспорт...")
        self.errors_export_btn.setIcon(self.style().standardIcon(QStyle.SP_DialogSaveButton))
        self.errors_export_btn.clicked.connect(self._export_errors)
        buttons_layout.addWidget(self.errors_export_btn)
        
        errors_layout.addLayout(buttons_layout)
        
        # Статистика
        self.errors_stats_label = QLabel("Ошибок не найдено")
        self.errors_stats_label.setFont(QFont("Arial", 9))
        errors_layout.addWidget(self.errors_stats_label)
        
        # Хранилище данных ошибок
        self.errors_data = []
        
        tabs.addTab(self.errors_tab, "Ошибки")
        
        # Вкладка с просмотром Excel
        self.excel_viewer = ExcelViewer()
        tabs.addTab(self.excel_viewer, "Просмотр формы")
        
        return tabs
    
    def connect_signals(self):
        """Подключение сигналов"""
        self.controller.projects_updated.connect(self.update_projects_list)
        self.controller.project_loaded.connect(self.on_project_loaded)
        self.controller.calculation_completed.connect(self.on_calculation_completed)
        self.controller.export_completed.connect(self.on_export_completed)
        self.controller.error_occurred.connect(self.on_error_occurred)
    
    def update_projects_list(self, _projects):
        """Обновление дерева проектов по новой архитектуре MainController.build_project_tree"""
        from PyQt5.QtWidgets import QTreeWidgetItem

        self.projects_tree.clear()

        # Получаем структурированные данные от контроллера
        tree_data = self.controller.build_project_tree()

        for year_entry in tree_data:
            year_label = f"Год {year_entry['year']}"
            year_item = QTreeWidgetItem([year_label])
            self.projects_tree.addTopLevelItem(year_item)

            for proj in year_entry["projects"]:
                proj_item = QTreeWidgetItem([proj["name"]])
                # Сохраняем ID проекта на уровне узла проекта
                proj_item.setData(0, Qt.UserRole, proj["id"])
                year_item.addChild(proj_item)

                # Формы/периоды/ревизии (показываем даже пустые, с заглушками)
                if proj.get("forms"):
                    for form in proj["forms"]:
                        form_label = f"{form['form_name']} ({form['form_code']})"
                        form_item = QTreeWidgetItem([form_label])
                        proj_item.addChild(form_item)

                        periods = form.get("periods") or []
                        if not periods:
                            form_item.addChild(QTreeWidgetItem(["Нет периодов"]))
                            continue

                        for period in periods:
                            period_label = period.get("period_name") or period.get("period_code") or "—"
                            period_item = QTreeWidgetItem([period_label])
                            form_item.addChild(period_item)

                            revisions = period.get("revisions") or []
                            if revisions:
                                for rev in revisions:
                                    status_icon = "✅" if rev["status"] == "calculated" else "📝"
                                    rev_text = f"{status_icon} рев. {rev['revision']}"
                                    rev_item = QTreeWidgetItem([rev_text])
                                    rev_item.setData(0, Qt.UserRole, rev.get("project_id"))
                                    revision_id = rev.get("revision_id")
                                    rev_item.setData(0, Qt.UserRole + 1, revision_id)
                                    if revision_id:
                                        logger.debug(
                                            f"Сохранена ревизия в дереве: "
                                            f"revision_id={revision_id}, project_id={rev.get('project_id')}, revision={rev.get('revision')}"
                                        )
                                    period_item.addChild(rev_item)
                            else:
                                period_item.addChild(QTreeWidgetItem(["Нет ревизий"]))
                else:
                    # Совсем нет форм — заглушка
                    placeholder = QTreeWidgetItem(["Нет ревизий"])
                    proj_item.addChild(placeholder)

        # Разворачиваем верхние уровни (год, проект, форма, период)
        # Ревизии остаются свернутыми по умолчанию
        for i in range(self.projects_tree.topLevelItemCount()):
            year_item = self.projects_tree.topLevelItem(i)
            year_item.setExpanded(True)
            for j in range(year_item.childCount()):
                proj_item = year_item.child(j)
                proj_item.setExpanded(True)
                for k in range(proj_item.childCount()):
                    form_item = proj_item.child(k)
                    form_item.setExpanded(True)
                    for m in range(form_item.childCount()):
                        period_item = form_item.child(m)
                        period_item.setExpanded(True)

    def on_project_tree_double_clicked(self, item, column):
        """Обработка двойного клика по дереву проектов"""
        # Поднимаемся по дереву, чтобы найти project_id/revision_id даже при клике на заглушки
        def _resolve_ids(it):
            proj_id = None
            rev_id = None
            cur = it
            while cur:
                if proj_id is None:
                    proj_id = cur.data(0, Qt.UserRole)
                if rev_id is None:
                    rev_id = cur.data(0, Qt.UserRole + 1)
                if proj_id is not None and rev_id is not None:
                    break
                cur = cur.parent()
            return proj_id, rev_id

        project_id, revision_id = _resolve_ids(item)
        
        if not project_id:
            return
        
        # Определяем, является ли узел ревизией (ревизия имеет revision_id и является дочерним элементом периода)
        is_revision = False
        if revision_id is not None and revision_id != 0:
            # Проверяем структуру дерева: ревизия является дочерним элементом периода
            parent = item.parent()
            if parent and item.childCount() == 0:
                # Период является дочерним элементом формы
                grandparent = parent.parent() if parent else None
                if grandparent:
                    grandparent_text = grandparent.text(0).lower()
                    if "форма" in grandparent_text or "(" in grandparent_text:
                        is_revision = True
        
        if is_revision:
            # Подтягиваем параметры формы из ревизии для последующей загрузки файлов
            self.controller.set_form_params_from_revision(revision_id)
            # Загружаем конкретную ревизию
            logger.info(f"Загрузка ревизии {revision_id} для проекта {project_id}")
            self.controller.load_revision(revision_id, project_id)
        else:
            # Клик по проекту/форме/периоду/заглушке — выбираем проект, чтобы можно было загрузить новую форму
            if project_id:
                logger.debug(f"Выбор проекта {project_id}")
                self.controller.project_controller.load_project(project_id)
            else:
                logger.warning("Проект не определён для выбранного узла")

    def show_project_context_menu(self, position):
        """Контекстное меню для дерева проектов"""
        item = self.projects_tree.itemAt(position)
        if not item:
            return
        project_id = item.data(0, Qt.UserRole)
        revision_id = item.data(0, Qt.UserRole + 1)

        # Если нет ID проекта — контекстное меню не показываем
        if not project_id:
            return

        # Определяем, является ли узел ревизией
        # Структура дерева: Год -> Проект -> Форма -> Период -> Ревизия
        # Ревизия - это узел, который является дочерним элементом периода
        # и не имеет дочерних элементов
        is_revision = False
        
        # Проверяем структуру дерева: ревизия является дочерним элементом периода
        parent = item.parent()
        if parent and item.childCount() == 0:
            # Период является дочерним элементом формы
            grandparent = parent.parent() if parent else None
            if grandparent:
                # Проверяем, что дедушка - это форма (содержит "форма" или "(")
                grandparent_text = grandparent.text(0).lower()
                if "форма" in grandparent_text or "(" in grandparent_text:
                    # Родитель - период, значит текущий узел - ревизия
                    is_revision = True

        menu = QMenu()
        edit_action = None
        edit_rev_action = None
        delete_rev_action = None
        delete_project_action = None

        # Если это узел ревизии
        if is_revision:
            # Для ревизии нужен revision_id для редактирования/удаления
            if revision_id is not None:
                edit_rev_action = menu.addAction("Редактировать ревизию")
                delete_rev_action = menu.addAction("Удалить ревизию")
            # Если revision_id не установлен (виртуальная ревизия из старой модели),
            # действия редактирования/удаления недоступны
        else:
            # Для узла проекта (не ревизии) показываем действия проекта
            edit_action = menu.addAction("Редактировать проект")
            delete_project_action = menu.addAction("Удалить проект")

        action = menu.exec_(self.projects_tree.mapToGlobal(position))

        if action == edit_action:
            self.edit_project(project_id)
        elif edit_rev_action is not None and action == edit_rev_action and revision_id:
            self.edit_revision(revision_id, project_id)
        elif delete_rev_action is not None and action == delete_rev_action and revision_id:
            reply = QMessageBox.question(
                self,
                "Подтверждение",
                "Вы уверены, что хотите удалить выбранную ревизию?",
                QMessageBox.Yes | QMessageBox.No,
            )
            if reply == QMessageBox.Yes:
                self.controller.delete_form_revision(revision_id)
                # После удаления ревизии обновляем дерево
                self.update_projects_list(None)
        elif action == delete_project_action:
            reply = QMessageBox.question(
                self,
                "Подтверждение",
                "Вы уверены, что хотите удалить проект (все ревизии)?",
                QMessageBox.Yes | QMessageBox.No,
            )
            if reply == QMessageBox.Yes:
                self.controller.delete_project(project_id)

    def edit_project(self, project_id: int):
        """Редактирование проекта через диалог"""
        try:
            # Загружаем проект в контроллер (установит current_project)
            self.controller.load_project(project_id)

            from views.project_dialog import ProjectDialog

            dlg = ProjectDialog(self)
            # Заполняем диалог текущим проектом
            if self.controller.current_project:
                dlg.set_project(self.controller.current_project)

            if dlg.exec_():
                project_data = dlg.get_project_data()
                if self.controller.update_project(project_data):
                    self.status_bar.showMessage(
                        f"Проект '{self.controller.current_project.name}' обновлён"
                    )
                    # Обновляем дерево проектов
                    self.update_projects_list(None)
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Ошибка редактирования проекта: {e}")
    
    def edit_revision(self, revision_id: int, project_id: int):
        """Редактирование ревизии через диалог"""
        try:
            from views.revision_dialog import RevisionDialog

            dlg = RevisionDialog(self.controller.db_manager, self)
            # Загружаем данные ревизии
            revision = self.controller.db_manager.get_form_revision_by_id(revision_id)
            if not revision:
                QMessageBox.warning(self, "Ошибка", "Ревизия не найдена")
                return
            
            dlg.set_revision(revision, project_id)

            if dlg.exec_():
                revision_data = dlg.get_revision_data()
                if self.controller.update_form_revision(revision_id, revision_data):
                    self.status_bar.showMessage("Ревизия обновлена")
                    # Дерево проектов обновится автоматически через сигнал projects_updated
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Ошибка редактирования ревизии: {e}")
    
    def on_project_loaded(self, project):
        """Обработка загруженного проекта"""
        try:
            # Убеждаемся, что прогресс-бар скрыт
            self.progress_bar.setVisible(False)

            # --------------------------------------------------
            # Определяем текущую ревизию и связанную мета‑информацию
            # --------------------------------------------------
            rev_id = getattr(self.controller, "current_revision_id", None)
            form_text = "—"
            revision_text = "—"
            status_text = "—"
            period_text = "—"

            excel_path = None

            if rev_id:
                try:
                    db = self.controller.db_manager
                    revision = db.get_form_revision_by_id(rev_id)
                    if revision:
                        # Ревизия и статус
                        revision_text = revision.revision or "—"
                        from models.base_models import ProjectStatus  # локальный импорт, чтобы избежать циклов
                        if isinstance(revision.status, ProjectStatus):
                            status_text = revision.status.value
                        else:
                            # На случай строкового статуса
                            status_text = str(revision.status or "—")

                        # Путь к файлу для Excel‑просмотра
                        excel_path = revision.file_path or None

                        # Находим связанную форму и её тип / период
                        project_forms = db.load_project_forms(project.id)
                        pf = next((p for p in project_forms if p.id == revision.project_form_id), None)
                        if pf:
                            # Тип формы
                            form_types_meta = {ft.id: ft for ft in db.load_form_types_meta()}
                            ft_meta = form_types_meta.get(pf.form_type_id)
                            if ft_meta:
                                # Показываем и код, и читаемое имя, если есть
                                if ft_meta.name:
                                    form_text = f"{ft_meta.name} ({ft_meta.code})"
                                else:
                                    form_text = ft_meta.code
                            # Период
                            if pf.period_id:
                                periods = db.load_periods()
                                period_ref = next((p for p in periods if p.id == pf.period_id), None)
                                if period_ref:
                                    period_text = period_ref.name or period_ref.code or period_text
                    else:
                        # Если ревизия по ID не найдена — fallback на старые поля проекта
                        revision_text = project.revision or "—"
                        status_text = getattr(project.status, "value", str(project.status)) if project.status else "—"
                        form_text = getattr(project.form_type, "value", str(project.form_type)) if project.form_type else "—"
                except Exception as e:
                    logger.error(f"Ошибка получения информации о ревизии: {e}", exc_info=True)
                    # Fallback на старые поля проекта
                    revision_text = project.revision or "—"
                    status_text = getattr(project.status, "value", str(project.status)) if project.status else "—"
                    form_text = getattr(project.form_type, "value", str(project.form_type)) if project.form_type else "—"
            else:
                # Проект без выбранной ревизии (старые проекты или только что созданные)
                form_text = getattr(project.form_type, "value", str(project.form_type)) if project.form_type else "—"
                revision_text = project.revision or "—"
                status_text = getattr(project.status, "value", str(project.status)) if project.status else "—"

            # МО — берём из справочника по municipality_id проекта
            municipality_text = "—"
            try:
                if hasattr(project, "municipality_id") and project.municipality_id:
                    db = self.controller.db_manager
                    municip_list = db.load_municipalities()
                    municip_ref = next((m for m in municip_list if m.id == project.municipality_id), None)
                    if municip_ref:
                        municipality_text = municip_ref.name or municipality_text
            except Exception as e:
                logger.warning(f"Ошибка получения МО для проекта {project.id}: {e}", exc_info=True)

            # Обновляем информацию о проекте
            info_text = (
                f"<b>Проект:</b> {project.name}<br>"
                f"<b>Форма:</b> {form_text}<br>"
                f"<b>Ревизия:</b> {revision_text}<br>"
                f"<b>МО:</b> {municipality_text}<br>"
                f"<b>Период:</b> {period_text}<br>"
                f"<b>Статус:</b> {status_text}<br>"
                f"<b>Создан:</b> {project.created_at.strftime('%d.%m.%Y %H:%M')}"
            )
            self.project_info_label.setText(info_text)

            # Обновляем состояние кнопок ревизии
            self.update_revision_buttons_state(rev_id is not None)

            # Загружаем данные в древовидное представление
            self.load_project_data_to_tree(project)

            # Загружаем метаданные
            self.load_metadata(project)
            
            # Обновляем вкладку ошибок
            self.load_errors_to_tab(project.data)

            # Загружаем файл в просмотрщик Excel:
            # Используем исходный файл ревизии (form_revisions.file_path), а не экспортированный
            # Экспортированный файл сохраняется отдельно и не должен заменять исходный
            if excel_path and os.path.exists(excel_path):
                # excel_path уже содержит путь к исходному файлу ревизии из revision_record.file_path
                self.excel_viewer.load_excel_file(excel_path)
            # Если файл не найден, просто не загружаем его

            self.status_bar.showMessage(f"Проект '{project.name}' загружен")
        except Exception as e:
            error_msg = f"Ошибка при загрузке проекта: {e}"
            logger.error(error_msg, exc_info=True)
            self.status_bar.setVisible(True)
            self.status_bar.showMessage(error_msg)
            self.progress_bar.setVisible(False)
    
    def _get_tree_widgets(self):
        """Получить все виджеты дерева (в главном окне и открепленных)"""
        widgets = []
        # Виджет в главном окне
        if hasattr(self, 'data_tree') and self.data_tree:
            widgets.append(self.data_tree)
        
        # Виджеты в открепленных окнах
        if "Древовидные данные" in self.detached_windows:
            detached_window = self.detached_windows["Древовидные данные"]
            tab_widget = detached_window.get_tab_widget()
            if tab_widget:
                for child in tab_widget.findChildren(QTreeWidget):
                    if child not in widgets:
                        widgets.append(child)
        
        return widgets if widgets else []
    
    def _get_errors_widgets(self):
        """Получить все виджеты ошибок с их фильтрами и метками (в главном окне и открепленных)"""
        widgets_info = []
        # Виджет в главном окне
        if hasattr(self, 'errors_tab') and self.errors_tab and hasattr(self, 'errors_table'):
            widgets_info.append({
                'table': self.errors_table,
                'filter': self.errors_section_filter,
                'stats': self.errors_stats_label
            })
        
        # Виджеты в открепленных окнах
        if "Ошибки" in self.detached_windows:
            detached_window = self.detached_windows["Ошибки"]
            tab_widget = detached_window.get_tab_widget()
            if tab_widget:
                # Ищем таблицу, фильтр и метку статистики в открепленном окне
                errors_table = None
                errors_filter = None
                errors_stats = None
                for child in tab_widget.findChildren(QTableWidget):
                    errors_table = child
                    break
                for child in tab_widget.findChildren(QComboBox):
                    errors_filter = child
                    break
                for child in tab_widget.findChildren(QLabel):
                    if "ошибок" in child.text().lower():
                        errors_stats = child
                        break
                if errors_table:
                    widgets_info.append({
                        'table': errors_table,
                        'filter': errors_filter,
                        'stats': errors_stats
                    })
        
        return widgets_info
    
    def _get_metadata_widgets(self):
        """Получить все виджеты метаданных (в главном окне и открепленных)"""
        widgets = []
        # Виджет в главном окне
        if hasattr(self, 'metadata_text') and self.metadata_text:
            widgets.append(self.metadata_text)
        
        # Виджеты в открепленных окнах
        if "Метаданные" in self.detached_windows:
            detached_window = self.detached_windows["Метаданные"]
            tab_widget = detached_window.get_tab_widget()
            if tab_widget:
                for child in tab_widget.findChildren(QTextEdit):
                    if child not in widgets:
                        widgets.append(child)
        
        return widgets
    
    def load_project_data_to_tree(self, project):
        """Загрузка данных проекта в древовидное представление"""
        try:
            if not project:
                self.status_bar.showMessage("Проект не выбран")
                return
            
            if not project.data:
                self.status_bar.showMessage("В проекте нет данных для отображения")
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
                self.status_bar.showMessage("Ошибка: виджеты дерева не инициализированы")
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
            self.configure_tree_headers(self.current_section)
            
            section_key = section_map.get(self.current_section)
            if section_key and section_key in project.data:
                data = project.data[section_key]
                if data and len(data) > 0:
                    # Для раздела "Расходы" подсвечиваем строку 450, сравнивая
                    # план/исполнение с пересчитанным результатом исполнения бюджета
                    # (дефицит/профицит), который теперь берём из calculated_deficit_proficit.
                    if (
                        self.current_section == "Расходы"
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
                        # Сначала настраиваем заголовки, чтобы кастомный заголовок был установлен
                        self._configure_tree_headers_for_widget(tree_widget, self.current_section)
                        # Затем загружаем данные
                        self.build_tree_from_data(data, tree_widget)
                    
                    # Обновляем высоту заголовка после загрузки данных
                    # Обновляем синхронно и через таймер для надежности
                    self._update_tree_header_height_for_all()
                    QTimer.singleShot(100, lambda: self._update_tree_header_height_for_all())
                    # Обновляем вкладку ошибок
                    self.load_errors_to_tab(project.data)
                    # Применяем скрытие нулевых столбцов, если чекбокс включен
                    if hasattr(self, 'hide_zero_columns_checkbox') and self.hide_zero_columns_checkbox.isChecked():
                        QTimer.singleShot(150, lambda: self.apply_hide_zero_columns())
                    self.status_bar.showMessage(f"Загружено {len(data)} записей в разделе '{self.current_section}'")
                else:
                    self.status_bar.showMessage(f"В разделе '{self.current_section}' нет данных для отображения")
            else:
                self.status_bar.showMessage(f"Раздел '{self.current_section}' не найден в данных проекта")
        except Exception as e:
            error_msg = f"Ошибка загрузки данных в дерево: {e}"
            logger.error(error_msg, exc_info=True)
            self.status_bar.showMessage(error_msg)

    def load_errors_to_tab(self, project_data):
        """Загрузка ошибок расчетов во вкладку ошибок"""
        self.errors_data = []
        
        if not project_data:
            # Обновляем все таблицы ошибок
            for widget_info in self._get_errors_widgets():
                self._update_errors_table(
                    widget_info.get('table'),
                    widget_info.get('filter'),
                    widget_info.get('stats')
                )
            return
        
        # Проверяем разделы
        sections = {
            "Доходы": "доходы_data",
            "Расходы": "расходы_data",
            "Источники финансирования": "источники_финансирования_data",
            "Консолидируемые расчеты": "консолидируемые_расчеты_data"
        }
        
        for section_name, section_key in sections.items():
            section_data = project_data.get(section_key, [])
            if not section_data:
                continue
            
            if section_name == "Консолидируемые расчеты":
                self._check_consolidated_errors(section_data, section_name)
            else:
                self._check_budget_errors(section_data, section_name)
        
        # Обновляем все таблицы ошибок
        for widget_info in self._get_errors_widgets():
            self._update_errors_table(
                widget_info.get('table'),
                widget_info.get('filter'),
                widget_info.get('stats')
            )
    
    def _check_budget_errors(self, data, section_name: str):
        """Проверка ошибок для бюджетных разделов (доходы, расходы, источники)"""
        budget_cols = Form0503317Constants.BUDGET_COLUMNS
        
        for item in data:
            level = item.get('уровень', 0)
            # Проверяем только уровни < 6
            if level >= 6:
                continue
            
            name = item.get('наименование_показателя', '')
            code = item.get('код_строки', '')
            
            approved_data = item.get('утвержденный', {}) or {}
            executed_data = item.get('исполненный', {}) or {}
            
            for col in budget_cols:
                # Проверка утвержденных значений
                original_approved = approved_data.get(col, 0) or 0
                calculated_approved = item.get(f'расчетный_утвержденный_{col}', original_approved)
                
                if self._is_value_different(original_approved, calculated_approved):
                    diff = self._calculate_error_difference(original_approved, calculated_approved)
                    self.errors_data.append({
                        'section': section_name,
                        'name': name,
                        'code': code,
                        'level': level,
                        'type': 'Утвержденный',
                        'column': col,
                        'original': original_approved,
                        'calculated': calculated_approved,
                        'difference': diff
                    })
                
                # Проверка исполненных значений
                original_executed = executed_data.get(col, 0) or 0
                calculated_executed = item.get(f'расчетный_исполненный_{col}', original_executed)
                
                if self._is_value_different(original_executed, calculated_executed):
                    diff = self._calculate_error_difference(original_executed, calculated_executed)
                    self.errors_data.append({
                        'section': section_name,
                        'name': name,
                        'code': code,
                        'level': level,
                        'type': 'Исполненный',
                        'column': col,
                        'original': original_executed,
                        'calculated': calculated_executed,
                        'difference': diff
                    })
    
    def _check_consolidated_errors(self, data, section_name: str):
        """Проверка ошибок для консолидированных расчетов"""
        cons_cols = Form0503317Constants.CONSOLIDATED_COLUMNS
        
        for item in data:
            level = item.get('уровень', 0)
            # Для консолидированных расчетов проверяем все уровни для столбца ИТОГО,
            # и уровни < 6 для остальных столбцов
            name = item.get('наименование_показателя', '')
            code = item.get('код_строки', '')
            
            cons_data = item.get('поступления', {}) or {}
            
            for col in cons_cols:
                # Оригинальное значение
                if isinstance(cons_data, dict) and col in cons_data:
                    original_value = cons_data.get(col, 0) or 0
                else:
                    original_value = item.get(f'поступления_{col}', 0) or 0
                
                # Расчетное значение
                calculated_value = item.get(f'расчетный_поступления_{col}')
                if calculated_value is None:
                    calculated_value = original_value
                
                # Проверяем несоответствие
                is_total_column = (col == 'ИТОГО')
                should_check = (level < 6) or is_total_column
                
                if should_check and self._is_value_different(original_value, calculated_value):
                    diff = self._calculate_error_difference(original_value, calculated_value)
                    self.errors_data.append({
                        'section': section_name,
                        'name': name,
                        'code': code,
                        'level': level,
                        'type': 'Поступления',
                        'column': col,
                        'original': original_value,
                        'calculated': calculated_value,
                        'difference': diff
                    })
    
    def _calculate_error_difference(self, original: float, calculated: float) -> float:
        """Вычисление разницы между значениями"""
        try:
            original_val = float(original) if original not in (None, "", "x") else 0.0
            calculated_val = float(calculated) if calculated not in (None, "", "x") else 0.0
            return calculated_val - original_val
        except (ValueError, TypeError):
            return 0.0
    
    def _update_errors_table(self, errors_table=None, section_filter_widget=None, stats_label=None):
        """Обновление таблицы с ошибками"""
        if errors_table is None:
            errors_table = self.errors_table
        if section_filter_widget is None:
            section_filter_widget = self.errors_section_filter
        if stats_label is None:
            stats_label = self.errors_stats_label
        
        # Фильтрация по разделу
        section_filter = section_filter_widget.currentText() if section_filter_widget else "Все"
        filtered_data = self.errors_data
        if section_filter != "Все":
            filtered_data = [e for e in self.errors_data if e['section'] == section_filter]
        
        # Заполнение таблицы
        errors_table.setRowCount(len(filtered_data))
        
        error_color = QColor("#FF6B6B")
        
        for row_idx, error in enumerate(filtered_data):
            # Раздел
            errors_table.setItem(row_idx, 0, QTableWidgetItem(error['section']))
            
            # Наименование
            name_item = QTableWidgetItem(error['name'])
            name_item.setForeground(QBrush(error_color))
            errors_table.setItem(row_idx, 1, name_item)
            
            # Код строки
            errors_table.setItem(row_idx, 2, QTableWidgetItem(str(error['code'])))
            
            # Уровень
            errors_table.setItem(row_idx, 3, QTableWidgetItem(str(error['level'])))
            
            # Тип
            errors_table.setItem(row_idx, 4, QTableWidgetItem(error['type']))
            
            # Колонка
            errors_table.setItem(row_idx, 5, QTableWidgetItem(error['column']))
            
            # Оригинальное значение
            orig_text = self._format_error_value(error['original'])
            orig_item = QTableWidgetItem(orig_text)
            errors_table.setItem(row_idx, 6, orig_item)
            
            # Расчетное значение
            calc_text = self._format_error_value(error['calculated'])
            calc_item = QTableWidgetItem(calc_text)
            calc_item.setForeground(QBrush(error_color))
            errors_table.setItem(row_idx, 7, calc_item)
            
            # Разница
            diff_text = self._format_error_value(error['difference'])
            diff_item = QTableWidgetItem(diff_text)
            diff_item.setForeground(QBrush(error_color))
            errors_table.setItem(row_idx, 8, diff_item)
        
        # Убеждаемся, что режим изменения размера столбцов установлен (на случай, если он был сброшен)
        header = errors_table.horizontalHeader()
        header.setStretchLastSection(False)
        for i in range(9):
            header.setSectionResizeMode(i, QHeaderView.Interactive)
        
        # Обновление статистики
        if stats_label:
            total_count = len(self.errors_data)
            filtered_count = len(filtered_data)
            if section_filter == "Все":
                stats_label.setText(f"Всего ошибок: {total_count}")
            else:
                stats_label.setText(f"Ошибок в разделе '{section_filter}': {filtered_count} (всего: {total_count})")
    
    def _format_error_value(self, value) -> str:
        """Форматирование значения для отображения"""
        if value in (None, "", "x"):
            return ""
        try:
            return f"{float(value):,.2f}"
        except (ValueError, TypeError):
            return str(value)
    
    def _export_errors(self):
        """Экспорт ошибок в файл"""
        import csv
        
        if not self.errors_data:
            QMessageBox.information(self, "Информация", "Нет ошибок для экспорта")
            return
        
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Экспорт ошибок расчетов",
            "ошибки_расчетов.csv",
            "CSV files (*.csv);;All files (*.*)"
        )
        
        if not file_path:
            return
        
        try:
            with open(file_path, 'w', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f, delimiter=';')
                # Заголовки
                writer.writerow([
                    "Раздел", "Наименование", "Код строки", "Уровень",
                    "Тип", "Колонка", "Оригинальное", "Расчетное", "Разница"
                ])
                # Данные
                for error in self.errors_data:
                    writer.writerow([
                        error['section'],
                        error['name'],
                        error['code'],
                        error['level'],
                        error['type'],
                        error['column'],
                        self._format_error_value(error['original']),
                        self._format_error_value(error['calculated']),
                        self._format_error_value(error['difference'])
                    ])
            
            QMessageBox.information(self, "Успех", f"Ошибки экспортированы в файл:\n{file_path}")
        except Exception as e:
            logger.error(f"Ошибка экспорта: {e}", exc_info=True)
            QMessageBox.critical(self, "Ошибка", f"Не удалось экспортировать ошибки:\n{e}")
    
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
            display_headers = self.tree_headers
        if mapping is None:
            mapping = self.tree_column_mapping
        
        # Устанавливаем делегат для переноса текста в ячейках
        tree_widget.setItemDelegate(WordWrapItemDelegate())
        # Отключаем единую высоту строк, чтобы высота подстраивалась под содержимое
        tree_widget.setUniformRowHeights(False)
        
        # Применяем текущий размер шрифта
        font = tree_widget.font()
        font.setPointSize(self.font_size)
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
        header_font.setPointSize(self.header_font_size)
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
            tree_widget = self.data_tree
        try:
            header = tree_widget.header()
            font_metrics = header.fontMetrics()
            max_height = 0
            
            # Получаем заголовки из headerItem
            header_item = tree_widget.headerItem()
            if header_item:
                # Проходим по всем заголовкам и вычисляем максимальную высоту с учетом переноса
                for idx in range(tree_widget.columnCount()):
                    if tree_widget.isColumnHidden(idx):
                        continue
                    
                    # Получаем текст из headerItem
                    text = header_item.text(idx) if idx < tree_widget.columnCount() else ""
                    if not text and idx < len(self.tree_headers):
                        text = self.tree_headers[idx]
                    
                    if text:
                        # Получаем ширину столбца
                        width = max(header.sectionSize(idx), 50)
                        
                        # Создаем документ для расчета высоты с учетом переноса
                        from PyQt5.QtGui import QTextDocument, QTextOption
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
                for idx, text in enumerate(self.tree_headers):
                    if text and not tree_widget.isColumnHidden(idx):
                        width = max(header.sectionSize(idx), 50)
                        
                        # Создаем документ для расчета высоты
                        from PyQt5.QtGui import QTextDocument, QTextOption
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
                header = self.data_tree.header()
                font_metrics = header.fontMetrics()
                header.setFixedHeight(font_metrics.lineSpacing() + 6)
            except:
                pass
    
    def _on_tree_header_section_resized(self, logicalIndex, oldSize, newSize):
        """Обработчик изменения размера столбца заголовка дерева"""
        # Обновляем высоту заголовка при изменении размера столбца
        try:
            QTimer.singleShot(100, self._update_tree_header_height)
        except Exception as e:
            logger.warning(f"Ошибка в _on_tree_header_section_resized: {e}", exc_info=True)

    def hide_zero_columns_in_tree(self, section_key: str, data):
        """
        Скрытие столбцов дерева, в которых итоговое значение равно 0.
        Логика аналогична табличному представлению.
        """
        if not data:
            return

        if section_key == "консолидируемые_расчеты_data":
            cons_cols = Form0503317Constants.CONSOLIDATED_COLUMNS
            mapping = self.tree_column_mapping or {}
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

            header = self.data_tree.header()
            zero_cols = []
            for i, col_name in enumerate(cons_cols):
                val = totals.get(col_name, 0)
                if isinstance(val, (int, float)) and abs(val) < 1e-9:
                    col_index = value_start + i
                    if 0 <= col_index < self.data_tree.columnCount():
                        zero_cols.append(col_index)

            # Сужаем «нулевые» колонки до минимальной ширины и очищаем заголовки
            header_item = self.data_tree.headerItem()
            for col_index in zero_cols:
                header.resizeSection(col_index, 2)  # минимальная ширина
                if header_item:
                    header_item.setText(col_index, "")
                    header_item.setToolTip(col_index, "")
            return

        # Доходы, расходы, источники
        # Паттерн: итоговые строки для первых трёх разделов
        # оканчиваются на "всего" (регистр не важен).
        budget_cols = Form0503317Constants.BUDGET_COLUMNS
        mapping = self.tree_column_mapping or {}
        if mapping.get("type") != "budget":
            return

        total_item = None
        for item in data:
            name = str(item.get("наименование_показателя", "")).strip().lower()
            # Ищем первую строку, где встречается слово "всего"
            # (без жёсткого условия «только в конце», чтобы не зависеть
            #  от возможных двоеточий, уточнений и т.п.)
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
        show_approved = self.current_data_type in ("Утвержденный", "Оба")
        show_executed = self.current_data_type in ("Исполненный", "Оба")

        header = self.data_tree.header()
        zero_cols = set()
        for i, col_name in enumerate(budget_cols):
            a_val = approved.get(col_name, 0) or 0
            e_val = executed.get(col_name, 0) or 0
            if isinstance(a_val, (int, float)) and isinstance(e_val, (int, float)):
                if abs(a_val) < 1e-9 and abs(e_val) < 1e-9:
                    appr_idx = approved_start + i
                    exec_idx = executed_start + i
                    if show_approved and 0 <= appr_idx < self.data_tree.columnCount():
                        zero_cols.add(appr_idx)
                    if show_executed and 0 <= exec_idx < self.data_tree.columnCount():
                        zero_cols.add(exec_idx)

        # Сужаем «нулевые» колонки до минимальной ширины и очищаем заголовки
        header_item = self.data_tree.headerItem()
        for col_index in zero_cols:
            header.resizeSection(col_index, 2)  # минимальная ширина
            if header_item:
                header_item.setText(col_index, "")
                header_item.setToolTip(col_index, "")

    def apply_tree_data_type_visibility(self):
        """Скрывает столбцы дерева в зависимости от выбранного типа данных"""
        if not self.tree_column_mapping:
            return

        column_total = len(self.tree_headers)
        
        # Применяем ко всем деревьям
        for tree_widget in self._get_tree_widgets():
            for col in range(column_total):
                tree_widget.setColumnHidden(col, False)

            if self.tree_column_mapping.get("type") != "budget":
                continue

            approved_start = self.tree_column_mapping.get("approved_start", 0)
            executed_start = self.tree_column_mapping.get("executed_start", 0)
            budget_cols = self.tree_column_mapping.get("budget_columns", [])

            approved_range = range(approved_start, approved_start + len(budget_cols))
            executed_range = range(executed_start, executed_start + len(budget_cols))

            show_approved = self.current_data_type in ("Утвержденный", "Оба")
            show_executed = self.current_data_type in ("Исполненный", "Оба")

            for idx in approved_range:
                tree_widget.setColumnHidden(idx, not show_approved)
            for idx in executed_range:
                tree_widget.setColumnHidden(idx, not show_executed)

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
    
    def build_tree_from_data(self, data, tree_widget=None):
        """Построение дерева из данных"""
        try:
            if tree_widget is None:
                tree_widget = self.data_tree
            
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
            
            # Обновляем размеры столбцов после загрузки данных
            if items_created > 0:
                header = tree_widget.header()
                # Обновляем размеры столбцов
                for idx in range(tree_widget.columnCount()):
                    if not tree_widget.isColumnHidden(idx):
                        if idx == 0:  # Столбец "Наименование" - устанавливаем ширину с учетом отступов дерева
                            # Получаем отступы дерева и добавляем запас
                            indentation = tree_widget.indentation()
                            # Добавляем запас на отступы (примерно 6 уровней * отступ + небольшой запас)
                            indent_reserve = indentation * 6 + 50  # Запас на отступы и дополнительные элементы
                            header.resizeSection(idx, 400 + indent_reserve)
                        elif idx == 1:  # Столбец "Код строки" - устанавливаем фиксированную ширину 80px
                            header.setSectionResizeMode(idx, QHeaderView.Fixed)
                            header.resizeSection(idx, 80)
                        elif idx == 2:  # Столбец "Код классификации" - устанавливаем фиксированную ширину 200px
                            header.resizeSection(idx, 200)
                        elif idx == 3:  # Столбец "Уровень" - устанавливаем фиксированную ширину 50px
                            header.setSectionResizeMode(idx, QHeaderView.Fixed)
                            header.resizeSection(idx, 50)
                        else:
                            # Остальные столбцы - фиксированная ширина 150px
                            header.setSectionResizeMode(idx, QHeaderView.Fixed)
                            header.resizeSection(idx, 150)
                # Обновляем высоту заголовка
                QTimer.singleShot(100, lambda tw=tree_widget: self._update_tree_header_height(tw))
            
            if items_created > 0 and tree_widget == self.data_tree:
                msg = f"Построено дерево: {items_created} элементов"
                if items_failed > 0:
                    msg += f", ошибок: {items_failed}"
                self.status_bar.showMessage(msg)
        except Exception as e:
            error_msg = f"Ошибка построения дерева: {e}"
            logger.error(error_msg, exc_info=True)
            if tree_widget == self.data_tree:
                self.status_bar.showMessage(error_msg)
    
    def create_tree_item(self, item, level_colors, tree_widget=None):
        """Создание элемента дерева"""
        try:
            if tree_widget is None:
                tree_widget = self.data_tree
            
            level = item.get('уровень', 0)

            column_count = tree_widget.columnCount()
            if column_count == 0:
                # Если колонок нет, создаем хотя бы одну
                tree_widget.setColumnCount(1)
                column_count = 1
            
            tree_item = QTreeWidgetItem([""] * column_count)
            
            # Основные данные
            name = str(item.get('наименование_показателя', ''))
            code_line = str(item.get('код_строки', ''))
            class_code = str(item.get('код_классификации_форматированный', item.get('код_классификации', '')))

            if column_count > 0:
                tree_item.setText(0, name)
            if column_count > 1:
                tree_item.setText(1, code_line)
            if column_count > 2:
                tree_item.setText(2, class_code)
            if column_count > 3:
                tree_item.setText(3, str(level))

            mapping = self.tree_column_mapping or {}
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
                            if approved_start + idx < column_count:
                                tree_item.setText(approved_start + idx, approved_value)
                                tree_item.setForeground(approved_start + idx, QBrush(error_color))
                        else:
                            approved_value = self.format_budget_value(original_approved)
                            if approved_start + idx < column_count:
                                tree_item.setText(approved_start + idx, approved_value)
                        
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
                            if executed_start + idx < column_count:
                                tree_item.setText(executed_start + idx, executed_value)
                                tree_item.setForeground(executed_start + idx, QBrush(error_color))
                        else:
                            executed_value = self.format_budget_value(original_executed)
                            if executed_start + idx < column_count:
                                tree_item.setText(executed_start + idx, executed_value)
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
                            if value_start + idx < column_count:
                                tree_item.setText(value_start + idx, display_value)
                                tree_item.setForeground(value_start + idx, QBrush(error_color))
                        else:
                            # Обычное отображение без несоответствий
                            if value_start + idx < column_count:
                                tree_item.setText(value_start + idx, self.format_budget_value(original_value))
                    except Exception as e:
                        logger.warning(f"Ошибка обработки несоответствий для консолидируемых расчетов, колонка {col}: {e}", exc_info=True)
                        pass
            
            # Устанавливаем цвет фона для всех столбцов
            try:
                if level in level_colors:
                    color = QColor(level_colors[level])
                    brush = QBrush(color)
                    # Применяем цвет ко всем столбцам
                    for i in range(column_count):
                        tree_item.setBackground(i, brush)
            except Exception as e:
                logger.warning(f"Ошибка установки цвета фона для уровня {level}: {e}", exc_info=True)
                pass
            
            # Устанавливаем подсказки (колонка -> заголовок)
            try:
                for idx, tip in enumerate(self.tree_header_tooltips):
                    if idx < tree_item.columnCount() and idx < len(self.tree_header_tooltips):
                        current_text = tree_item.text(idx)
                        if current_text:
                            tree_item.setToolTip(idx, f"{tip}: {current_text}")
                        else:
                            tree_item.setToolTip(idx, tip)
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
            column_count = max(self.data_tree.columnCount(), 1)
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
    
    def load_metadata(self, project):
        """Загрузка метаданных для выбранной ревизии"""
        # Метаданные должны быть только у ревизии, а не у проекта
        # Проверяем, что загружена ревизия (current_revision_id установлен)
        rev_id = getattr(self.controller, "current_revision_id", None)
        
        # Получаем все виджеты метаданных
        metadata_widgets = self._get_metadata_widgets()
        
        if not rev_id:
            # Если ревизия не загружена, метаданные не отображаем
            for metadata_widget in metadata_widgets:
                metadata_widget.setHtml("")
            return
        
        # Метаданные берём из данных проекта (которые загружаются из ревизии)
        if not project or not project.data:
            for metadata_widget in metadata_widgets:
                metadata_widget.setHtml("")
            return
        
        meta_info = project.data.get('meta_info', {})
        if not meta_info:
            for metadata_widget in metadata_widgets:
                metadata_widget.setHtml("")
            return
        
        metadata_text = ""
        for key, value in meta_info.items():
            metadata_text += f"<b>{key}:</b> {value}<br>"
        
        # Обновляем все виджеты метаданных
        for metadata_widget in metadata_widgets:
            metadata_widget.setHtml(metadata_text)
    
    def on_section_changed(self, section_name):
        """Обработка смены раздела"""
        self.current_section = section_name
        # Сбрасываем столбец выделения при смене раздела
        self.selection_start_column = None
        if self.controller.current_project:
            self.load_project_data_to_tree(self.controller.current_project)
            # Применяем скрытие нулевых столбцов, если чекбокс включен
            if hasattr(self, 'hide_zero_columns_checkbox') and self.hide_zero_columns_checkbox.isChecked():
                QTimer.singleShot(200, lambda: self.apply_hide_zero_columns())
    
    def on_data_type_changed(self, data_type):
        """Обработка смены типа данных"""
        self.current_data_type = data_type
        self.apply_tree_data_type_visibility()
        # Применяем скрытие нулевых столбцов, если чекбокс включен
        if hasattr(self, 'hide_zero_columns_checkbox') and self.hide_zero_columns_checkbox.isChecked():
            self.apply_hide_zero_columns()
        if self.controller.current_project:
            self.load_project_data_to_tree(self.controller.current_project)
    
    def on_hide_zero_columns_changed(self, state):
        """Обработка изменения состояния чекбокса 'Скрыть нулевые столбцы'"""
        if state == Qt.Checked:
            self.apply_hide_zero_columns()
        else:
            self.show_all_columns()
    
    def apply_hide_zero_columns(self):
        """Применить скрытие нулевых столбцов"""
        if not (self.controller.current_project and self.controller.current_project.data):
            logger.debug("apply_hide_zero_columns: нет проекта или данных")
            return

        section_map = {
            "Доходы": "доходы_data",
            "Расходы": "расходы_data",
            "Источники финансирования": "источники_финансирования_data",
            "Консолидируемые расчеты": "консолидируемые_расчеты_data"
        }
        section_key = section_map.get(self.current_section)
        if not section_key or section_key not in self.controller.current_project.data:
            logger.debug(f"apply_hide_zero_columns: раздел {self.current_section} не найден")
            return

        data = self.controller.current_project.data[section_key]
        if not data:
            logger.debug(f"apply_hide_zero_columns: нет данных для раздела {section_key}")
            return

        logger.debug(f"apply_hide_zero_columns: применяю скрытие для раздела {section_key}, записей: {len(data)}")

        # Сначала показываем все столбцы
        self.show_all_columns()
        
        # Применяем отображение колонок в зависимости от выбранного типа данных
        self.apply_tree_data_type_visibility()

        # Затем применяем скрытие нулевых столбцов (после применения видимости по типу данных)
        self.hide_zero_columns_in_tree(section_key, data)
    
    def expand_all_tree(self):
        """Развернуть все узлы дерева"""
        for tree_widget in self._get_tree_widgets():
            tree_widget.expandAll()
    
    def collapse_all_tree(self):
        """Свернуть все узлы дерева"""
        for tree_widget in self._get_tree_widgets():
            tree_widget.collapseAll()
    
    def on_tree_item_expanded(self, item):
        """Обработка разворачивания узла дерева"""
        pass
    
    def on_tree_item_collapsed(self, item):
        """Обработка сворачивания узла дерева"""
        pass
    
    def show_tree_context_menu(self, position):
        """Контекстное меню для дерева"""
        item = self.data_tree.itemAt(position)
        if not item:
            return
        
        menu = QMenu()
        copy_action = menu.addAction("Копировать значение")
        
        action = menu.exec_(self.data_tree.mapToGlobal(position))
        
        if action == copy_action:
            self.copy_tree_item_value(item)

    def show_tree_header_context_menu(self, position):
        """Контекстное меню для заголовков дерева (скрытие/отображение столбцов)"""
        header = self.data_tree.header()
        col = header.logicalIndexAt(position)
        if col < 0:
            return

        menu = QMenu(self)
        hide_action = menu.addAction("Скрыть столбец")
        show_all_action = menu.addAction("Показать все столбцы")
        chosen = menu.exec_(header.mapToGlobal(position))

        if chosen == hide_action:
            # Не скрываем первый столбец с названием
            if col > 0:
                self.data_tree.setColumnHidden(col, True)
        elif chosen == show_all_action:
            for i in range(self.data_tree.columnCount()):
                self.data_tree.setColumnHidden(i, False)
    
    def show_all_columns(self):
        """Показать все столбцы в дереве и вернуть им нормальные ширины/заголовки"""
        # Проще всего — переинициализировать заголовки и ширины так же,
        # как это делается при загрузке данных
        tree_widgets = self._get_tree_widgets()
        for tree_widget in tree_widgets:
            if tree_widget:
                self._configure_tree_headers_for_widget(tree_widget, self.current_section)

        # Снова применяем фильтр по типу данных (утверждённый/исполненный/оба)
        self.apply_tree_data_type_visibility()

    def hide_zero_columns_global(self):
        """Сворачивает все столбцы с нулевыми значениями в итоговой строке
        
        Этот метод оставлен для совместимости с меню/тулбаром.
        Теперь он устанавливает чекбокс и вызывает apply_hide_zero_columns.
        """
        if hasattr(self, 'hide_zero_columns_checkbox'):
            self.hide_zero_columns_checkbox.setChecked(True)
        else:
            # Если чекбокс еще не создан, используем старую логику
            self.apply_hide_zero_columns()
    
    def copy_tree_item_value(self, item):
        """Копировать значение из дерева"""
        if item:
            text = item.text(0)  # Копируем значение из первого столбца
            clipboard = QApplication.clipboard()
            clipboard.setText(text)
    
    
    def show_new_project_dialog(self):
        """Показать диалог создания проекта"""
        dialog = ProjectDialog(self)
        if dialog.exec_():
            project_data = dialog.get_project_data()
            project = self.controller.create_project(project_data)
            if project:
                QMessageBox.information(self, "Успех", f"Проект '{project.name}' создан")
    
    def show_reference_dialog(self, ref_type: str = None):
        """Показать диалог загрузки справочника"""
        dialog = ReferenceDialog(self, ref_type)
        if dialog.exec_():
            ref_data = dialog.get_reference_data()
            success = self.controller.load_reference_file(
                ref_data['file_path'],
                ref_data['reference_type'],
                ref_data['name']
            )
            if success:
                QMessageBox.information(self, "Успех", "Справочник загружен")
    
    def show_reference_viewer(self):
        """Показать просмотрщик справочников в отдельном окне"""
        from PyQt5.QtWidgets import QMainWindow, QToolBar

        if self.reference_window is None:
            self.reference_window = QMainWindow(self)
            self.reference_window.setWindowTitle("Справочники")
            self.reference_window.resize(900, 600)
            # Включаем стандартные кнопки окна (включая максимизацию)
            self.reference_window.setWindowFlags(self.reference_window.windowFlags() | Qt.WindowMaximizeButtonHint)
            self.reference_window.is_fullscreen = False

            self.reference_viewer = ReferenceViewer()
            self.reference_window.setCentralWidget(self.reference_viewer)
            
            # Обработка F11 для полноэкранного режима
            # Создаем обработчик событий для окна справочников
            def key_press_handler(event):
                if event.key() == Qt.Key_F11:
                    self._toggle_reference_fullscreen()
                else:
                    QMainWindow.keyPressEvent(self.reference_window, event)
            
            self.reference_window.keyPressEvent = key_press_handler

        # Устанавливаем callback для обновления данных из контроллера
        def refresh_callback():
            # Обновляем справочники в контроллере
            self.controller.refresh_references()
            # Обновляем данные в окне справочников
            self.reference_viewer.load_references(self.controller.references)
        
        self.reference_viewer.refresh_callback = refresh_callback
        
        # Загружаем актуальные справочники и показываем окно
        self.reference_viewer.load_references(self.controller.references)
        self.reference_window.show()
        self.reference_window.raise_()
        self.reference_window.activateWindow()
    
    def _toggle_reference_fullscreen(self):
        """Переключение полноэкранного режима для окна справочников"""
        if self.reference_window is None:
            return
        
        if self.reference_window.is_fullscreen:
            self.reference_window.showNormal()
            self.reference_window.is_fullscreen = False
        else:
            self.reference_window.showFullScreen()
            self.reference_window.is_fullscreen = True
    

    def show_config_dictionaries(self):
        """Показать диалог редактирования справочников конфигурации"""
        dlg = DictionariesDialog(self.controller.db_manager, self)
        dlg.exec_()

    def on_projects_side_button_clicked(self):
        """Обработчик клика по боковой кнопке панели проектов"""
        if not self.projects_inner_panel:
            return
        # Инвертируем состояние по видимости внутренней панели
        self.toggle_projects_panel(not self.projects_inner_panel.isVisible())

    def toggle_projects_panel(self, checked: bool = None):
        """Показать/скрыть панель проектов"""
        if not self.main_splitter or not self.projects_inner_panel:
            return
        
        # Если checked не указан, инвертируем текущее состояние
        if checked is None:
            checked = not self.projects_inner_panel.isVisible()
        else:
            # Обновляем состояние меню
            for action in self.menuBar().actions():
                if action.text() == "&Вид":
                    for sub_action in action.menu().actions():
                        if sub_action.text() == "&Панель проектов":
                            sub_action.setChecked(checked)
                            break
                    break

        if not checked:
            # Запоминаем текущую ширину панели перед схлопыванием
            sizes = self.main_splitter.sizes()
            if sizes and sizes[0] > 0:
                self.projects_panel_last_size = sizes[0]

            # Скрываем содержимое, оставляя узкую кнопку
            self.projects_inner_panel.setVisible(False)
            if self.projects_toggle_button:
                self.projects_toggle_button.setText("▶")

            handle_width = self.projects_toggle_button.width() if self.projects_toggle_button else 20
            self.main_splitter.setSizes([handle_width, max(400, self.width() - handle_width)])
        else:
            # Показываем внутреннюю панель
            self.projects_inner_panel.setVisible(True)
            if self.projects_toggle_button:
                self.projects_toggle_button.setText("◀")

            total_width = max(self.width(), self.projects_panel_last_size + 400)
            self.main_splitter.setSizes(
                [self.projects_panel_last_size, total_width - self.projects_panel_last_size]
            )
    
    def load_form_file(self):
        """Загрузка файла формы"""
        # Если проект не выбран, пытаемся выбрать из текущего выделения в дереве
        if not self.controller.current_project:
            item = self.projects_tree.currentItem()
            if item:
                proj_id = item.data(0, Qt.UserRole) or (item.parent().data(0, Qt.UserRole) if item.parent() else None)
                if proj_id:
                    self.controller.project_controller.load_project(proj_id)
        if not self.controller.current_project:
            QMessageBox.warning(self, "Ошибка", "Сначала выберите или создайте проект")
            return
        
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Выберите файл формы",
            "",
            "Excel files (*.xlsx *.xls);;All files (*.*)"
        )
        
        if file_path:
            # Перед загрузкой файла спрашиваем тип формы, период и ревизию
            defaults = self.controller.get_pending_form_params() if hasattr(self.controller, "get_pending_form_params") else {}
            params_dialog = FormLoadDialog(self.controller.db_manager, self, defaults=defaults)
            if params_dialog.exec_() != QDialog.Accepted:
                return

            form_params = params_dialog.get_form_params()

            # Сохраняем выбранные пользователем параметры формы в контроллере
            if self.controller.current_project:
                form_code = form_params["form_code"]
                revision = form_params["revision"]
                period_code = form_params["period_code"]

                self.controller.set_current_form_params(
                    form_code=form_code,
                    revision=revision,
                    period_code=period_code,
                )

            self.progress_bar.setVisible(True)
            self.progress_bar.setRange(0, 0)
            self.status_bar.showMessage("Загрузка файла формы...")

            QTimer.singleShot(100, lambda: self._process_form_file(file_path))
    
    def _process_form_file(self, file_path):
        """Обработка файла формы"""
        try:
            success = self.controller.load_form_file(file_path)
            if success:
                # Перезагружаем данные проекта после загрузки формы
                if self.controller.current_project:
                    self.load_project_data_to_tree(self.controller.current_project)
                QMessageBox.information(self, "Успех", "Форма загружена и распарсена")
                self.status_bar.showMessage("Форма успешно загружена")
            else:
                QMessageBox.warning(self, "Ошибка", "Не удалось загрузить форму")
                self.status_bar.showMessage("Ошибка загрузки формы")
        except Exception as e:
            error_msg = f"Ошибка обработки файла формы: {e}"
            logger.error(error_msg, exc_info=True)
            QMessageBox.critical(self, "Ошибка", error_msg)
            self.status_bar.showMessage(error_msg)
        finally:
            self.progress_bar.setVisible(False)
    
    def update_revision_buttons_state(self, has_revision: bool):
        """Обновление состояния кнопок ревизии в зависимости от наличия выбранной ревизии"""
        if hasattr(self, 'recalculate_btn'):
            self.recalculate_btn.setEnabled(has_revision)
        if hasattr(self, 'export_calculated_btn'):
            self.export_calculated_btn.setEnabled(has_revision)
        if hasattr(self, 'show_errors_btn'):
            self.show_errors_btn.setEnabled(has_revision)
        if hasattr(self, 'documents_menu_btn'):
            self.documents_menu_btn.setEnabled(has_revision)
    
    def calculate_sums(self):
        """Расчет агрегированных сумм"""
        if not self.controller.current_project or not self.controller.current_revision_id:
            QMessageBox.warning(self, "Ошибка", "Сначала выберите проект и загрузите ревизию формы")
            return
        
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, 0)
        
        QTimer.singleShot(100, self.controller.calculate_sums)
        QTimer.singleShot(1000, self._do_refresh_projects)
        
    
    def on_calculation_completed(self, results):
        """Обработка завершения расчета"""
        self.progress_bar.setVisible(False)
        QMessageBox.information(self, "Успех", "Расчет завершен")
        
        # Обновляем отображение данных
        if self.controller.current_project:
            self.load_project_data_to_tree(self.controller.current_project)
            # Обновляем вкладку ошибок
            self.load_errors_to_tab(self.controller.current_project.data)
    
    def export_validation(self):
        """Экспорт формы с проверкой (старый метод, оставлен для совместимости)"""
        self.export_calculated_table()
    
    def export_calculated_table(self):
        """Экспорт пересчитанной таблицы"""
        if not self.controller.current_project or not self.controller.current_revision_id:
            QMessageBox.warning(self, "Ошибка", "Сначала выберите проект и загрузите ревизию формы")
            return
        
        # Получаем информацию о ревизии для имени файла
        rev_id = self.controller.current_revision_id
        revision = self.controller.db_manager.get_form_revision_by_id(rev_id)
        revision_text = revision.revision if revision else "unknown"
        
        output_path, _ = QFileDialog.getSaveFileName(
            self,
            "Сохранить пересчитанную форму",
            f"{self.controller.current_project.name}_рев{revision_text}_пересчет.xlsx",
            "Excel files (*.xlsx)"
        )
        
        if output_path:
            self.progress_bar.setVisible(True)
            self.progress_bar.setRange(0, 0)
            
            QTimer.singleShot(100, lambda: self._process_export(output_path))
    
    def _process_export(self, output_path):
        """Обработка экспорта"""
        success = self.controller.export_validation(output_path)
        self.progress_bar.setVisible(False)
        
        if success:
            # Сохраняем путь к последнему экспортированному файлу
            self.last_exported_file = output_path
            self.open_last_file_action.setEnabled(True)
            if hasattr(self, 'open_last_file_btn'):
                self.open_last_file_btn.setEnabled(True)
            
            # Предлагаем открыть файл
            reply = QMessageBox.question(
                self,
                "Успех",
                f"Форма экспортирована: {output_path}\n\nОткрыть файл?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.Yes
            )
            if reply == QMessageBox.Yes:
                self.open_file(output_path)
        else:
            QMessageBox.warning(self, "Ошибка", "Не удалось экспортировать форму")
    
    def on_export_completed(self, file_path):
        """Обработка завершения экспорта"""
        self.status_bar.showMessage(f"Форма экспортирована: {file_path}")
    
    def on_error_occurred(self, error_message):
        """Обработка ошибки"""
        self.progress_bar.setVisible(False)
        QMessageBox.critical(self, "Ошибка", error_message)
        self.status_bar.showMessage(f"Ошибка: {error_message}")
    
    def refresh_projects(self):
        """Обновление списка проектов"""
        # Показываем прогресс-бар во время обновления
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, 0)  # Неопределенный прогресс
        self.status_bar.showMessage("Обновление списка проектов...")
        
        # Обновляем данные с небольшой задержкой, чтобы UI успел обновиться
        # Это позволяет показать прогресс-бар до начала загрузки
        QTimer.singleShot(10, self._do_refresh_projects)
    
    def _do_refresh_projects(self):
        """Выполнение обновления списка проектов"""
        try:
            # Обновляем только список проектов, не перезагружая текущий проект
            # Это предотвращает зависание из-за пересчета уровней
            projects = self.controller.project_controller.load_projects()
            self.controller.projects_updated.emit(projects)
            
            # Обновляем справочники отдельно, чтобы не блокировать UI
            self.controller.refresh_references()
            
            self.status_bar.showMessage("Список проектов обновлен")
        except Exception as e:
            self.status_bar.showMessage(f"Ошибка обновления: {str(e)}")
            QMessageBox.critical(self, "Ошибка", f"Ошибка обновления списка проектов: {str(e)}")
        finally:
            self.progress_bar.setVisible(False)
    
    def edit_current_project(self):
        """Редактировать текущий проект"""
        if not self.controller.current_project or not self.controller.current_project.id:
            QMessageBox.warning(self, "Предупреждение", "Проект не выбран")
            return
        self.edit_project(self.controller.current_project.id)
    
    def delete_current_project(self):
        """Удалить текущий проект"""
        if not self.controller.current_project or not self.controller.current_project.id:
            QMessageBox.warning(self, "Предупреждение", "Проект не выбран")
            return
        
        reply = QMessageBox.question(
            self, 
            "Подтверждение удаления",
            f"Вы уверены, что хотите удалить проект '{self.controller.current_project.name}'?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            self.controller.delete_project(self.controller.current_project.id)
            QMessageBox.information(self, "Успех", "Проект удален")
    
    def on_font_size_changed(self, size: int):
        """Обработка изменения размера шрифта данных"""
        self.font_size = size
        self.apply_font_sizes()
    
    def on_header_font_size_changed(self, size: int):
        """Обработка изменения размера шрифта заголовков"""
        self.header_font_size = size
        self.apply_font_sizes()
    
    def apply_font_sizes(self):
        """Применение размеров шрифтов ко всем деревьям"""
        # Получаем все виджеты дерева
        tree_widgets = self._get_tree_widgets()
        
        for tree_widget in tree_widgets:
            if tree_widget:
                # Применяем размер шрифта к дереву данных
                font = tree_widget.font()
                font.setPointSize(self.font_size)
                tree_widget.setFont(font)
                
                # Применяем размер шрифта к заголовкам
                header = tree_widget.header()
                if header:
                    header_font = header.font()
                    header_font.setPointSize(self.header_font_size)
                    header.setFont(header_font)
                    
                    # Обновляем высоту заголовка с учетом нового размера шрифта
                    self._update_tree_header_height(tree_widget)
                
                # Обновляем делегат, если он использует шрифт
                delegate = tree_widget.itemDelegate()
                if delegate:
                    # Делегат будет использовать шрифт из option, который берется из виджета
                    tree_widget.update()
        
        # Обновляем отображение
        QApplication.processEvents()
    
    def toggle_fullscreen(self, checked: bool):
        """Переключить полноэкранный режим"""
        # Проверяем, активна ли вкладка ошибок
        if self.tabs_panel and self.tabs_panel.currentWidget() == self.errors_tab:
            self._toggle_errors_tab_fullscreen()
        else:
            if checked:
                self.showFullScreen()
            else:
                self.showNormal()
    
    def _toggle_errors_tab_fullscreen(self):
        """Переключение полноэкранного режима для вкладки ошибок"""
        if self.errors_tab_fullscreen:
            # Выходим из полноэкранного режима
            self.errors_tab_fullscreen = False
            self.showNormal()
        else:
            # Входим в полноэкранный режим
            self.errors_tab_fullscreen = True
            self.showFullScreen()
    
    def keyPressEvent(self, event):
        """Обработка нажатий клавиш"""
        if event.key() == Qt.Key_F11:
            # Если активна вкладка ошибок, переключаем её полноэкранный режим
            if self.tabs_panel and self.tabs_panel.currentWidget() == self.errors_tab:
                self._toggle_errors_tab_fullscreen()
            else:
                # Иначе переключаем полноэкранный режим главного окна
                self.toggle_fullscreen(not self.isFullScreen())
        else:
            super().keyPressEvent(event)
    
    def on_tree_item_clicked(self, item, column):
        """Обработчик клика по элементу дерева - запоминаем столбец начала выделения"""
        self.selection_start_column = column
        # Также обновляем сумму сразу после клика
        QTimer.singleShot(10, self.on_tree_selection_changed)
    
    def on_tree_selection_changed(self):
        """Обработчик изменения выделения - подсчитываем сумму"""
        selected_items = self.data_tree.selectedItems()
        
        if not selected_items:
            # Если ничего не выбрано, очищаем статус
            self.status_bar.showMessage("Готов к работе")
            return
        
        # Определяем столбец: сначала используем сохраненный, если нет - текущий столбец
        column_index = self.selection_start_column
        if column_index is None:
            # Пытаемся определить столбец из текущего элемента
            current_item = self.data_tree.currentItem()
            if current_item:
                # Используем столбец текущего элемента
                column_index = self.data_tree.currentColumn()
                if column_index < 0:
                    column_index = 0
            else:
                # Если не можем определить, используем первый столбец данных (после базовых)
                mapping = self.tree_column_mapping or {}
                column_type = mapping.get("type", "base")
                if column_type == "budget":
                    column_index = mapping.get("approved_start", 4)
                elif column_type == "consolidated":
                    column_index = mapping.get("value_start", 4)
                else:
                    column_index = 4  # По умолчанию
        
        # Определяем тип столбца и получаем данные
        mapping = self.tree_column_mapping or {}
        column_type = mapping.get("type", "base")
        
        total = 0.0
        count = 0
        column_name = ""
        
        # Определяем название столбца для отображения
        if column_index < len(self.tree_headers):
            column_name = self.tree_headers[column_index]
        
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
            self.status_bar.showMessage(message)
        else:
            self.status_bar.showMessage("Готов к работе")
    
    def show_about(self):
        """Показать информацию о программе"""
        QMessageBox.about(
            self,
            "О программе",
            "<h2>Система обработки бюджетных форм</h2>"
            "<p>Версия 1.0</p>"
            "<p>Приложение для обработки и анализа бюджетных форм, "
            "включая формы 0503317 и другие.</p>"
            "<p><b>Основные возможности:</b></p>"
            "<ul>"
            "<li>Загрузка и парсинг бюджетных форм</li>"
            "<li>Расчет агрегированных сумм</li>"
            "<li>Валидация данных</li>"
            "<li>Работа со справочниками</li>"
            "<li>Экспорт с проверкой</li>"
            "</ul>"
        )
    
    def show_calculation_errors(self):
        """Показать вкладку с ошибками расчетов"""
        if not self.controller.current_project or not self.controller.current_revision_id:
            QMessageBox.warning(self, "Ошибка", "Сначала выберите проект и загрузите ревизию формы")
            return
        
        if not self.controller.current_project.data:
            QMessageBox.warning(self, "Предупреждение", "Нет данных для анализа ошибок")
            return
        
        # Загружаем ошибки из текущих данных проекта
        self.load_errors_to_tab(self.controller.current_project.data)
        
        # Переключаемся на вкладку ошибок
        tabs = self.tabs_panel
        if tabs:
            for i in range(tabs.count()):
                if tabs.tabText(i) == "Ошибки":
                    tabs.setCurrentIndex(i)
                    break
    
    def show_shortcuts(self):
        """Показать список горячих клавиш"""
        shortcuts_text = """
        <h2>Горячие клавиши</h2>
        <table border="1" cellpadding="5">
        <tr><th>Действие</th><th>Клавиша</th></tr>
        <tr><td>Новый проект</td><td><b>Ctrl+N</b></td></tr>
        <tr><td>Загрузить форму</td><td><b>Ctrl+O</b></td></tr>
        <tr><td>Экспорт проверки</td><td><b>Ctrl+E</b></td></tr>
        <tr><td>Выход</td><td><b>Ctrl+Q</b></td></tr>
        <tr><td>Редактировать проект</td><td><b>Ctrl+P</b></td></tr>
        <tr><td>Удалить проект</td><td><b>Ctrl+Delete</b></td></tr>
        <tr><td>Обновить список</td><td><b>F5</b></td></tr>
        <tr><td>Пересчитать суммы</td><td><b>F9</b></td></tr>
        <tr><td>Ошибки расчетов</td><td><b>Ctrl+Shift+E</b></td></tr>
        <tr><td>Скрыть нулевые столбцы</td><td><b>Ctrl+H</b></td></tr>
        <tr><td>Просмотр справочников</td><td><b>Ctrl+R</b></td></tr>
        <tr><td>Справочники конфигурации</td><td><b>Ctrl+D</b></td></tr>
        <tr><td>Панель проектов</td><td><b>Ctrl+1</b></td></tr>
        <tr><td>Полноэкранный режим</td><td><b>F11</b></td></tr>
        </table>
        """
        msg = QMessageBox(self)
        msg.setWindowTitle("Горячие клавиши")
        msg.setText(shortcuts_text)
        msg.setTextFormat(Qt.RichText)
        msg.exec_()
    
    def show_document_dialog(self):
        """Показать диалог формирования документов"""
        if not self.controller.current_project or not self.controller.current_revision_id:
            QMessageBox.warning(self, "Ошибка", "Сначала выберите проект и загрузите ревизию формы")
            return
        
        dialog = DocumentDialog(self)
        dialog.exec_()
    
    def parse_solution_document(self):
        """Обработка решения о бюджете"""
        if not self.controller.current_project:
            QMessageBox.warning(self, "Ошибка", "Сначала выберите проект")
            return
        
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Выберите файл решения о бюджете",
            "",
            "Word Documents (*.docx *.doc);;All Files (*)"
        )
        
        if file_path:
            try:
                result = self.controller.parse_solution_document(file_path)
                if result:
                    QMessageBox.information(
                        self,
                        "Успех",
                        f"Решение обработано:\n"
                        f"Доходов: {len(result.get('приложение1', []))}\n"
                        f"Расходов (общие): {len(result.get('приложение2', []))}\n"
                        f"Расходов (по ГРБС): {len(result.get('приложение3', []))}"
                    )
                else:
                    QMessageBox.warning(self, "Ошибка", "Не удалось обработать решение")
            except Exception as e:
                logger.error(f"Ошибка обработки решения: {e}", exc_info=True)
                QMessageBox.critical(self, "Ошибка", f"Ошибка обработки решения:\n{str(e)}")
    
    def open_file(self, file_path: str):
        """Открыть файл в системе"""
        if not file_path or not os.path.exists(file_path):
            QMessageBox.warning(self, "Ошибка", f"Файл не найден: {file_path}")
            return
        
        try:
            system = platform.system()
            if system == "Windows":
                os.startfile(file_path)
            elif system == "Darwin":  # macOS
                subprocess.run(["open", file_path])
            else:  # Linux
                subprocess.run(["xdg-open", file_path])
            self.status_bar.showMessage(f"Файл открыт: {file_path}")
        except Exception as e:
            logger.error(f"Ошибка открытия файла: {e}", exc_info=True)
            QMessageBox.critical(self, "Ошибка", f"Не удалось открыть файл:\n{str(e)}")
    
    def open_file_dialog(self):
        """Диалог выбора файла для открытия"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Выберите файл для открытия",
            "",
            "Все поддерживаемые файлы (*.doc *.docx *.xls *.xlsx);;"
            "Word Documents (*.doc *.docx);;"
            "Excel Files (*.xls *.xlsx);;"
            "All Files (*.*)"
        )
        
        if file_path:
            self.open_file(file_path)
    
    def open_last_exported_file(self):
        """Открыть последний экспортированный файл"""
        if self.last_exported_file and os.path.exists(self.last_exported_file):
            self.open_file(self.last_exported_file)
        else:
            QMessageBox.warning(
                self,
                "Ошибка",
                "Последний экспортированный файл не найден или не был сохранен"
            )
    
    def show_tab_context_menu(self, position):
        """Контекстное меню для вкладок"""
        # position - это позиция клика относительно QTabWidget
        # Проверяем, что клик был именно на tabBar
        tab_bar = self.tabs_panel.tabBar()
        tab_bar_pos = tab_bar.mapFrom(self.tabs_panel, position)
        tab_index = tab_bar.tabAt(tab_bar_pos)
        
        # Если не нашли вкладку по позиции, пробуем найти по текущей выбранной
        if tab_index < 0:
            tab_index = self.tabs_panel.currentIndex()
            if tab_index < 0:
                return
        
        tab_name = self.tabs_panel.tabText(tab_index)
        if not tab_name:
            return
        
        menu = QMenu(self)
        
        # Проверяем, откреплена ли вкладка
        if tab_name in self.detached_windows:
            attach_action = menu.addAction("Вернуть во вкладки")
            attach_action.setIcon(self.style().standardIcon(QStyle.SP_DialogApplyButton))
            action = menu.exec_(self.tabs_panel.mapToGlobal(position))
            if action == attach_action:
                self.attach_tab(tab_name, None)
        else:
            detach_action = menu.addAction("Открыть в отдельном окне")
            detach_action.setIcon(self.style().standardIcon(QStyle.SP_TitleBarNormalButton))
            action = menu.exec_(self.tabs_panel.mapToGlobal(position))
            if action == detach_action:
                self.detach_tab(tab_index, tab_name)
    
    def detach_tab(self, tab_index, tab_name):
        """Открепление вкладки в отдельное окно"""
        # Получаем виджет вкладки
        tab_widget = self.tabs_panel.widget(tab_index)
        if not tab_widget:
            return
        
        # Сохраняем текущий размер виджета
        widget_size = tab_widget.size()
        
        # Удаляем вкладку из главного окна (но не удаляем сам виджет)
        self.tabs_panel.removeTab(tab_index)
        
        # Убеждаемся, что виджет видим и имеет правильный размер
        tab_widget.setParent(None)
        tab_widget.setVisible(True)
        if widget_size.isValid() and widget_size.width() > 0 and widget_size.height() > 0:
            tab_widget.resize(widget_size)
        
        # Создаем отдельное окно
        detached_window = DetachedTabWindow(tab_widget, tab_name, self)
        self.detached_windows[tab_name] = detached_window
        
        # Показываем окно
        detached_window.show()
        detached_window.raise_()
        detached_window.activateWindow()
    
    def attach_tab(self, tab_name, tab_widget=None):
        """Возврат вкладки в главное окно"""
        logger.debug(f"attach_tab вызван для вкладки '{tab_name}'")
        
        # Проверяем, есть ли эта вкладка в открепленных окнах
        if tab_name not in self.detached_windows:
            # Если вкладка уже не в словаре, возможно она уже была возвращена
            # Проверяем, не находится ли она уже в tabs_panel
            for i in range(self.tabs_panel.count()):
                if self.tabs_panel.tabText(i) == tab_name:
                    logger.debug(f"Вкладка '{tab_name}' уже находится в tabs_panel")
                    return
            logger.warning(f"Вкладка '{tab_name}' не найдена в detached_windows и не найдена в tabs_panel")
            return
        
        detached_window = self.detached_windows[tab_name]
        
        # Получаем виджет из окна (теперь это центральный виджет напрямую)
        if tab_widget is None:
            tab_widget = detached_window.centralWidget()
        
        if not tab_widget:
            logger.error(f"Не удалось получить виджет для вкладки '{tab_name}'")
            # Если виджет не найден, просто удаляем запись
            try:
                detached_window.setProperty("attaching", True)
                detached_window.close()
            except:
                pass
            if tab_name in self.detached_windows:
                del self.detached_windows[tab_name]
            return
        
        logger.debug(f"Виджет для вкладки '{tab_name}' получен: {type(tab_widget).__name__}")
        
        # Сохраняем размер виджета
        widget_size = tab_widget.size()
        logger.debug(f"Размер виджета: {widget_size.width()}x{widget_size.height()}")
        
        # Устанавливаем флаг, чтобы closeEvent не вызывал attach_tab повторно
        detached_window.setProperty("attaching", True)
        
        # Удаляем запись из словаря перед добавлением вкладки обратно
        # Это предотвратит повторные вызовы attach_tab
        if tab_name in self.detached_windows:
            del self.detached_windows[tab_name]
        
        # Определяем позицию вкладки по имени
        tab_positions = {
            "Древовидные данные": 0,
            "Метаданные": 1,
            "Ошибки": 2,
            "Просмотр формы": 3
        }
        position = tab_positions.get(tab_name, self.tabs_panel.count())
        
        logger.debug(f"Добавление вкладки '{tab_name}' в позицию {position}, текущее количество вкладок: {self.tabs_panel.count()}")
        logger.debug(f"Виджет имеет layout: {tab_widget.layout() is not None}")
        logger.debug(f"Виджет имеет родителя: {tab_widget.parent() is not None}, тип родителя: {type(tab_widget.parent()).__name__ if tab_widget.parent() else 'None'}")
        
        # ВАЖНО: Не удаляем виджет из окна до добавления в tabs_panel
        # QTabWidget.insertTab() автоматически установит правильного родителя
        # и удалит виджет из старого родителя
        
        # Убеждаемся, что виджет видим
        tab_widget.setVisible(True)
        
        # Восстанавливаем размер, если он был валидным
        if widget_size.isValid() and widget_size.width() > 0 and widget_size.height() > 0:
            tab_widget.resize(widget_size)
        
        # Добавляем вкладку обратно в главное окно
        # insertTab автоматически установит правильного родителя и удалит из старого
        try:
            inserted_index = self.tabs_panel.insertTab(position, tab_widget, tab_name)
            logger.debug(f"Вкладка вставлена на индекс {inserted_index}, новое количество вкладок: {self.tabs_panel.count()}")
            
            # Проверяем, что вкладка действительно добавлена
            if inserted_index >= 0 and inserted_index < self.tabs_panel.count():
                actual_tab_name = self.tabs_panel.tabText(inserted_index)
                logger.debug(f"Проверка: вкладка на индексе {inserted_index} имеет имя '{actual_tab_name}'")
                
                # Проверяем, что виджет действительно установлен как виджет вкладки
                widget_at_index = self.tabs_panel.widget(inserted_index)
                logger.debug(f"Виджет на индексе {inserted_index}: {type(widget_at_index).__name__ if widget_at_index else 'None'}, совпадает с tab_widget: {widget_at_index == tab_widget}")
                
                # Убеждаемся, что вкладка видна
                self.tabs_panel.setCurrentIndex(inserted_index)
                self.tabs_panel.setTabVisible(inserted_index, True)
                
                # Теперь можно удалить виджет из окна, так как он уже в tabs_panel
                try:
                    detached_window.setCentralWidget(None)
                    logger.debug("Центральный виджет удален из окна после добавления в tabs_panel")
                except Exception as e:
                    logger.warning(f"Ошибка при удалении центрального виджета: {e}")
                
                # Принудительно обновляем отображение
                tab_widget.show()
                tab_widget.update()
                self.tabs_panel.update()
                
                # Принудительно перерисовываем
                QApplication.processEvents()
            else:
                logger.error(f"Ошибка: вкладка не была добавлена правильно. inserted_index={inserted_index}, count={self.tabs_panel.count()}")
        except Exception as e:
            logger.error(f"Ошибка при добавлении вкладки в tabs_panel: {e}", exc_info=True)
        
        # Закрываем окно после того, как вкладка добавлена
        try:
            detached_window.close()
        except Exception as e:
            logger.warning(f"Ошибка при закрытии окна: {e}")
        
        logger.info(f"Вкладка '{tab_name}' успешно возвращена в главное окно на позицию {position}")
"""Кастомные делегаты для виджетов"""
from PyQt5.QtWidgets import QStyledItemDelegate, QStyle
from PyQt5.QtCore import Qt, QSize
from PyQt5.QtGui import (QTextDocument, QTextOption, QTextCharFormat, 
                        QTextCursor, QColor, QBrush, QFont, QPainter)


class WordWrapItemDelegate(QStyledItemDelegate):
    """Делегат для переноса текста в ячейках дерева. Внутри закрашенной ячейки отступ одинаковый для всех уровней."""

    # Одинаковые отступы внутри ячейки для всех уровней; дерево само рисует ветки/иконки
    CELL_LEFT_MARGIN = 4
    CELL_RIGHT_MARGIN = 4
    # Базовый запас: в paint виджет передаёт option.rect уже, чем header.sectionSize(0). На нижних уровнях ещё уже (ветки дерева).
    FIRST_COLUMN_WIDTH_RESERVE = 24

    def __init__(self, tree_widget=None, parent=None):
        super().__init__(parent)
        self._tree_widget = tree_widget

    def _get_item_level(self, index) -> int:
        """Уровень элемента (0–6) для расчёта запаса ширины на нижних уровнях."""
        try:
            model = index.model()
            if model:
                li = model.index(index.row(), 3, index.parent())
                if li.isValid():
                    t = model.data(li, Qt.DisplayRole)
                    if t:
                        return max(0, min(6, int(str(t))))
                parent = index.parent()
                level = 0
                while parent.isValid():
                    level += 1
                    parent = parent.parent()
                return level
        except Exception:
            pass
        return 0

    def _get_column_width(self, column: int, option, widget) -> int:
        """Получение ширины столбца. Для столбца 0 (Наименование) всегда берём из заголовка,
        чтобы при ресайзе высота строки считалась по новой ширине, а не по устаревшему option.rect.
        """
        column_width = 200  # Значение по умолчанию
        
        if widget and hasattr(widget, 'header'):
            header = widget.header()
            if column >= 0:
                column_width = max(header.sectionSize(column), 50)
        
        # Для столбца «Наименование» (0) не подменять шириной из option — при ресайзе rect ещё старый,
        # из-за этого высота не пересчитывается вовремя и текст «съезжает»
        if column != 0 and option.rect.width() > 0:
            column_width = option.rect.width()
        
        return column_width
    
    def _paint_background(self, painter, option, index):
        """Отрисовка фона ячейки
        
        Args:
            painter: Объект для отрисовки
            option: Опции отрисовки
            index: Индекс элемента
        """
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
    
    def _paint_code_column(self, painter, option, index, text: str):
        """Отрисовка колонки "Код классификации" (без переноса)
        
        Args:
            painter: Объект для отрисовки
            option: Опции отрисовки
            index: Индекс элемента
            text: Текст для отрисовки
        """
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
    
    def _paint_text_column(self, painter, option, index, text: str, right_padding: int = 0, left_indent: int = 0):
        """Отрисовка текстовой колонки с переносом. Одинаковые отступы внутри ячейки."""
        doc = QTextDocument()
        doc.setDefaultFont(option.font)
        doc.setPlainText(str(text))
        doc.setDocumentMargin(2)
        text_option = QTextOption()
        text_option.setWrapMode(QTextOption.WrapAtWordBoundaryOrAnywhere)
        doc.setDefaultTextOption(text_option)
        width = max(option.rect.width() - left_indent - right_padding, 20)
        doc.setTextWidth(width)
        
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
        
        text_rect = option.rect.adjusted(left_indent, 0, -right_padding, 0)
        top_pad = 2
        painter.save()
        painter.translate(text_rect.left(), text_rect.top() + top_pad)
        doc.drawContents(painter)
        painter.restore()
    
    def paint(self, painter, option, index):
        if not index.isValid():
            return
        
        # Для эксперимента: проверяем, скрыт ли столбец, и закрашиваем его черным
        column = index.column()
        widget = option.widget
        if widget and hasattr(widget, 'isColumnHidden'):
            if widget.isColumnHidden(column):
                return
        
        # Настраиваем опции отрисовки (нужно сделать до проверки текста)
        option = option.__class__(option)
        self.initStyleOption(option, index)
        
        # Получаем текст из модели
        text = index.data(Qt.DisplayRole) or ""
        
        # Рисуем фон даже если текст пустой (для окраски по уровням)
        self._paint_background(painter, option, index)
        
        # Если текст пустой, только рисуем фон и выходим
        if not text:
            return
        
        # Получаем номер столбца
        column = index.column()
        
        # Для столбца "Код классификации" рисуем текст без переноса
        if column == 2:
            self._paint_code_column(painter, option, index, text)
            return
        
        left_margin = self.CELL_LEFT_MARGIN if column == 0 else 0
        right_margin = self.CELL_RIGHT_MARGIN if column == 0 else 0
        self._paint_text_column(painter, option, index, text, right_margin, left_margin)
    
    def sizeHint(self, option, index):
        if not index.isValid():
            return QSize()
        
        text = index.data(Qt.DisplayRole) or ""
        if not text:
            return QSize(0, option.fontMetrics.height())
        
        column = index.column()
        widget = option.widget or self._tree_widget
        column_width = self._get_column_width(column, option, widget)
        
        left_margin = self.CELL_LEFT_MARGIN if column == 0 else 0
        right_margin = self.CELL_RIGHT_MARGIN if column == 0 else 0
        
        # Для столбца "Код классификации" (индекс 2) используем ширину текста без переноса
        if column == 2:
            # Возвращаем размер текста без переноса
            text_width = option.fontMetrics.horizontalAdvance(str(text))
            return QSize(text_width, option.fontMetrics.height())
        
        # Документ для расчёта размера с переносом. Для столбца 0 ширина = столбец минус одинаковые отступы (текст не выходит за границы ячейки)
        doc = QTextDocument()
        doc.setDefaultFont(option.font)
        doc.setPlainText(str(text))
        doc.setDocumentMargin(2)
        text_option = QTextOption()
        text_option.setWrapMode(QTextOption.WrapAtWordBoundaryOrAnywhere)
        doc.setDefaultTextOption(text_option)
        if column == 0:
            item_level = self._get_item_level(index)
            indentation = (widget.indentation() if widget and hasattr(widget, 'indentation') else None) or 20
            level_reserve = item_level * indentation
            reserve = self.FIRST_COLUMN_WIDTH_RESERVE + level_reserve
            available_width = column_width - left_margin - right_margin - reserve
        else:
            available_width = column_width - left_margin - right_margin
        doc.setTextWidth(max(available_width, 60))
        doc_height = int(doc.size().height())
        fm = option.fontMetrics
        vertical_padding = max(4, fm.lineSpacing() // 2)  # минимальный запас по высоте
        return QSize(int(doc.idealWidth()), doc_height + vertical_padding)

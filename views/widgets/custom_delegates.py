"""Кастомные делегаты для виджетов"""
from PyQt5.QtWidgets import QStyledItemDelegate, QStyle
from PyQt5.QtCore import Qt, QSize
from PyQt5.QtGui import (QTextDocument, QTextOption, QTextCharFormat, 
                        QTextCursor, QColor, QBrush, QFont, QPainter)


class WordWrapItemDelegate(QStyledItemDelegate):
    """Делегат для переноса текста в ячейках дерева"""
    
    def _calculate_item_level(self, index) -> int:
        """Вычисление уровня элемента для внутреннего отступа справа
        
        Args:
            index: Индекс элемента
        
        Returns:
            Уровень элемента (0-6)
        """
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
        
        return item_level
    
    def _calculate_right_padding(self, item_level: int, indentation: int) -> int:
        """Вычисление внутреннего отступа справа с учетом всех уровней
        
        Args:
            item_level: Уровень элемента
            indentation: Отступ дерева
        
        Returns:
            Отступ справа в пикселях
        """
        if item_level > 0:
            # Сумма отступов всех уровней: indentation * (0 + 1 + 2 + ... + item_level)
            # Формула суммы арифметической прогрессии: n * (n + 1) / 2
            return indentation * item_level * (item_level + 1) // 2
        return 0
    
    def _get_column_width(self, column: int, option, widget) -> int:
        """Получение ширины столбца
        
        Args:
            column: Индекс столбца
            option: Опции отрисовки
            widget: Виджет дерева
        
        Returns:
            Ширина столбца в пикселях
        """
        column_width = 200  # Значение по умолчанию
        
        if widget and hasattr(widget, 'header'):
            header = widget.header()
            if column >= 0:
                column_width = max(header.sectionSize(column), 50)
        
        # Если ширина из option доступна, используем её
        if option.rect.width() > 0:
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
    
    def _paint_text_column(self, painter, option, index, text: str, right_padding: int = 0):
        """Отрисовка текстовой колонки с переносом
        
        Args:
            painter: Объект для отрисовки
            option: Опции отрисовки
            index: Индекс элемента
            text: Текст для отрисовки
            right_padding: Отступ справа в пикселях
        """
        # Создаем документ для переноса текста
        doc = QTextDocument()
        doc.setDefaultFont(option.font)
        doc.setPlainText(str(text))
        
        # Настраиваем перенос текста
        text_option = QTextOption()
        text_option.setWrapMode(QTextOption.WrapAtWordBoundaryOrAnywhere)
        doc.setDefaultTextOption(text_option)
        
        # Устанавливаем ширину документа равной ширине ячейки
        width = option.rect.width()
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
    
    def paint(self, painter, option, index):
        if not index.isValid():
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
        
        # Для остальных столбцов используем документ с переносом
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
                
                # Вычисляем уровень элемента и отступ справа
                item_level = self._calculate_item_level(index)
                right_padding = self._calculate_right_padding(item_level, indentation)
            else:
                width = 400  # Значение по умолчанию
        
        # Отрисовываем текстовую колонку с переносом
        self._paint_text_column(painter, option, index, text, right_padding)
    
    def sizeHint(self, option, index):
        if not index.isValid():
            return QSize()
        
        text = index.data(Qt.DisplayRole) or ""
        if not text:
            return QSize(0, option.fontMetrics.height())
        
        # Получаем ширину столбца
        column = index.column()
        column_width = self._get_column_width(column, option, option.widget)
        
        right_padding = 0  # Отступ справа для столбца "Наименование"
        
        # Для столбца "Наименование" (индекс 0) используем ширину с учетом отступов
        if column == 0:
            widget = option.widget
            if widget and hasattr(widget, 'indentation'):
                indentation = widget.indentation()
                indent_reserve = indentation * 6 + 50
                column_width = 400 + indent_reserve
                
                # Вычисляем уровень элемента и отступ справа
                item_level = self._calculate_item_level(index)
                right_padding = self._calculate_right_padding(item_level, indentation)
        
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

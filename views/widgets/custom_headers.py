"""Кастомные заголовки для таблиц"""
from PyQt5.QtWidgets import QHeaderView, QStyleOptionHeader
from PyQt5.QtCore import Qt, QRect
from PyQt5.QtGui import QTextDocument, QTextOption, QPainter


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

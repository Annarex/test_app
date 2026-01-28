"""Кастомные заголовки для таблиц"""
from PyQt5.QtWidgets import QHeaderView, QStyleOptionHeader
from PyQt5.QtCore import Qt, QRect, QTimer
from PyQt5.QtGui import QTextDocument, QTextOption, QPainter


class WrapHeaderView(QHeaderView):
    """Кастомный заголовок с поддержкой переноса текста"""
    
    def __init__(self, orientation=Qt.Horizontal, parent=None):
        super().__init__(orientation, parent)
        self.setTextElideMode(Qt.ElideNone)
        self._header_texts = {}  # Кэш текстов заголовков
    
    def moveSection(self, fromVisualIndex: int, toVisualIndex: int):
        """После перемещения секции принудительно задаём скрытым столбцам ширину 0."""
        super().moveSection(fromVisualIndex, toVisualIndex)
        self._enforce_hidden_sections_zero_width()
        QTimer.singleShot(0, self._enforce_hidden_sections_zero_width)
        QTimer.singleShot(50, self._enforce_hidden_sections_zero_width)
    
    def _enforce_hidden_sections_zero_width(self):
        """Для всех скрытых секций: режим Fixed и ширина 0. Временно разрешаем min=0."""
        old_min = self.minimumSectionSize()
        self.setMinimumSectionSize(0)
        tree = self.parent()
        try:
            for i in range(self.count()):
                if self.isSectionHidden(i):
                    self.setSectionResizeMode(i, QHeaderView.Fixed)
                    self.resizeSection(i, 1)
                    self.resizeSection(i, 0)
            if tree is not None and hasattr(tree, 'setColumnWidth'):
                for i in range(self.count()):
                    if self.isSectionHidden(i):
                        tree.setColumnWidth(i, 1)
                        tree.setColumnWidth(i, 0)
        finally:
            self.setMinimumSectionSize(max(old_min, 1) if old_min > 0 else 1)
        self.updateGeometries()
        if tree is not None:
            if hasattr(tree, 'viewport'):
                tree.viewport().update()
            if hasattr(tree, 'updateGeometry'):
                tree.updateGeometry()
            if hasattr(tree, 'doItemsLayout'):
                tree.doItemsLayout()
    
    def setHeaderTexts(self, texts):
        """Устанавливает тексты заголовков для кэширования"""
        self._header_texts = texts
    
    def sectionSize(self, logicalIndex):
        """Переопределяем размер секции - возвращаем 0 для скрытых столбцов"""
        if self.isSectionHidden(logicalIndex):
            # Принудительно возвращаем 0 для скрытых столбцов
            # Даже если базовый метод возвращает другое значение
            return 0
        size = super().sectionSize(logicalIndex)
        # Дополнительная проверка: если размер > 0, но столбец скрыт, возвращаем 0
        if size > 0 and self.isSectionHidden(logicalIndex):
            return 0
        return size
    
    def sectionPosition(self, logicalIndex):
        """Переопределяем позицию секции - пропускаем скрытые столбцы при вычислении позиций"""
        # Вычисляем позицию, суммируя размеры всех видимых столбцов перед данным индексом
        # Используем self.sectionSize() чтобы учитывать наши переопределения (возврат 0 для скрытых)
        pos = 0
        for i in range(logicalIndex):
            pos += self.sectionSize(i)  # Используем self, а не super, чтобы учитывать переопределения
        return pos
    
    def sectionViewportPosition(self, logicalIndex):
        """Переопределяем позицию секции во viewport - пропускаем скрытые столбцы"""
        # Позиция во viewport = позиция в контенте минус смещение прокрутки
        return self.sectionPosition(logicalIndex) - self.offset()
    
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
        """Возвращаем размер заголовка с учетом переноса текста и только видимых столбцов."""
        size = super().sizeHint()
        # Ширина = сумма только видимых (sectionSize для скрытых у нас 0)
        if self.orientation() == Qt.Horizontal:
            w = sum(self.sectionSize(i) for i in range(self.count()))
            size.setWidth(w)
        
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

"""Помощник для работы с высотой заголовков дерева"""
from PyQt5.QtWidgets import QTreeWidget
from PyQt5.QtGui import QTextDocument, QTextOption
from logger import logger


class TreeHeaderLayoutHelper:
    """Класс для управления высотой заголовков дерева"""
    
    def __init__(self, main_window):
        """
        Args:
            main_window: Ссылка на главное окно для доступа к свойствам
        """
        self.main_window = main_window
    
    def update_header_height(self, tree_widget: QTreeWidget):
        """Обновляет высоту заголовка дерева с учетом автоматического переноса текста
        
        Args:
            tree_widget: Виджет дерева
        """
        try:
            header = tree_widget.header()
            font_metrics = header.fontMetrics()
            max_height = 0
            
            # Получаем заголовки из headerItem
            header_item = tree_widget.headerItem()
            tree_headers = getattr(self.main_window, 'tree_headers', [])
            
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
                        height = self._calculate_header_height(header, text, idx)
                        max_height = max(max_height, height)
            else:
                # Если нет headerItem, используем tree_headers
                for idx, text in enumerate(tree_headers):
                    if text and not tree_widget.isColumnHidden(idx):
                        height = self._calculate_header_height(header, text, idx)
                        max_height = max(max_height, height)
            
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
    
    def _calculate_header_height(self, header, text: str, column_index: int) -> float:
        """Вычисление высоты заголовка для конкретной колонки
        
        Args:
            header: Заголовок дерева
            text: Текст заголовка
            column_index: Индекс колонки
        
        Returns:
            Высота заголовка в пикселях
        """
        # Получаем ширину столбца
        width = max(header.sectionSize(column_index), 50)
        
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
        return doc.size().height()

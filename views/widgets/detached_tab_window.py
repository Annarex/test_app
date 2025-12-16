"""Окно для открепленных вкладок"""
from PyQt5.QtWidgets import QMainWindow, QWidget
from logger import logger


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

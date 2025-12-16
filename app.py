import sys
import os
from pathlib import Path

from PyQt5.QtWidgets import QApplication
from PyQt5.QtCore import QTimer
from PyQt5.QtGui import QFont

from controllers import main_controller
from views.main_window import MainWindow
from logger import logger

def setup_application():
    """Настройка приложения"""
    # Создаем необходимые директории
    data_dir = Path("data")
    
    # Настраиваем приложение
    app = QApplication(sys.argv)
    
    # Устанавливаем шрифт по умолчанию
    font = QFont("Arial", 10)
    app.setFont(font)
    
    # Создаем и показываем главное окно
    main_window = MainWindow()
    main_window.show()
    
    return app, main_window

def main():
    """Главная функция приложения"""
    try:
        app, main_window = setup_application()
        
        # Запускаем приложение
        return app.exec_()
        
    except Exception as e:
        logger.error(f"Ошибка запуска приложения: {e}", exc_info=True)
        return 1

if __name__ == "__main__":
    sys.exit(main())
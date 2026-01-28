import sys
import os
from pathlib import Path

from PyQt5.QtWidgets import QApplication
from PyQt5.QtCore import QTimer
from PyQt5.QtGui import QFont

from controllers import main_controller
from views.main_window import MainWindow
from styles import apply_styles_to_app
from logger import logger

def setup_application():
    """Настройка приложения"""

    app = QApplication(sys.argv)

    font = QFont("Arial", 10)
    app.setFont(font)
    
    if not apply_styles_to_app(app):
        logger.warning("Не удалось загрузить стили приложения")
    
    main_window = MainWindow()
    
    QTimer.singleShot(10, main_window._center_window)
    return app, main_window

def main():
    """Главная функция приложения"""
    try:
        app, main_window = setup_application()
        main_window.show()
        return app.exec_()
        
    except Exception as e:
        logger.error(f"Ошибка запуска приложения: {e}", exc_info=True)
        return 1

if __name__ == "__main__":
    sys.exit(main())
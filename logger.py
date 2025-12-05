"""
Модуль для настройки логирования приложения.
Логи выводятся в консоль и в файл.
"""
import logging
import sys
from pathlib import Path
from datetime import datetime


def setup_logger(name: str = "budget_app", log_file: str = None) -> logging.Logger:
    """
    Настройка logger'а для вывода в консоль и в файл.
    
    Args:
        name: Имя logger'а
        log_file: Путь к файлу лога. Если None, используется logs/app.log
    
    Returns:
        Настроенный logger
    """
    # Создаем директорию для логов, если её нет
    if log_file is None:
        log_dir = Path("logs")
        log_dir.mkdir(exist_ok=True)
        log_file = log_dir / f"app_{datetime.now().strftime('%Y%m%d')}.log"
    else:
        log_file = Path(log_file)
        log_file.parent.mkdir(parents=True, exist_ok=True)
    
    # Создаем logger
    logger = logging.getLogger(name)
    logger.setLevel(logging.DEBUG)
    
    # Удаляем существующие обработчики, чтобы избежать дублирования
    logger.handlers.clear()
    
    # Формат сообщений
    formatter = logging.Formatter(
        '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    
    # Обработчик для консоли
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)
    
    # Обработчик для файла
    file_handler = logging.FileHandler(log_file, encoding='utf-8')
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)
    
    return logger


# Создаем глобальный logger для использования во всем приложении
logger = setup_logger()


"""
Модуль для загрузки и применения стилей из CSS файла
"""
from pathlib import Path
from typing import Optional
from logger import logger


# Кэш для загруженных стилей
_cached_styles: Optional[str] = None


def get_styles_path(theme: str = "styles") -> Path:
    """Возвращает путь к файлу стилей
    
    Args:
        theme: Имя темы (без расширения .qss)
    
    Returns:
        Path к файлу стилей
    """
    return Path(__file__).parent / f"{theme}.qss"


def load_styles(theme: str = "styles", use_cache: bool = True) -> str:
    """Загружает стили из CSS файла
    
    Args:
        theme: Имя темы (без расширения .qss)
        use_cache: Использовать кэш для повторных загрузок
    
    Returns:
        Строка со стилями или пустая строка при ошибке
    """
    global _cached_styles
    
    # Возвращаем из кэша, если доступно
    if use_cache and _cached_styles is not None and theme == "styles":
        return _cached_styles
    
    styles_path = get_styles_path(theme)
    if styles_path.exists():
        try:
            with open(styles_path, 'r', encoding='utf-8') as f:
                styles = f.read()
                # Кэшируем только основную тему
                if theme == "styles" and use_cache:
                    _cached_styles = styles
                return styles
        except Exception as e:
            logger.error(f"Ошибка загрузки стилей из {styles_path}: {e}")
            return ""
    else:
        logger.warning(f"Файл стилей не найден: {styles_path}")
        return ""


def clear_style_cache():
    """Очищает кэш стилей (полезно при переключении тем)"""
    global _cached_styles
    _cached_styles = None


def apply_styles_to_app(app) -> bool:
    """Применяет стили глобально ко всему приложению
    
    Args:
        app: Экземпляр QApplication
    
    Returns:
        True если стили успешно применены, False в противном случае
    """
    styles = load_styles()
    if styles:
        app.setStyleSheet(styles)
        return True
    return False


def set_tab_bar_min_width(tab_bar, min_width: int) -> None:
    """Устанавливает минимальную ширину для вкладок QTabBar
    
    Используется для динамических стилей, когда min-width вычисляется
    на основе ширины текста.
    
    Args:
        tab_bar: Экземпляр QTabBar
        min_width: Минимальная ширина в пикселях
    """
    current_style = tab_bar.styleSheet()
    tab_style = f"QTabBar::tab {{ min-width: {min_width}px; }}"
    tab_bar.setStyleSheet(current_style + "\n" + tab_style)

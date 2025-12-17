"""Утилиты для работы с числовыми значениями"""
from typing import Union


def is_value_different(original: Union[float, int, str, None], 
                      calculated: Union[float, int, str, None], 
                      tolerance: float = 0.00001) -> bool:
    """Проверка различия значений с учетом округления
    
    Безопасно обрабатывает строки, None и спец-значение 'x',
    приводя сравниваемые значения к float там, где это возможно.
    Оба значения округляются до 5 знаков после запятой перед сравнением.
    
    Args:
        original: Оригинальное значение
        calculated: Расчетное значение
        tolerance: Допустимая погрешность (по умолчанию 0.00001)
    
    Returns:
        True, если значения различаются больше чем на tolerance
    """
    try:
        original_val = float(original) if original not in (None, "", "x") else 0.0
        calculated_val = float(calculated) if calculated not in (None, "", "x") else 0.0
        # Округляем до 5 знаков после запятой для точного сравнения
        original_rounded = round(original_val, 5)
        calculated_rounded = round(calculated_val, 5)
        return abs(original_rounded - calculated_rounded) > tolerance
    except (TypeError, ValueError):
        # Если значения некорректные и не приводятся к числу, считаем их равными
        return False


def format_numeric_value(value: Union[float, int, str, None]) -> str:
    """Форматирование числового значения для отображения
    
    Args:
        value: Значение для форматирования
    
    Returns:
        Отформатированная строка
    """
    if value in (None, "", "x"):
        return str(value) if value else ""
    try:
        return f"{float(value):,.2f}"
    except (ValueError, TypeError):
        return str(value)


def safe_float(value: Union[float, int, str, None], default: float = 0.0) -> float:
    """Безопасное преобразование значения в float
    
    Args:
        value: Значение для преобразования
        default: Значение по умолчанию, если преобразование невозможно
    
    Returns:
        Преобразованное значение или default
    """
    if value in (None, "", "x"):
        return default
    try:
        return float(value)
    except (ValueError, TypeError):
        return default


def calculate_error_difference(original: Union[float, int, str, None], 
                              calculated: Union[float, int, str, None]) -> float:
    """Вычисление разницы между значениями
    
    Args:
        original: Оригинальное значение
        calculated: Расчетное значение
    
    Returns:
        Разница (calculated - original)
    """
    try:
        original_val = safe_float(original, 0.0)
        calculated_val = safe_float(calculated, 0.0)
        return calculated_val - original_val
    except (ValueError, TypeError):
        return 0.0

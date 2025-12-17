"""Утилиты для работы с формами"""
import pandas as pd
import re
from typing import Union


def column_to_index(column_letter: str) -> int:
    """Конвертация буквы колонки в индекс
    
    Args:
        column_letter: Буква колонки (например, 'A', 'AA', 'AB')
    
    Returns:
        Индекс колонки (начиная с 0)
    """
    column_letter = column_letter.upper()
    index = 0
    for char in column_letter:
        index = index * 26 + (ord(char) - ord('A') + 1)
    return index - 1


def get_cell_value(sheet: pd.DataFrame, row: int, col: int) -> str:
    """Получение значения ячейки из DataFrame
    
    Args:
        sheet: DataFrame с данными листа Excel
        row: Номер строки (начиная с 0)
        col: Номер колонки (начиная с 0)
    
    Returns:
        Строковое значение ячейки или пустая строка
    """
    if row < len(sheet) and col < len(sheet.columns):
        value = sheet.iloc[row, col]
        return str(value) if pd.notna(value) else ""
    return ""


def get_numeric_value(sheet: pd.DataFrame, row: int, col: int) -> Union[float, str]:
    """Получение числового значения из DataFrame
    
    Args:
        sheet: DataFrame с данными листа Excel
        row: Номер строки (начиная с 0)
        col: Номер колонки (начиная с 0)
    
    Returns:
        Числовое значение или 'x' для специального значения, или 0.0
    """
    if row < len(sheet) and col < len(sheet.columns):
        value = sheet.iloc[row, col]
        if pd.notna(value):
            try:
                if str(value).lower() == 'x':
                    return 'x'
                return float(value)
            except (ValueError, TypeError):
                return 0.0
    return 0.0


def clean_dbk_code(code: str) -> str:
    """Очистка кода классификации
    
    Args:
        code: Код классификации (может содержать пробелы)
    
    Returns:
        Очищенный код (20 символов, дополненный нулями)
    """
    if pd.isna(code) or not isinstance(code, str):
        return ""
    
    clean_code = code.replace(' ', '')
    
    if len(clean_code) < 20 and clean_code.isdigit():
        clean_code = clean_code.zfill(20)
    
    return clean_code


def format_classification_code(code: str, section_type: str) -> str:
    """Форматирование кода классификации с пробелами
    
    Args:
        code: Код классификации (20 символов)
        section_type: Тип раздела ('доходы', 'расходы', 'источники_финансирования')
    
    Returns:
        Отформатированный код с пробелами
    """
    if len(code) != 20:
        return code
        
    if section_type == 'доходы':
        return re.sub(r'(\d{3})(\d{1})(\d{2})(\d{5})(\d{2})(\d{4})(\d{3})', r'\1 \2 \3 \4 \5 \6 \7', code)
    elif section_type == 'расходы':
        return re.sub(r'(\d{3})(\d{4})(\d{10})(\d{3})', r'\1 \2 \3 \4', code)
    elif section_type == 'источники_финансирования':
        return re.sub(r'(\d{3})(\d{2})(\d{2})(\d{2})(\d{2})(\d{2})(\d{4})(\d{3})', r'\1 \2 \3 \4 \5 \6 \7 \8', code)
    return code


def extract_original_value_from_cell(cell_value) -> float:
    """Извлечение оригинального значения из ячейки Excel
    
    Обрабатывает случаи, когда в ячейке уже есть формат "orig (calc)" 
    из предыдущего экспорта. Извлекает первое число (оригинальное значение).
    
    Args:
        cell_value: Значение ячейки Excel (может быть числом или строкой)
    
    Returns:
        Оригинальное числовое значение
    """
    if cell_value is None:
        return 0.0
    
    # Если значение уже число, возвращаем его
    if isinstance(cell_value, (int, float)):
        return float(cell_value)
    
    # Если значение строка, пытаемся извлечь число
    if isinstance(cell_value, str):
        # Проверяем формат "orig (calc)" - извлекаем первое число
        match = re.match(r'^([-+]?\d+\.?\d*)\s*\(', cell_value)
        if match:
            try:
                return float(match.group(1))
            except (ValueError, TypeError):
                pass
        
        # Если не формат "orig (calc)", пытаемся преобразовать всю строку в число
        try:
            return float(cell_value)
        except (ValueError, TypeError):
            pass
    
    return 0.0

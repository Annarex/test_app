"""
Модуль для обработки и сравнения кодов.
Содержит функции для нормализации кодов и их сравнения.
"""
from Levenshtein import distance as levenshtein_distance
from typing import Optional


def normalize_classification_code(code: str) -> str:
    """
    Нормализует код классификации для сравнения.
    Если длина кода равна 20, убирает первые 3 символа.
    
    Args:
        code: Исходный код классификации
    
    Returns:
        Нормализованный код
    """
    if not code:
        return ""
    
    # Удаляем все пробелы и обрезаем пробелы по краям
    code = str(code).strip().replace(' ', '')
    # Если длина кода равна 20, убираем первые 3 символа
    if len(code) == 20:
        return code[3:]
    return code


def compare_codes(
    orig_code: str,
    reference_codes: list[str],
    max_distance: int = 25
) -> tuple[bool, str, int, list[int], list[tuple[int, int, str, str]], Optional[int]]:
    """
    Сравнивает код классификации с кодами из справочника.
    
    Args:
        orig_code: Исходный код классификации
        reference_codes: Список кодов из справочника
        max_distance: Максимальное допустимое расстояние Левенштейна
    
    Returns:
        Кортеж из 6 элементов:
        - есть ли ошибка (bool): True если есть ошибка
        - эталонный код из справочника (str): пустая строка если не найден
        - расстояние Левенштейна (int)
        - список индексов позиций с ошибками (list[int]): для подсветки
        - список исправлений (list[tuple]): (start, end, correct_text, kind)
        - индекс найденного кода в справочнике (Optional[int]): None если не найден
    """
    from .text_comparator import find_differences
    
    # Нормализуем исходный код
    normalized_orig = normalize_classification_code(orig_code)
    
    if not normalized_orig:
        return False, "", 0, [], [], None
    
    # Ищем лучшее совпадение
    best_idx: Optional[int] = None
    best_ref = ""
    best_dist = max_distance + 1
    
    for idx_ref, ref_code in enumerate(reference_codes):
        # Удаляем все пробелы из кода справочника перед сравнением
        ref_code_clean = str(ref_code).strip().replace(' ', '')
        dist = levenshtein_distance(normalized_orig, ref_code_clean)
        if dist <= max_distance and dist < best_dist:
            best_dist = dist
            best_idx = idx_ref
            best_ref = ref_code_clean
            if dist == 0:
                break
    
    # Если совпадение плохое или не найдено
    if best_dist > max_distance or best_idx is None:
        return True, "", best_dist, list(range(len(normalized_orig))), [], None
    
    # Если точное совпадение - нет ошибки
    if best_dist == 0:
        return False, best_ref, 0, [], [], best_idx
    
    # Находим различия для подсветки
    diff_idx, corrections = find_differences(normalized_orig, best_ref)
    
    # Есть ошибка, если есть различия
    has_error = len(diff_idx) > 0 or any(kind == "insert" for _, _, _, kind in corrections)
    
    return has_error, best_ref, best_dist, diff_idx, corrections, best_idx


"""
Модуль для поиска ошибок в данных.
Содержит функции для сравнения исходных данных со справочником и поиска различий.
"""
from Levenshtein import distance as levenshtein_distance
from typing import List, Tuple, Optional, NamedTuple
from .text_comparator import find_differences
from .code_processor import compare_codes


class CodeErrorInfo(NamedTuple):
    """Информация об ошибке в коде классификации."""
    has_error: bool
    ref_code: str
    distance: int
    diff_indices: list[int]
    corrections: list[tuple[int, int, str, str]]
    code_ref_idx: Optional[int]


class ErrorInfo(NamedTuple):
    """Структурированная информация об ошибке в наименовании и/или коде."""
    original_text: str  # Исходное наименование
    reference_text: str  # Эталонное наименование из справочника
    distance: int  # Расстояние Левенштейна по наименованию
    diff_indices: list[int]  # Индексы позиций с ошибками
    corrections: list[tuple[int, int, str, str]]  # Исправления
    original_index: int  # Индекс в исходных данных
    reference_index: Optional[int]  # Индекс в справочнике (None если не найден)
    code_error: Optional[CodeErrorInfo]  # Информация об ошибке в коде (None если не проверялся)


def find_best_match(
    orig: str,
    reference_codes: List[str],
    max_distance: int = 25
) -> Tuple[Optional[int], str, int]:
    """
    Находит лучшее совпадение для исходной строки в справочнике.
    Сначала проверяет, есть ли строки, которые начинаются с искомого текста.
    
    Args:
        orig: Исходная строка для поиска
        reference_codes: Список строк справочника
        max_distance: Максимальное допустимое расстояние Левенштейна
    
    Returns:
        Кортеж из:
        - индекс лучшего совпадения в справочнике (None если не найдено)
        - текст лучшего совпадения (пустая строка если не найдено)
        - расстояние Левенштейна
    """
    best_idx = None
    best_ref = ""
    best_dist = max_distance + 1
    
    # Сначала проверяем, есть ли строки, которые начинаются с искомого текста
    # Это быстрее, чем расчет Левенштейна
    starts_with_matches = []
    for idx_reference, reference_code in enumerate(reference_codes):
        if reference_code.startswith(orig):
            starts_with_matches.append((idx_reference, reference_code))
    
    # Если нашли совпадения "начинается с", выбираем первое (или можно выбрать самое короткое)
    if starts_with_matches:
        # Выбираем первое совпадение (или можно выбрать самое короткое для лучшего совпадения)
        best_idx, best_ref = starts_with_matches[0]
        # Расстояние считаем как разницу в длине (сколько символов лишних)
        best_dist = len(best_ref) - len(orig)
        # Если точное совпадение, расстояние 0
        if best_dist == 0:
            return best_idx, best_ref, 0
        # Если разница в длине в пределах max_distance, возвращаем это совпадение
        if best_dist <= max_distance:
            return best_idx, best_ref, best_dist
    
    # Если не нашли совпадений "начинается с", ищем по Левенштейну
    for idx_reference, reference_code in enumerate(reference_codes):
        dist = levenshtein_distance(orig, reference_code)
        # Если расстояние лучше текущего лучшего
        if dist < best_dist:
            best_dist = dist
            best_idx = idx_reference
            best_ref = reference_code
            # Если нашли идеальное совпадение (dist == 0), можно прекратить поиск
            if dist == 0:
                break
    
    # Если не нашли ничего в пределах max_distance, все равно возвращаем лучшее из всех
    return best_idx, best_ref, best_dist


def find_errors(
    orig_names: List[str],
    reference_names: List[str],
    orig_codes: List[str] = None,
    reference_codes: List[str] = None,
    max_distance: int = 25,
    progress_callback = None
) -> List[ErrorInfo]:
    """
    Находит все ошибки в исходных данных по сравнению со справочником.
    Сравнивает как наименования, так и коды (если указаны).
    
    Логика проверки:
    1. Ищем лучшее совпадение по наименованию
    2. Сразу проверяем код найденного варианта
    3. Если коды не равны, пытаемся найти другой вариант с таким же наименованием, но с совпадающим кодом
    4. Если не найден, возвращаемся к максимально похожему варианту
    
    Args:
        orig_names: Список исходных наименований
        reference_names: Список наименований справочника
        orig_codes: Список исходных кодов (опционально)
        reference_codes: Список кодов справочника (опционально)
        max_distance: Максимальное допустимое расстояние Левенштейна
        progress_callback: Функция для вывода прогресса (current, total)
    
    Returns:
        Список объектов ErrorInfo с информацией об ошибках
    """
    from .code_processor import normalize_classification_code
    
    # Создаем индекс для быстрого поиска по наименованиям (name -> list of indices)
    name_to_indices = {}
    for idx, name in enumerate(reference_names):
        if name not in name_to_indices:
            name_to_indices[name] = []
        name_to_indices[name].append(idx)
    
    errors = []
    total = len(orig_names)
    
    for idx_orig, orig in enumerate(orig_names):
        # Вывод прогресса
        if progress_callback:
            progress_callback(idx_orig + 1, total)
        
        orig_code = orig_codes[idx_orig] if orig_codes and idx_orig < len(orig_codes) else ""
        
        # Нормализуем код сразу для дальнейшей работы
        from .code_processor import normalize_classification_code
        normalized_orig_code = normalize_classification_code(orig_code) if orig_code else ""
        skip_code_comparison = not normalized_orig_code  # Флаг: пропускать сравнение кодов
        
        # Шаг 1: Сначала проверяем точное совпадение (быстрее)
        best_idx_reference = None
        best_ref = ""
        best_dist = max_distance + 1
        
        if orig in name_to_indices:
            # Берем первый индекс с таким наименованием
            if name_to_indices[orig]:
                best_idx_reference = name_to_indices[orig][0]
                best_ref = orig
                best_dist = 0
        
        # Если точного совпадения нет, ищем лучшее совпадение по наименованию
        if best_dist > 0:
            best_idx_in_active, best_ref, best_dist = find_best_match(orig, reference_names, max_distance)
            if best_idx_in_active is not None:
                best_idx_reference = best_idx_in_active
        
        # Если совпадение вообще не найдено (best_idx_reference == None), добавляем с пустым эталоном
        if best_idx_reference is None:
            ref_show = ""
            diff_idx: list[int] = list(range(len(orig)))  # подсветить всю строку
            final_idx_reference = None  # нет соответствующей строки в справочнике
            corrections: list[tuple[int, int, str, str]] = []  # нет осмысленных исправлений
            code_error_info = None
            
            # Если есть коды И НЕ установлен флаг пропуска (код не 'x'), проверяем их отдельно
            if not skip_code_comparison and normalized_orig_code and reference_codes:
                has_code_error, ref_code, code_dist, code_diff_idx, code_corrections, code_ref_idx = compare_codes(
                    normalized_orig_code, reference_codes, max_distance
                )
                if has_code_error:
                    # Если несовпадение кода больше 50% символов, не выводим исправления
                    from .code_processor import normalize_classification_code
                    normalized_orig_code = normalize_classification_code(orig_code)
                    code_max_len = max(len(normalized_orig_code), len(ref_code)) if ref_code else len(normalized_orig_code)
                    code_mismatch_percent = (code_dist / code_max_len * 100) if code_max_len > 0 else 0
                    
                    if code_mismatch_percent > 50:
                        code_corrections = []
                    
                    code_error_info = CodeErrorInfo(
                        has_error=has_code_error,
                        ref_code=ref_code,
                        distance=code_dist,
                        diff_indices=code_diff_idx,
                        corrections=code_corrections,
                        code_ref_idx=code_ref_idx
                    )
            
            errors.append(ErrorInfo(
                original_text=orig,
                reference_text=ref_show,
                distance=best_dist,
                diff_indices=diff_idx,
                corrections=corrections,
                original_index=idx_orig,
                reference_index=final_idx_reference,
                code_error=code_error_info
            ))
            continue
        
        # Шаг 2: Проверяем код найденного варианта (если коды указаны)
        final_idx_reference = best_idx_reference
        ref_show = best_ref
        code_error_info = None
        exact_match_found = False  # Флаг точного совпадения (наименование + код)
        
        # Если код был 'x' (skip_code_comparison=True), проверяем только наименование
        if skip_code_comparison:
            # Код 'x' - игнорируем коды, проверяем только наименование
            if best_dist == 0 and best_idx_reference is not None:
                exact_match_found = True
        elif best_idx_reference is not None and normalized_orig_code and reference_codes:
            # Проверяем совпадение кода только если код НЕ 'x'
            if normalized_orig_code:
                # Получаем код найденного варианта
                ref_code_at_idx = reference_codes[best_idx_reference] if best_idx_reference < len(reference_codes) else ""
                # Нормализуем код из справочника для корректного сравнения
                normalized_ref_code_at_idx = normalize_classification_code(ref_code_at_idx)
                
                # Сравниваем коды
                if normalized_orig_code == normalized_ref_code_at_idx:
                    # Точное совпадение по коду и наименованию
                    if best_dist == 0:
                        exact_match_found = True
                else:
                    # Коды не совпадают - ищем среди всех вариантов с таким же наименованием
                    # Используем предварительно созданный индекс для быстрого поиска
                    matching_indices = name_to_indices.get(best_ref, [])
                    
                    # Ищем среди вариантов с таким же наименованием тот, у которого код совпадает
                    found_match = False
                    for match_idx in matching_indices:
                        if match_idx < len(reference_codes):
                            ref_code_candidate = reference_codes[match_idx]
                            # Нормализуем код из справочника для корректного сравнения
                            normalized_ref_code = normalize_classification_code(ref_code_candidate)
                            if normalized_orig_code == normalized_ref_code:
                                # Нашли вариант с совпадающим кодом
                                final_idx_reference = match_idx
                                found_match = True
                                if best_dist == 0:  # И наименование точно совпадает
                                    exact_match_found = True
                                break
                    
                    # Если не нашли совпадение по коду среди вариантов с таким же наименованием,
                    # проверяем ВСЕ варианты справочника на полное совпадение по коду
                    if not found_match:
                        for check_idx in range(len(reference_codes)):
                            ref_code_candidate = reference_codes[check_idx]
                            # Нормализуем код из справочника для корректного сравнения
                            normalized_ref_code = normalize_classification_code(ref_code_candidate)
                            if normalized_orig_code == normalized_ref_code:
                                # Нашли полное совпадение по коду (но с другим текстом)
                                # Используем этот вариант вместо ближайшего по тексту
                                final_idx_reference = check_idx
                                best_idx_reference = check_idx
                                ref_show = reference_names[check_idx] if check_idx < len(reference_names) else ""
                                best_ref = ref_show
                                # Пересчитываем distance для нового текста
                                best_dist = levenshtein_distance(orig, ref_show)
                                found_match = True
                                # code_error_info НЕ создаем, т.к. код совпадает идеально
                                # Но ошибка по тексту будет показана ниже, если текст отличается
                                break
                    
                    # Если не нашли совпадение по коду, используем исходный вариант (максимально похожий по наименованию)
                    # и создаем информацию об ошибке в коде
                    if not found_match:
                        # Сравниваем нормализованный исходный код с нормализованным кодом из справочника для подсветки
                        code_diff_idx, code_corrections = find_differences(normalized_orig_code, normalized_ref_code_at_idx)
                        has_code_error = len(code_diff_idx) > 0 or any(kind == "insert" for _, _, _, kind in code_corrections)
                        
                        if has_code_error:
                            # Вычисляем расстояние Левенштейна для информации
                            code_dist = levenshtein_distance(normalized_orig_code, normalized_ref_code_at_idx)
                            
                            # Если несовпадение больше 50% символов, не выводим исправления для кода
                            code_max_len = max(len(normalized_orig_code), len(normalized_ref_code_at_idx)) if normalized_ref_code_at_idx else len(normalized_orig_code)
                            code_mismatch_percent = (code_dist / code_max_len * 100) if code_max_len > 0 else 0
                            
                            if code_mismatch_percent > 50:
                                # Очищаем исправления для кода
                                code_corrections = []
                            
                            code_error_info = CodeErrorInfo(
                                has_error=has_code_error,
                                ref_code=normalized_ref_code_at_idx,
                                distance=code_dist,
                                diff_indices=code_diff_idx,
                                corrections=code_corrections,
                                code_ref_idx=best_idx_reference
                            )
        
        # Шаг 3: Проверяем различия в наименовании
        diff_idx, corrections = find_differences(orig, ref_show)
        
        # Если несовпадение больше 50% символов, не выводим исправления
        # Считаем процент несовпадения от максимальной длины строк
        max_len = max(len(orig), len(ref_show)) if ref_show else len(orig)
        mismatch_percent = (best_dist / max_len * 100) if max_len > 0 else 0
        
        if mismatch_percent > 50:
            # Очищаем исправления, оставляем только индексы ошибок
            corrections = []
        
        has_insert_only = any(kind == "insert" for _, _, _, kind in corrections)
        has_name_error = len(diff_idx) > 0 or has_insert_only
        
        # Если нет ошибок ни в наименовании, ни в коде - пропускаем
        if not has_name_error and not code_error_info:
            continue
        
        # Сохраняем информацию об ошибке
        errors.append(ErrorInfo(
            original_text=orig,
            reference_text=ref_show,
            distance=best_dist,
            diff_indices=diff_idx,
            corrections=corrections,
            original_index=idx_orig,
            reference_index=final_idx_reference,
            code_error=code_error_info
        ))
    
    return errors


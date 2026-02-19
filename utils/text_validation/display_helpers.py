"""Вспомогательные функции для отображения данных валидации текстов"""
import sqlite3
import pandas as pd
from typing import List, Tuple, Callable, Optional
from logger import logger


def richtext_to_html(orig_text: str, diff_idx: list, corrections: list, invert_colors: bool = False) -> str:
    """
    Преобразование подсвеченного текста в HTML для отображения в Qt таблице
    
    Args:
        orig_text: Исходный текст
        diff_idx: Список индексов позиций с ошибками
        corrections: Список исправлений (start, end, correct_text, kind)
        invert_colors: Если True, отличия подсвечиваются зеленым (для альтернатив)
    
    Returns:
        HTML строка с подсветкой различий
    """
    if not orig_text:
        return ""
    
    # Функция для экранирования HTML символов
    def escape_html(text):
        return (text.replace('&', '&amp;')
                   .replace('<', '&lt;')
                   .replace('>', '&gt;')
                   .replace('"', '&quot;')
                   .replace("'", '&#39;'))
    
    # Разделяем исправления на:
    # - corrections_map: для replace/delete (привязаны к диапазону ошибок)
    # - insert_map: для insert (отсутствует в orig, есть только в ref)
    corrections_map = {}
    insert_map = {}
    for start, end, correct_text, kind in corrections:
        if kind in ("replace", "delete"):
            # Привязываем правильный текст к началу ошибочного диапазона
            corrections_map[start] = correct_text
        elif kind == "insert":
            # Вставка: позиция между символами 0..len(orig)
            if start not in insert_map:
                insert_map[start] = []
            insert_map[start].append(correct_text)
    
    html_parts = []
    diff_idx_set = set(diff_idx)
    
    # Сначала обрабатываем все insert в начале строки
    if 0 in insert_map:
        for txt in insert_map[0]:
            if invert_colors:
                # Для альтернатив: insert = удаленное из оригинала (красным перечеркнутым)
                html_parts.append(f'<span style="color:red;text-decoration:line-through">{escape_html(txt)}</span>')
            else:
                # Для основной таблицы: insert = что нужно добавить (зеленым)
                html_parts.append(f'<span style="color:green;font-weight:bold;text-decoration:underline">{escape_html(txt)}</span>')
    
    i = 0
    while i < len(orig_text):
        if i in diff_idx_set:
            # Нашли ошибку (replace/delete): показываем правильный текст (зеленым),
            # затем неправильный фрагмент (красным).
            error_start = i
            while i < len(orig_text) and i in diff_idx_set:
                i += 1
            error_end = i
            error_text = orig_text[error_start:error_end]
            
            # Цвета для подсветки
            if invert_colors:
                # Для альтернатив: отличия в альтернативе = правильные варианты (зеленый)
                error_color = "green"
                correct_color = "red"  # То что было в оригинале (красным перечеркнутым)
                correct_decoration = "line-through"
            else:
                # Для основной таблицы: ошибки = красный, правильное = зеленый
                error_color = "red"
                correct_color = "green"
                correct_decoration = "underline"
            
            # Показываем текст из corrections (что в оригинале/эталоне)
            if error_start in corrections_map:
                correct_text = corrections_map[error_start]
                if correct_text:
                    if invert_colors:
                        # Для альтернатив: показываем что было в оригинале (красным перечеркнутым)
                        html_parts.append(f'<span style="color:{correct_color};text-decoration:{correct_decoration}">{escape_html(correct_text)}</span>')
                    else:
                        # Для основной таблицы: показываем правильный вариант (зеленым)
                        html_parts.append(f'<span style="color:{correct_color};font-weight:bold;text-decoration:{correct_decoration}">{escape_html(correct_text)}</span>')
            
            # Отличающийся текст подсвечиваем
            html_parts.append(f'<span style="color:{error_color};font-weight:bold;text-decoration:underline">{escape_html(error_text)}</span>')
        else:
            # Правильный текст - обычным цветом
            start_ok = i
            while i < len(orig_text) and i not in diff_idx_set:
                # Проверяем, есть ли insert после текущего символа
                next_pos = i + 1
                if next_pos in insert_map:
                    # Добавляем текущий символ
                    if start_ok <= i:
                        text_ok = orig_text[start_ok:i+1]
                        if text_ok:
                            html_parts.append(escape_html(text_ok))
                    # Добавляем insert после этого символа
                    for txt in insert_map[next_pos]:
                        if invert_colors:
                            # Для альтернатив: insert = удаленное из оригинала (красным перечеркнутым)
                            html_parts.append(f'<span style="color:red;text-decoration:line-through">{escape_html(txt)}</span>')
                        else:
                            # Для основной таблицы: insert = что нужно добавить
                            html_parts.append(f'<span style="color:green;font-weight:bold;text-decoration:underline">{escape_html(txt)}</span>')
                    start_ok = i + 1
                i += 1
            text_ok = orig_text[start_ok:i]
            if text_ok:
                html_parts.append(escape_html(text_ok))
    
    # Обрабатываем вставки в самом конце строки
    if len(orig_text) in insert_map:
        for txt in insert_map[len(orig_text)]:
            if invert_colors:
                # Для альтернатив: insert = удаленное из оригинала (красным перечеркнутым)
                html_parts.append(f'<span style="color:red;text-decoration:line-through">{escape_html(txt)}</span>')
            else:
                # Для основной таблицы: insert = что нужно добавить
                html_parts.append(f'<span style="color:green;font-weight:bold;text-decoration:underline">{escape_html(txt)}</span>')
    
    return ''.join(html_parts)


def find_alternative_variants(
    original_text: str,
    db_path: str,
    table_name: str,
    filter_date: str,
    ppocode_filter,
    code_extractor: Callable[[pd.DataFrame], List[str]],
    deduplicate: bool = True,
    top_n: int = 5
) -> List[Tuple[int, str, str, int, str, str, str]]:
    """
    Поиск альтернативных вариантов из справочника по Levenshtein distance
    
    Args:
        original_text: Исходный текст для поиска
        db_path: Путь к базе данных
        table_name: Имя таблицы справочника
        filter_date: Дата для фильтрации
        ppocode_filter: Фильтр по ОКТМО (строка или список)
        code_extractor: Функция для извлечения кодов из DataFrame
        deduplicate: Если True, загружает справочник с дедупликацией
        top_n: Количество лучших вариантов для возврата
    
    Returns:
        List[(distance, text, code, ref_index, startdate, ppocode, pponame)]
    """
    from Levenshtein import distance as levenshtein_distance
    from utils.db_utils import get_filtered_view
    
    try:
        # Загружаем справочник
        conn = sqlite3.connect(db_path)
        reference_df = get_filtered_view(
            conn, table_name,
            filter_date=filter_date,
            filter_ppocode=ppocode_filter,
            deduplicate=deduplicate
        )
        conn.close()
        
        if reference_df.empty:
            return []
        
        reference_names = reference_df['name'].astype(str).str.strip().tolist()
        reference_codes = code_extractor(reference_df)
        reference_data = reference_df.to_dict('records')
        
    except Exception as e:
        logger.error(f"Ошибка загрузки альтернатив: {e}", exc_info=True)
        return []
    
    # Вычисляем distance для всех записей
    variants = []
    for idx, ref_name in enumerate(reference_names):
        dist = levenshtein_distance(original_text, ref_name)
        ref_code = reference_codes[idx] if idx < len(reference_codes) else ""
        startdate = reference_data[idx].get('startdate', '') if idx < len(reference_data) else ""
        ppocode = reference_data[idx].get('ppocode', '') if idx < len(reference_data) else ""
        pponame = reference_data[idx].get('pponame', '') if idx < len(reference_data) else ""
        variants.append((dist, ref_name, ref_code, idx, startdate, ppocode, pponame))
    
    # Сортируем по distance и берем топ-N
    variants.sort(key=lambda x: x[0])
    return variants[:top_n]

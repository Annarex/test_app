"""
Утилиты для определения уровня кодов классификации доходов и расходов.
Логика перенесена из Osnova/app_budgetclastypeinc_merged.py для интеграции в проект.
"""
import pandas as pd
from typing import Optional, Tuple, List


def find_nearest_code(
    target_code: str, 
    candidate_codes: List[str], 
    candidate_indices: List[int], 
    candidate_levels: Optional[List[int]] = None
) -> Tuple[int, str]:
    """
    Находит ближайший код к целевому по минимальной абсолютной разнице.
    Код-кандидат должен быть численно меньше целевого кода.
    При одинаковых кодах выбирает запись с level != 1.
    
    Args:
        target_code: Целевой код для поиска
        candidate_codes: Список кодов-кандидатов
        candidate_indices: Список индексов кандидатов в исходном DataFrame
        candidate_levels: Список уровней кандидатов (опционально)
        
    Returns:
        Кортеж (индекс ближайшего кандидата, код) или (-1, '') если не найден
    """
    if not candidate_codes:
        return (-1, '')
    
    try:
        target_num = int(target_code)
    except ValueError:
        return (-1, '')
    
    candidates_with_min_diff = []
    for i, (code, orig_idx) in enumerate(zip(candidate_codes, candidate_indices)):
        try:
            candidate_num = int(code)
            if candidate_num < target_num:
                diff = abs(target_num - candidate_num)
                level = candidate_levels[i] if candidate_levels and i < len(candidate_levels) else None
                candidates_with_min_diff.append((diff, orig_idx, code, level))
        except ValueError:
            continue
    
    if not candidates_with_min_diff:
        return (-1, '')
    
    min_diff = min(c[0] for c in candidates_with_min_diff)
    min_diff_candidates = [c for c in candidates_with_min_diff if c[0] == min_diff]
    
    if len(min_diff_candidates) > 1 and candidate_levels:
        candidates_not_level_1 = [c for c in min_diff_candidates if c[3] is not None and c[3] != 1]
        if candidates_not_level_1:
            _, best_idx, best_code, _ = candidates_not_level_1[0]
            return (best_idx, best_code)
    
    _, best_idx, best_code, _ = min_diff_candidates[0]
    return (best_idx, best_code)


def recalculate_levels_for_income_codes(
    df: pd.DataFrame,
    base_ppocode: str = '00000000'
) -> pd.DataFrame:
    """
    Переопределяет уровни для записей доходов на основе базовых записей.
    
    Логика:
    - Базовые записи (ppocode == base_ppocode) не переопределяются
    - Для записей с level == 1 ищется ближайшая базовая запись по concatenated_code
    - Новый уровень = уровень базовой записи + 1
    - Если базовая запись не найдена, устанавливается уровень 0
    
    Args:
        df: DataFrame с данными из VIEW v_budgetclastypeinc_merged
             Должен содержать колонки: 'ppocode', 'concatenated_code', 'level', 'id'
        base_ppocode: Код базового участника БП (по умолчанию '00000000' для ФУ)
        
    Returns:
        DataFrame с обновленными уровнями
    """
    df = df.copy()
    df['level'] = pd.to_numeric(df['level'], errors='coerce')
    df = df.sort_values('concatenated_code', ascending=True, ignore_index=True)
    
    base_mask = df['ppocode'] == base_ppocode
    base_rows = df[base_mask].copy()
    base_indices = base_rows.index.tolist()
    base_codes = base_rows['concatenated_code'].tolist()
    base_levels = base_rows['level'].tolist() if 'level' in base_rows.columns else None
    
    # Переопределяем уровни только для non-base записей на основе базовых
    if len(base_rows) > 0:
        non_base_mask = ~base_mask
        non_base_indices = df[non_base_mask].index.tolist()
        
        # Обновляем уровни только для записей с level == 1
        for idx in non_base_indices:
            current_level = df.loc[idx, 'level']
            # Пропускаем записи, где уровень не равен 1
            if pd.notna(current_level) and int(current_level) != 1:
                continue
            
            target_code = df.loc[idx, 'concatenated_code']
            nearest_base_idx, nearest_code = find_nearest_code(
                target_code, base_codes, base_indices, base_levels
            )
            
            if nearest_base_idx >= 0 and nearest_code:
                base_level = df.loc[nearest_base_idx, 'level']
                if pd.isna(base_level):
                    base_level = 0
                else:
                    base_level = int(base_level)
                new_level = base_level + 1
                df.loc[idx, 'level'] = new_level
            else:
                # Если нет ближайшей базовой записи, устанавливаем уровень 0
                new_level = 0
                df.loc[idx, 'level'] = new_level
    
    return df


def get_level_updates_from_dataframe(
    df: pd.DataFrame,
    source_table_map: dict
) -> List[Tuple[str, int, int, str]]:
    """
    Формирует список обновлений уровней для базы данных.
    
    Args:
        df: DataFrame с обновленными уровнями (результат recalculate_levels_for_income_codes)
        source_table_map: Словарь {(id, ppocode): table_name} для определения исходной таблицы
        
    Returns:
        Список кортежей (table_name, id, level, ppocode) для обновления в БД
    """
    updates = []
    
    # Определяем, какие записи были изменены (только для записей с level == 1, которые были переопределены)
    base_ppocode = '00000000'
    base_mask = df['ppocode'] == base_ppocode
    non_base_mask = ~base_mask
    
    for idx in df[non_base_mask].index:
        record_id = int(df.loc[idx, 'id'])
        ppocode = df.loc[idx, 'ppocode']
        level = df.loc[idx, 'level']
        
        key = (record_id, ppocode)
        table_name = source_table_map.get(key, 'budgetclassubtypincmo')
        
        # Добавляем только записи, которые были переопределены (level != 1 или level == 0)
        if pd.notna(level):
            level_int = int(level)
            if level_int != 1:  # Были переопределены
                updates.append((table_name, record_id, level_int, ppocode))
    
    return updates

"""
Утилиты для работы с базой данных SQLite, включая функции для фильтрации VIEW по дате.
Перенесено из Osnova/db_utils.py для интеграции в проект.
"""
import sqlite3
import pandas as pd
from typing import Optional, List, Union

# Правила сортировки для различных таблиц справочников
SORT_RULES = {
    'oktmo': ['code', 'startdate', 'enddate'],
    'budgetclastypeinc': ['ppocode', 'inctypecode', 'incsubtypecode', 'analyticalgroupcode', 'startdate', 'enddate', 'year'],
    'budgetclassubtypincmo': ['ppocode', 'inctypecode', 'incsubtypecode', 'analyticalgroupcode', 'startdate', 'enddate', 'year'],
    'budgetclasgabs': ['ppocode', 'code', 'startdate', 'enddate', 'year'],
    'budgetclasgabsmo': ['ppocode', 'code', 'startdate', 'enddate', 'year'],
    'budgetclasgrbs': ['ppocode', 'code', 'startdate', 'enddate', 'year'],
    'budgetclasgrbsmo': ['ppocode', 'code', 'startdate', 'enddate', 'year'],
    'budgetclascosts': ['ppocode', 'grbscode', 'rzpr', 'kcsr', 'kvr', 'startdate', 'enddate', 'year'],
    'budgetclascostsmo': ['ppocode', 'grbscode', 'rzpr', 'kcsr', 'kvr', 'startdate', 'enddate', 'year'],
    'budgetclasgaiffb': ['ppocode', 'code', 'startdate', 'enddate', 'year'],
    'budgetclasgaifmo': ['ppocode', 'code', 'startdate', 'enddate', 'year'],
}


def get_partition_by_fields(view_name: str) -> List[str]:
    """
    Определяет поля для PARTITION BY на основе имени VIEW.
    
    Args:
        view_name: Имя VIEW (например, 'v_budgetclastypeinc_merged')
        
    Returns:
        Список полей для группировки при дедупликации
    """
    if view_name.startswith('v_') and view_name.endswith('_merged'):
        table_name = view_name[2:-7]
    else:
        table_name = view_name

    partition_fields = [field for field in SORT_RULES.get(table_name, []) 
                       if field not in {'startdate', 'enddate', 'year'}]
      
    return [] if not partition_fields else partition_fields


def get_filtered_view(
    conn: sqlite3.Connection,
    view_name: str,
    filter_date: Optional[str] = None,
    partition_by: Union[str, List[str]] = None
) -> pd.DataFrame:
    """
    Получает данные из VIEW с фильтрацией по дате и дедупликацией.
    
    Логика работы:
    1. Определяет поля для PARTITION BY на основе SORT_RULES (исключая 'startdate', 'enddate', 'year')
    2. Фильтрует записи по дате: startdate <= filter_date <= enddate (или enddate пустой)
    3. Применяет дедупликацию: выбирает запись с максимальным startdate для каждой комбинации полей из partition_by
    
    Args:
        conn: SQLite соединение
        view_name: Имя VIEW (например, 'v_budgetclastypeinc_merged')
        filter_date: Дата для фильтрации в формате 'YYYY-MM-DD' (опционально, по умолчанию текущая дата)
        partition_by: Поля для группировки при дедупликации (опционально, определяется автоматически)
        
    Returns:
        DataFrame с отфильтрованными и дедуплицированными данными
    """
    # Определяем поля для PARTITION BY
    if partition_by is None:
        partition_by = get_partition_by_fields(view_name)
    elif isinstance(partition_by, str):
        partition_by = [partition_by]
    
    # Формируем список полей для PARTITION BY
    partition_fields = ', '.join(partition_by) if partition_by else '1'
    
    # Определяем дату для фильтрации
    if filter_date is not None:
        date_filter = "date(?)"
        params = [filter_date, filter_date]
    else:
        date_filter = "date('now')"
        params = None
    
    # Формируем SQL запрос с фильтрацией по дате и дедупликацией
    # Используем date() для нормализации дат (на случай если в БД даты в формате 'YYYY-MM-DD HH:MM:SS')
    query = f"""
        SELECT * FROM (
            SELECT 
                *,
                ROW_NUMBER() OVER (
                    PARTITION BY {partition_fields}
                    ORDER BY startdate DESC
                ) AS rn
            FROM {view_name}
            WHERE (startdate IS NULL OR date(startdate) <= {date_filter})
              AND (enddate IS NULL OR enddate = '' OR date(enddate) >= {date_filter})
        ) WHERE rn = 1
    """
    
    if params:
        return pd.read_sql_query(query, conn, params=params)
    else:
        return pd.read_sql_query(query, conn)

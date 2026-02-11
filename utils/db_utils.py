"""
Утилиты для работы с базой данных SQLite, включая функции для фильтрации таблиц и VIEW по дате.
Перенесено из Osnova/db_utils.py для интеграции в проект.
"""
import sqlite3
import pandas as pd
from typing import Optional, List, Union, Sequence

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
    'budgetclassources': ['gaifcode', 'ppocode', 'code', 'startdate', 'enddate', 'year'],
    'budgetclassourcesmo': ['gaifcode', 'ppocode', 'code', 'startdate', 'enddate', 'year'],
}


def get_partition_by_fields(table_name: str) -> List[str]:
    """
    Определяет поля для PARTITION BY на основе имени таблицы или VIEW.
    
    Args:
        table_name: Имя таблицы или VIEW (например, 'budgetclastypeinc' или 'v_budgetclastypeinc_merged')
                    В SQLite VIEW и таблицы используются одинаково в SQL запросах.
        
    Returns:
        Список полей для группировки при дедупликации
    """
    # Если это VIEW с префиксом v_ и суффиксом _merged, извлекаем имя базовой таблицы
    if table_name.startswith('v_') and table_name.endswith('_merged'):
        base_table_name = table_name[2:-7]
    else:
        base_table_name = table_name

    partition_fields = [field for field in SORT_RULES.get(base_table_name, []) 
                       if field not in {'startdate', 'enddate', 'year'}]
      
    return [] if not partition_fields else partition_fields


def _has_column(conn: sqlite3.Connection, table_name: str, column_name: str) -> bool:
    """
    Проверяет наличие указанного поля в таблице или VIEW.
    
    Args:
        conn: SQLite соединение
        table_name: Имя таблицы или VIEW
        column_name: Имя проверяемого поля
        
    Returns:
        True если поле существует, False иначе
    """
    try:
        cursor = conn.cursor()
        cursor.execute(f'SELECT {column_name} FROM {table_name} LIMIT 0')
        return True
    except (sqlite3.OperationalError, sqlite3.DatabaseError):
        return False


def get_filtered_view(
    conn: sqlite3.Connection,
    table_name: str,
    filter_date: Optional[str] = None,
    filter_ppocode: Optional[Union[str, Sequence[str]]] = None,
    partition_by: Union[str, List[str]] = None,
    join_npa: bool = False,
    deduplicate: bool = True
) -> pd.DataFrame:
    """
    Получает данные из таблицы или VIEW с фильтрацией по дате, по ppocode (ОКТМО) и опциональной дедупликацией.
    
    Работает как для обычных таблиц онлайн справочников (например, 'budgetclastypeinc'),
    так и для объединенных VIEW (например, 'v_budgetclastypeinc_merged').
    В SQLite VIEW и таблицы используются одинаково в SQL запросах.
    
    Логика работы:
    1. Определяет поля для PARTITION BY на основе SORT_RULES (исключая 'startdate', 'enddate', 'year')
    2. Фильтрует записи по дате: startdate <= filter_date <= enddate (или enddate пустой)
    3. Если задан filter_ppocode и у таблицы/VIEW есть колонка ppocode — фильтрует:
       - строка или один элемент: ppocode = ?
       - список: ppocode IN (?, ?, ...) (например, ФУ 00000000 + ОКТМО из ревизии)
    4. Применяет дедупликацию (если deduplicate=True): выбирает запись с максимальным startdate (и year, если присутствует) 
       для каждой комбинации полей из partition_by. Если deduplicate=False, возвращает все записи.
    5. Опционально добавляет JOIN к таблице npa для получения данных нормативно-правовых актов
    
    Args:
        conn: SQLite соединение
        table_name: Имя таблицы или VIEW (например, 'budgetclastypeinc' или 'v_budgetclastypeinc_merged')
        filter_date: Дата для фильтрации в формате 'YYYY-MM-DD' (опционально, по умолчанию текущая дата)
        filter_ppocode: Код(ы) участника БП (ОКТМО): одна строка или список строк (ФУ + ОКТМО из ревизии)
        partition_by: Поля для группировки при дедупликации (опционально, определяется автоматически)
        join_npa: Если True, добавляет LEFT JOIN к таблице npa для получения данных НПА
        deduplicate: Если True (по умолчанию), оставляет только одну запись на группу. Если False, возвращает все записи.
        
    Returns:
        DataFrame с отфильтрованными и дедуплицированными данными
    """
    
    # Определяем поля для PARTITION BY
    if partition_by is None:
        partition_by = get_partition_by_fields(table_name)
    elif isinstance(partition_by, str):
        partition_by = [partition_by]

    table_prefix = 't.' if join_npa else ''
    partition_fields = ', '.join([f'{table_prefix}{field}' for field in partition_by]) if partition_by else '1'
    
    # Формируем ORDER BY
    has_year = _has_column(conn, table_name, 'year')
    order_by = f'ORDER BY {f"{table_prefix}year DESC, " if has_year else ""}{table_prefix}startdate DESC'

    # Определяем дату для фильтрации
    date_filter = "date(?)" if filter_date else "date('now')"
    params: List = []
    if filter_date:
        params.extend([filter_date, filter_date])

    # Фильтр по ppocode: одна строка -> = ?, список -> IN (?, ?, ...)
    ppocode_list: List[str] = []
    if filter_ppocode is not None:
        if isinstance(filter_ppocode, (list, tuple)):
            ppocode_list = [str(p).strip() for p in filter_ppocode if str(p).strip()]
        elif str(filter_ppocode).strip():
            ppocode_list = [str(filter_ppocode).strip()]
    use_ppocode = bool(ppocode_list) and _has_column(conn, table_name, 'ppocode')
    if use_ppocode:
        params.extend(ppocode_list)
    
    # FROM и SELECT для JOIN к npa
    from_clause = table_name
    select_npa = ''
    if join_npa and _has_column(conn, table_name, 'npa_id'):
        cursor = conn.cursor()
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND LOWER(name)='npa'")
        npa_row = cursor.fetchone()
        if npa_row:
            npa_tbl = f'{npa_row[0]}'
            cursor.execute(f"PRAGMA table_info({npa_row[0]})")
            npa_cols = [row[1] for row in cursor.fetchall() if row[1] != 'id']
            if npa_cols:
                select_npa = ', ' + ', '.join([f'n.{col} AS npa_{col}' for col in npa_cols])
                from_clause = f'{table_name} AS t LEFT JOIN {npa_tbl} AS n ON t.npa_id = n.id'

    if join_npa:
        from_clause = f'{table_name} AS t' if not select_npa else from_clause
        table_prefix = 't.'

    where_parts = [
        f"({table_prefix}startdate IS NULL OR date({table_prefix}startdate) <= {date_filter})",
        f"({table_prefix}enddate IS NULL OR {table_prefix}enddate = '' OR date({table_prefix}enddate) >= {date_filter})",
    ]
    if use_ppocode:
        if len(ppocode_list) == 1:
            where_parts.append(f"{table_prefix}ppocode = ?")
        else:
            placeholders = ", ".join("?" for _ in ppocode_list)
            where_parts.append(f"{table_prefix}ppocode IN ({placeholders})")

    inner_select = ('t.*' + select_npa) if join_npa else '*'
    where_clause = " AND ".join(where_parts)
    # Во внешнем SELECT — только *: результат подзапроса не имеет алиаса t/n, иначе "no such table: t".
    # Условие rn: = 1 для дедупликации (только первая запись), <> 0 для всех записей
    rn_condition = "rn = 1" if deduplicate else "rn <> 0"
    query = f"""
        SELECT * FROM (
            SELECT {inner_select},
                ROW_NUMBER() OVER (PARTITION BY {partition_fields} {order_by}) AS rn
            FROM {from_clause}
            WHERE {where_clause}
        ) WHERE {rn_condition}
    """

    # Подставляем параметры в SQL для отладки
    query_with_params = query
    for param in (params if params else []):
        param_str = f"'{param}'" if isinstance(param, str) else str(param)
        query_with_params = query_with_params.replace('?', param_str, 1)
    
    result = pd.read_sql_query(query, conn, params=params if params else None)
    result = result.drop(columns=['rn'], errors='ignore')
    #result.to_csv('debug_filtered_view.csv', index=False, encoding='utf-8-sig')  # Сохранение для отладки
    return result

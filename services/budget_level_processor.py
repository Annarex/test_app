"""
Обработка и переопределение уровней для бюджетных справочников доходов.
Логика перенесена из Osnova/app_budgetclastypeinc_merged.py для интеграции в проект.
"""
import pandas as pd
import sqlite3
import json
from datetime import datetime
from typing import Optional, Dict, List, Tuple, Set
from contextlib import contextmanager
from logger import logger
from utils.db_utils import get_filtered_view
from utils.level_utils import (
    recalculate_levels_for_income_codes,
    get_level_updates_from_dataframe
)


TABLE_FU = 'budgetclastypeinc'
TABLE_MO = 'budgetclassubtypincmo'
VIEW_MERGED = 'v_budgetclastypeinc_merged'
DEFAULT_BASE_PPOCODE = '00000000'


class BudgetLevelProcessor:
    """Процессор для обработки и переопределения уровней кодов доходов"""
    
    def __init__(self, db_path: str = "budget_forms.db"):
        """
        Args:
            db_path: Путь к базе данных SQLite
        """
        self.db_path = db_path
    
    @contextmanager
    def _get_connection(self):
        """Контекстный менеджер для работы с соединением БД"""
        conn = sqlite3.connect(self.db_path)
        try:
            yield conn
        finally:
            conn.close()
    
    def _determine_source_table(
        self, 
        cursor: sqlite3.Cursor, 
        record_id: int, 
        ppocode: str
    ) -> str:
        """
        Определяет таблицу-источник для записи по (id, ppocode).
        
        Args:
            cursor: Курсор БД
            record_id: ID записи
            ppocode: Код участника БП
            
        Returns:
            Имя таблицы-источника
        """
        cursor.execute(
            "SELECT 1 FROM budgetclastypeinc WHERE id = ? AND ppocode = ?", 
            (record_id, ppocode)
        )
        return TABLE_FU if cursor.fetchone() else TABLE_MO
    
    def _build_source_table_map_optimized(
        self, 
        cursor: sqlite3.Cursor, 
        df: pd.DataFrame
    ) -> Dict[Tuple[int, str], str]:
        """Строит карту (id, ppocode) -> table_name одним SQL запросом
        
        Разбивает запрос на батчи для избежания ошибки "too many SQL variables"
        """
        unique_keys = df[['id', 'ppocode']].drop_duplicates()
        
        if len(unique_keys) == 0:
            return {}
        
        keys_list = [
            (int(row['id']), row['ppocode']) 
            for _, row in unique_keys.iterrows()
        ]
        
        # SQLite имеет ограничение на количество переменных (обычно ~999 или 32766)
        # Используем батчи по 400 пар (800 параметров) для безопасности
        BATCH_SIZE = 400
        fu_records = set()
        
        # Обрабатываем ключи батчами
        for i in range(0, len(keys_list), BATCH_SIZE):
            batch = keys_list[i:i + BATCH_SIZE]
            
            placeholders = ','.join(['(?, ?)'] * len(batch))
            query = f"""
                SELECT id, ppocode 
                FROM {TABLE_FU} 
                WHERE (id, ppocode) IN (VALUES {placeholders})
            """
            
            params = []
            for record_id, ppocode in batch:
                params.extend([record_id, ppocode])
            
            cursor.execute(query, params)
            batch_fu_records = {(row[0], row[1]) for row in cursor.fetchall()}
            fu_records.update(batch_fu_records)
        
        source_table_map = {}
        for record_id, ppocode in keys_list:
            key = (record_id, ppocode)
            source_table_map[key] = TABLE_FU if key in fu_records else TABLE_MO
        
        return source_table_map
    
    def _build_source_table_map(
        self, 
        cursor: sqlite3.Cursor, 
        df: pd.DataFrame
    ) -> Dict[Tuple[int, str], str]:
        return self._build_source_table_map_optimized(cursor, df)
    
    def process_and_update_levels(
        self, 
        base_ppocode: str = DEFAULT_BASE_PPOCODE
    ) -> pd.DataFrame:
        """
        Загружает данные из VIEW, переопределяет уровни, обновляет в БД и возвращает результат.
        
        Args:
            base_ppocode: Код базового участника БП (по умолчанию '00000000' для ФУ)
            
        Returns:
            DataFrame с обновленными уровнями
        """        
        with self._get_connection() as conn:
            cursor = conn.cursor()
            
            logger.info(f"Загрузка данных из VIEW {VIEW_MERGED}...")
            df = pd.read_sql_query(f"SELECT * FROM {VIEW_MERGED}", conn)
            
            if df.empty:
                logger.warning(f"VIEW {VIEW_MERGED} пуст")
                return df
            
            source_table_map = self._build_source_table_map(cursor, df)
            df['source_table'] = df.apply(
                lambda row: source_table_map.get(
                    (int(row['id']), row['ppocode']), 
                    TABLE_MO
                ), 
                axis=1
            )
            
            df = recalculate_levels_for_income_codes(df, base_ppocode)
            updates = get_level_updates_from_dataframe(df, source_table_map)
            
            if updates:
                self._update_levels_in_db(cursor, updates)
                conn.commit()
            
            return df
    
    def _update_levels_in_db(
        self, 
        cursor: sqlite3.Cursor, 
        updates: List[Tuple[str, int, int, str]]
    ) -> None:
        """
        Обновляет уровни в базе данных.
        
        Args:
            cursor: Курсор БД
            updates: Список кортежей (table_name, id, level, ppocode)
        """
        logger.info(f"Обновление уровней в базе данных ({len(updates)} записей)...")
        updated_count = 0
        not_found_count = 0
        error_count = 0
        
        updates_by_table: Dict[str, List[Tuple[int, int, str]]] = {}
        for table_name, record_id, level_value, ppocode in updates:
            if table_name not in updates_by_table:
                updates_by_table[table_name] = []
            updates_by_table[table_name].append((record_id, level_value, ppocode))
        
        for table_name, table_updates in updates_by_table.items():
            for record_id, level_value, ppocode in table_updates:
                try:
                    cursor.execute(
                        f"SELECT id, level FROM {table_name} WHERE id = ? AND ppocode = ?", 
                        (record_id, ppocode)
                    )
                    existing = cursor.fetchone()
                    
                    if existing:
                        old_level = existing[1]
                        cursor.execute(
                            f"UPDATE {table_name} SET level = ? WHERE id = ? AND ppocode = ?", 
                            (str(level_value), record_id, ppocode)
                        )
                        if cursor.rowcount > 0:
                            updated_count += 1
                            if updated_count <= 5:
                                logger.info(
                                    f"  Обновлено: {table_name}.id={record_id}, "
                                    f"ppocode={ppocode}, level: {old_level} -> {level_value}"
                                )
                        else:
                            not_found_count += 1
                    else:
                        not_found_count += 1
                        if not_found_count <= 5:
                            logger.warning(f"  Не найдено: {table_name}.id={record_id}, ppocode={ppocode}")
                except sqlite3.Error as e:
                    error_count += 1
                    if error_count <= 5:
                        logger.error(
                            f"  Ошибка при обновлении {table_name}.id={record_id}, "
                            f"ppocode={ppocode}: {e}"
                        )
        
        logger.info(
            f"Итого: обновлено {updated_count}, не найдено {not_found_count}, "
            f"ошибок {error_count} из {len(updates)} записей"
        )
    
    def _load_npa_data(
        self, 
        conn: sqlite3.Connection, 
        df: pd.DataFrame
    ) -> Tuple[Dict[Tuple[int, str], Optional[int]], Dict[str, Dict]]:
        cursor = conn.cursor()
        npa_id_map = {}
        npa_data_cache = {}
        
        unique_combinations = df[['id', 'ppocode']].drop_duplicates()
        
        if len(unique_combinations) == 0:
            return npa_id_map, npa_data_cache
        
        source_table_map = self._build_source_table_map_optimized(cursor, df)
        npa_ids_to_load: Set[int] = set()
        
        for _, row in unique_combinations.iterrows():
            record_id = int(row['id'])
            ppocode = row['ppocode']
            key = (record_id, ppocode)
            table_name = source_table_map.get(key, TABLE_MO)
            
            cursor.execute(
                f"SELECT npa_id FROM {table_name} WHERE id = ? AND ppocode = ?", 
                (record_id, ppocode)
            )
            result = cursor.fetchone()
            npa_id_value = result[0] if result and result[0] is not None else None
            npa_id_map[key] = npa_id_value
            
            if npa_id_value is not None:
                npa_id_str = str(npa_id_value)
                npa_ids = [id.strip() for id in npa_id_str.split(',') if id.strip()]
                for npa_id in npa_ids:
                    try:
                        npa_ids_to_load.add(int(npa_id))
                    except ValueError:
                        continue
        
        if npa_ids_to_load:
            # Разбиваем запрос на батчи для избежания ошибки "too many SQL variables"
            BATCH_SIZE = 500  # Безопасный размер батча для SQLite
            npa_ids_list = list(npa_ids_to_load)
            
            for i in range(0, len(npa_ids_list), BATCH_SIZE):
                batch = npa_ids_list[i:i + BATCH_SIZE]
                placeholders = ','.join(['?'] * len(batch))
                query = f"""
                    SELECT id, name, numdoc, approvaldate, kindname 
                    FROM npa 
                    WHERE id IN ({placeholders})
                """
                cursor.execute(query, batch)
                
                for row in cursor.fetchall():
                    npa_id = str(row[0])
                    npa_data_cache[npa_id] = {
                        'name': row[1] or '',
                        'numdoc': row[2] or '',
                        'approvaldate': row[3] or '',
                        'kindname': row[4] or ''
                    }
        
        return npa_id_map, npa_data_cache
    
    def _format_npa(
        self, 
        row: pd.Series, 
        npa_id_map: Dict[Tuple[int, str], Optional[int]], 
        npa_data_cache: Dict[str, Dict]
    ) -> str:
        """
        Преобразует данные НПА в JSON массив.
        
        Args:
            row: Строка DataFrame
            npa_id_map: Словарь {(id, ppocode): npa_id}
            npa_data_cache: Кэш данных НПА
            
        Returns:
            JSON строка с массивом НПА
        """
        npa_list = []
        key = (int(row['id']), row['ppocode'])
        npa_id_value = npa_id_map.get(key)
        
        if npa_id_value is not None:
            npa_id_str = str(npa_id_value)
            npa_ids = [id.strip() for id in npa_id_str.split(',') if id.strip()]
            for npa_id in npa_ids:
                npa_item = npa_data_cache.get(npa_id)
                if npa_item:
                    npa_list.append(npa_item)
        
        return json.dumps(npa_list, ensure_ascii=False) if npa_list else '[]'

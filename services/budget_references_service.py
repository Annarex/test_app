"""
Сервис для заполнения базы данных SQLite справочниками бюджетной системы из API.
Логика перенесена из Osnova/app_fill_database.py для интеграции в проект.
"""
import time
import requests
import sqlite3
import json
from datetime import datetime
from typing import Optional, Dict, List
from logger import logger

# URL справочников
URL_OKTMO = "http://budget.gov.ru/epbs/registry/7710568760-OKTMO/data"
URL_BUDGETCLASTYPEINC = "http://budget.gov.ru/epbs/registry/7710568760-BUDGETCLASTYPEINC/data"
URL_BUDGETCLASSUBTYPINCMO = "http://budget.gov.ru/epbs/registry/7710568760-BUDGETCLASSUBTYPINCMO/data"
URL_BUDGETCLASGABS = "http://budget.gov.ru/epbs/registry/7710568760-BUDGETCLASGABS/data"
URL_BUDGETCLASGABSMO = "http://budget.gov.ru/epbs/registry/7710568760-BUDGETCLASGABSMO/data"
URL_BUDGETCLASGRBS = "http://budget.gov.ru/epbs/registry/7710568760-BUDGETCLASGRBS/data"
URL_BUDGETCLASGRBSMO = "http://budget.gov.ru/epbs/registry/7710568760-BUDGETCLASGRBSMO/data"
URL_BUDGETCLASCOSTS = "http://budget.gov.ru/epbs/registry/7710568760-BUDGETCLASCOSTS/data"
URL_BUDGETCLASCOSTSMO = "http://budget.gov.ru/epbs/registry/7710568760-BUDGETCLASCOSTSMO/data"
URL_BUDGETCLASGAIFFB = "http://budget.gov.ru/epbs/registry/7710568760-BUDGETCLASGAIFFB/data"
URL_BUDGETCLASGAIFMO = "http://budget.gov.ru/epbs/registry/7710568760-BUDGETCLASGAIFMO/data"
URL_BUDGETCLASSOURCES = "http://budget.gov.ru/epbs/registry/7710568760-BUDGETCLASSOURCES/data"
URL_BUDGETCLASSOURCESMO = "http://budget.gov.ru/epbs/registry/7710568760-BUDGETCLASSOURCESMO/data"
URL_BUDGETCLASKVR = "http://budget.gov.ru/epbs/registry/7710568760-BUDGETCLASKVR/data"
URL_BUDGETCLASRZPR = "https://budget.gov.ru/epbs/registry/7710568760-BUDGETCLASRZPR/data"
URL_BUDGETCLASKCSR = "https://budget.gov.ru/epbs/registry/7710568760-BUDGETCLASKCSR/data"
URL_BUDGETCLASKCSRMO = "https://budget.gov.ru/epbs/registry/7710568760-BUDGETCLASKCSRMO/data"

PAGE_SIZE = 1000

URL_TO_TABLE = {
    URL_OKTMO: 'oktmo',
    URL_BUDGETCLASTYPEINC: 'budgetclastypeinc',
    URL_BUDGETCLASSUBTYPINCMO: 'budgetclassubtypincmo',
    URL_BUDGETCLASGABS: 'budgetclasgabs',
    URL_BUDGETCLASGABSMO: 'budgetclasgabsmo',
    URL_BUDGETCLASGRBS: 'budgetclasgrbs',
    URL_BUDGETCLASGRBSMO: 'budgetclasgrbsmo',
    URL_BUDGETCLASCOSTS: 'budgetclascosts',
    URL_BUDGETCLASCOSTSMO: 'budgetclascostsmo',
    URL_BUDGETCLASGAIFFB: 'budgetclasgaiffb',
    URL_BUDGETCLASGAIFMO: 'budgetclasgaifmo',
    URL_BUDGETCLASSOURCES: 'budgetclassources',
    URL_BUDGETCLASSOURCESMO: 'budgetclassourcesmo',
    URL_BUDGETCLASKVR: 'budgetclaskvr',
    URL_BUDGETCLASRZPR: 'budgetclasrzpr',
    URL_BUDGETCLASKCSR: 'budgetclaskcsr',
    URL_BUDGETCLASKCSRMO: 'budgetclaskcsrmo',
}

FIELD_MAPPING = {}


def get_default_filters(default_region: str = "21") -> Dict[str, Optional[Dict]]:
    """
    Возвращает словарь фильтров по умолчанию для всех справочников.
    
    Args:
        default_region: Код региона по умолчанию (для фильтрации МО справочников)
        
    Returns:
        Словарь {table_name: filters_dict} с фильтрами для каждого справочника
    """
    return {
        'oktmo': {"regioncode": default_region, "status": "ACTIVE"},
        'budgetclastypeinc': None,  # ФУ - без фильтров
        'budgetclassubtypincmo': {"ppocode": f"{default_region}______"},
        'budgetclasgabs': {"ppocode": f"{default_region}______"},
        'budgetclasgabsmo': {"ppocode": f"{default_region}______"},
        'budgetclasgrbs': {"ppocode": f"{default_region}______"},
        'budgetclasgrbsmo': {"ppocode": f"{default_region}______"},
        'budgetclascosts': {"ppocode": f"{default_region}______"},
        'budgetclascostsmo': {"ppocode": f"{default_region}______"},
        'budgetclasgaiffb': None,  # ФУ - без фильтров
        'budgetclasgaifmo': {"ppocode": f"{default_region}______"},
        'budgetclassources': {"ppocode": f"{default_region}______"},
        'budgetclassourcesmo': {"ppocode": f"{default_region}______"},
        'budgetclaskvr': None,
        'budgetclasrzpr': None,
        'budgetclaskcsr': {"ppocode": f"{default_region}______"},
        'budgetclaskcsrmo': {"ppocode": f"{default_region}______"},
    }


class BudgetReferencesService:
    """Сервис для работы с бюджетными справочниками из API"""
    
    def __init__(self, db_path: str = "budget_forms.db"):
        """
        Args:
            db_path: Путь к базе данных SQLite
        """
        self.db_path = db_path
    
    def get_record_count(self, url: str, filters: Optional[Dict] = None) -> int:
        """
        Получает общее количество записей из API без загрузки всех данных.
        Делает запрос с минимальными параметрами (pageSize=10, offset=0) и извлекает recordCount.
        
        Args:
            url: URL API
            filters: Словарь фильтров для запроса
            
        Returns:
            Общее количество записей (0 если не удалось получить или произошла ошибка)
        """
        params = {
            "pageSize": 10,
            "offset": 0,
        }
        
        if filters:
            for key, value in filters.items():
                params[f"filter{key}"] = value
        
        try:
            r = requests.get(url, params=params, timeout=30)
            r.raise_for_status()
            data = r.json()
            
            # Извлекаем recordCount из ответа
            # Проверяем разные возможные варианты структуры ответа
            if isinstance(data, dict):
                # Стандартный вариант: recordCount в корне
                if "recordCount" in data:
                    count = data["recordCount"]
                    if count is not None:
                        return int(count)
                
                # Альтернативные варианты структуры
                if "total" in data:
                    count = data["total"]
                    if count is not None:
                        return int(count)
                
                if "count" in data:
                    count = data["count"]
                    if count is not None:
                        return int(count)
                
                # Если есть data/items, считаем их количество
                if "data" in data and isinstance(data["data"], list):
                    return len(data["data"])
                
                if "items" in data and isinstance(data["items"], list):
                    return len(data["items"])
            
            # Если это список, возвращаем его длину
            if isinstance(data, list):
                return len(data)
            
            logger.warning(f"[get_record_count] Не удалось найти recordCount в ответе API для {url}. Структура ответа: {type(data)}")
            logger.debug(f"[get_record_count] Ответ API: {str(data)[:500]}")
            return 0
            
        except requests.RequestException as e:
            logger.error(f"[get_record_count] Ошибка при получении количества записей из {url}: {e}")
            return 0
        except (ValueError, KeyError, TypeError) as e:
            logger.error(f"[get_record_count] Ошибка при парсинге ответа API для {url}: {e}")
            return 0
        except Exception as e:
            logger.error(f"[get_record_count] Неожиданная ошибка при получении количества записей из {url}: {e}", exc_info=True)
            return 0
    
    def fetch_chunk(self, url: str, offset: int, page_size: int = PAGE_SIZE, filters: Optional[Dict] = None) -> List[Dict]:
        """
        Загружает одну страницу данных из API.
        
        Args:
            url: URL API
            offset: Смещение для пагинации
            page_size: Размер страницы
            filters: Словарь фильтров для запроса
            
        Returns:
            Список записей
        """
        params = {
            "pageSize": page_size,
            "offset": offset,
        }
        
        if filters:
            for key, value in filters.items():
                params[f"filter{key}"] = value
        
        try:
            r = requests.get(url, params=params, timeout=30)
            r.raise_for_status()
            data = r.json()
            
            if isinstance(data, dict):
                if "data" in data:
                    return data["data"]
                if "items" in data:
                    return data["items"]
            if isinstance(data, list):
                return data
            return []
        except requests.RequestException as e:
            logger.error(f"Ошибка при загрузке данных из {url}: {e}")
            return []
    
    def fetch_all_rows(self, url: str, filters: Optional[Dict] = None, progress_callback=None, cancel_check=None) -> List[Dict]:
        """
        Загружает все строки из API с пагинацией.
        
        Args:
            url: URL API
            filters: Словарь фильтров для запроса
            progress_callback: Функция обратного вызова для обновления прогресса (current, total)
            cancel_check: Функция для проверки отмены загрузки (должна возвращать True если нужно отменить)
            
        Returns:
            Список всех записей
        """
        all_rows = []
        offset = 0
        
        # Получаем общее количество записей для прогресса
        total_count = None
        if progress_callback:
            total_count = self.get_record_count(url, filters)
        
        logger.info(f"Получение данных из {url}{' с фильтрами: '+str(filters) if filters else ' без фильтров'}")
        if total_count:
            logger.info(f"Всего записей в API: {total_count}")
        
        while True:
            # Проверяем отмену перед каждой итерацией
            if cancel_check and cancel_check():
                logger.info(f"Загрузка отменена пользователем на смещении {offset}")
                break
            
            rows = self.fetch_chunk(url, offset, filters=filters)
            if not rows:
                break
            all_rows.extend(rows)
            
            # Обновляем прогресс (вызываем даже если total_count неизвестен)
            if progress_callback:
                progress_callback(len(all_rows), total_count if total_count else 0)
            
            logger.info(f"смещение={offset}, получено {len(rows)}, всего {len(all_rows)}")
            if len(rows) < PAGE_SIZE:
                break
            offset += PAGE_SIZE
            time.sleep(0.2)
        
        # Финальный вызов callback
        if progress_callback:
            progress_callback(len(all_rows), total_count if total_count else len(all_rows))
        
        return all_rows
    
    def get_table_columns(self, cursor: sqlite3.Cursor, table_name: str) -> List[str]:
        """
        Получает список столбцов таблицы.
        
        Args:
            cursor: Курсор БД
            table_name: Имя таблицы
            
        Returns:
            Список имен столбцов
        """
        cursor.execute(f"PRAGMA table_info({table_name})")
        return [row[1] for row in cursor.fetchall()]
    
    def get_or_insert_npa(self, cursor: sqlite3.Cursor, npa_item: Dict) -> int:
        """
        Получает или вставляет НПА, возвращает его Код НПА.
        
        Args:
            cursor: Курсор БД
            npa_item: Словарь с данными НПА
            
        Returns:
            Код НПА
        """
        name = npa_item.get('name', '') or ''
        numdoc = npa_item.get('numdoc', '') or ''
        approvaldate = npa_item.get('approvaldate', '') or ''
        kindname = npa_item.get('kindname', '') or ''
        
        cursor.execute("""
            SELECT id FROM npa 
            WHERE name = ? AND numdoc = ? AND approvaldate = ? AND kindname = ?
        """, (name, numdoc, approvaldate, kindname))
        
        result = cursor.fetchone()
        if result:
            return result[0]
        
        cursor.execute("""
            INSERT INTO npa (name, numdoc, approvaldate, kindname)
            VALUES (?, ?, ?, ?)
        """, (name, numdoc, approvaldate, kindname))
        
        return cursor.lastrowid
    
    def insert_records(self, cursor: sqlite3.Cursor, table_name: str, records: List[Dict]) -> int:
        """
        Вставляет записи в указанную таблицу.
        
        Для таблиц с direct_id (budgetclascosts, budgetclascostsmo):
        - Использует INSERT OR REPLACE
        - Включает поле 'id' в список полей
        
        Для обычных таблиц:
        - Использует INSERT OR IGNORE (на случай если очистка не сработала)
        - Исключает поле 'id' (AUTOINCREMENT)
        
        Args:
            cursor: Курсор БД
            table_name: Имя таблицы
            records: Список записей для вставки
            
        Returns:
            Количество вставленных записей
        """
        if not records:
            return 0
        
        # Определяем тип таблицы
        use_direct_id = table_name in {'budgetclascosts', 'budgetclascostsmo'}
        
        # Получаем список колонок таблицы
        table_columns = self.get_table_columns(cursor, table_name)
        
        # Получаем поля из первой записи (исключаем 'npa', он обрабатывается отдельно)
        first_record = records[0]
        npa_field = 'npa' if 'npa' in first_record else None
        api_fields = [f for f in first_record.keys() if f != 'npa']
        
        # Строим маппинг полей API -> БД
        field_mapping = FIELD_MAPPING.get(table_name, {})
        field_to_api = {}
        db_fields = []
        
        for api_field in api_fields:
            db_field = field_mapping.get(api_field, api_field)
            
            # Проверяем, существует ли поле в таблице
            if db_field not in table_columns:
                continue
            
            # Для таблиц с direct_id включаем все поля (включая 'id')
            # Для обычных таблиц исключаем 'id' (AUTOINCREMENT)
            if use_direct_id:
                # Включаем все поля
                if db_field not in field_to_api:
                    field_to_api[db_field] = api_field
                    db_fields.append(db_field)
            else:
                # Исключаем 'id'
                if db_field != 'id':
                    if db_field not in field_to_api:
                        field_to_api[db_field] = api_field
                        db_fields.append(db_field)
        
        if not db_fields:
            logger.error(f"[{table_name}] Нет полей для вставки!")
            return 0
        
        # Формируем SQL запрос
        placeholders = ', '.join(['?' for _ in db_fields])
        field_names = ', '.join(db_fields)
        
        if use_direct_id:
            insert_sql = f"INSERT OR REPLACE INTO {table_name} ({field_names}) VALUES ({placeholders})"
        else:
            insert_sql = f"INSERT OR IGNORE INTO {table_name} ({field_names}) VALUES ({placeholders})"
        
        logger.info(f"[{table_name}] Вставка {len(records)} записей (use_direct_id={use_direct_id}, полей={len(db_fields)})")
        
        # Вставляем записи
        inserted_count = 0
        error_count = 0
        
        for record in records:
            # Формируем значения для вставки
            values = []
            for db_field in db_fields:
                api_field = field_to_api[db_field]
                value = record.get(api_field, '')
                
                # Обрабатываем разные типы значений
                if value is None:
                    value = ''
                elif isinstance(value, (list, dict)):
                    value = json.dumps(value, ensure_ascii=False)
                else:
                    value = str(value)
                
                values.append(value)
            
            try:
                # Вставляем запись
                cursor.execute(insert_sql, values)
                
                # Получаем ID записи для обновления NPA
                if use_direct_id and 'id' in record:
                    record_id = record['id']
                else:
                    record_id = cursor.lastrowid
                
                # Обрабатываем NPA если есть
                if npa_field and npa_field in record and record[npa_field]:
                    npa_ids = []
                    for npa_item in record[npa_field]:
                        npa_id = self.get_or_insert_npa(cursor, npa_item)
                        npa_ids.append(str(npa_id))
                    
                    if npa_ids:
                        cursor.execute(f"UPDATE {table_name} SET npa_id = ? WHERE id = ?", 
                                     (','.join(npa_ids), record_id))
                
                inserted_count += 1
                
            except sqlite3.Error as e:
                error_count += 1
                if error_count <= 5:  # Логируем только первые 5 ошибок
                    logger.error(f"[{table_name}] Ошибка при вставке записи: {e}")
                    if error_count == 5:
                        logger.warning(f"[{table_name}] Дальнейшие ошибки вставки будут пропущены в логах")
        
        if error_count > 0:
            logger.warning(f"[{table_name}] Вставлено: {inserted_count}, ошибок: {error_count}")
        else:
            logger.info(f"[{table_name}] Успешно вставлено: {inserted_count} записей")
        
        return inserted_count
    
    def fill_table_from_url(
        self, 
        url: str, 
        table_name: str, 
        filters: Optional[Dict] = None, 
        clear_existing: bool = False,
        api_count: Optional[int] = None,
        db_count: Optional[int] = None,
        progress_callback=None,
        cancel_check=None
    ) -> int:
        """
        Заполняет таблицу данными из URL.
        
        Логика работы:
        1. Проверяет количество записей в API и БД
        2. Если api_count == db_count - пропускает загрузку
        3. Если api_count != db_count И api_count != 0 - очищает БД и загружает новые данные
        4. Загружает все записи из API
        5. Вставляет записи в БД (INSERT OR REPLACE для direct_id, INSERT OR IGNORE для остальных)
        
        Args:
            url: URL API
            table_name: Имя таблицы в БД
            filters: Словарь фильтров для запроса
            clear_existing: Очищать ли существующие данные принудительно (если True, игнорирует проверки)
            api_count: Количество записей в API
            db_count: Текущее количество записей в БД
            progress_callback: Функция обратного вызова для обновления прогресса (current, total)
            cancel_check: Функция для проверки отмены загрузки
            
        Returns:
            Количество вставленных записей (0 если загрузка не требовалась)
        """
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        try:
            # Шаг 1: Преобразуем значения в int
            api_count_int = None
            db_count_int = None
            
            if api_count is not None:
                try:
                    api_count_int = int(api_count)
                except (ValueError, TypeError):
                    logger.warning(f"Не удалось преобразовать api_count в int для {table_name}: {api_count}")
            
            if db_count is not None:
                try:
                    db_count_int = int(db_count)
                except (ValueError, TypeError):
                    logger.warning(f"Не удалось преобразовать db_count в int для {table_name}: {db_count}")
            
            logger.info(f"[{table_name}] Начало обновления: API={api_count_int}, БД={db_count_int}, clear_existing={clear_existing}")
            
            # Шаг 2: Проверяем, нужно ли обновлять
            # Количество должно быть проверено заранее перед запуском потока
            # Если api_count=0 и db_count=0, это означает ошибку проверки - загружаем данные без проверки
            if not clear_existing and api_count_int is not None and db_count_int is not None:
                # Если оба равны 0, это ошибка проверки - пропускаем проверку и загружаем данные
                if api_count_int == 0 and db_count_int == 0:
                    logger.warning(f"[{table_name}] Количество не удалось проверить (оба равны 0), загружаем данные без проверки")
                elif api_count_int > 0 and api_count_int == db_count_int:
                    logger.info(f"[{table_name}] Количество совпадает ({api_count_int}). Загрузка не требуется.")
                    from models.database import DatabaseManager
                    db_manager = DatabaseManager(self.db_path)
                    db_manager.update_budget_reference_date(table_name, api_count_int)
                    return 0
            
            # Шаг 4: Определяем, нужно ли очищать таблицу
            should_clear = False
            
            if clear_existing:
                # Принудительная очистка
                should_clear = True
                logger.info(f"[{table_name}] Принудительная очистка (clear_existing=True)")
            elif api_count_int is not None and db_count_int is not None:
                # Автоматическое определение:
                # Очищаем если: api_count != db_count AND api_count != 0
                should_clear = (api_count_int != db_count_int) and (api_count_int != 0)
                logger.info(f"[{table_name}] Автоопределение очистки: API={api_count_int}, БД={db_count_int}, очистка={'ДА' if should_clear else 'НЕТ'}")
            else:
                # Не удалось определить - не очищаем
                logger.warning(f"[{table_name}] Не удалось определить необходимость очистки (API={api_count_int}, БД={db_count_int})")
                should_clear = False
            
            # Шаг 5: Очищаем таблицу если нужно
            if should_clear:
                logger.info(f"[{table_name}] Очистка таблицы...")
                try:
                    cursor.execute(f"SELECT COUNT(*) FROM {table_name}")
                    count_before = cursor.fetchone()[0]
                    
                    cursor.execute(f"DELETE FROM {table_name}")
                    cursor.execute(f"DELETE FROM sqlite_sequence WHERE name='{table_name}'")
                    conn.commit()
                    
                    cursor.execute(f"SELECT COUNT(*) FROM {table_name}")
                    count_after = cursor.fetchone()[0]
                    
                    logger.info(f"[{table_name}] Очистка завершена: было {count_before}, осталось {count_after}")
                    
                    if count_after > 0:
                        logger.error(f"[{table_name}] ОШИБКА: После очистки осталось {count_after} записей!")
                except Exception as e:
                    logger.error(f"[{table_name}] Ошибка при очистке: {e}", exc_info=True)
                    conn.rollback()
                    raise
            
            # Шаг 6: Загружаем все записи из API
            logger.info(f"[{table_name}] Загрузка данных из API...")
            records = self.fetch_all_rows(url, filters=filters, progress_callback=progress_callback, cancel_check=cancel_check)
            
            if not records:
                logger.warning(f"[{table_name}] Нет данных для загрузки")
                return 0
            
            logger.info(f"[{table_name}] Загружено {len(records)} записей из API")
            
            # Шаг 7: Вставляем записи в БД
            logger.info(f"[{table_name}] Вставка записей в БД...")
            inserted = self.insert_records(cursor, table_name, records)
            conn.commit()
            
            # Шаг 8: Обновляем дату последнего обновления
            from models.database import DatabaseManager
            db_manager = DatabaseManager(self.db_path)
            db_manager.update_budget_reference_date(table_name, inserted)
            
            logger.info(f"[{table_name}] Обновление завершено: вставлено {inserted} записей")
            return inserted
            
        except Exception as e:
            logger.error(f"[{table_name}] Критическая ошибка при обновлении: {e}", exc_info=True)
            conn.rollback()
            raise
        finally:
            conn.close()
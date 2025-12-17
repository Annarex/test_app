"""
Контроллер для обработки решений о бюджете
Реализует логику из 1С: Решение
"""
from PyQt5.QtCore import QObject, pyqtSignal
from typing import Dict, Any, Optional, List
from pathlib import Path
from datetime import datetime
import docx
import pandas as pd
import sqlite3
from logger import logger

from models.database import DatabaseManager
from models.utils.code_utils import parse_income_code, parse_expense_code, build_expense_code


class SolutionController(QObject):
    """Контроллер для обработки решений о бюджете из Word документов"""
    
    # Сигналы
    solution_parsed = pyqtSignal(dict)
    error_occurred = pyqtSignal(str)
    
    def __init__(self, db_manager: DatabaseManager):
        super().__init__()
        self.db_manager = db_manager
    
    def parse_solution_document(self, file_path: str, project_id: int) -> Optional[Dict[str, Any]]:
        """
        Парсинг Word документа с решением о бюджете
        
        Args:
            file_path: Путь к Word документу
            project_id: ID проекта для сохранения данных
        
        Returns:
            Словарь с распарсенными данными или None при ошибке
        """
        try:
            doc = docx.Document(file_path)
            
            # Извлекаем все таблицы из документа
            tables = doc.tables
            if not tables:
                self.error_occurred.emit("В документе не найдено таблиц")
                return None
            
            result = {
                'приложение1': [],  # Доходы
                'приложение2': [],  # Расходы (общие)
                'приложение3': []   # Расходы по ГРБС
            }
            
            # Обрабатываем каждую таблицу
            for table in tables:
                table_data = self._extract_table_data(table)
                
                # Определяем тип приложения по содержимому
                app_type = self._identify_application_type(table_data)
                
                if app_type == 'приложение1':
                    result['приложение1'] = self._process_appendix1(table_data, project_id)
                elif app_type == 'приложение2':
                    result['приложение2'] = self._process_appendix2(table_data, project_id)
                elif app_type == 'приложение3':
                    result['приложение3'] = self._process_appendix3(table_data, project_id)
            
            self.solution_parsed.emit(result)
            logger.info(f"Решение обработано: {len(result['приложение1'])} доходов, "
                       f"{len(result['приложение2'])} расходов")
            
            return result
            
        except Exception as e:
            error_msg = f"Ошибка парсинга решения: {e}"
            logger.error(error_msg, exc_info=True)
            self.error_occurred.emit(error_msg)
            return None
    
    def _extract_table_data(self, table: docx.table.Table) -> List[List[str]]:
        """Извлечение данных из таблицы Word"""
        data = []
        for row in table.rows:
            row_data = []
            for cell in row.cells:
                # Очищаем текст от лишних символов
                text = cell.text.strip()
                text = text.replace('\n', '').replace('\r', '')
                row_data.append(text)
            data.append(row_data)
        return data
    
    def _identify_application_type(self, table_data: List[List[str]]) -> Optional[str]:
        """Определение типа приложения по содержимому таблицы"""
        if not table_data:
            return None
        
        # Проверяем первую строку на наличие ключевых слов
        first_row_text = ' '.join(table_data[0]).lower()
        
        if 'код классификации доходов бюджетов' in first_row_text:
            return 'приложение1'
        elif 'код целевой статьи' in first_row_text or 'код классификации расходов' in first_row_text:
            # Проверяем наличие колонки ГРБС
            has_grbs = any('грбс' in ' '.join(row).lower() for row in table_data[:3])
            if has_grbs:
                return 'приложение3'
            else:
                return 'приложение2'
        
        return None
    
    def _process_appendix1(self, table_data: List[List[str]], project_id: int) -> List[Dict[str, Any]]:
        """
        Обработка Приложения 1 (Доходы)
        Логика из 1С: Приложение1НаСервере
        Столбец1 - заголовок, Столбец2 - код, Столбец3 - наименование, Столбец4-6 - суммы
        """
        result = []
        t = 1  # Счетчик для группировки
        
        # Обрабатываем строки данных
        for row in table_data:
            if len(row) < 4:
                continue
            
            # Столбец2 (индекс 1) должен содержать символ '0' (код 48) и длина > 10
            столбец2 = row[1].strip() if len(row) > 1 else ''
            
            # Проверяем условие из 1С: содержит '0' и длина > 10
            if '0' in столбец2 and len(столбец2) > 10:
                # Очищаем код от пробелов и "000 "
                н_код = столбец2.replace(' ', '').replace('000', '')
                
                # Ищем код в справочнике доходов
                код_д = self._find_income_code(н_код)
                
                # Если код найден в справочнике, используем его уровень
                # Если не найден, пропускаем строку (не создаем новую запись)
                if not код_д:
                    # Код не найден в справочнике - пропускаем строку
                    continue
                
                уровень = код_д.get('уровень', 0)
                
                # Извлекаем суммы из столбцов 4, 5, 6 (индексы 3, 4, 5)
                сумма1 = 0
                сумма2 = 0
                сумма3 = 0
                
                if len(row) > 3:
                    сумма1 = self._parse_number(row[3])
                if len(row) > 4:
                    сумма2 = self._parse_number(row[4])
                if len(row) > 5:
                    сумма3 = self._parse_number(row[5])
                
                # Наименование из столбца 3 (индекс 2)
                наименование = row[2].strip() if len(row) > 2 else ''
                
                # Если уровень = 2, увеличиваем счетчик t
                if уровень == 2:
                    t += 1
                
                # Сохраняем только если уровень = 1 или 2
                if уровень in [1, 2]:
                    код_значение = код_д.get('название', '') if уровень in [1, 2] else н_код
                    
                    result.append({
                        'код': н_код,
                        'наименование': наименование,
                        'уровень': уровень,
                        'ТТ': t,
                        'сумма1': сумма1,
                        'сумма2': сумма2,
                        'сумма3': сумма3,
                        'код_значение': код_значение
                    })
        
        # Группируем по уровню и ТТ (аналог Свернуть в 1С)
        grouped_result = self._group_by_lvl_tt(result)
        
        # Разбираем коды на компоненты для каждого элемента
        for item in grouped_result:
            код = item.get('код', '')
            if код:
                # Очищаем код от пробелов
                код_чистый = код.replace(' ', '').replace('\xa0', '')
                # Если код не 20 разрядов, пытаемся дополнить нулями слева
                if len(код_чистый) < 20:
                    код_чистый = код_чистый.zfill(20)
                elif len(код_чистый) > 20:
                    код_чистый = код_чистый[:20]
                
                # Разбираем код на компоненты
                if len(код_чистый) == 20:
                    components = parse_income_code(код_чистый)
                    if components:
                        item.update(components)
                        # Обновляем код на полный 20-разрядный
                        item['код'] = код_чистый
        
        return grouped_result
    
    def _process_appendix2(self, table_data: List[List[str]], project_id: int) -> List[Dict[str, Any]]:
        """
        Обработка Приложения 2 (Расходы общие)
        Логика из 1С: Приложение2
        Столбец1 - наименование, Столбец2 - код Р, Столбец3 - код ПР, 
        Столбец4 - код ЦС, Столбец5 - код ВР, Столбец6-8 - суммы
        """
        result = []
        
        # Обрабатываем строки данных
        for row in table_data:
            if len(row) < 6:
                continue
            
            # Столбец2 (индекс 1) должен содержать '0' или '1' и длина = 2
            столбец2 = row[1].strip() if len(row) > 1 else ''
            
            # Проверяем условие из 1С
            if ('0' in столбец2 or '1' in столбец2) and len(столбец2) == 2:
                # Формируем наименование = столбец2 + столбец3 + столбец4 (без пробелов) + столбец5
                столбец3 = row[2].strip() if len(row) > 2 else ''
                столбец4 = row[3].strip().replace(' ', '') if len(row) > 3 else ''
                столбец5 = row[4].strip() if len(row) > 4 else ''
                
                наименование = столбец2 + столбец3 + столбец4 + столбец5
                
                # Ищем код в справочнике расходов
                код_р = self._find_expense_code(наименование)
                
                if код_р and not код_р.get('пустой', True):
                    уровень = код_р.get('уровень', 0)
                    код_р_значение = столбец2
                    
                    # Извлекаем суммы из столбцов 6, 7, 8 (индексы 5, 6, 7)
                    сумма1 = self._parse_number(row[5]) if len(row) > 5 else 0
                    сумма2 = self._parse_number(row[6]) if len(row) > 6 else 0
                    сумма3 = self._parse_number(row[7]) if len(row) > 7 else 0
                    
                    # Формируем полный код из компонентов (ГРБС по умолчанию '000')
                    ГРБС = '000'  # В приложении 2 ГРБС не указан, используем значение по умолчанию
                    полный_код = build_expense_code(ГРБС, код_р_значение, столбец3, столбец4, столбец5)
                    
                    # Разбираем код на компоненты для проверки
                    components = parse_expense_code(полный_код)
                    
                    result.append({
                        'ГРБС': ГРБС,
                        'код_Р': код_р_значение,
                        'код_ПР': столбец3,
                        'код_ЦС': столбец4,
                        'код_ВР': столбец5,
                        'уровень': уровень,
                        'сумма1': сумма1,
                        'сумма2': сумма2,
                        'сумма3': сумма3,
                        'наименование': наименование,
                        'полный_код': полный_код
                    })
                    
                    # Добавляем компоненты из разбора
                    if components:
                        result[-1].update(components)
                else:
                    # Если код не найден в справочнике, пропускаем строку
                    # (не создаем новую запись)
                    pass
        
        # Группируем по уровню, код_Р, код_ПР, код_ЦС, код_ВР
        grouped_result = self._group_by_expense_codes(result)
        
        return grouped_result
    
    def _process_appendix3(self, table_data: List[List[str]], project_id: int) -> List[Dict[str, Any]]:
        """
        Обработка Приложения 3 (Расходы по ГРБС)
        Логика из 1С: Приложение3
        Столбец1 - наименование, Столбец2 - ГРБС, Столбец3 - код Р, 
        Столбец4 - код ПР, Столбец5 - код ЦС, Столбец6 - код ВР, Столбец7 - сумма
        """
        result = []
        
        # Обрабатываем строки данных
        for row in table_data:
            if len(row) < 7:
                continue
            
            # Столбец3 (индекс 2) должен содержать '0' или '1' и длина = 2
            столбец3 = row[2].strip() if len(row) > 2 else ''
            
            # Проверяем условие из 1С
            if ('0' in столбец3 or '1' in столбец3) and len(столбец3) == 2:
                # ГРБС из столбца 2 (индекс 1)
                грбс = row[1].strip() if len(row) > 1 else ''
                
                # Формируем наименование = столбец3 + столбец4 + столбец5 (без пробелов) + столбец6
                столбец4 = row[3].strip() if len(row) > 3 else ''
                столбец5 = row[4].strip().replace(' ', '') if len(row) > 4 else ''
                столбец6 = row[5].strip() if len(row) > 5 else ''
                
                наименование = столбец3 + столбец4 + столбец5 + столбец6
                
                # Ищем код в справочнике расходов
                код_р = self._find_expense_code(наименование)
                
                if код_р and not код_р.get('пустой', True):
                    уровень = код_р.get('уровень', 0)
                    код_р_значение = столбец3
                    
                    # Извлекаем сумму из столбца 7 (индекс 6)
                    сумма1 = self._parse_number(row[6]) if len(row) > 6 else 0
                    
                    # Формируем полный код из компонентов
                    полный_код = build_expense_code(грбс, код_р_значение, столбец4, столбец5, столбец6)
                    
                    # Разбираем код на компоненты для проверки
                    components = parse_expense_code(полный_код)
                    
                    result.append({
                        'ГРБС': грбс,
                        'код_Р': код_р_значение,
                        'код_ПР': столбец4,
                        'код_ЦС': столбец5,
                        'код_ВР': столбец6,
                        'уровень': уровень,
                        'сумма1': сумма1,
                        'наименование': наименование,
                        'полный_код': полный_код
                    })
                    
                    # Добавляем компоненты из разбора
                    if components:
                        result[-1].update(components)
        
        # Группируем по уровню, ГРБС, код_Р
        grouped_result = self._group_by_grbs(result)
        
        return grouped_result
    
    def _find_income_code(self, code: str) -> Optional[Dict[str, Any]]:
        """
        Поиск кода в справочнике доходов - логика из 1С
        
        Если код не найден, возвращает None (не создает новую запись).
        Уровень определяется из справочника.
        """
        try:
            # Очищаем код от пробелов для поиска
            code_clean = code.replace(' ', '').replace('\xa0', '')
            
            # Сначала пытаемся найти в новой таблице ref_income_codes
            import sqlite3
            with sqlite3.connect(self.db_manager.db_path) as conn:
                cursor = conn.cursor()
                cursor.execute(
                    "SELECT name FROM sqlite_master WHERE type='table' AND name='ref_income_codes'"
                )
                if cursor.fetchone():
                    # Используем новую таблицу
                    cursor.execute(
                        'SELECT код, название, уровень, наименование FROM ref_income_codes WHERE код = ?',
                        (code_clean,)
                    )
                    row = cursor.fetchone()
                    if row:
                        return {
                            'код': row[0],
                            'название': row[1] or row[3] if len(row) > 3 else row[1] or '',
                            'уровень': int(row[2]) if row[2] is not None else 0,
                            'наименование': row[3] if len(row) > 3 and row[3] else row[1] or ''
                        }
            
            # Если не найдено в новой таблице, используем старую
            income_df = self.db_manager.load_income_reference_df()
            if income_df is not None and not income_df.empty:
                match = income_df[income_df['код_классификации_ДБ'] == code_clean]
                if not match.empty:
                    row = match.iloc[0]
                    return {
                        'код': row.get('код_классификации_ДБ', code_clean),
                        'название': row.get('наименование', ''),
                        'уровень': int(row.get('уровень_кода', 0)) if pd.notna(row.get('уровень_кода')) else 0,
                        'наименование': row.get('наименование', '')
                    }
        except Exception as e:
            logger.warning(f"Ошибка поиска кода дохода {code}: {e}")
        
        # Код не найден - возвращаем None (не создаем новую запись)
        return None
    
    def _find_or_create_income_code(self, code: str, наименование: str = '') -> Dict[str, Any]:
        """Поиск кода в справочнике или создание новой записи"""
        код_д = self._find_income_code(code)
        
        if not код_д:
            # Создаем новую запись в справочнике
            код_д = self._create_income_code(code, наименование)
            logger.info(f"Создана новая запись в справочнике доходов: {code}")
        
        return код_д
    
    def _create_income_code(self, code: str, наименование: str = '') -> Dict[str, Any]:
        """Создание новой записи в справочнике доходов"""
        # Определяем уровень по структуре кода
        уровень = self._determine_income_level(code)
        
        # Сохраняем в БД
        try:
            with sqlite3.connect(self.db_manager.db_path) as conn:
                cursor = conn.cursor()
                
                # Проверяем, существует ли таблица ref_income_codes
                cursor.execute(
                    "SELECT name FROM sqlite_master WHERE type='table' AND name='ref_income_codes'"
                )
                if cursor.fetchone():
                    # Сохраняем в ref_income_codes
                    cursor.execute('''
                        INSERT OR IGNORE INTO ref_income_codes (код, название, уровень, наименование)
                        VALUES (?, ?, ?, ?)
                    ''', (code, наименование or '', уровень, наименование or ''))
                    conn.commit()
                else:
                    # Если таблицы нет, пытаемся сохранить в income_reference_records
                    cursor.execute('''
                        INSERT OR IGNORE INTO income_reference_records (code, name, level)
                        VALUES (?, ?, ?)
                    ''', (code, наименование or '', уровень))
                    conn.commit()
        except Exception as e:
            logger.warning(f"Ошибка создания записи в справочнике доходов: {e}")
        
        return {
            'код': code,
            'название': наименование or '',
            'уровень': уровень,
            'наименование': наименование or ''
        }
    
    def _determine_income_level(self, code: str) -> int:
        """
        Определение уровня кода доходов по структуре кода
        
        Args:
            code: 20-разрядный код доходов
            
        Returns:
            Уровень кода (1-6)
        """
        if not code or len(code.replace(' ', '')) != 20:
            return 0
        
        code = code.replace(' ', '').replace('\xa0', '')
        
        # Логика определения уровня по структуре кода доходов:
        # код_ГАДБ(3) + код_группы_ДБ(1) + код_подгруппы_ДБ(2) + 
        # код_статьи_подстатьи_ДБ(5) + код_элемента_ДБ(2) + 
        # код_группы_ПДБ(4) + код_группы_АПДБ(3)
        
        # Уровень 1: если код_группы_ДБ != '0'
        if code[3:4] != '0':
            # Уровень 2: если код_подгруппы_ДБ != '00'
            if code[4:6] != '00':
                # Уровень 3: если код_статьи_подстатьи_ДБ != '00000'
                if code[6:11] != '00000':
                    # Уровень 4: если код_элемента_ДБ != '00'
                    if code[11:13] != '00':
                        # Уровень 5: если код_группы_ПДБ != '0000'
                        if code[13:17] != '0000':
                            # Уровень 6: если код_группы_АПДБ != '000'
                            if code[17:20] != '000':
                                return 6
                            return 5
                        return 4
                    return 3
                return 2
            return 1
        
        return 0
    
    def _find_expense_code(self, code: str) -> Optional[Dict[str, Any]]:
        """
        Поиск кода в справочнике расходов - логика из 1С
        
        Если код не найден, возвращает None (не создает новую запись).
        Уровень определяется из справочника.
        """
        try:
            # Используем справочник расходов из БД
            expense_df = self.db_manager.load_expense_reference_df()
            if expense_df is None or expense_df.empty:
                return {'пустой': True}
            
            # Очищаем код от пробелов для поиска
            code_clean = code.replace(' ', '').replace('\xa0', '')
            
            # Ищем код в DataFrame по полю 'наименование' (это составной код)
            # Также можно искать по полю 'код' (полный 20-разрядный код)
            match = expense_df[
                (expense_df['наименование'] == code_clean) | 
                (expense_df['код'] == code_clean)
            ]
            
            if not match.empty:
                row = match.iloc[0]
                return {
                    'код': row.get('код', code_clean),
                    'название': row.get('наименование', ''),
                    'уровень': int(row.get('уровень', 0)) if pd.notna(row.get('уровень')) else 0,
                    'код_Р': row.get('код_Р', ''),
                    'код_ПР': row.get('код_ПР', ''),
                    'код_ЦС': row.get('код_ЦС', ''),
                    'код_ВР': row.get('код_ВР', ''),
                    'пустой': False
                }
        except Exception as e:
            logger.warning(f"Ошибка поиска кода расхода {code}: {e}")
        
        # Код не найден - возвращаем пустой результат (не создаем новую запись)
        return {'пустой': True}
    
    def _find_or_create_expense_code(
        self, 
        code: str, 
        код_Р: str = '',
        код_ПР: str = '',
        код_ЦС: str = '',
        код_ВР: str = '',
        наименование: str = ''
    ) -> Dict[str, Any]:
        """Поиск кода в справочнике или создание новой записи"""
        код_р = self._find_expense_code(code)
        
        if not код_р or код_р.get('пустой', True):
            # Создаем новую запись в справочнике
            код_р = self._create_expense_code(code, код_Р, код_ПР, код_ЦС, код_ВР, наименование)
            logger.info(f"Создана новая запись в справочнике расходов: {code}")
        
        return код_р
    
    def _create_expense_code(
        self,
        code: str,
        код_Р: str = '',
        код_ПР: str = '',
        код_ЦС: str = '',
        код_ВР: str = '',
        наименование: str = ''
    ) -> Dict[str, Any]:
        """Создание новой записи в справочнике расходов"""
        # Определяем уровень по структуре кода
        уровень = self._determine_expense_level(code)
        
        # Сохраняем в БД
        try:
            with sqlite3.connect(self.db_manager.db_path) as conn:
                cursor = conn.cursor()
                
                # Проверяем, существует ли таблица ref_expense_codes
                cursor.execute(
                    "SELECT name FROM sqlite_master WHERE type='table' AND name='ref_expense_codes'"
                )
                if cursor.fetchone():
                    # Сохраняем в ref_expense_codes
                    cursor.execute('''
                        INSERT OR IGNORE INTO ref_expense_codes 
                        (код, название, уровень, код_Р, код_ПР, код_ЦС, код_ВР)
                        VALUES (?, ?, ?, ?, ?, ?, ?)
                    ''', (code, наименование or '', уровень, код_Р, код_ПР, код_ЦС, код_ВР))
                    conn.commit()
        except Exception as e:
            logger.warning(f"Ошибка создания записи в справочнике расходов: {e}")
        
        return {
            'код': code,
            'название': наименование or '',
            'уровень': уровень,
            'код_Р': код_Р,
            'код_ПР': код_ПР,
            'код_ЦС': код_ЦС,
            'код_ВР': код_ВР,
            'пустой': False
        }
    
    def _determine_expense_level(self, code: str) -> int:
        """
        Определение уровня кода расходов по структуре кода
        
        Args:
            code: 20-разрядный код расходов
            
        Returns:
            Уровень кода (0-5)
        """
        if not code or len(code.replace(' ', '')) != 20:
            return 0
        
        code = code.replace(' ', '').replace('\xa0', '')
        
        # Логика определения уровня по структуре кода расходов:
        # ГРБС(3) + код_Р(2) + код_ПР(2) + код_ЦС(10) + код_ВР(3)
        
        level = 0
        
        if code[3:5] != '00':
            level = 1
        
        if level == 1 and code[5:7] != '00':
            level = 2
        
        if level == 2 and code[17] != '0':
            level = 3
        
        if level == 3 and code[18] != '0':
            level = 4
        
        if level == 4 and code[19] != '0':
            level = 5
        
        return level
    
    def _parse_number(self, text: str) -> float:
        """Парсинг числа из текста - логика из 1С"""
        if not text:
            return 0.0
        
        # Убираем пробелы (как в 1С: СтрЗаменить(..., " ", ""))
        text = text.replace(' ', '').replace(',', '.')
        
        try:
            return float(text)
        except ValueError:
            return 0.0
    
    def _group_by_lvl_tt(self, data: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """Группировка по уровню и ТТ (аналог Свернуть в 1С)"""
        grouped = {}
        for item in data:
            key = (item.get('уровень', 0), item.get('ТТ', 0))
            if key not in grouped:
                grouped[key] = {
                    'уровень': item.get('уровень', 0),
                    'ТТ': item.get('ТТ', 0),
                    'сумма1': 0,
                    'сумма2': 0,
                    'сумма3': 0,
                    'код': item.get('код', ''),
                    'наименование': item.get('наименование', '')
                }
            grouped[key]['сумма1'] += item.get('сумма1', 0)
            grouped[key]['сумма2'] += item.get('сумма2', 0)
            grouped[key]['сумма3'] += item.get('сумма3', 0)
        
        return list(grouped.values())
    
    def _group_by_expense_codes(self, data: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """Группировка по уровню, код_Р, код_ПР, код_ЦС, код_ВР"""
        grouped = {}
        for item in data:
            key = (
                item.get('уровень', 0),
                item.get('код_Р', ''),
                item.get('код_ПР', ''),
                item.get('код_ЦС', ''),
                item.get('код_ВР', '')
            )
            if key not in grouped:
                grouped[key] = item.copy()
                grouped[key]['сумма1'] = 0
                grouped[key]['сумма2'] = 0
                grouped[key]['сумма3'] = 0
            grouped[key]['сумма1'] += item.get('сумма1', 0)
            grouped[key]['сумма2'] += item.get('сумма2', 0)
            grouped[key]['сумма3'] += item.get('сумма3', 0)
        
        return list(grouped.values())
    
    def _group_by_grbs(self, data: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """Группировка по уровню, ГРБС, код_Р"""
        grouped = {}
        for item in data:
            key = (
                item.get('уровень', 0),
                item.get('ГРБС', ''),
                item.get('код_Р', '')
            )
            if key not in grouped:
                grouped[key] = item.copy()
                grouped[key]['сумма1'] = 0
            grouped[key]['сумма1'] += item.get('сумма1', 0)
        
        return list(grouped.values())
    
    def save_solution_data(
        self,
        project_id: int,
        solution_data: Dict[str, Any],
        file_path: str
    ) -> Optional[int]:
        """
        Сохранение данных решения в БД с учетом структуры кодов
        
        Args:
            project_id: ID проекта
            solution_data: Словарь с распарсенными данными решения
            file_path: Путь к файлу решения
            
        Returns:
            ID сохраненного решения или None при ошибке
        """
        try:
            with sqlite3.connect(self.db_manager.db_path) as conn:
                cursor = conn.cursor()
                
                # Сохраняем основную запись решения
                cursor.execute('''
                    INSERT INTO solution_data (project_id, solution_file_path, parsed_at)
                    VALUES (?, ?, ?)
                ''', (project_id, file_path, datetime.now().isoformat()))
                solution_id = cursor.lastrowid
                
                # Сохраняем Приложение 1 (доходы)
                for item in solution_data.get('приложение1', []):
                    # Проверяем наличие полей для компонентов кода
                    код_ГАДБ = item.get('код_ГАДБ', '')
                    код_группы_ДБ = item.get('код_группы_ДБ', '')
                    код_подгруппы_ДБ = item.get('код_подгруппы_ДБ', '')
                    код_статьи_подстатьи_ДБ = item.get('код_статьи_подстатьи_ДБ', '')
                    код_элемента_ДБ = item.get('код_элемента_ДБ', '')
                    код_группы_ПДБ = item.get('код_группы_ПДБ', '')
                    код_группы_АПДБ = item.get('код_группы_АПДБ', '')
                    
                    # Проверяем, есть ли поля для компонентов в таблице
                    cursor.execute("PRAGMA table_info(solution_income_data)")
                    columns = [col[1] for col in cursor.fetchall()]
                    has_components = 'код_ГАДБ' in columns
                    
                    if has_components:
                        cursor.execute('''
                            INSERT INTO solution_income_data 
                            (solution_id, код, наименование, уровень, ТТ, сумма1, сумма2, сумма3,
                             код_ГАДБ, код_группы_ДБ, код_подгруппы_ДБ, код_статьи_подстатьи_ДБ,
                             код_элемента_ДБ, код_группы_ПДБ, код_группы_АПДБ)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        ''', (
                            solution_id,
                            item.get('код', ''),
                            item.get('наименование', ''),
                            item.get('уровень', 0),
                            item.get('ТТ', 0),
                            item.get('сумма1', 0),
                            item.get('сумма2', 0),
                            item.get('сумма3', 0),
                            код_ГАДБ,
                            код_группы_ДБ,
                            код_подгруппы_ДБ,
                            код_статьи_подстатьи_ДБ,
                            код_элемента_ДБ,
                            код_группы_ПДБ,
                            код_группы_АПДБ
                        ))
                    else:
                        # Если поля компонентов еще не добавлены, сохраняем только основные поля
                        cursor.execute('''
                            INSERT INTO solution_income_data 
                            (solution_id, код, наименование, уровень, ТТ, сумма1, сумма2, сумма3)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                        ''', (
                            solution_id,
                            item.get('код', ''),
                            item.get('наименование', ''),
                            item.get('уровень', 0),
                            item.get('ТТ', 0),
                            item.get('сумма1', 0),
                            item.get('сумма2', 0),
                            item.get('сумма3', 0)
                        ))
                
                # Сохраняем Приложение 2 (расходы)
                for item in solution_data.get('приложение2', []):
                    # Проверяем наличие поля ГРБС в таблице
                    cursor.execute("PRAGMA table_info(solution_expense_data)")
                    columns = [col[1] for col in cursor.fetchall()]
                    has_grbs = 'ГРБС' in columns
                    
                    if has_grbs:
                        cursor.execute('''
                            INSERT INTO solution_expense_data 
                            (solution_id, ГРБС, код_Р, код_ПР, код_ЦС, код_ВР, уровень, сумма1, сумма2, сумма3, наименование)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        ''', (
                            solution_id,
                            item.get('ГРБС', '000'),
                            item.get('код_Р', ''),
                            item.get('код_ПР', ''),
                            item.get('код_ЦС', ''),
                            item.get('код_ВР', ''),
                            item.get('уровень', 0),
                            item.get('сумма1', 0),
                            item.get('сумма2', 0),
                            item.get('сумма3', 0),
                            item.get('наименование', '')
                        ))
                    else:
                        # Если поле ГРБС еще не добавлено, сохраняем без него
                        cursor.execute('''
                            INSERT INTO solution_expense_data 
                            (solution_id, код_Р, код_ПР, код_ЦС, код_ВР, уровень, сумма1, сумма2, сумма3, наименование)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        ''', (
                            solution_id,
                            item.get('код_Р', ''),
                            item.get('код_ПР', ''),
                            item.get('код_ЦС', ''),
                            item.get('код_ВР', ''),
                            item.get('уровень', 0),
                            item.get('сумма1', 0),
                            item.get('сумма2', 0),
                            item.get('сумма3', 0),
                            item.get('наименование', '')
                        ))
                
                # Сохраняем Приложение 3 (расходы по ГРБС)
                for item in solution_data.get('приложение3', []):
                    cursor.execute('''
                        INSERT INTO solution_expense_grbs_data 
                        (solution_id, ГРБС, код_Р, код_ПР, код_ЦС, код_ВР, уровень, сумма1, наименование)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                    ''', (
                        solution_id,
                        item.get('ГРБС', ''),
                        item.get('код_Р', ''),
                        item.get('код_ПР', ''),
                        item.get('код_ЦС', ''),
                        item.get('код_ВР', ''),
                        item.get('уровень', 0),
                        item.get('сумма1', 0),
                        item.get('наименование', '')
                    ))
                
                conn.commit()
                logger.info(f"Данные решения сохранены в БД: solution_id={solution_id}")
                return solution_id
                
        except Exception as e:
            error_msg = f"Ошибка сохранения данных решения: {e}"
            logger.error(error_msg, exc_info=True)
            self.error_occurred.emit(error_msg)
            return None
    
    def aggregate_solution_data(self, solution_id: int) -> Dict[str, Any]:
        """
        Агрегация данных решения по уровням
        
        Args:
            solution_id: ID решения в БД
            
        Returns:
            Словарь с агрегированными данными:
            {
                'доходы': агрегированные данные доходов,
                'расходы': агрегированные данные расходов
            }
        """
        try:
            # Загружаем данные из БД
            income_data = self._load_solution_income_data(solution_id)
            expense_data = self._load_solution_expense_data(solution_id)
            
            # Агрегируем по уровням
            aggregated_income = self._aggregate_by_level(income_data, 'доходы')
            aggregated_expense = self._aggregate_by_level(expense_data, 'расходы')
            
            return {
                'доходы': aggregated_income,
                'расходы': aggregated_expense
            }
        except Exception as e:
            logger.error(f"Ошибка агрегации данных решения: {e}", exc_info=True)
            return {'доходы': [], 'расходы': []}
    
    def _load_solution_income_data(self, solution_id: int) -> List[Dict[str, Any]]:
        """Загрузка данных доходов решения из БД"""
        try:
            with sqlite3.connect(self.db_manager.db_path) as conn:
                cursor = conn.cursor()
                cursor.execute('''
                    SELECT код, наименование, уровень, ТТ, сумма1, сумма2, сумма3
                    FROM solution_income_data
                    WHERE solution_id = ?
                    ORDER BY уровень, ТТ
                ''', (solution_id,))
                
                rows = cursor.fetchall()
                return [
                    {
                        'код': row[0],
                        'наименование': row[1],
                        'уровень': row[2],
                        'ТТ': row[3],
                        'сумма1': row[4],
                        'сумма2': row[5],
                        'сумма3': row[6]
                    }
                    for row in rows
                ]
        except Exception as e:
            logger.error(f"Ошибка загрузки данных доходов решения: {e}", exc_info=True)
            return []
    
    def _load_solution_expense_data(self, solution_id: int) -> List[Dict[str, Any]]:
        """Загрузка данных расходов решения из БД"""
        try:
            with sqlite3.connect(self.db_manager.db_path) as conn:
                cursor = conn.cursor()
                cursor.execute('''
                    SELECT код_Р, код_ПР, код_ЦС, код_ВР, уровень, сумма1, сумма2, сумма3, наименование
                    FROM solution_expense_data
                    WHERE solution_id = ?
                    ORDER BY уровень, код_Р, код_ПР, код_ЦС, код_ВР
                ''', (solution_id,))
                
                rows = cursor.fetchall()
                return [
                    {
                        'код_Р': row[0],
                        'код_ПР': row[1],
                        'код_ЦС': row[2],
                        'код_ВР': row[3],
                        'уровень': row[4],
                        'сумма1': row[5],
                        'сумма2': row[6],
                        'сумма3': row[7],
                        'наименование': row[8]
                    }
                    for row in rows
                ]
        except Exception as e:
            logger.error(f"Ошибка загрузки данных расходов решения: {e}", exc_info=True)
            return []
    
    def _aggregate_by_level(self, data: List[Dict[str, Any]], data_type: str) -> List[Dict[str, Any]]:
        """
        Агрегация данных по уровням (свертка сумм от дочерних уровней к родительским)
        
        Args:
            data: Список данных для агрегации
            data_type: Тип данных ('доходы' или 'расходы')
            
        Returns:
            Список агрегированных данных
        """
        if not data:
            return []
        
        # Сортируем данные по уровню (от большего к меньшему для правильной агрегации)
        sorted_data = sorted(data, key=lambda x: x.get('уровень', 0), reverse=True)
        
        # Словарь для хранения агрегированных сумм по уровням
        aggregated = {}
        
        for item in sorted_data:
            уровень = item.get('уровень', 0)
            
            # Инициализируем ключ для уровня, если его еще нет
            if уровень not in aggregated:
                aggregated[уровень] = {}
            
            # Формируем ключ для группировки (зависит от типа данных)
            if data_type == 'доходы':
                # Для доходов группируем по коду и ТТ
                key = (item.get('код', ''), item.get('ТТ', 0))
            else:
                # Для расходов группируем по компонентам кода
                key = (
                    item.get('код_Р', ''),
                    item.get('код_ПР', ''),
                    item.get('код_ЦС', ''),
                    item.get('код_ВР', '')
                )
            
            if key not in aggregated[уровень]:
                aggregated[уровень][key] = {
                    'уровень': уровень,
                    'сумма1': 0,
                    'сумма2': 0,
                    'сумма3': 0
                }
                # Копируем остальные поля
                for k, v in item.items():
                    if k not in ['сумма1', 'сумма2', 'сумма3']:
                        aggregated[уровень][key][k] = v
            
            # Суммируем значения
            aggregated[уровень][key]['сумма1'] += item.get('сумма1', 0)
            aggregated[уровень][key]['сумма2'] += item.get('сумма2', 0)
            aggregated[уровень][key]['сумма3'] += item.get('сумма3', 0)
        
        # Формируем итоговый список, сортируя по уровню (от меньшего к большему)
        result = []
        for уровень in sorted(aggregated.keys()):
            for key_data in aggregated[уровень].values():
                result.append(key_data)
        
        return result


"""
Контроллер для обработки решений о бюджете
Реализует логику из 1С: Решение
"""
from PyQt5.QtCore import QObject, pyqtSignal
from typing import Dict, Any, Optional, List
from pathlib import Path
import docx
import pandas as pd
from logger import logger

from models.database import DatabaseManager


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
                
                if код_д and код_д.get('уровень'):
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
                    
                    result.append({
                        'код_Р': код_р_значение,
                        'код_ПР': столбец3,
                        'код_ЦС': столбец4,
                        'код_ВР': столбец5,
                        'уровень': уровень,
                        'сумма1': сумма1,
                        'сумма2': сумма2,
                        'сумма3': сумма3,
                        'наименование': наименование
                    })
                else:
                    # Если код не найден, пытаемся найти по частям или создать новый
                    # TODO: реализовать создание нового кода в справочнике
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
                    
                    result.append({
                        'ГРБС': грбс,
                        'код_Р': код_р_значение,
                        'код_ПР': столбец4,
                        'код_ЦС': столбец5,
                        'код_ВР': столбец6,
                        'уровень': уровень,
                        'сумма1': сумма1,
                        'наименование': наименование
                    })
        
        # Группируем по уровню, ГРБС, код_Р
        grouped_result = self._group_by_grbs(result)
        
        return grouped_result
    
    def _find_income_code(self, code: str) -> Optional[Dict[str, Any]]:
        """Поиск кода в справочнике доходов - логика из 1С"""
        try:
            # Используем справочник доходов из БД
            income_df = self.db_manager.load_income_reference_df()
            if income_df is None or income_df.empty:
                return None
            
            # Ищем код в DataFrame
            match = income_df[income_df['код_классификации_ДБ'] == code]
            if not match.empty:
                row = match.iloc[0]
                return {
                    'код': row.get('код_классификации_ДБ', code),
                    'название': row.get('наименование', ''),
                    'уровень': int(row.get('уровень_кода', 0)) if pd.notna(row.get('уровень_кода')) else 0,
                    'наименование': row.get('наименование', '')
                }
        except Exception as e:
            logger.warning(f"Ошибка поиска кода дохода {code}: {e}")
        
        return None
    
    def _find_expense_code(self, code: str) -> Optional[Dict[str, Any]]:
        """Поиск кода в справочнике расходов - логика из 1С"""
        try:
            import pandas as pd
            # TODO: добавить метод load_expense_reference_df в DatabaseManager
            # Пока используем прямое обращение к БД
            import sqlite3
            conn = sqlite3.connect(self.db_manager.db_path)
            
            # Ищем в таблице расходов (нужно создать таблицу ref_expense_codes)
            # Пока возвращаем None, если таблица не существует
            cursor = conn.cursor()
            cursor.execute(
                "SELECT name FROM sqlite_master WHERE type='table' AND name='ref_expense_codes'"
            )
            if cursor.fetchone():
                cursor.execute(
                    'SELECT код, название, уровень, код_Р, код_ПР, код_ЦС, код_ВР FROM ref_expense_codes WHERE наименование = ?',
                    (code,)
                )
                row = cursor.fetchone()
                conn.close()
                if row:
                    return {
                        'код': row[0],
                        'название': row[1],
                        'уровень': row[2],
                        'код_Р': row[3],
                        'код_ПР': row[4],
                        'код_ЦС': row[5],
                        'код_ВР': row[6],
                        'пустой': False
                    }
            else:
                conn.close()
                # Если таблица не существует, возвращаем пустой результат
                return {'пустой': True}
        except Exception as e:
            logger.warning(f"Ошибка поиска кода расхода {code}: {e}")
        
        return {'пустой': True}
    
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


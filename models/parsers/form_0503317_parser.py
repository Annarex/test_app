"""Парсер Excel файлов для формы 0503317"""
import pandas as pd
import re
from typing import Dict, List, Any, Optional
from logger import logger
from models.constants.form_0503317_constants import Form0503317Constants
from models.utils.form_utils import (
    column_to_index, get_cell_value, get_numeric_value,
    clean_dbk_code, format_classification_code
)


class Form0503317Parser:
    """Парсер для Excel файлов формы 0503317"""
    
    def __init__(self, constants: Form0503317Constants):
        """
        Args:
            constants: Константы формы 0503317
        """
        self.constants = constants
        self.reference_data_доходы = None
        self.reference_data_источники = None
    
    def parse_excel(
        self, 
        file_path: str, 
        reference_data_доходы: Optional[pd.DataFrame] = None, 
        reference_data_источники: Optional[pd.DataFrame] = None
    ) -> Dict[str, Any]:
        """Парсинг Excel файла формы 0503317
        
        Args:
            file_path: Путь к Excel файлу
            reference_data_доходы: DataFrame со справочником доходов
            reference_data_источники: DataFrame со справочником источников
        
        Returns:
            Словарь с распарсенными данными:
            - meta_info: метаданные формы
            - доходы_data: данные доходов
            - расходы_data: данные расходов
            - источники_финансирования_data: данные источников
            - консолидируемые_расчеты_data: данные консолидированных расчетов
        """
        self.reference_data_доходы = reference_data_доходы
        self.reference_data_источники = reference_data_источники
        
        # Инициализация результатов
        meta_info = {}
        доходы_data = []
        расходы_data = []
        источники_финансирования_data = []
        консолидируемые_расчеты_data = []
        zero_columns = {}
        
        # Загрузка листов
        sheets = {}
        for sheet_name in self.constants.SHEETS:
            try:
                sheets[sheet_name] = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
                logger.debug(f"Лист '{sheet_name}' загружен")
            except Exception as e:
                logger.warning(f"Лист '{sheet_name}' не найден: {e}")
        
        # Извлечение метаданных
        if 'стр. 1-2' in sheets:
            meta_info = self._extract_metadata(sheets['стр. 1-2'])
        
        # Извлечение данных разделов
        for section_type in self.constants.SECTION_SHEETS.keys():
            sheet_name = self.constants.SECTION_SHEETS[section_type]
            if sheet_name in sheets:
                if section_type == 'консолидируемые_расчеты':
                    section_data, section_zero_cols = self._extract_consolidated_data(sheets[sheet_name])
                    консолидируемые_расчеты_data = section_data
                else:
                    section_data, section_zero_cols = self._extract_section_data(sheets[sheet_name], section_type)
                    if section_type == 'доходы':
                        доходы_data = section_data
                    elif section_type == 'расходы':
                        расходы_data = section_data
                    elif section_type == 'источники_финансирования':
                        источники_финансирования_data = section_data
                    
                    if section_zero_cols:
                        zero_columns[section_type] = section_zero_cols
        
        return {
            'meta_info': meta_info,
            'доходы_data': доходы_data,
            'расходы_data': расходы_data,
            'источники_финансирования_data': источники_финансирования_data,
            'консолидируемые_расчеты_data': консолидируемые_расчеты_data,
            'zero_columns': zero_columns
        }
    
    def _extract_metadata(self, sheet: pd.DataFrame) -> Dict[str, str]:
        """Извлечение метаданных из листа
        
        Args:
            sheet: DataFrame с данными листа
        
        Returns:
            Словарь с метаданными
        """
        return {
            'Наименование формы': get_cell_value(sheet, 2, 1) + ' ' + get_cell_value(sheet, 3, 1),
            'Наименование финансового органа': get_cell_value(sheet, 6, 3),
            'Наименование бюджета': get_cell_value(sheet, 7, 3),
            'Периодичность': get_cell_value(sheet, 8, 17),
            'Форма по ОКУД': get_cell_value(sheet, 4, 17),
            'Дата': get_cell_value(sheet, 5, 17),
            'код ОКПО': get_cell_value(sheet, 6, 17),
            'код ОКТМО': get_cell_value(sheet, 7, 17),
            'код ОКЕИ': get_cell_value(sheet, 9, 17),
        }
    
    def _find_section_start(self, sheet: pd.DataFrame, section_type: str, search_column: int = 0) -> Optional[int]:
        """Поиск начала раздела в листе
        
        Args:
            sheet: DataFrame с данными листа
            section_type: Тип раздела
            search_column: Колонка для поиска
        
        Returns:
            Номер строки начала раздела или None
        """
        pattern = self.constants.SECTION_PATTERNS.get(section_type)
        if not pattern:
            return None
            
        for i in range(len(sheet)):
            cell_value = str(sheet.iloc[i, search_column]) if pd.notna(sheet.iloc[i, search_column]) else ''
            if re.search(pattern, cell_value, re.IGNORECASE):
                return i
                
        return None
    
    def _extract_section_data(self, sheet: pd.DataFrame, section_type: str) -> tuple:
        """Извлечение данных раздела
        
        Args:
            sheet: DataFrame с данными листа
            section_type: Тип раздела
        
        Returns:
            Кортеж (список данных, список нулевых колонок)
        """
        start_row = self._find_section_start(sheet, section_type)
        
        if start_row is None:
            logger.warning(f"Раздел '{section_type}' не найден")
            return [], []
            
        logger.debug(f"Раздел '{section_type}' начинается со строки {start_row + 1}")
        header_row = start_row + 3
        return self._extract_table_data(sheet, header_row, section_type)
    
    def _extract_table_data(self, sheet: pd.DataFrame, header_row: int, section_type: str) -> tuple:
        """Извлечение табличных данных
        
        Args:
            sheet: DataFrame с данными листа
            header_row: Номер строки заголовка
            section_type: Тип раздела
        
        Returns:
            Кортеж (список данных, список нулевых колонок)
        """
        mapping = self.constants.COLUMN_MAPPING[section_type]
        budget_columns = self.constants.BUDGET_COLUMNS
        
        total_row_data = None
        data_start_row = header_row + 2
        data = []

        for row_idx in range(data_start_row, len(sheet)):
            row_data = self._extract_row_data(sheet, row_idx, mapping, budget_columns, section_type)
            if row_data and row_data['наименование_показателя']:
                if self._is_total_row(row_data, section_type):
                    total_row_data = row_data
                
                data.append(row_data)
        
        # Пересчет уровней для источников финансирования
        if section_type == 'источники_финансирования':
            self._recalculate_sources_levels(data)
        
        # Определение нулевых колонок
        zero_columns = []
        if total_row_data:
            zero_columns = self._get_zero_columns(total_row_data, budget_columns)
            logger.debug(f"Для раздела '{section_type}' нулевые столбцы: {zero_columns}")
        
        return data, zero_columns
    
    def _extract_row_data(
        self, 
        sheet: pd.DataFrame, 
        row_idx: int, 
        mapping: dict, 
        budget_columns: list, 
        section_type: str
    ) -> Optional[dict]:
        """Извлечение данных строки
        
        Args:
            sheet: DataFrame с данными листа
            row_idx: Индекс строки
            mapping: Маппинг колонок
            budget_columns: Список бюджетных колонок
            section_type: Тип раздела
        
        Returns:
            Словарь с данными строки или None
        """
        common_data = {}
        for i, col in enumerate(mapping['common_cols']):
            col_idx = column_to_index(col)
            value = get_cell_value(sheet, row_idx, col_idx)
            
            if i == 0:
                value = value.replace("в том числе:\n", "").replace("\n", "").strip()
                common_data['наименование_показателя'] = value
            elif i == 1:
                value = value.replace("\n", "").strip()
                common_data['код_строки'] = value
            elif i == 2:
                value = value.replace("\n", "").strip()
                clean_code = clean_dbk_code(value)
                common_data['код_классификации'] = clean_code
                common_data['код_классификации_форматированный'] = format_classification_code(clean_code, section_type)

        if not common_data.get('наименование_показателя'):
            return None

        level = self._determine_level(
            common_data['код_классификации'], 
            section_type, 
            common_data['наименование_показателя']
        )
        common_data['уровень'] = level

        утвержденный = self._extract_budget_data(sheet, row_idx, mapping['утвержденный'], budget_columns)
        исполненный = self._extract_budget_data(sheet, row_idx, mapping['исполненный'], budget_columns)

        return {
            **common_data,
            'раздел': section_type,
            'утвержденный': утвержденный,
            'исполненный': исполненный,
            'исходная_строка': row_idx + 1,
        }
    
    def _extract_budget_data(
        self, 
        sheet: pd.DataFrame, 
        row_idx: int, 
        columns: list, 
        budget_columns: list
    ) -> dict:
        """Извлечение данных по бюджету
        
        Args:
            sheet: DataFrame с данными листа
            row_idx: Индекс строки
            columns: Список колонок для извлечения
            budget_columns: Список бюджетных колонок
        
        Returns:
            Словарь с данными по бюджету
        """
        data = {}
        for i, col in enumerate(columns):
            col_idx = column_to_index(col)
            value = get_numeric_value(sheet, row_idx, col_idx)
            data[budget_columns[i]] = value
        return data
    
    def _determine_level(self, classification_code: str, section_type: str, name: str = "") -> int:
        """Определение уровня строки
        
        Args:
            classification_code: Код классификации
            section_type: Тип раздела
            name: Наименование показателя
        
        Returns:
            Уровень строки (0-6)
        """
        # Для доходов и источников используем справочники
        if section_type == 'доходы':
            if classification_code == '00000000000000000000':
                return 0
            return self._get_level_from_reference(classification_code, section_type)
        
        if section_type == 'источники_финансирования':
            if classification_code == '00000000000000000000':
                return 0
            name_lower = name.lower()
            if ('источники финансирования дефицита бюджетов - всего' in name_lower or
                'источники внутреннего финансирования' in name_lower or
                'источники внешнего финансирования' in name_lower):
                return 1
            return self._get_level_from_reference(classification_code, section_type)
        
        if section_type == 'расходы':
            return self._determine_expenditure_level(classification_code)
        
        return 0
    
    def _determine_expenditure_level(self, code: str) -> int:
        """Определение уровня для расходов
        
        Args:
            code: Код классификации (20 символов)
        
        Returns:
            Уровень (0-5)
        """
        if len(code) != 20:
            return 0
            
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
    
    def _get_level_from_reference(self, classification_code: str, section_type: str) -> int:
        """Получение уровня из справочника
        
        Args:
            classification_code: Код классификации
            section_type: Тип раздела
        
        Returns:
            Уровень из справочника или 0
        """
        try:
            if section_type == 'доходы' and self.reference_data_доходы is not None:
                match = self.reference_data_доходы[
                    self.reference_data_доходы['код_классификации_ДБ'] == classification_code
                ]
                if not match.empty:
                    level = match.iloc[0]['уровень_кода']
                    return int(level) if pd.notna(level) else 0
            
            elif section_type == 'источники_финансирования' and self.reference_data_источники is not None:
                match = self.reference_data_источники[
                    self.reference_data_источники['код_классификации_ИФДБ'] == classification_code
                ]
                if not match.empty:
                    level = match.iloc[0]['уровень_кода']
                    return int(level) if pd.notna(level) else 0
        
        except Exception as e:
            logger.warning(f"Ошибка получения уровня из справочника: {e}", exc_info=True)
        
        return 0
    
    def _is_total_row(self, row_data: dict, section_type: str) -> bool:
        """Проверка, является ли строка итоговой
        
        Args:
            row_data: Данные строки
            section_type: Тип раздела
        
        Returns:
            True, если строка итоговая
        """
        pattern = self.constants.TOTAL_PATTERNS.get(section_type)
        if not pattern:
            return False
            
        name = row_data['наименование_показателя'].lower()
        if section_type == 'консолидируемые_расчеты':
            return 'Всего выбытий' in row_data['наименование_показателя']
        
        return re.search(pattern, name, re.IGNORECASE) is not None
    
    def _get_zero_columns(self, total_row_data: dict, budget_columns: list) -> list:
        """Получение нулевых столбцов
        
        Столбец считается нулевым, если в итоговой строке оба значения 
        (утвержденный и исполненный) равны 0.
        
        Args:
            total_row_data: Данные итоговой строки
            budget_columns: Список бюджетных колонок
        
        Returns:
            Список индексов нулевых столбцов
        """
        zero_columns = []
        
        approved = total_row_data.get('утвержденный', {}) or {}
        executed = total_row_data.get('исполненный', {}) or {}
        
        for i, budget_col in enumerate(budget_columns):
            a_val = approved.get(budget_col, 0) or 0
            e_val = executed.get(budget_col, 0) or 0
            
            # Проверяем, что оба значения равны 0 (или близки к 0)
            if isinstance(a_val, (int, float)) and isinstance(e_val, (int, float)):
                if abs(a_val) < 1e-9 and abs(e_val) < 1e-9:
                    # Добавляем индекс для утвержденного столбца
                    zero_columns.append(i)
                    # Добавляем индекс для исполненного столбца
                    zero_columns.append(i + len(budget_columns))
        
        return zero_columns
    
    def _recalculate_sources_levels(self, data: List[dict]):
        """Пересчет уровней для источников финансирования
        
        Args:
            data: Список данных источников финансирования (изменяется in-place)
        """
        if not data:
            return
        
        for item in data:
            name = item['наименование_показателя'].lower()
            
            if 'источники финансирования дефицита бюджетов - всего' in name:
                item['уровень'] = 0
                item['is_total_sources'] = True
            elif 'источники внутреннего финансирования' in name:
                item['уровень'] = 1
                item['is_internal_total'] = True
            elif 'источники внешнего финансирования' in name:
                item['уровень'] = 1
                item['is_external_total'] = True
    
    def _extract_consolidated_data(self, sheet: pd.DataFrame) -> tuple:
        """Извлечение данных консолидируемых расчетов
        
        Args:
            sheet: DataFrame с данными листа
        
        Returns:
            Кортеж (список данных, список нулевых колонок - всегда пустой для консолидированных)
        """
        start_row = self._find_section_start(sheet, 'консолидируемые_расчеты', search_column=1)
        
        if start_row is None:
            logger.warning("Таблица консолидируемых расчетов не найдена")
            return [], []
        
        table_start_row = start_row + 2
        data = self._extract_consolidated_table_data(sheet, table_start_row)
        return data, []
    
    def _extract_consolidated_table_data(self, sheet: pd.DataFrame, start_row: int) -> List[dict]:
        """Извлечение данных таблицы консолидируемых расчетов
        
        Args:
            sheet: DataFrame с данными листа
            start_row: Начальная строка таблицы
        
        Returns:
            Список данных консолидируемых расчетов
        """
        mapping = self.constants.COLUMN_MAPPING['консолидируемые_расчеты']
        consolidated_columns = self.constants.CONSOLIDATED_COLUMNS
        
        current_row = start_row
        data_parts = []
        
        while current_row < len(sheet):
            cell_value = str(sheet.iloc[current_row, 1]) if pd.notna(sheet.iloc[current_row, 1]) else ''
            
            if cell_value == 'Наименование показателя':
                data_start_row = current_row + 3
                part_data = self._extract_consolidated_part_data(sheet, data_start_row, mapping, consolidated_columns)
                data_parts.extend(part_data)
                current_row = data_start_row + len(part_data)
            else:
                current_row += 1
        
        logger.info(f"Извлечено {len(data_parts)} записей из таблицы консолидируемых расчетов")
        return data_parts
    
    def _extract_consolidated_part_data(
        self, 
        sheet: pd.DataFrame, 
        start_row: int, 
        mapping: dict, 
        consolidated_columns: list
    ) -> List[dict]:
        """Извлечение данных части таблицы консолидируемых расчетов
        
        Args:
            sheet: DataFrame с данными листа
            start_row: Начальная строка части
            mapping: Маппинг колонок
            consolidated_columns: Список колонок консолидированных расчетов
        
        Returns:
            Список данных части таблицы
        """
        data = []
        current_row = start_row
        
        while current_row < len(sheet):
            name_cell = sheet.iloc[current_row, 1] if pd.notna(sheet.iloc[current_row, 1]) else ''
            
            if not name_cell:
                next_cell = sheet.iloc[current_row + 1, 1] if current_row + 1 < len(sheet) and pd.notna(sheet.iloc[current_row + 1, 1]) else ''
                if next_cell == 'Наименование показателя':
                    break
            
            row_data = self._extract_consolidated_row_data(sheet, current_row, mapping, consolidated_columns)
            if row_data and row_data['наименование_показателя']:
                data.append(row_data)
            
            current_row += 1
        
        return data
    
    def _extract_consolidated_row_data(
        self, 
        sheet: pd.DataFrame, 
        row_idx: int, 
        mapping: dict, 
        consolidated_columns: list
    ) -> Optional[dict]:
        """Извлечение данных строки консолидируемых расчетов
        
        Args:
            sheet: DataFrame с данными листа
            row_idx: Индекс строки
            mapping: Маппинг колонок
            consolidated_columns: Список колонок консолидированных расчетов
        
        Returns:
            Словарь с данными строки или None
        """
        common_data = {}
        for i, col_offset in enumerate([0, 1]):
            col_idx = 1 + col_offset
            value = get_cell_value(sheet, row_idx, col_idx)
            
            if i == 0:
                value = value.replace("\n", "").strip()
                common_data['наименование_показателя'] = value
            elif i == 1:
                value = value.replace("\n", "").strip()
                common_data['код_строки'] = value

        if not common_data.get('наименование_показателя') or not common_data.get('код_строки'):
            return None

        level = self._determine_consolidated_level(common_data['код_строки'])
        common_data['уровень'] = level

        поступления = {}
        for i, col_offset in enumerate(range(11)):
            col_idx = 3 + col_offset
            value = get_numeric_value(sheet, row_idx, col_idx)
            поступления[consolidated_columns[i]] = value

        return {
            **common_data,
            'раздел': 'консолидируемые_расчеты',
            'поступления': поступления,
            'исходная_строка': row_idx + 1,
        }
    
    def _determine_consolidated_level(self, code: str) -> int:
        """Определение уровня для консолидируемых расчетов
        
        Args:
            code: Код строки
        
        Returns:
            Уровень (0-2)
        """
        if not code or len(code) < 3:
            return 0
        
        if code == '899':
            return 0
        
        if code[0] == '9' and code[1] in '0123456789' and code[2] == '0':
            return 1
        
        if code[0] == '9' and code[1] in '0123456789' and code[2] != '0':
            return 2
        
        return 0

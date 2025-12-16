import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
import re
import shutil
from typing import Dict, List, Any, Optional
from datetime import datetime
from .base_models import BaseFormModel, FormType
from logger import logger

class Form0503317Constants:
    """Константы для формы 0503317"""
    
    SHEETS = ['стр. 1-2', 'стр. 3-4', 'стр. 5-6', 'стр. 7-11']
    
    SECTION_SHEETS = {
        'доходы': 'стр. 1-2',
        'расходы': 'стр. 3-4', 
        'источники_финансирования': 'стр. 5-6',
        'консолидируемые_расчеты': 'стр. 7-11'
    }
    
    SECTION_PATTERNS = {
        'доходы': r'1\.\s*Доходы бюджета',
        'расходы': r'2\.\s*Расходы бюджета', 
        'источники_финансирования': r'3\.\s*Источники финансирования дефицита бюджета',
        'консолидируемые_расчеты': r'4\.\s*Таблица консолидируемых расчетов'
    }
    
    TOTAL_PATTERNS = {
        'доходы': r'доходы бюджета.*всего',
        'расходы': r'расходы бюджета.*всего',
        'источники_финансирования': r'источники финансирования дефицита бюджетов.*всего',
        'консолидируемые_расчеты': r'Всего выбытий'
    }
    
    BUDGET_COLUMNS = [
        'консолидированный бюджет субъекта Российской Федерации и территориального государственного внебюджетного фонда',
        'суммы, подлежащие исключению в рамках консолидированного бюджета субъекта Российской Федерации и бюджета территориального государственного внебюджетного фонда',
        'консолидированный бюджет субъекта Российской Федерации',
        'суммы, подлежащие исключению в рамках консолидированного бюджета Российской Федерации',
        'бюджет субъекта Российской Федерации',
        'бюджеты внутригородских муниципальных образований городов федерального значения',
        'бюджеты муниципальных округов',
        'бюджеты городских округов',
        'бюджеты городских округов с внутригородским делением',
        'бюджеты внутригородских районов',
        'бюджеты муниципальных районов',
        'бюджеты городских поселений',
        'бюджеты сельских поселений',
        'бюджет территориального государственного внебюджетного фонда'
    ]
    
    CONSOLIDATED_COLUMNS = [
        'бюджет субъекта Российской Федерации',
        'бюджеты внутригородских муниципальных образований городов федерального значения',
        'бюджеты муниципальных округов',
        'бюджеты городских округов',
        'бюджеты городских округов с внутригородским делением',
        'бюджеты внутригородских районов',
        'бюджеты муниципальных районов',
        'бюджеты городских поселений',
        'бюджеты сельских поселений',
        'бюджет территориального государственного внебюджетного фонда',
        'ИТОГО'
    ]
    
    COLUMN_MAPPING = {
        'доходы': {
            'common_cols': ['A', 'B', 'C'],
            'утвержденный': ['E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R'],
            'исполненный': ['W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ']
        },
        'расходы': {
            'common_cols': ['A', 'B', 'C'],
            'утвержденный': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q'],
            'исполненный': ['U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH']
        },
        'источники_финансирования': {
            'common_cols': ['A', 'B', 'C'],
            'утвержденный': ['E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R'],
            'исполненный': ['W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ']
        },
        'консолидируемые_расчеты': {
            'common_cols': ['B', 'C'],
            'поступления': ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N']
        }
    }

class StyleConstants:
    """Константы стилей"""
    
    LEVEL_COLORS = {
        0: "E6E6FA", 1: "68e368", 2: "98FB98", 3: "FFFF99", 
        4: "FFB366", 5: "FF9999", 6: "FFCCCC"
    }
    
    ERROR_FILL = PatternFill(start_color="FF4444", end_color="FF4444", fill_type="solid")
    HEADER_FONT = Font(bold=True)
    HEADER_FILL = PatternFill(start_color="DDDDDD", end_color="DDDDDD", fill_type="solid")
    POSITIVE_FILL = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")
    NEGATIVE_FILL = PatternFill(start_color="FF9999", end_color="FF9999", fill_type="solid")

class Form0503317(BaseFormModel):
    """Модель для формы 0503317"""
    
    def __init__(self, revision: str = "1.0", column_mapping: Optional[dict] = None):
        super().__init__(FormType.FORM_0503317, revision)
        self.constants = Form0503317Constants()
        if column_mapping:
            # Позволяем переопределять mapping колонок из справочника типов форм
            self.constants.COLUMN_MAPPING = column_mapping
        self.meta_info = {}
        self.доходы_data = []
        self.расходы_data = []
        self.источники_финансирования_data = []
        self.консолидируемые_расчеты_data = []
        self.reference_data_доходы = None
        self.reference_data_источники = None
        self.zero_columns = {}
        self.show_error_values = True
        self.доходы_всего = None
        self.расходы_всего = None
        self.calculated_deficit_proficit = None  # Переименовано из результат_исполнения_data для ясности
    
    def get_form_constants(self):
        return self.constants
    
    def parse_excel(self, file_path: str, reference_data_доходы: pd.DataFrame = None, reference_data_источники: pd.DataFrame = None) -> Dict[str, Any]:
        """Парсинг Excel файла формы 0503317"""
        # Сбрасываем предыдущее состояние, чтобы при повторной загрузке формы
        # данные не дублировались
        self.meta_info = {}
        self.доходы_data = []
        self.расходы_data = []
        self.источники_финансирования_data = []
        self.консолидируемые_расчеты_data = []
        self.zero_columns = {}
        self.доходы_всего = None
        self.расходы_всего = None
        self.calculated_deficit_proficit = None  # Переименовано из результат_исполнения_data для ясности

        self.reference_data_доходы = reference_data_доходы
        self.reference_data_источники = reference_data_источники
        
        sheets = {}
        for sheet_name in self.constants.SHEETS:
            try:
                sheets[sheet_name] = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
                logger.debug(f"Лист '{sheet_name}' загружен")
            except Exception as e:
                logger.warning(f"Лист '{sheet_name}' не найден: {e}")
        
        if 'стр. 1-2' in sheets:
            self._extract_metadata(sheets['стр. 1-2'])
        
        for section_type in self.constants.SECTION_SHEETS.keys():
            sheet_name = self.constants.SECTION_SHEETS[section_type]
            if sheet_name in sheets:
                if section_type == 'консолидируемые_расчеты':
                    self._extract_consolidated_data(sheets[sheet_name])
                else:
                    self._extract_section_data(sheets[sheet_name], section_type)
        
        self._calculate_deficit_proficit()
        
        return {
            'meta_info': self.meta_info,
            'доходы_data': self.доходы_data,
            'расходы_data': self.расходы_data,
            'источники_финансирования_data': self.источники_финансирования_data,
            'консолидируемые_расчеты_data': self.консолидируемые_расчеты_data,
            'calculated_deficit_proficit': self.calculated_deficit_proficit
        }
    
    def calculate_sums(self) -> Dict[str, Any]:
        """Расчет агрегированных сумм"""
        # Расчет для доходов
        if self.доходы_data:
            df_доходы = self._prepare_dataframe_for_calculation(self.доходы_data, self.constants.BUDGET_COLUMNS)
            df_доходы_with_sums = self._calculate_budget_sums(df_доходы, self.constants.BUDGET_COLUMNS)
            self.доходы_data = df_доходы_with_sums.to_dict('records')
        
        # Расчет для расходов
        if self.расходы_data:
            df_расходы = self._prepare_dataframe_for_calculation(self.расходы_data, self.constants.BUDGET_COLUMNS)
            df_расходы_with_sums = self._calculate_budget_sums(df_расходы, self.constants.BUDGET_COLUMNS)
            self.расходы_data = df_расходы_with_sums.to_dict('records')
        
        # Расчет для источников финансирования
        if self.источники_финансирования_data:
            df_источники = self._prepare_dataframe_for_calculation(self.источники_финансирования_data, self.constants.BUDGET_COLUMNS)
            df_источники_with_sums = self._calculate_budget_sums(df_источники, self.constants.BUDGET_COLUMNS)
            self.источники_финансирования_data = df_источники_with_sums.to_dict('records')
        
        # Расчет для консолидируемых расчетов
        if self.консолидируемые_расчеты_data:
            df_консолидируемые = self._prepare_consolidated_dataframe_for_calculation(
                self.консолидируемые_расчеты_data, self.constants.CONSOLIDATED_COLUMNS)
            df_консолидируемые_with_sums = self._calculate_consolidated_sums(df_консолидируемые)
            self.консолидируемые_расчеты_data = df_консолидируемые_with_sums.to_dict('records')
        
        # calculate_sums возвращает только данные разделов для пересчета
        # meta_info и результат_исполнения_data не должны возвращаться здесь,
        # они уже есть в форме и сохраняются отдельно при сохранении ревизии
        return {
            'доходы_data': self.доходы_data,
            'расходы_data': self.расходы_data,
            'источники_финансирования_data': self.источники_финансирования_data,
            'консолидируемые_расчеты_data': self.консолидируемые_расчеты_data
        }

    def recalculate_levels_with_references(
        self,
        form_data: Dict[str, Any],
        reference_data_доходы: Optional[pd.DataFrame] = None,
        reference_data_источники: Optional[pd.DataFrame] = None
    ) -> Dict[str, Any]:
        """
        Пересчет уровней строк на основе актуальных справочников.
        Используется при повторной загрузке уже сохраненного проекта.
        """
        self.reference_data_доходы = reference_data_доходы
        self.reference_data_источники = reference_data_источники

        # Доходы
        доходы_data = form_data.get('доходы_data', [])
        for item in доходы_data:
            code = item.get('код_классификации', '')
            name = item.get('наименование_показателя', '')
            item['уровень'] = self._determine_level(code, 'доходы', name)

        # Источники финансирования
        источники_data = form_data.get('источники_финансирования_data', [])
        for item in источники_data:
            code = item.get('код_классификации', '')
            name = item.get('наименование_показателя', '')
            item['уровень'] = self._determine_level(code, 'источники_финансирования', name)

        # Расходы зависят только от кода, справочники не нужны
        расходы_data = form_data.get('расходы_data', [])
        for item in расходы_data:
            code = item.get('код_классификации', '')
            item['уровень'] = self._determine_expenditure_level(code)

        form_data['доходы_data'] = доходы_data
        form_data['расходы_data'] = расходы_data
        form_data['источники_финансирования_data'] = источники_data
        return form_data

    def load_saved_data(self, form_data: Dict[str, Any]):
        """
        Инициализация внутренних структур формы из сохраненных данных проекта.
        Требуется при повторной загрузке проекта, чтобы экспорт/проверка работали корректно.
        """
        self.meta_info = form_data.get('meta_info', {})
        self.доходы_data = form_data.get('доходы_data', []) or []
        self.расходы_data = form_data.get('расходы_data', []) or []
        self.источники_финансирования_data = form_data.get('источники_финансирования_data', []) or []
        self.консолидируемые_расчеты_data = form_data.get('консолидируемые_расчеты_data', []) or []

        # Восстанавливаем вспомогательные структуры
        self.zero_columns = {}
        if self.доходы_data:
            total_row = next((item for item in self.доходы_data if 'всего' in item.get('наименование_показателя', '').lower()), None)
            if total_row:
                self.zero_columns['доходы'] = self._get_zero_columns(total_row, self.constants.BUDGET_COLUMNS)

        if self.расходы_data:
            total_row = next((item for item in self.расходы_data if 'всего' in item.get('наименование_показателя', '').lower()), None)
            if total_row:
                self.zero_columns['расходы'] = self._get_zero_columns(total_row, self.constants.BUDGET_COLUMNS)

        if self.источники_финансирования_data:
            total_row = next((item for item in self.источники_финансирования_data if 'всего' in item.get('наименование_показателя', '').lower()), None)
            if total_row:
                self.zero_columns['источники_финансирования'] = self._get_zero_columns(total_row, self.constants.BUDGET_COLUMNS)

        # Источники требуют корректировки уровней
        self._recalculate_sources_levels()

        # Пересчитываем итоговые значения для дефицита/профицита, если их нет
        if not self.calculated_deficit_proficit:
            self._calculate_deficit_proficit()
    
    def validate_data(self) -> List[Dict[str, Any]]:
        """Валидация данных"""
        errors = []
        
        # Проверка агрегации сумм
        for section, data in [
            ('Доходы', self.доходы_data),
            ('Расходы', self.расходы_data),
            ('Источники финансирования', self.источники_финансирования_data),
            ('Консолидируемые расчеты', self.консолидируемые_расчеты_data)
        ]:
            if data:
                section_errors = self._validate_section_aggregation(section, data)
                errors.extend(section_errors)
        
        return errors
    
    def export_validation(self, original_file_path: str, output_file_path: str) -> str:
        """Экспорт формы с проверкой"""
        shutil.copy2(original_file_path, output_file_path)
        wb = openpyxl.load_workbook(output_file_path)
        
        # Обработка всех разделов
        sections_data = [
            ('доходы', self.доходы_data),
            ('расходы', self.расходы_data),
            ('источники_финансирования', self.источники_финансирования_data),
            ('консолидируемые_расчеты', self.консолидируемые_расчеты_data)
        ]
        
        for section_type, data in sections_data:
            if data:
                if section_type == 'консолидируемые_расчеты':
                    self._process_consolidated_section_in_original_form(wb, data)
                else:
                    self._process_section_in_original_form(wb, data, section_type)

        # Пересчитываем дефицит/профицит по текущим данным формы перед проверкой.
        # calculated_deficit_proficit больше не загружается из метаданных, поэтому
        # всегда считаем его из итоговых строк доходов и расходов.
        self._calculate_deficit_proficit()

        self._validate_deficit_proficit(wb)
        wb.save(output_file_path)
        return output_file_path
    
    # Вспомогательные методы
    def _extract_metadata(self, sheet: pd.DataFrame):
        """Извлечение метаданных"""
        self.meta_info = {
            'Наименование формы': self._get_cell_value(sheet, 2, 1) + ' ' + self._get_cell_value(sheet, 3, 1),
            'Наименование финансового органа': self._get_cell_value(sheet, 6, 3),
            'Наименование бюджета': self._get_cell_value(sheet, 7, 3),
            'Периодичность': self._get_cell_value(sheet, 8, 17),
            'Форма по ОКУД': self._get_cell_value(sheet, 4, 17),
            'Дата': self._get_cell_value(sheet, 5, 17),
            'код ОКПО': self._get_cell_value(sheet, 6, 17),
            'код ОКТМО': self._get_cell_value(sheet, 7, 17),
            'код ОКЕИ': self._get_cell_value(sheet, 9, 17),
        }
    
    def _find_section_start(self, sheet: pd.DataFrame, section_type: str, search_column: int = 0) -> int:
        """Поиск начала раздела"""
        pattern = self.constants.SECTION_PATTERNS.get(section_type)
        if not pattern:
            return None
            
        for i in range(len(sheet)):
            cell_value = str(sheet.iloc[i, search_column]) if pd.notna(sheet.iloc[i, search_column]) else ''
            if re.search(pattern, cell_value, re.IGNORECASE):
                return i
                
        return None
    
    def _extract_consolidated_data(self, sheet: pd.DataFrame):
        """Извлечение данных консолидируемых расчетов"""
        start_row = self._find_section_start(sheet, 'консолидируемые_расчеты', search_column=1)
        
        if start_row is None:
            logger.warning("Таблица консолидируемых расчетов не найдена")
            return
        
        table_start_row = start_row + 2
        self._extract_consolidated_table_data(sheet, table_start_row)
    
    def _extract_consolidated_table_data(self, sheet: pd.DataFrame, start_row: int):
        """Извлечение данных таблицы консолидируемых расчетов"""
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
        
        self.консолидируемые_расчеты_data = data_parts
        logger.info(f"Извлечено {len(data_parts)} записей из таблицы консолидируемых расчетов")
    
    def _extract_consolidated_part_data(self, sheet: pd.DataFrame, start_row: int, mapping: dict, consolidated_columns: list) -> list:
        """Извлечение данных части таблицы консолидируемых расчетов"""
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
    
    def _extract_consolidated_row_data(self, sheet: pd.DataFrame, row_idx: int, mapping: dict, consolidated_columns: list) -> dict:
        """Извлечение данных строки консолидируемых расчетов"""
        common_data = {}
        for i, col_offset in enumerate([0, 1]):
            col_idx = 1 + col_offset
            value = self._get_cell_value(sheet, row_idx, col_idx)
            
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
            value = self._get_numeric_value(sheet, row_idx, col_idx)
            поступления[consolidated_columns[i]] = value

        return {
            **common_data,
            'раздел': 'консолидируемые_расчеты',
            'поступления': поступления,
            'исходная_строка': row_idx + 1,
        }
    
    def _determine_consolidated_level(self, code: str) -> int:
        """Определение уровня для консолидируемых расчетов"""
        if not code or len(code) < 3:
            return 0
        
        if code == '899':
            return 0
        
        if code[0] == '9' and code[1] in '0123456789' and code[2] == '0':
            return 1
        
        if code[0] == '9' and code[1] in '0123456789' and code[2] != '0':
            return 2
        
        return 0
    
    def _extract_section_data(self, sheet: pd.DataFrame, section_type: str):
        """Извлечение данных раздела"""
        search_column = 1 if section_type == 'консолидируемые_расчеты' else 0
        start_row = self._find_section_start(sheet, section_type, search_column)
        
        if start_row is None:
            logger.warning(f"Раздел '{section_type}' не найден")
            return
            
        logger.debug(f"Раздел '{section_type}' начинается со строки {start_row + 1}")
        header_row = start_row + 3
        self._extract_table_data(sheet, header_row, section_type)
    
    def _extract_table_data(self, sheet: pd.DataFrame, header_row: int, section_type: str):
        """Извлечение табличных данных"""
        mapping = self.constants.COLUMN_MAPPING[section_type]
        budget_columns = self.constants.BUDGET_COLUMNS
        
        total_row_data = None
        data_start_row = header_row + 2

        for row_idx in range(data_start_row, len(sheet)):
            row_data = self._extract_row_data(sheet, row_idx, mapping, budget_columns, section_type)
            if row_data and row_data['наименование_показателя']:
                if self._is_total_row(row_data, section_type):
                    total_row_data = row_data
                
                getattr(self, f'{section_type}_data').append(row_data)
        
        if section_type == 'источники_финансирования':
            self._recalculate_sources_levels()
        
        if total_row_data:
            zero_columns = self._get_zero_columns(total_row_data, budget_columns)
            self.zero_columns[section_type] = zero_columns
            logger.debug(f"Для раздела '{section_type}' нулевые столбцы: {zero_columns}")
    
    def _extract_row_data(self, sheet: pd.DataFrame, row_idx: int, mapping: dict, budget_columns: list, section_type: str) -> dict:
        """Извлечение данных строки"""
        common_data = {}
        for i, col in enumerate(mapping['common_cols']):
            col_idx = self._column_to_index(col)
            value = self._get_cell_value(sheet, row_idx, col_idx)
            
            if i == 0:
                value = value.replace("в том числе:\n", "").replace("\n", "").strip()
                common_data['наименование_показателя'] = value
            elif i == 1:
                value = value.replace("\n", "").strip()
                common_data['код_строки'] = value
            elif i == 2:
                value = value.replace("\n", "").strip()
                clean_code = self._clean_dbk_code(value)
                common_data['код_классификации'] = clean_code
                common_data['код_классификации_форматированный'] = self._format_classification_code(clean_code, section_type)

        if not common_data.get('наименование_показателя'):
            return None

        level = self._determine_level(common_data['код_классификации'], section_type, common_data['наименование_показателя'])
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
    
    def _extract_budget_data(self, sheet: pd.DataFrame, row_idx: int, columns: list, budget_columns: list) -> dict:
        """Извлечение данных по бюджету"""
        data = {}
        for i, col in enumerate(columns):
            col_idx = self._column_to_index(col)
            value = self._get_numeric_value(sheet, row_idx, col_idx)
            data[budget_columns[i]] = value
        return data
    
    def _determine_level(self, classification_code: str, section_type: str, name: str = "") -> int:
        """Определение уровня строки"""
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
        """Определение уровня для расходов"""
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
        """Получение уровня из справочника"""
        try:
            if section_type == 'доходы' and self.reference_data_доходы is not None:
                # Ищем в DataFrame
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
        """Проверка, является ли строка итоговой"""
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
    
    def _recalculate_sources_levels(self):
        """Пересчет уровней для источников финансирования"""
        if not self.источники_финансирования_data:
            return
        
        for i, item in enumerate(self.источники_финансирования_data):
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
    
    def _find_total_row(self, data: list, pattern: str) -> dict:
        """Поиск итоговой строки по паттерну"""
        import re
        for item in data:
            name = str(item.get('наименование_показателя', '')).lower()
            if re.search(pattern, name, re.IGNORECASE):
                return item
        return None
    
    def _calculate_deficit_proficit_from_original(self, original_доходы_data: list = None, original_расходы_data: list = None):
        """Расчет дефицита/профицита из исходных данных (до пересчета)"""
        # Используем переданные исходные данные или текущие данные формы
        доходы_data = original_доходы_data if original_доходы_data is not None else self.доходы_data
        расходы_data = original_расходы_data if original_расходы_data is not None else self.расходы_data
        
        # Ищем итоговые строки в исходных данных
        доходы_всего = None
        расходы_всего = None
        
        pattern_доходы = self.constants.TOTAL_PATTERNS.get('доходы', r'доходы бюджета.*всего')
        pattern_расходы = self.constants.TOTAL_PATTERNS.get('расходы', r'расходы бюджета.*всего')
        
        доходы_всего = self._find_total_row(доходы_data, pattern_доходы)
        расходы_всего = self._find_total_row(расходы_data, pattern_расходы)
        
        if not доходы_всего or not расходы_всего:
            return None
        
        budget_columns = self.constants.BUDGET_COLUMNS
        утвержденный = {}
        исполненный = {}
        
        # Берем значения из данных с учетом пересчета:
        # если есть расчетные итоговые значения (расчетный_утвержденный/исполненный),
        # используем их, иначе — оригинальные.
        for budget_col in budget_columns:
            доходы_утвержденный = (
                доходы_всего.get(f'расчетный_утвержденный_{budget_col}')
                if доходы_всего.get(f'расчетный_утвержденный_{budget_col}') is not None
                else доходы_всего.get('утвержденный', {}).get(budget_col, 0)
            ) or 0
            доходы_исполненный = (
                доходы_всего.get(f'расчетный_исполненный_{budget_col}')
                if доходы_всего.get(f'расчетный_исполненный_{budget_col}') is not None
                else доходы_всего.get('исполненный', {}).get(budget_col, 0)
            ) or 0
            расходы_утвержденный = (
                расходы_всего.get(f'расчетный_утвержденный_{budget_col}')
                if расходы_всего.get(f'расчетный_утвержденный_{budget_col}') is not None
                else расходы_всего.get('утвержденный', {}).get(budget_col, 0)
            ) or 0
            расходы_исполненный = (
                расходы_всего.get(f'расчетный_исполненный_{budget_col}')
                if расходы_всего.get(f'расчетный_исполненный_{budget_col}') is not None
                else расходы_всего.get('исполненный', {}).get(budget_col, 0)
            ) or 0
            
            утвержденный[budget_col] = расходы_утвержденный - доходы_утвержденный
            исполненный[budget_col] = расходы_исполненный - доходы_исполненный
        
        self.calculated_deficit_proficit = {
            'утвержденный': утвержденный,
            'исполненный': исполненный
        }
        
        return self.calculated_deficit_proficit
    
    def _calculate_deficit_proficit(self):
        """Расчет дефицита/профицита (использует текущие данные формы)"""
        # Вызываем новый метод без исходных данных (использует текущие данные)
        return self._calculate_deficit_proficit_from_original()
    
    def _prepare_dataframe_for_calculation(self, data: list, budget_columns: list) -> pd.DataFrame:
        """Подготовка DataFrame для вычислений"""
        df = pd.DataFrame(data)
        
        for budget_col in budget_columns:
            approved_values = [row['утвержденный'][budget_col] for row in data]
            executed_values = [row['исполненный'][budget_col] for row in data]
            
            df[f'утвержденный_{budget_col}'] = approved_values
            df[f'исполненный_{budget_col}'] = executed_values
            # Оптимизация: используем .values.copy() для создания независимой копии данных, а не копирование Series
            df[f'расчетный_утвержденный_{budget_col}'] = approved_values.copy()
            df[f'расчетный_исполненный_{budget_col}'] = executed_values.copy()
        
        return df
    
    def _prepare_consolidated_dataframe_for_calculation(self, data: list, consolidated_columns: list) -> pd.DataFrame:
        """Подготовка DataFrame для консолидируемых расчетов"""
        df = pd.DataFrame(data)
        
        for col in consolidated_columns:
            values = [row['поступления'][col] for row in data]
            df[f'поступления_{col}'] = values
            # Оптимизация: используем копию списка вместо копирования Series
            df[f'расчетный_поступления_{col}'] = values.copy()
        
        return df
    
    def _calculate_budget_sums(self, df: pd.DataFrame, budget_columns: list) -> pd.DataFrame:
        """Расчет бюджетных сумм"""
        result_df = df.copy()
        
        if result_df.iloc[0]['раздел'] == 'источники_финансирования':
            return self._calculate_sources_sums(result_df, budget_columns)
        
        return self._calculate_standard_sums(result_df, budget_columns)
    
    def _calculate_standard_sums(self, df: pd.DataFrame, budget_columns: list) -> pd.DataFrame:
        """Стандартный расчет сумм"""
        result_df = df.copy()
        
        # Оптимизация: предварительно группируем индексы по уровням, чтобы избежать фильтрации в цикле
        level_groups = {}
        for idx, level in result_df['уровень'].items():
            if level not in level_groups:
                level_groups[level] = []
            level_groups[level].append(idx)
        
        for current_level in range(6, -1, -1):
            if current_level not in level_groups:
                continue
            level_indices = level_groups[current_level]
            
            for idx in level_indices:
                start_idx, end_idx = self._find_child_boundaries(result_df, idx, current_level)
                
                for budget_col in budget_columns:
                    self._sum_children_for_budget_column(result_df, idx, start_idx, end_idx, budget_col, current_level)
        
        return result_df
    
    def _calculate_sources_sums(self, df: pd.DataFrame, budget_columns: list) -> pd.DataFrame:
        """Расчет сумм для источников финансирования"""
        # Оптимизация: _calculate_standard_sums уже делает копию, поэтому не нужно копировать здесь
        result_df = self._calculate_standard_sums(df, budget_columns)
        
        internal_total_indices = []
        external_total_indices = []
        total_sources_index = None
        
        # Оптимизация: используем векторизованные операции для суммирования вместо iterrows()
        # Фильтруем строки уровня 1 с нужными кодами
        level_1_mask = result_df['уровень'] == 1
        internal_mask = level_1_mask & result_df['код_классификации'].str.startswith('00001', na=False)
        external_mask = level_1_mask & result_df['код_классификации'].str.startswith('00002', na=False)
        
        # Векторизованное суммирование для внутренних источников
        internal_sum_approved = {}
        internal_sum_executed = {}
        for budget_col in budget_columns:
            internal_sum_approved[budget_col] = result_df.loc[internal_mask, f'утвержденный_{budget_col}'].sum()
            internal_sum_executed[budget_col] = result_df.loc[internal_mask, f'исполненный_{budget_col}'].sum()
        
        # Векторизованное суммирование для внешних источников
        external_sum_approved = {}
        external_sum_executed = {}
        for budget_col in budget_columns:
            external_sum_approved[budget_col] = result_df.loc[external_mask, f'утвержденный_{budget_col}'].sum()
            external_sum_executed[budget_col] = result_df.loc[external_mask, f'исполненный_{budget_col}'].sum()
        
        # Поиск индексов (требует iterrows, но это быстрый проход только для поиска)
        for idx, row in result_df.iterrows():
            name = row['наименование_показателя'].lower()
            if 'источники финансирования дефицита бюджетов - всего' in name:
                total_sources_index = idx
            elif 'источники внутреннего финансирования' in name:
                internal_total_indices.append(idx)
            elif 'источники внешнего финансирования' in name:
                external_total_indices.append(idx)
        
        for idx in internal_total_indices:
            for budget_col in budget_columns:
                result_df.at[idx, f'расчетный_утвержденный_{budget_col}'] = round(internal_sum_approved[budget_col], 5)
                result_df.at[idx, f'расчетный_исполненный_{budget_col}'] = round(internal_sum_executed[budget_col], 5)
        
        for idx in external_total_indices:
            for budget_col in budget_columns:
                result_df.at[idx, f'расчетный_утвержденный_{budget_col}'] = round(external_sum_approved[budget_col], 5)
                result_df.at[idx, f'расчетный_исполненный_{budget_col}'] = round(external_sum_executed[budget_col], 5)
        
        if total_sources_index is not None:
            for budget_col in budget_columns:
                total_approved = internal_sum_approved[budget_col] + external_sum_approved[budget_col]
                total_executed = internal_sum_executed[budget_col] + external_sum_executed[budget_col]
                result_df.at[total_sources_index, f'расчетный_утвержденный_{budget_col}'] = round(total_approved, 5)
                result_df.at[total_sources_index, f'расчетный_исполненный_{budget_col}'] = round(total_executed, 5)
        
        return result_df
    
    def _calculate_consolidated_sums(self, df: pd.DataFrame) -> pd.DataFrame:
        """Расчет сумм для консолидируемых расчетов"""
        result_df = df.copy()
        
        # Бюджетные столбцы (без ИТОГО)
        budget_columns = self.constants.CONSOLIDATED_COLUMNS[:-1]
        # Столбец ИТОГО
        total_column = self.constants.CONSOLIDATED_COLUMNS[-1]
        
        # Сначала вычисляем уровни 2 и 1
        for current_level in range(2, 0, -1):
            level_indices = result_df[result_df['уровень'] == current_level].index
            
            for idx in level_indices:
                start_idx, end_idx = self._find_child_boundaries(result_df, idx, current_level)
                
                # Пересчитываем бюджетные столбцы
                for col in budget_columns:
                    self._sum_consolidated_children(result_df, idx, start_idx, end_idx, col, current_level)
                
                # Пересчитываем ИТОГО как сумму всех бюджетных столбцов
                calculated_total = 0.0
                for col in budget_columns:
                    calculated_value = result_df.at[idx, f'расчетный_поступления_{col}']
                    if calculated_value is None:
                        calculated_value = result_df.at[idx, f'поступления_{col}']
                    if calculated_value != 'x' and calculated_value is not None:
                        calculated_total += calculated_value
                result_df.at[idx, f'расчетный_поступления_{total_column}'] = round(calculated_total, 5)
        
        # Затем вычисляем уровень 0 по пересчитанным значениям уровня 1
        level0_indices = result_df[result_df['уровень'] == 0].index
        for idx in level0_indices:
            # Пересчитываем бюджетные столбцы
            for col in budget_columns:
                self._sum_level1_for_level0(result_df, idx, col)
            
            # Пересчитываем ИТОГО как сумму всех бюджетных столбцов
            calculated_total = 0.0
            for col in budget_columns:
                calculated_value = result_df.at[idx, f'расчетный_поступления_{col}']
                if calculated_value is None:
                    calculated_value = result_df.at[idx, f'поступления_{col}']
                if calculated_value != 'x' and calculated_value is not None:
                    calculated_total += calculated_value
            result_df.at[idx, f'расчетный_поступления_{total_column}'] = round(calculated_total, 5)
        
        return result_df
    
    def _sum_level1_for_level0(self, df: pd.DataFrame, parent_idx: int, column: str):
        """Суммирование значений уровня 1 для уровня 0"""
        level1_sum = 0.0
        has_level1 = False
        
        for idx, row in df.iterrows():
            if row['уровень'] == 1 and row['код_строки'].startswith('9'):
                calculated_value = row.get(f'расчетный_поступления_{column}')
                if calculated_value is None:
                    calculated_value = row[f'поступления_{column}']
                
                if calculated_value != 'x' and calculated_value is not None:
                    level1_sum += calculated_value
                    has_level1 = True
        
        if has_level1:
            df.at[parent_idx, f'расчетный_поступления_{column}'] = round(level1_sum, 5)
    
    def _sum_consolidated_children(self, df: pd.DataFrame, parent_idx: int, start_idx: int, end_idx: int, column: str, current_level: int):
        """Суммирование дочерних элементов для консолидируемых расчетов"""
        child_sum = 0.0
        has_children = False
        
        for child_idx in range(start_idx, end_idx):
            if df.loc[child_idx, 'уровень'] == current_level + 1:
                calculated_value = df.loc[child_idx, f'расчетный_поступления_{column}']
                if calculated_value is None:
                    calculated_value = df.loc[child_idx, f'поступления_{column}']
                
                if calculated_value != 'x' and calculated_value is not None:
                    child_sum += calculated_value
                    has_children = True
        
        if has_children:
            df.at[parent_idx, f'расчетный_поступления_{column}'] = round(child_sum, 5)
    
    def _find_child_boundaries(self, df: pd.DataFrame, current_idx: int, current_level: int) -> tuple:
        """Поиск границ дочерних элементов"""
        start_idx = current_idx + 1
        end_idx = len(df)
        
        for next_idx in range(start_idx, len(df)):
            if df.loc[next_idx, 'уровень'] <= current_level:
                end_idx = next_idx
                break
        
        return start_idx, end_idx
    
    def _sum_children_for_budget_column(self, df: pd.DataFrame, parent_idx: int, start_idx: int, end_idx: int, budget_col: str, current_level: int):
        """Суммирование дочерних элементов для бюджетной колонки"""
        approved_child_sum = 0.0
        executed_child_sum = 0.0
        has_children = False
        
        for child_idx in range(start_idx, end_idx):
            if df.loc[child_idx, 'уровень'] == current_level + 1:
                approved_child_sum += df.loc[child_idx, f'расчетный_утвержденный_{budget_col}']
                executed_child_sum += df.loc[child_idx, f'расчетный_исполненный_{budget_col}']
                has_children = True
        
        if has_children:
            df.at[parent_idx, f'расчетный_утвержденный_{budget_col}'] = round(approved_child_sum, 5)
            df.at[parent_idx, f'расчетный_исполненный_{budget_col}'] = round(executed_child_sum, 5)
    
    def _validate_section_aggregation(self, section_name: str, data: List[Dict]) -> List[Dict[str, Any]]:
        """Валидация агрегации раздела"""
        errors = []
        
        for item in data:
            level = item.get('уровень', 0)
            if level < 6:  # Проверяем только агрегирующие уровни
                # Проверка бюджетных колонок
                for budget_col in self.constants.BUDGET_COLUMNS:
                    original_approved = item.get('утвержденный', {}).get(budget_col, 0)
                    calculated_approved = item.get('расчетный_утвержденный_{budget_col}', original_approved)
                    
                    if abs(original_approved - calculated_approved) > 0.00001:
                        errors.append({
                            'section': section_name,
                            'row': item.get('наименование_показателя', ''),
                            'level': level,
                            'column': f'утвержденный_{budget_col}',
                            'original': original_approved,
                            'calculated': calculated_approved
                        })
        
        return errors
    
    def _process_consolidated_section_in_original_form(self, wb: openpyxl.Workbook, data: list):
        """Обработка консолидируемых расчетов в исходной форме"""
        if not data:
            return
        
        sheet_name = self.constants.SECTION_SHEETS['консолидируемые_расчеты']
        if sheet_name not in wb.sheetnames:
            return
        
        ws = wb[sheet_name]
        consolidated_columns = self.constants.CONSOLIDATED_COLUMNS
        
        df = self._prepare_consolidated_dataframe_for_calculation(data, consolidated_columns)
        df_with_sums = self._calculate_consolidated_sums(df)
        self._apply_consolidated_validation_to_original_cells(ws, df_with_sums, consolidated_columns)
    
    def _process_section_in_original_form(self, wb: openpyxl.Workbook, data: list, section_name: str):
        """Обработка раздела в исходной форме"""
        if not data:
            return
        
        sheet_name = self.constants.SECTION_SHEETS.get(section_name)
        if sheet_name not in wb.sheetnames:
            return
        
        ws = wb[sheet_name]
        budget_columns = self.constants.BUDGET_COLUMNS
        
        df = self._prepare_dataframe_for_calculation(data, budget_columns)
        df_with_sums = self._calculate_budget_sums(df, budget_columns)
        self._apply_validation_to_original_cells(ws, df_with_sums, budget_columns, section_name)
    
    def _apply_consolidated_validation_to_original_cells(self, ws, df: pd.DataFrame, consolidated_columns: list):
        """Применение проверки к ячейкам консолидируемых расчетов"""
        for row_idx, (_, row_data) in enumerate(df.iterrows(), 0):
            self._apply_consolidated_row_validation(ws, row_data, consolidated_columns)
    
    def _apply_consolidated_row_validation(self, ws, row_data: dict, consolidated_columns: list):
        """Применение проверки к строке консолидируемых расчетов"""
        level = row_data['уровень']
        original_row = row_data['исходная_строка']
        mapping = self.constants.COLUMN_MAPPING['консолидируемые_расчеты']
        
        if level in StyleConstants.LEVEL_COLORS:
            self._fill_consolidated_row_with_level_color(ws, original_row, level)
        
        budget_columns = consolidated_columns[:-1]
        for i, budget_col in enumerate(budget_columns):
            self._validate_consolidated_cells(ws, row_data, budget_col, i, mapping, original_row, level)
        
        self._validate_sum_column(ws, row_data, consolidated_columns, mapping, original_row, level)
    
    def _fill_consolidated_row_with_level_color(self, ws, row: int, level: int):
        """Закрашивание строки консолидируемых расчетов"""
        level_fill = PatternFill(
            start_color=StyleConstants.LEVEL_COLORS[level], 
            end_color=StyleConstants.LEVEL_COLORS[level], 
            fill_type="solid"
        )
        
        for col_idx in range(2, 15):
            cell = ws.cell(row=row, column=col_idx)
            cell.fill = level_fill
    
    def _validate_consolidated_cells(self, ws, row_data: dict, budget_col: str, col_index: int, mapping: dict, original_row: int, level: int):
        """Проверка ячеек консолидируемых расчетов"""
        receipt_col = mapping['поступления'][col_index]
        receipt_col_idx = self._column_to_index(receipt_col) + 1
        receipt_cell = ws.cell(row=original_row, column=receipt_col_idx)
        
        original_receipt = row_data[f'поступления_{budget_col}']
        calculated_receipt = row_data[f'расчетный_поступления_{budget_col}']
        
        if original_receipt == 'x':
            return
        
        if self._is_value_different(original_receipt, calculated_receipt) and level < 2:
            receipt_cell.fill = StyleConstants.ERROR_FILL
            if self.show_error_values:
                try:
                    orig_val = float(original_receipt or 0)
                except (TypeError, ValueError):
                    orig_val = 0.0
                try:
                    calc_val = float(calculated_receipt or 0)
                except (TypeError, ValueError):
                    calc_val = 0.0
                receipt_cell.value = f"{orig_val:.2f} ({calc_val:.2f})"
    
    def _validate_sum_column(self, ws, row_data: dict, consolidated_columns: list, mapping: dict, original_row: int, level: int):
        """Проверка столбца 'ИТОГО'"""
        sum_col_index = len(consolidated_columns) - 1
        sum_col = mapping['поступления'][sum_col_index]
        sum_col_idx = self._column_to_index(sum_col) + 1
        sum_cell = ws.cell(row=original_row, column=sum_col_idx)
        
        # Последний столбец в CONSOLIDATED_COLUMNS содержит агрегированное значение (ИТОГО)
        sum_column_name = consolidated_columns[-1]
        original_sum = row_data.get(f'поступления_{sum_column_name}')
        
        # Используем уже рассчитанное значение, если оно есть
        calculated_sum = row_data.get(f'расчетный_поступления_{sum_column_name}')
        
        # Если расчетное значение не найдено, пересчитываем сумму из бюджетных столбцов
        if calculated_sum is None:
            calculated_sum = 0.0
            budget_columns = consolidated_columns[:-1]
            
            for budget_col in budget_columns:
                calculated_value = row_data.get(f'расчетный_поступления_{budget_col}')
                if calculated_value is None:
                    calculated_value = row_data.get(f'поступления_{budget_col}', 0)
                
                if calculated_value != 'x' and calculated_value is not None:
                    calculated_sum += calculated_value
        
        if original_sum == 'x':
            return
        
        if self._is_value_different(original_sum, calculated_sum) and level < 2:
            sum_cell.fill = StyleConstants.ERROR_FILL
            if self.show_error_values:
                try:
                    orig_val = float(original_sum or 0)
                except (TypeError, ValueError):
                    orig_val = 0.0
                try:
                    calc_val = float(calculated_sum or 0)
                except (TypeError, ValueError):
                    calc_val = 0.0
                sum_cell.value = f"{orig_val:.2f} ({calc_val:.2f})"
    
    def _validate_deficit_proficit(self, wb: openpyxl.Workbook):
        """Проверка дефицита/профицита"""
        if not self.calculated_deficit_proficit:
            return
        
        sheet_name = self.constants.SECTION_SHEETS['расходы']
        if sheet_name not in wb.sheetnames:
            return
        
        ws = wb[sheet_name]
        budget_columns = self.constants.BUDGET_COLUMNS
        mapping = self.constants.COLUMN_MAPPING['расходы']
        
        target_row = None
        for row in range(1, ws.max_row + 1):
            cell_value = ws.cell(row=row, column=1).value
            if cell_value and 'результат исполнения бюджета' in str(cell_value).lower():
                target_row = row
                code_cell = ws.cell(row=row, column=2).value
                if code_cell and '450' in str(code_cell):
                    break
        
        if not target_row:
            return
        
        for i, budget_col in enumerate(budget_columns):
            approved_col = mapping['утвержденный'][i]
            approved_col_idx = self._column_to_index(approved_col) + 1
            approved_cell = ws.cell(row=target_row, column=approved_col_idx)
            
            original_approved = approved_cell.value or 0
            calculated_approved = self.calculated_deficit_proficit['утвержденный'][budget_col]
            
            if self._is_value_different(original_approved, calculated_approved):
                approved_cell.fill = StyleConstants.ERROR_FILL
                if self.show_error_values:
                    try:
                        orig_val = float(original_approved or 0)
                    except (TypeError, ValueError):
                        orig_val = 0.0
                    try:
                        calc_val = float(calculated_approved or 0)
                    except (TypeError, ValueError):
                        calc_val = 0.0
                    approved_cell.value = f"{orig_val:.2f} ({calc_val:.2f})"
            
            executed_col = mapping['исполненный'][i]
            executed_col_idx = self._column_to_index(executed_col) + 1
            executed_cell = ws.cell(row=target_row, column=executed_col_idx)
            
            original_executed = executed_cell.value or 0
            calculated_executed = self.calculated_deficit_proficit['исполненный'][budget_col]
            
            if self._is_value_different(original_executed, calculated_executed):
                executed_cell.fill = StyleConstants.ERROR_FILL
                if self.show_error_values:
                    try:
                        orig_val = float(original_executed or 0)
                    except (TypeError, ValueError):
                        orig_val = 0.0
                    try:
                        calc_val = float(calculated_executed or 0)
                    except (TypeError, ValueError):
                        calc_val = 0.0
                    executed_cell.value = f"{orig_val:.2f} ({calc_val:.2f})"
    
    def _apply_validation_to_original_cells(self, ws, df: pd.DataFrame, budget_columns: list, section_name: str):
        """Применение проверки к исходным ячейкам"""
        zero_columns = self.zero_columns.get(section_name, [])
        
        for row_idx, (_, row_data) in enumerate(df.iterrows(), 0):
            self._apply_row_validation(ws, row_data, budget_columns, zero_columns, section_name)
        
        self._hide_zero_columns(ws, section_name, zero_columns)
    
    def _apply_row_validation(self, ws, row_data: dict, budget_columns: list, zero_columns: list, section_name: str):
        """Применение проверки к строке"""
        level = row_data['уровень']
        original_row = row_data['исходная_строка']
        mapping = self.constants.COLUMN_MAPPING[section_name]
        
        if level in StyleConstants.LEVEL_COLORS:
            self._fill_row_with_level_color(ws, original_row, level)
        
        for i, budget_col in enumerate(budget_columns):
            if i in zero_columns or (i + len(budget_columns)) in zero_columns:
                continue
            
            self._validate_budget_cells(ws, row_data, budget_col, i, mapping, original_row, level)
    
    def _fill_row_with_level_color(self, ws, row: int, level: int):
        """Закрашивание строки"""
        level_fill = PatternFill(
            start_color=StyleConstants.LEVEL_COLORS[level], 
            end_color=StyleConstants.LEVEL_COLORS[level], 
            fill_type="solid"
        )
        
        for col_idx in range(1, 37):
            cell = ws.cell(row=row, column=col_idx)
            cell.fill = level_fill
    
    def _validate_budget_cells(self, ws, row_data: dict, budget_col: str, col_index: int, mapping: dict, original_row: int, level: int):
        """Проверка бюджетных ячеек"""
        approved_col = mapping['утвержденный'][col_index]
        approved_col_idx = self._column_to_index(approved_col) + 1
        approved_cell = ws.cell(row=original_row, column=approved_col_idx)
        
        original_approved = row_data[f'утвержденный_{budget_col}']
        calculated_approved = row_data[f'расчетный_утвержденный_{budget_col}']
        
        if self._is_value_different(original_approved, calculated_approved) and level < 6:
            approved_cell.fill = StyleConstants.ERROR_FILL
            if self.show_error_values:
                try:
                    orig_val = float(original_approved or 0)
                except (TypeError, ValueError):
                    orig_val = 0.0
                try:
                    calc_val = float(calculated_approved or 0)
                except (TypeError, ValueError):
                    calc_val = 0.0
                approved_cell.value = f"{orig_val:.2f} ({calc_val:.2f})"
        
        executed_col = mapping['исполненный'][col_index]
        executed_col_idx = self._column_to_index(executed_col) + 1
        executed_cell = ws.cell(row=original_row, column=executed_col_idx)
        
        original_executed = row_data[f'исполненный_{budget_col}']
        calculated_executed = row_data[f'расчетный_исполненный_{budget_col}']
        
        if self._is_value_different(original_executed, calculated_executed) and level < 6:
            executed_cell.fill = StyleConstants.ERROR_FILL
            if self.show_error_values:
                try:
                    orig_val = float(original_executed or 0)
                except (TypeError, ValueError):
                    orig_val = 0.0
                try:
                    calc_val = float(calculated_executed or 0)
                except (TypeError, ValueError):
                    calc_val = 0.0
                executed_cell.value = f"{orig_val:.2f} ({calc_val:.2f})"
    
    def _hide_zero_columns(self, ws, section_name: str, zero_columns: list):
        """Скрытие нулевых столбцов"""
        mapping = self.constants.COLUMN_MAPPING[section_name]
        budget_columns = self.constants.BUDGET_COLUMNS
        
        for col_index in zero_columns:
            if col_index < len(budget_columns):
                col_letter = mapping['утвержденный'][col_index]
            else:
                col_letter = mapping['исполненный'][col_index - len(budget_columns)]
            
            col_idx = self._column_to_index(col_letter) + 1
            ws.column_dimensions[openpyxl.utils.get_column_letter(col_idx)].hidden = True
    
    def _is_value_different(self, original: float, calculated: float) -> bool:
        """Проверка различия значений"""
        return abs(original - calculated) > 0.00001
    
    def _column_to_index(self, column_letter: str) -> int:
        """Конвертация буквы колонки в индекс"""
        column_letter = column_letter.upper()
        index = 0
        for char in column_letter:
            index = index * 26 + (ord(char) - ord('A') + 1)
        return index - 1
    
    def _get_cell_value(self, sheet: pd.DataFrame, row: int, col: int) -> str:
        """Получение значения ячейки"""
        if row < len(sheet) and col < len(sheet.columns):
            value = sheet.iloc[row, col]
            return str(value) if pd.notna(value) else ""
        return ""
    
    def _get_numeric_value(self, sheet: pd.DataFrame, row: int, col: int) -> float:
        """Получение числового значения"""
        if row < len(sheet) and col < len(sheet.columns):
            value = sheet.iloc[row, col]
            if pd.notna(value):
                try:
                    if str(value).lower() == 'x':
                        return 'x'
                    return float(value)
                except (ValueError, TypeError):
                    return 0.0
        return 0.0
    
    def _clean_dbk_code(self, code: str) -> str:
        """Очистка кода классификации"""
        if pd.isna(code) or not isinstance(code, str):
            return ""
        
        clean_code = code.replace(' ', '')
        
        if len(clean_code) < 20 and clean_code.isdigit():
            clean_code = clean_code.zfill(20)
        
        return clean_code
    
    def _format_classification_code(self, code: str, section_type: str) -> str:
        """Форматирование кода классификации"""
        if len(code) != 20:
            return code
            
        if section_type == 'доходы':
            return re.sub(r'(\d{3})(\d{1})(\d{2})(\d{5})(\d{2})(\d{4})(\d{3})', r'\1 \2 \3 \4 \5 \6 \7', code)
        elif section_type == 'расходы':
            return re.sub(r'(\d{3})(\d{4})(\d{10})(\d{3})', r'\1 \2 \3 \4', code)
        elif section_type == 'источники_финансирования':
            return re.sub(r'(\d{3})(\d{2})(\d{2})(\d{2})(\d{2})(\d{2})(\d{4})(\d{3})', r'\1 \2 \3 \4 \5 \6 \7 \8', code)
        return code
"""Калькулятор сумм для формы 0503317"""
import pandas as pd
import re
from typing import Dict, List, Any, Optional
from logger import logger
from models.constants.form_0503317_constants import Form0503317Constants


class Form0503317Calculator:
    """Калькулятор для расчета агрегированных сумм формы 0503317"""
    
    def __init__(self, constants: Form0503317Constants):
        """
        Args:
            constants: Константы формы 0503317
        """
        self.constants = constants
    
    def calculate_sums(self, form_data: Dict[str, Any]) -> Dict[str, Any]:
        """Расчет агрегированных сумм для всех разделов
        
        Args:
            form_data: Словарь с данными формы:
                - доходы_data: список данных доходов
                - расходы_data: список данных расходов
                - источники_финансирования_data: список данных источников
                - консолидируемые_расчеты_data: список данных консолидированных расчетов
        
        Returns:
            Словарь с пересчитанными данными разделов
        """
        result = {}
        
        # Расчет для доходов
        if form_data.get('доходы_data'):
            df_доходы = self._prepare_dataframe_for_calculation(
                form_data['доходы_data'], 
                self.constants.BUDGET_COLUMNS
            )
            df_доходы_with_sums = self._calculate_budget_sums(df_доходы, self.constants.BUDGET_COLUMNS)
            result['доходы_data'] = df_доходы_with_sums.to_dict('records')
        
        # Расчет для расходов
        if form_data.get('расходы_data'):
            df_расходы = self._prepare_dataframe_for_calculation(
                form_data['расходы_data'], 
                self.constants.BUDGET_COLUMNS
            )
            df_расходы_with_sums = self._calculate_budget_sums(df_расходы, self.constants.BUDGET_COLUMNS)
            result['расходы_data'] = df_расходы_with_sums.to_dict('records')
        
        # Расчет для источников финансирования
        if form_data.get('источники_финансирования_data'):
            df_источники = self._prepare_dataframe_for_calculation(
                form_data['источники_финансирования_data'], 
                self.constants.BUDGET_COLUMNS
            )
            df_источники_with_sums = self._calculate_budget_sums(df_источники, self.constants.BUDGET_COLUMNS)
            result['источники_финансирования_data'] = df_источники_with_sums.to_dict('records')
        
        # Расчет для консолидируемых расчетов
        if form_data.get('консолидируемые_расчеты_data'):
            df_консолидируемые = self._prepare_consolidated_dataframe_for_calculation(
                form_data['консолидируемые_расчеты_data'], 
                self.constants.CONSOLIDATED_COLUMNS
            )
            df_консолидируемые_with_sums = self._calculate_consolidated_sums(df_консолидируемые)
            result['консолидируемые_расчеты_data'] = df_консолидируемые_with_sums.to_dict('records')
        
        return result
    
    def calculate_deficit_proficit(
        self, 
        доходы_data: List[dict], 
        расходы_data: List[dict]
    ) -> Optional[Dict[str, Dict[str, float]]]:
        """Расчет дефицита/профицита
        
        Args:
            доходы_data: Список данных доходов
            расходы_data: Список данных расходов
        
        Returns:
            Словарь с дефицитом/профицитом или None
        """
        # Ищем итоговые строки
        pattern_доходы = self.constants.TOTAL_PATTERNS.get('доходы', r'доходы бюджета.*всего')
        pattern_расходы = self.constants.TOTAL_PATTERNS.get('расходы', r'расходы бюджета.*всего')
        
        доходы_всего = self._find_total_row(доходы_data, pattern_доходы)
        расходы_всего = self._find_total_row(расходы_data, pattern_расходы)
        
        if not доходы_всего or not расходы_всего:
            return None
        
        budget_columns = self.constants.BUDGET_COLUMNS
        утвержденный = {}
        исполненный = {}
        
        # Берем значения из данных с учетом пересчета
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
            
            утвержденный[budget_col] = round(расходы_утвержденный - доходы_утвержденный, 5)
            исполненный[budget_col] = round(расходы_исполненный - доходы_исполненный, 5)
        
        return {
            'утвержденный': утвержденный,
            'исполненный': исполненный
        }
    
    def _find_total_row(self, data: List[dict], pattern: str) -> Optional[dict]:
        """Поиск итоговой строки по паттерну
        
        Args:
            data: Список данных
            pattern: Регулярное выражение для поиска
        
        Returns:
            Данные итоговой строки или None
        """
        for item in data:
            name = str(item.get('наименование_показателя', '')).lower()
            if re.search(pattern, name, re.IGNORECASE):
                return item
        return None
    
    def _prepare_dataframe_for_calculation(self, data: List[dict], budget_columns: List[str]) -> pd.DataFrame:
        """Подготовка DataFrame для вычислений
        
        Args:
            data: Список данных раздела
            budget_columns: Список бюджетных колонок
        
        Returns:
            DataFrame с подготовленными данными
        """
        df = pd.DataFrame(data)
        
        for budget_col in budget_columns:
            approved_values = [row['утвержденный'][budget_col] for row in data]
            executed_values = [row['исполненный'][budget_col] for row in data]
            
            df[f'утвержденный_{budget_col}'] = approved_values
            df[f'исполненный_{budget_col}'] = executed_values
            # Используем копию списка для создания независимой копии данных
            df[f'расчетный_утвержденный_{budget_col}'] = approved_values.copy()
            df[f'расчетный_исполненный_{budget_col}'] = executed_values.copy()
        
        return df
    
    def _prepare_consolidated_dataframe_for_calculation(
        self, 
        data: List[dict], 
        consolidated_columns: List[str]
    ) -> pd.DataFrame:
        """Подготовка DataFrame для консолидируемых расчетов
        
        Args:
            data: Список данных консолидированных расчетов
            consolidated_columns: Список колонок консолидированных расчетов
        
        Returns:
            DataFrame с подготовленными данными
        """
        df = pd.DataFrame(data)
        
        for col in consolidated_columns:
            values = [row['поступления'][col] for row in data]
            df[f'поступления_{col}'] = values
            # Используем копию списка вместо копирования Series
            df[f'расчетный_поступления_{col}'] = values.copy()
        
        return df
    
    def _calculate_budget_sums(self, df: pd.DataFrame, budget_columns: List[str]) -> pd.DataFrame:
        """Расчет бюджетных сумм
        
        Args:
            df: DataFrame с данными
            budget_columns: Список бюджетных колонок
        
        Returns:
            DataFrame с пересчитанными суммами
        """
        result_df = df.copy()
        
        if result_df.iloc[0]['раздел'] == 'источники_финансирования':
            return self._calculate_sources_sums(result_df, budget_columns)
        
        return self._calculate_standard_sums(result_df, budget_columns)
    
    def _calculate_standard_sums(self, df: pd.DataFrame, budget_columns: List[str]) -> pd.DataFrame:
        """Стандартный расчет сумм
        
        Args:
            df: DataFrame с данными
            budget_columns: Список бюджетных колонок
        
        Returns:
            DataFrame с пересчитанными суммами
        """
        result_df = df.copy()
        
        # Оптимизация: предварительно группируем индексы по уровням
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
                    self._sum_children_for_budget_column(
                        result_df, idx, start_idx, end_idx, budget_col, current_level
                    )
        
        return result_df
    
    def _calculate_sources_sums(self, df: pd.DataFrame, budget_columns: List[str]) -> pd.DataFrame:
        """Расчет сумм для источников финансирования
        
        Args:
            df: DataFrame с данными
            budget_columns: Список бюджетных колонок
        
        Returns:
            DataFrame с пересчитанными суммами
        """
        # Сначала стандартный расчет
        result_df = self._calculate_standard_sums(df, budget_columns)
        
        internal_total_indices = []
        external_total_indices = []
        total_sources_index = None
        
        # Векторизованные операции для суммирования
        level_1_mask = result_df['уровень'] == 1
        internal_mask = level_1_mask & result_df['код_классификации'].str.startswith('00001', na=False)
        external_mask = level_1_mask & result_df['код_классификации'].str.startswith('00002', na=False)
        
        # Векторизованное суммирование для внутренних источников
        internal_sum_approved = {}
        internal_sum_executed = {}
        for budget_col in budget_columns:
            internal_sum_approved[budget_col] = round(
                result_df.loc[internal_mask, f'утвержденный_{budget_col}'].sum(), 5
            )
            internal_sum_executed[budget_col] = round(
                result_df.loc[internal_mask, f'исполненный_{budget_col}'].sum(), 5
            )
        
        # Векторизованное суммирование для внешних источников
        external_sum_approved = {}
        external_sum_executed = {}
        for budget_col in budget_columns:
            external_sum_approved[budget_col] = round(
                result_df.loc[external_mask, f'утвержденный_{budget_col}'].sum(), 5
            )
            external_sum_executed[budget_col] = round(
                result_df.loc[external_mask, f'исполненный_{budget_col}'].sum(), 5
            )
        
        # Поиск индексов итоговых строк
        for idx, row in result_df.iterrows():
            name = row['наименование_показателя'].lower()
            if 'источники финансирования дефицита бюджетов - всего' in name:
                total_sources_index = idx
            elif 'источники внутреннего финансирования' in name:
                internal_total_indices.append(idx)
            elif 'источники внешнего финансирования' in name:
                external_total_indices.append(idx)
        
        # Обновление итоговых строк
        for idx in internal_total_indices:
            for budget_col in budget_columns:
                result_df.at[idx, f'расчетный_утвержденный_{budget_col}'] = round(
                    internal_sum_approved[budget_col], 5
                )
                result_df.at[idx, f'расчетный_исполненный_{budget_col}'] = round(
                    internal_sum_executed[budget_col], 5
                )
        
        for idx in external_total_indices:
            for budget_col in budget_columns:
                result_df.at[idx, f'расчетный_утвержденный_{budget_col}'] = round(
                    external_sum_approved[budget_col], 5
                )
                result_df.at[idx, f'расчетный_исполненный_{budget_col}'] = round(
                    external_sum_executed[budget_col], 5
                )
        
        if total_sources_index is not None:
            for budget_col in budget_columns:
                total_approved = round(
                    internal_sum_approved[budget_col] + external_sum_approved[budget_col], 5
                )
                total_executed = round(
                    internal_sum_executed[budget_col] + external_sum_executed[budget_col], 5
                )
                result_df.at[total_sources_index, f'расчетный_утвержденный_{budget_col}'] = total_approved
                result_df.at[total_sources_index, f'расчетный_исполненный_{budget_col}'] = total_executed
        
        return result_df
    
    def _calculate_consolidated_sums(self, df: pd.DataFrame) -> pd.DataFrame:
        """Расчет сумм для консолидируемых расчетов
        
        Args:
            df: DataFrame с данными консолидированных расчетов
        
        Returns:
            DataFrame с пересчитанными суммами
        """
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
                        calculated_total += round(float(calculated_value), 5)
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
        """Суммирование значений уровня 1 для уровня 0
        
        Args:
            df: DataFrame с данными
            parent_idx: Индекс родительской строки
            column: Название колонки
        """
        level1_sum = 0.0
        has_level1 = False
        
        for idx, row in df.iterrows():
            if row['уровень'] == 1 and row['код_строки'].startswith('9'):
                calculated_value = row.get(f'расчетный_поступления_{column}')
                if calculated_value is None:
                    calculated_value = row[f'поступления_{column}']
                
                if calculated_value != 'x' and calculated_value is not None:
                    level1_sum += round(float(calculated_value), 5)
                    has_level1 = True
        
        if has_level1:
            df.at[parent_idx, f'расчетный_поступления_{column}'] = round(level1_sum, 5)
    
    def _sum_consolidated_children(
        self, 
        df: pd.DataFrame, 
        parent_idx: int, 
        start_idx: int, 
        end_idx: int, 
        column: str, 
        current_level: int
    ):
        """Суммирование дочерних элементов для консолидируемых расчетов
        
        Args:
            df: DataFrame с данными
            parent_idx: Индекс родительской строки
            start_idx: Начальный индекс дочерних элементов
            end_idx: Конечный индекс дочерних элементов
            column: Название колонки
            current_level: Текущий уровень
        """
        child_sum = 0.0
        has_children = False
        
        for child_idx in range(start_idx, end_idx):
            if df.loc[child_idx, 'уровень'] == current_level + 1:
                calculated_value = df.loc[child_idx, f'расчетный_поступления_{column}']
                if calculated_value is None:
                    calculated_value = df.loc[child_idx, f'поступления_{column}']
                
                if calculated_value != 'x' and calculated_value is not None:
                    child_sum += round(float(calculated_value), 5)
                    has_children = True
        
        if has_children:
            df.at[parent_idx, f'расчетный_поступления_{column}'] = round(child_sum, 5)
    
    def _find_child_boundaries(self, df: pd.DataFrame, current_idx: int, current_level: int) -> tuple:
        """Поиск границ дочерних элементов
        
        Args:
            df: DataFrame с данными
            current_idx: Индекс текущей строки
            current_level: Текущий уровень
        
        Returns:
            Кортеж (start_idx, end_idx)
        """
        start_idx = current_idx + 1
        end_idx = len(df)
        
        for next_idx in range(start_idx, len(df)):
            if df.loc[next_idx, 'уровень'] <= current_level:
                end_idx = next_idx
                break
        
        return start_idx, end_idx
    
    def _sum_children_for_budget_column(
        self, 
        df: pd.DataFrame, 
        parent_idx: int, 
        start_idx: int, 
        end_idx: int, 
        budget_col: str, 
        current_level: int
    ):
        """Суммирование дочерних элементов для бюджетной колонки
        
        Args:
            df: DataFrame с данными
            parent_idx: Индекс родительской строки
            start_idx: Начальный индекс дочерних элементов
            end_idx: Конечный индекс дочерних элементов
            budget_col: Название бюджетной колонки
            current_level: Текущий уровень
        """
        approved_child_sum = 0.0
        executed_child_sum = 0.0
        has_children = False
        
        for child_idx in range(start_idx, end_idx):
            if df.loc[child_idx, 'уровень'] == current_level + 1:
                approved_val = df.loc[child_idx, f'расчетный_утвержденный_{budget_col}']
                executed_val = df.loc[child_idx, f'расчетный_исполненный_{budget_col}']
                approved_child_sum += round(float(approved_val) if approved_val is not None else 0.0, 5)
                executed_child_sum += round(float(executed_val) if executed_val is not None else 0.0, 5)
                has_children = True
        
        if has_children:
            df.at[parent_idx, f'расчетный_утвержденный_{budget_col}'] = round(approved_child_sum, 5)
            df.at[parent_idx, f'расчетный_исполненный_{budget_col}'] = round(executed_child_sum, 5)

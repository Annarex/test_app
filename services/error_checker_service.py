"""Сервис для проверки ошибок расчетов"""
from typing import List, Dict, Any
from logger import logger
from models.constants.form_0503317_constants import Form0503317Constants
from utils.numeric_utils import is_value_different, calculate_error_difference


class ErrorCheckerService:
    """Сервис для проверки ошибок расчетов (бизнес-логика без UI)"""
    
    def check_budget_errors(self, data: List[dict], section_name: str) -> List[Dict[str, Any]]:
        """Проверка ошибок для бюджетных разделов (доходы, расходы, источники)
        
        Args:
            data: Список данных раздела
            section_name: Название раздела
        
        Returns:
            Список ошибок
        """
        errors = []
        budget_cols = Form0503317Constants.BUDGET_COLUMNS
        
        for item in data:
            level = item.get('уровень', 0)
            # Проверяем только уровни < 6
            if level >= 6:
                continue
            
            name = item.get('наименование_показателя', '')
            code = item.get('код_строки', '')
            
            approved_data = item.get('утвержденный', {}) or {}
            executed_data = item.get('исполненный', {}) or {}
            
            for col in budget_cols:
                # Проверка утвержденных значений
                original_approved = approved_data.get(col, 0) or 0
                calculated_approved = item.get(f'расчетный_утвержденный_{col}', original_approved)
                
                if is_value_different(original_approved, calculated_approved):
                    diff = calculate_error_difference(original_approved, calculated_approved)
                    errors.append({
                        'section': section_name,
                        'name': name,
                        'code': code,
                        'level': level,
                        'type': 'Утвержденный',
                        'column': col,
                        'original': original_approved,
                        'calculated': calculated_approved,
                        'difference': diff
                    })
                
                # Проверка исполненных значений
                original_executed = executed_data.get(col, 0) or 0
                calculated_executed = item.get(f'расчетный_исполненный_{col}', original_executed)
                
                if is_value_different(original_executed, calculated_executed):
                    diff = calculate_error_difference(original_executed, calculated_executed)
                    errors.append({
                        'section': section_name,
                        'name': name,
                        'code': code,
                        'level': level,
                        'type': 'Исполненный',
                        'column': col,
                        'original': original_executed,
                        'calculated': calculated_executed,
                        'difference': diff
                    })
        
        return errors
    
    def check_consolidated_errors(self, data: List[dict], section_name: str) -> List[Dict[str, Any]]:
        """Проверка ошибок для консолидированных расчетов
        
        Args:
            data: Список данных консолидированных расчетов
            section_name: Название раздела
        
        Returns:
            Список ошибок
        """
        errors = []
        cons_cols = Form0503317Constants.CONSOLIDATED_COLUMNS
        
        for item in data:
            level = item.get('уровень', 0)
            # Для консолидированных расчетов проверяем все уровни для столбца ИТОГО,
            # и уровни < 6 для остальных столбцов
            name = item.get('наименование_показателя', '')
            code = item.get('код_строки', '')
            
            cons_data = item.get('поступления', {}) or {}
            
            for col in cons_cols:
                # Оригинальное значение
                if isinstance(cons_data, dict) and col in cons_data:
                    original_value = cons_data.get(col, 0) or 0
                else:
                    original_value = item.get(f'поступления_{col}', 0) or 0
                
                # Расчетное значение
                calculated_value = item.get(f'расчетный_поступления_{col}')
                if calculated_value is None:
                    calculated_value = original_value
                
                # Проверяем несоответствие
                is_total_column = (col == 'ИТОГО')
                should_check = (level < 6) or is_total_column
                
                if should_check and is_value_different(original_value, calculated_value):
                    diff = calculate_error_difference(original_value, calculated_value)
                    errors.append({
                        'section': section_name,
                        'name': name,
                        'code': code,
                        'level': level,
                        'type': 'Поступления',
                        'column': col,
                        'original': original_value,
                        'calculated': calculated_value,
                        'difference': diff
                    })
        
        return errors
    
    def check_deficit_proficit_errors(self, project_data: Dict[str, Any]) -> List[Dict[str, Any]]:
        """Проверка ошибок дефицита/профицита (строка 450 в разделе 'Расходы')
        
        Args:
            project_data: Данные проекта
        
        Returns:
            Список ошибок
        """
        errors = []
        calculated_deficit_proficit = project_data.get('calculated_deficit_proficit')
        if not calculated_deficit_proficit:
            return errors
        
        # Ищем строку 450 в разделе расходов
        расходы_data = project_data.get('расходы_data', [])
        if not расходы_data:
            return errors
        
        # Ищем строку с кодом 450
        строка_450 = None
        for item in расходы_data:
            if str(item.get('код_строки', '')).strip() == '450':
                строка_450 = item
                break
        
        if not строка_450:
            return errors
        
        budget_cols = Form0503317Constants.BUDGET_COLUMNS
        name = строка_450.get('наименование_показателя', 'Результат исполнения бюджета (дефицит/профицит)')
        code = строка_450.get('код_строки', '450')
        level = строка_450.get('уровень', 0)
        
        approved_data = строка_450.get('утвержденный', {}) or {}
        executed_data = строка_450.get('исполненный', {}) or {}
        
        calculated_approved = calculated_deficit_proficit.get('утвержденный', {}) or {}
        calculated_executed = calculated_deficit_proficit.get('исполненный', {}) or {}
        
        # Проверяем утвержденные значения
        for col in budget_cols:
            original_approved = approved_data.get(col, 0) or 0
            calc_approved = calculated_approved.get(col, 0) or 0
            
            if is_value_different(original_approved, calc_approved):
                diff = calculate_error_difference(original_approved, calc_approved)
                errors.append({
                    'section': 'Расходы',
                    'name': name,
                    'code': code,
                    'level': level,
                    'type': 'Утвержденный',
                    'column': col,
                    'original': original_approved,
                    'calculated': calc_approved,
                    'difference': diff
                })
        
        # Проверяем исполненные значения
        for col in budget_cols:
            original_executed = executed_data.get(col, 0) or 0
            calc_executed = calculated_executed.get(col, 0) or 0
            
            if is_value_different(original_executed, calc_executed):
                diff = calculate_error_difference(original_executed, calc_executed)
                errors.append({
                    'section': 'Расходы',
                    'name': name,
                    'code': code,
                    'level': level,
                    'type': 'Исполненный',
                    'column': col,
                    'original': original_executed,
                    'calculated': calc_executed,
                    'difference': diff
                })
        
        return errors

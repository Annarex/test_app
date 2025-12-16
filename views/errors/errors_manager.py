"""Управление ошибками расчетов"""
from PyQt5.QtWidgets import QTableWidgetItem, QComboBox, QLabel, QTableWidget, QMessageBox, QFileDialog
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QColor, QBrush
from logger import logger
from models.form_0503317 import Form0503317Constants


class ErrorsManager:
    """Класс для управления ошибками расчетов"""
    
    def __init__(self, main_window):
        """
        Args:
            main_window: Ссылка на главное окно для доступа к методам и свойствам
        """
        self.main_window = main_window
        self.errors_data = []
    
    def load_errors_to_tab(self, project_data):
        """Загрузка ошибок расчетов во вкладку ошибок"""
        self.errors_data = []
        
        if not project_data:
            # Обновляем все таблицы ошибок
            for widget_info in self._get_errors_widgets():
                self._update_errors_table(
                    widget_info.get('table'),
                    widget_info.get('filter'),
                    widget_info.get('stats')
                )
            return
        
        # Проверяем разделы
        sections = {
            "Доходы": "доходы_data",
            "Расходы": "расходы_data",
            "Источники финансирования": "источники_финансирования_data",
            "Консолидируемые расчеты": "консолидируемые_расчеты_data"
        }
        
        for section_name, section_key in sections.items():
            section_data = project_data.get(section_key, [])
            if not section_data:
                continue
            
            if section_name == "Консолидируемые расчеты":
                self._check_consolidated_errors(section_data, section_name)
            else:
                self._check_budget_errors(section_data, section_name)
        
        # Проверяем дефицит/профицит (строка 450 в разделе "Расходы")
        self._check_deficit_proficit_errors(project_data)
        
        # Обновляем все таблицы ошибок
        for widget_info in self._get_errors_widgets():
            self._update_errors_table(
                widget_info.get('table'),
                widget_info.get('filter'),
                widget_info.get('stats')
            )
    
    def _check_budget_errors(self, data, section_name: str):
        """Проверка ошибок для бюджетных разделов (доходы, расходы, источники)"""
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
                
                if self._is_value_different(original_approved, calculated_approved):
                    diff = self._calculate_error_difference(original_approved, calculated_approved)
                    self.errors_data.append({
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
                
                if self._is_value_different(original_executed, calculated_executed):
                    diff = self._calculate_error_difference(original_executed, calculated_executed)
                    self.errors_data.append({
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
    
    def _check_consolidated_errors(self, data, section_name: str):
        """Проверка ошибок для консолидированных расчетов"""
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
                
                if should_check and self._is_value_different(original_value, calculated_value):
                    diff = self._calculate_error_difference(original_value, calculated_value)
                    self.errors_data.append({
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
    
    def _check_deficit_proficit_errors(self, project_data):
        """Проверка ошибок дефицита/профицита (строка 450 в разделе 'Расходы')"""
        calculated_deficit_proficit = project_data.get('calculated_deficit_proficit')
        if not calculated_deficit_proficit:
            return
        
        # Ищем строку 450 в разделе расходов
        расходы_data = project_data.get('расходы_data', [])
        if not расходы_data:
            return
        
        # Ищем строку с кодом 450
        строка_450 = None
        for item in расходы_data:
            if str(item.get('код_строки', '')).strip() == '450':
                строка_450 = item
                break
        
        if not строка_450:
            return
        
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
            
            if self._is_value_different(original_approved, calc_approved):
                diff = self._calculate_error_difference(original_approved, calc_approved)
                self.errors_data.append({
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
            
            if self._is_value_different(original_executed, calc_executed):
                diff = self._calculate_error_difference(original_executed, calc_executed)
                self.errors_data.append({
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
    
    def _calculate_error_difference(self, original: float, calculated: float) -> float:
        """Вычисление разницы между значениями"""
        try:
            original_val = float(original) if original not in (None, "", "x") else 0.0
            calculated_val = float(calculated) if calculated not in (None, "", "x") else 0.0
            return calculated_val - original_val
        except (ValueError, TypeError):
            return 0.0
    
    def _is_value_different(self, original: float, calculated: float) -> bool:
        """Проверка различия значений"""
        try:
            original_val = float(original) if original not in (None, "", "x") else 0.0
            calculated_val = float(calculated) if calculated not in (None, "", "x") else 0.0
            return abs(original_val - calculated_val) > 0.00001
        except (ValueError, TypeError):
            return False
    
    def _update_errors_table(self, errors_table=None, section_filter_widget=None, stats_label=None):
        """Обновление таблицы с ошибками"""
        if errors_table is None:
            errors_table = self.main_window.errors_table
        if section_filter_widget is None:
            section_filter_widget = self.main_window.errors_section_filter
        if stats_label is None:
            stats_label = self.main_window.errors_stats_label
        
        # Фильтрация по разделу
        selected_section = section_filter_widget.currentText() if section_filter_widget else "Все"
        
        filtered_errors = self.errors_data
        if selected_section != "Все":
            filtered_errors = [e for e in self.errors_data if e['section'] == selected_section]
        
        # Очищаем таблицу
        errors_table.setRowCount(0)
        
        # Заполнение таблицы
        errors_table.setRowCount(len(filtered_errors))
        
        error_color = QColor("#FF6B6B")
        
        for row_idx, error in enumerate(filtered_errors):
            # Раздел
            errors_table.setItem(row_idx, 0, QTableWidgetItem(error['section']))
            
            # Наименование
            name_item = QTableWidgetItem(error['name'])
            name_item.setForeground(QBrush(error_color))
            errors_table.setItem(row_idx, 1, name_item)
            
            # Код строки
            errors_table.setItem(row_idx, 2, QTableWidgetItem(str(error['code'])))
            
            # Уровень
            errors_table.setItem(row_idx, 3, QTableWidgetItem(str(error['level'])))
            
            # Тип
            errors_table.setItem(row_idx, 4, QTableWidgetItem(error['type']))
            
            # Колонка
            errors_table.setItem(row_idx, 5, QTableWidgetItem(error['column']))
            
            # Оригинальное значение
            orig_text = self._format_error_value(error['original'])
            orig_item = QTableWidgetItem(orig_text)
            errors_table.setItem(row_idx, 6, orig_item)
            
            # Расчетное значение
            calc_text = self._format_error_value(error['calculated'])
            calc_item = QTableWidgetItem(calc_text)
            calc_item.setForeground(QBrush(error_color))
            errors_table.setItem(row_idx, 7, calc_item)
            
            # Разница
            diff_text = self._format_error_value(error['difference'])
            diff_item = QTableWidgetItem(diff_text)
            diff_item.setForeground(QBrush(error_color))
            errors_table.setItem(row_idx, 8, diff_item)
        
        # Убеждаемся, что режим изменения размера столбцов установлен
        from PyQt5.QtWidgets import QHeaderView
        header = errors_table.horizontalHeader()
        header.setStretchLastSection(False)
        for i in range(9):
            header.setSectionResizeMode(i, QHeaderView.Interactive)
        
        # Обновление статистики
        if stats_label:
            total_count = len(self.errors_data)
            filtered_count = len(filtered_errors)
            if selected_section == "Все":
                stats_label.setText(f"Всего ошибок: {total_count}")
            else:
                stats_label.setText(f"Ошибок в разделе '{selected_section}': {filtered_count} (всего: {total_count})")
    
    def _format_error_value(self, value) -> str:
        """Форматирование значения ошибки для отображения"""
        if value in (None, "", "x"):
            return str(value) if value else ""
        try:
            return f"{float(value):,.2f}"
        except (ValueError, TypeError):
            return str(value)
    
    def _export_errors(self):
        """Экспорт ошибок в файл"""
        import csv
        
        if not self.errors_data:
            QMessageBox.information(self.main_window, "Информация", "Нет ошибок для экспорта")
            return
        
        file_path, _ = QFileDialog.getSaveFileName(
            self.main_window,
            "Экспорт ошибок расчетов",
            "ошибки_расчетов.csv",
            "CSV files (*.csv);;All files (*.*)"
        )
        
        if not file_path:
            return
        
        try:
            with open(file_path, 'w', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f, delimiter=';')
                # Заголовки
                writer.writerow([
                    "Раздел", "Наименование", "Код строки", "Уровень",
                    "Тип", "Колонка", "Оригинальное", "Расчетное", "Разница"
                ])
                # Данные
                for error in self.errors_data:
                    writer.writerow([
                        error['section'],
                        error['name'],
                        error['code'],
                        error['level'],
                        error['type'],
                        error['column'],
                        self._format_error_value(error['original']),
                        self._format_error_value(error['calculated']),
                        self._format_error_value(error['difference'])
                    ])
            
            QMessageBox.information(self.main_window, "Успех", f"Ошибки экспортированы в файл:\n{file_path}")
        except Exception as e:
            logger.error(f"Ошибка экспорта: {e}", exc_info=True)
            QMessageBox.critical(self.main_window, "Ошибка", f"Не удалось экспортировать ошибки:\n{e}")
    
    def _get_errors_widgets(self):
        """Получить все виджеты ошибок с их фильтрами и метками (в главном окне и открепленных)"""
        widgets_info = []
        # Виджет в главном окне
        if hasattr(self.main_window, 'errors_tab') and self.main_window.errors_tab and hasattr(self.main_window, 'errors_table'):
            widgets_info.append({
                'table': self.main_window.errors_table,
                'filter': self.main_window.errors_section_filter,
                'stats': self.main_window.errors_stats_label
            })
        
        # Виджеты в открепленных окнах
        if hasattr(self.main_window, 'detached_windows') and "Ошибки" in self.main_window.detached_windows:
            detached_window = self.main_window.detached_windows["Ошибки"]
            tab_widget = detached_window.get_tab_widget()
            if tab_widget:
                # Ищем таблицу, фильтр и метку статистики в открепленном окне
                errors_table = None
                errors_filter = None
                errors_stats = None
                for child in tab_widget.findChildren(QTableWidget):
                    errors_table = child
                    break
                for child in tab_widget.findChildren(QComboBox):
                    errors_filter = child
                    break
                for child in tab_widget.findChildren(QLabel):
                    if "ошибок" in child.text().lower():
                        errors_stats = child
                        break
                if errors_table:
                    widgets_info.append({
                        'table': errors_table,
                        'filter': errors_filter,
                        'stats': errors_stats
                    })
        
        return widgets_info

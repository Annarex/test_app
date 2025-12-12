"""
Диалог для отображения всех ошибок расчетов по ревизии
"""
from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QTableWidget, QTableWidgetItem,
    QHeaderView, QPushButton, QLabel, QMessageBox, QComboBox, QAction, QApplication
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QColor, QBrush, QFont, QKeySequence
from typing import List, Dict, Any, Optional
from logger import logger
from models.form_0503317 import Form0503317Constants


class CalculationErrorsDialog(QDialog):
    """Диалог для отображения ошибок расчетов"""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Ошибки расчетов")
        self.setMinimumSize(1000, 600)
        self.errors_data = []
        self.is_fullscreen = False
        # Включаем стандартные кнопки окна (включая максимизацию)
        self.setWindowFlags(self.windowFlags() | Qt.WindowMaximizeButtonHint)
        
        self._init_ui()
    
    def _init_ui(self):
        """Инициализация интерфейса"""
        layout = QVBoxLayout(self)
        
        # Заголовок и фильтры
        header_layout = QHBoxLayout()
        
        info_label = QLabel("Ошибки расчетов (несоответствия между оригинальными и расчетными значениями):")
        info_label.setFont(QFont("Arial", 10, QFont.Bold))
        header_layout.addWidget(info_label)
        
        header_layout.addStretch()
        
        # Фильтр по разделу
        header_layout.addWidget(QLabel("Раздел:"))
        self.section_filter = QComboBox()
        self.section_filter.addItems(["Все", "Доходы", "Расходы", "Источники финансирования", "Консолидируемые расчеты"])
        self.section_filter.currentTextChanged.connect(self._apply_filters)
        header_layout.addWidget(self.section_filter)
        
        layout.addLayout(header_layout)
        
        # Таблица ошибок
        self.errors_table = QTableWidget()
        self.errors_table.setColumnCount(9)
        self.errors_table.setHorizontalHeaderLabels([
            "Раздел",
            "Наименование",
            "Код строки",
            "Уровень",
            "Тип",
            "Колонка",
            "Оригинальное",
            "Расчетное",
            "Разница"
        ])
        
        # Настройка таблицы
        header = self.errors_table.horizontalHeader()
        # Отключаем растягивание последнего столбца
        header.setStretchLastSection(False)
        # Используем Interactive режим для каждого столбца отдельно, чтобы можно было вручную изменять ширину
        for i in range(9):
            header.setSectionResizeMode(i, QHeaderView.Interactive)
        # Устанавливаем начальные размеры столбцов
        header.resizeSection(0, 120)  # Раздел
        header.resizeSection(1, 300)  # Наименование
        header.resizeSection(2, 100)  # Код строки
        header.resizeSection(3, 60)   # Уровень
        header.resizeSection(4, 120)  # Тип
        header.resizeSection(5, 100)  # Колонка
        header.resizeSection(6, 120)  # Оригинальное
        header.resizeSection(7, 120)  # Расчетное
        header.resizeSection(8, 120)  # Разница
        
        self.errors_table.setAlternatingRowColors(True)
        self.errors_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.errors_table.setEditTriggers(QTableWidget.NoEditTriggers)
        
        layout.addWidget(self.errors_table)
        
        # Кнопки
        buttons_layout = QHBoxLayout()
        buttons_layout.addStretch()
        
        refresh_btn = QPushButton("Обновить")
        refresh_btn.clicked.connect(self._refresh_errors)
        buttons_layout.addWidget(refresh_btn)
        
        export_btn = QPushButton("Экспорт...")
        export_btn.clicked.connect(self._export_errors)
        buttons_layout.addWidget(export_btn)
        
        close_btn = QPushButton("Закрыть")
        close_btn.clicked.connect(self.accept)
        buttons_layout.addWidget(close_btn)
        
        layout.addLayout(buttons_layout)
        
        # Статистика
        self.stats_label = QLabel("Ошибок не найдено")
        self.stats_label.setFont(QFont("Arial", 9))
        layout.addWidget(self.stats_label)
    
    def load_errors(self, project_data: Dict[str, Any]):
        """
        Загрузка ошибок расчетов из данных проекта
        
        Args:
            project_data: Словарь с данными проекта (доходы_data, расходы_data, и т.д.)
        """
        self.errors_data = []
        
        if not project_data:
            self._update_table()
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
        
        self._update_table()
    
    def _check_budget_errors(self, data: List[Dict], section_name: str):
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
                    diff = self._calculate_difference(original_approved, calculated_approved)
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
                    diff = self._calculate_difference(original_executed, calculated_executed)
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
    
    def _check_consolidated_errors(self, data: List[Dict], section_name: str):
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
                    diff = self._calculate_difference(original_value, calculated_value)
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
    
    def _is_value_different(self, original: float, calculated: float) -> bool:
        """Проверка различия значений"""
        try:
            original_val = float(original) if original not in (None, "", "x") else 0.0
            calculated_val = float(calculated) if calculated not in (None, "", "x") else 0.0
            return abs(original_val - calculated_val) > 0.00001
        except (ValueError, TypeError):
            return False
    
    def _calculate_difference(self, original: float, calculated: float) -> float:
        """Вычисление разницы между значениями"""
        try:
            original_val = float(original) if original not in (None, "", "x") else 0.0
            calculated_val = float(calculated) if calculated not in (None, "", "x") else 0.0
            return calculated_val - original_val
        except (ValueError, TypeError):
            return 0.0
    
    def _apply_filters(self):
        """Применение фильтров к таблице"""
        self._update_table()
    
    def _update_table(self):
        """Обновление таблицы с ошибками"""
        # Фильтрация по разделу
        section_filter = self.section_filter.currentText()
        filtered_data = self.errors_data
        if section_filter != "Все":
            filtered_data = [e for e in self.errors_data if e['section'] == section_filter]
        
        # Заполнение таблицы
        self.errors_table.setRowCount(len(filtered_data))
        
        error_color = QColor("#FF6B6B")
        
        for row_idx, error in enumerate(filtered_data):
            # Раздел
            self.errors_table.setItem(row_idx, 0, QTableWidgetItem(error['section']))
            
            # Наименование
            name_item = QTableWidgetItem(error['name'])
            name_item.setForeground(QBrush(error_color))
            self.errors_table.setItem(row_idx, 1, name_item)
            
            # Код строки
            self.errors_table.setItem(row_idx, 2, QTableWidgetItem(str(error['code'])))
            
            # Уровень
            self.errors_table.setItem(row_idx, 3, QTableWidgetItem(str(error['level'])))
            
            # Тип
            self.errors_table.setItem(row_idx, 4, QTableWidgetItem(error['type']))
            
            # Колонка
            self.errors_table.setItem(row_idx, 5, QTableWidgetItem(error['column']))
            
            # Оригинальное значение
            orig_text = self._format_value(error['original'])
            orig_item = QTableWidgetItem(orig_text)
            self.errors_table.setItem(row_idx, 6, orig_item)
            
            # Расчетное значение
            calc_text = self._format_value(error['calculated'])
            calc_item = QTableWidgetItem(calc_text)
            calc_item.setForeground(QBrush(error_color))
            self.errors_table.setItem(row_idx, 7, calc_item)
            
            # Разница
            diff_text = self._format_value(error['difference'])
            diff_item = QTableWidgetItem(diff_text)
            diff_item.setForeground(QBrush(error_color))
            self.errors_table.setItem(row_idx, 8, diff_item)
        
        # Обновление статистики
        total_count = len(self.errors_data)
        filtered_count = len(filtered_data)
        if section_filter == "Все":
            self.stats_label.setText(f"Всего ошибок: {total_count}")
        else:
            self.stats_label.setText(f"Ошибок в разделе '{section_filter}': {filtered_count} (всего: {total_count})")
        
        # Убеждаемся, что режим изменения размера столбцов установлен (на случай, если он был сброшен)
        header = self.errors_table.horizontalHeader()
        header.setStretchLastSection(False)
        for i in range(9):
            header.setSectionResizeMode(i, QHeaderView.Interactive)
    
    def _format_value(self, value) -> str:
        """Форматирование значения для отображения"""
        if value in (None, "", "x"):
            return ""
        try:
            return f"{float(value):,.2f}"
        except (ValueError, TypeError):
            return str(value)
    
    def _refresh_errors(self):
        """Обновление списка ошибок"""
        # Этот метод будет вызываться из главного окна для обновления данных
        if hasattr(self, '_refresh_callback') and self._refresh_callback:
            self._refresh_callback()
    
    def _export_errors(self):
        """Экспорт ошибок в файл"""
        from PyQt5.QtWidgets import QFileDialog
        import csv
        
        if not self.errors_data:
            QMessageBox.information(self, "Информация", "Нет ошибок для экспорта")
            return
        
        file_path, _ = QFileDialog.getSaveFileName(
            self,
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
                        self._format_value(error['original']),
                        self._format_value(error['calculated']),
                        self._format_value(error['difference'])
                    ])
            
            QMessageBox.information(self, "Успех", f"Ошибки экспортированы в файл:\n{file_path}")
        except Exception as e:
            logger.error(f"Ошибка экспорта: {e}", exc_info=True)
            QMessageBox.critical(self, "Ошибка", f"Не удалось экспортировать ошибки:\n{e}")
    
    def _toggle_fullscreen(self):
        """Переключение полноэкранного режима"""
        if self.is_fullscreen:
            self.showNormal()
            self.is_fullscreen = False
        else:
            self.showFullScreen()
            self.is_fullscreen = True
    
    def keyPressEvent(self, event):
        """Обработка нажатий клавиш"""
        if event.key() == Qt.Key_F11:
            self._toggle_fullscreen()
        else:
            super().keyPressEvent(event)


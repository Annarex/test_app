"""
Универсальный диалог для проверки текстов из Excel файла
Позволяет загружать данные из Excel, выбирать справочник и проверять на ошибки
"""
from PyQt5.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QPushButton,
                             QLabel, QTableWidget, QTableWidgetItem, QHeaderView,
                             QComboBox, QMessageBox, QFileDialog, QGroupBox,
                             QLineEdit, QDateEdit, QSpinBox, QCheckBox,
                             QProgressDialog, QTextBrowser, QSizePolicy, QWidget)
from PyQt5.QtCore import Qt, QDate
from PyQt5.QtGui import QColor
from datetime import datetime
import sqlite3
import pandas as pd
from logger import logger
from models.database import DatabaseManager
from utils.db_utils import get_filtered_view
from utils.text_validation.error_finder import find_errors
from utils.text_validation.text_comparator import find_differences


class ExcelPreviewDialog(QDialog):
    """Диалог предпросмотра данных Excel"""
    
    def __init__(self, excel_data, header_row, data_start_row, parent=None):
        super().__init__(parent)
        self.excel_data = excel_data
        self.header_row = header_row
        self.data_start_row = data_start_row
        
        self.setWindowTitle("Предпросмотр данных Excel")
        self.resize(1200, 700)
        self.init_ui()
        self.load_data()
    
    def init_ui(self):
        layout = QVBoxLayout(self)
        
        # Информация
        info_text = (
            f"Всего строк: <b>{len(self.excel_data)}</b> | "
            f"Столбцов: <b>{len(self.excel_data.columns)}</b> | "
            f"Заголовки: строка <b>{self.header_row}</b> | "
            f"Данные: с строки <b>{self.data_start_row}</b>"
        )
        info_label = QLabel(info_text)
        info_label.setStyleSheet("padding: 5px; background-color: #f0f0f0; border-radius: 3px;")
        layout.addWidget(info_label)
        
        # Таблица
        self.table = QTableWidget()
        self.table.setAlternatingRowColors(True)
        self.table.horizontalHeader().setStretchLastSection(True)
        layout.addWidget(self.table)
        
        # Подсказка
        hint = QLabel("💡 Коды с ведущими нулями (007, 0123) подсвечиваются желтым")
        hint.setStyleSheet("color: #666; font-size: 8pt; font-style: italic;")
        layout.addWidget(hint)
        
        # Кнопка закрытия
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        close_btn = QPushButton("Закрыть")
        close_btn.clicked.connect(self.accept)
        button_layout.addWidget(close_btn)
        layout.addLayout(button_layout)
    
    def load_data(self):
        """Загрузка данных в таблицу"""
        df = self.excel_data
        rows_to_show = min(len(df), 1000)  # Показываем до 1000 строк
        
        # Добавляем столбец с номерами строк
        self.table.setRowCount(rows_to_show)
        self.table.setColumnCount(len(df.columns) + 1)
        
        # Заголовки
        headers = ["№ строки"] + df.columns.tolist()
        self.table.setHorizontalHeaderLabels(headers)
        
        for row_idx in range(rows_to_show):
            # Номер строки в файле
            actual_row = self.data_start_row + row_idx
            row_num_item = QTableWidgetItem(str(actual_row))
            row_num_item.setFlags(row_num_item.flags() & ~Qt.ItemIsEditable)
            row_num_item.setBackground(QColor("#f0f0f0"))
            self.table.setItem(row_idx, 0, row_num_item)
            
            # Данные
            for col_idx, col_name in enumerate(df.columns):
                value = df.iloc[row_idx][col_name]
                display_value = str(value) if value and str(value).strip() else ""
                
                item = QTableWidgetItem(display_value)
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                
                # Подсветка кодов с ведущими нулями
                if display_value and len(display_value) >= 2:
                    if display_value[0] == '0' and any(c.isdigit() for c in display_value[1:]):
                        is_decimal = (
                            len(display_value) >= 3 and 
                            display_value[1] in ',.;' and 
                            display_value[2:].replace(',', '').replace('.', '').replace(';', '').isdigit()
                        )
                        if not is_decimal:
                            item.setBackground(QColor("#fff3cd"))
                            item.setToolTip("Код с ведущим нулем (сохранен как текст)")
                
                self.table.setItem(row_idx, col_idx + 1, item)
        
        self.table.resizeColumnsToContents()
        self.table.setColumnWidth(0, 80)


class UniversalTextValidationDialog(QDialog):
    """Универсальный диалог для проверки текстов из Excel"""
    
    def __init__(self, db_path: str, parent=None):
        super().__init__(parent)
        self.db_path = db_path
        self.db_manager = DatabaseManager(db_path)
        self.excel_data = None
        self.excel_file_path = None
        self.errors = []
        self.all_items = []
        self.reference_data = None
        self.reference_names = []
        self.reference_codes = []
        
        self.setWindowTitle("Проверка текстов из Excel")
        self.resize(1400, 900)
        self.setup_ui()
        self._load_oktmo_list()
        
    def setup_ui(self):
        """Создание интерфейса"""
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(12)
        
        # Панель настроек
        settings_group = QGroupBox("Параметры проверки")
        settings_layout = QVBoxLayout()
        settings_group.setLayout(settings_layout)
        
        # Excel файл
        excel_row = QHBoxLayout()
        excel_row.addWidget(QLabel("Excel файл:"))
        self.excel_path_label = QLabel("Файл не выбран")
        self.excel_path_label.setStyleSheet("color: #666; font-style: italic;")
        excel_row.addWidget(self.excel_path_label, 1)
        self.load_excel_btn = QPushButton("📂 Загрузить Excel")
        self.load_excel_btn.clicked.connect(self._load_excel)
        excel_row.addWidget(self.load_excel_btn)
        settings_layout.addLayout(excel_row)
        
        # Параметры загрузки Excel
        excel_params_row = QHBoxLayout()
        excel_params_row.addWidget(QLabel("Строка заголовков:"))
        self.header_row_spinbox = QSpinBox()
        self.header_row_spinbox.setRange(1, 100)
        self.header_row_spinbox.setValue(1)
        self.header_row_spinbox.setMaximumWidth(80)
        self.header_row_spinbox.setToolTip("Номер строки в файле, где находятся названия столбцов")
        self.header_row_spinbox.valueChanged.connect(self._on_header_row_changed)
        excel_params_row.addWidget(self.header_row_spinbox)
        
        excel_params_row.addSpacing(20)
        excel_params_row.addWidget(QLabel("Начало данных:"))
        self.data_start_row_spinbox = QSpinBox()
        self.data_start_row_spinbox.setRange(1, 100)
        self.data_start_row_spinbox.setValue(2)
        self.data_start_row_spinbox.setMaximumWidth(80)
        self.data_start_row_spinbox.setToolTip("Номер строки в файле, где начинаются данные")
        excel_params_row.addWidget(self.data_start_row_spinbox)
        
        self.reload_preview_btn = QPushButton("🔄 Обновить")
        self.reload_preview_btn.clicked.connect(self._reload_excel_preview)
        self.reload_preview_btn.setEnabled(False)
        excel_params_row.addWidget(self.reload_preview_btn)
        
        self.auto_detect_btn = QPushButton("🔍 Авто")
        self.auto_detect_btn.setToolTip("Автоматически определить начало таблицы")
        self.auto_detect_btn.clicked.connect(self._auto_detect_table_start)
        self.auto_detect_btn.setEnabled(False)
        excel_params_row.addWidget(self.auto_detect_btn)
        
        excel_params_row.addStretch()
        settings_layout.addLayout(excel_params_row)
        
        # Столбцы Excel
        columns_row = QHBoxLayout()
        columns_row.addWidget(QLabel("Столбец с текстом:"))
        self.text_column_combo = QComboBox()
        self.text_column_combo.setMinimumWidth(150)
        columns_row.addWidget(self.text_column_combo)
        
        columns_row.addSpacing(20)
        columns_row.addWidget(QLabel("Столбец с кодом:"))
        self.code_column_combo = QComboBox()
        self.code_column_combo.setMinimumWidth(150)
        self.code_column_combo.addItem("-- Без проверки кода --", None)
        columns_row.addWidget(self.code_column_combo)
        columns_row.addStretch()
        settings_layout.addLayout(columns_row)
        
        # Справочник
        reference_row = QHBoxLayout()
        reference_row.addWidget(QLabel("Справочник:"))
        self.reference_combo = QComboBox()
        self.reference_combo.addItem("Классификация доходов", "v_budgetclastypeinc_merged")
        self.reference_combo.addItem("Классификация расходов", "v_budgetclascosts_merged")
        self.reference_combo.addItem("Источники финансирования", "v_budgetclassources_merged")
        reference_row.addWidget(self.reference_combo, 1)
        
        reference_row.addSpacing(20)
        reference_row.addWidget(QLabel("Дата:"))
        self.date_edit = QDateEdit()
        self.date_edit.setCalendarPopup(True)
        self.date_edit.setDate(QDate.currentDate())
        self.date_edit.setDisplayFormat("dd.MM.yyyy")
        self.date_edit.dateChanged.connect(self._on_date_changed)
        reference_row.addWidget(self.date_edit)
        
        reference_row.addSpacing(20)
        reference_row.addWidget(QLabel("ОКТМО:"))
        self.oktmo_combo = QComboBox()
        self.oktmo_combo.setMinimumWidth(250)
        reference_row.addWidget(self.oktmo_combo)
        settings_layout.addLayout(reference_row)
        
        # Параметры проверки
        params_row = QHBoxLayout()
        params_row.addWidget(QLabel("Макс. расстояние:"))
        self.distance_spinbox = QSpinBox()
        self.distance_spinbox.setRange(1, 301)
        self.distance_spinbox.setValue(200)
        self.distance_spinbox.setMaximumWidth(80)
        self.distance_spinbox.valueChanged.connect(self._apply_filters)
        params_row.addWidget(self.distance_spinbox)
        
        params_row.addSpacing(20)
        self.show_all_checkbox = QCheckBox("Показать все тексты")
        self.show_all_checkbox.stateChanged.connect(self._apply_filters)
        params_row.addWidget(self.show_all_checkbox)
        
        params_row.addStretch()
        self.check_btn = QPushButton("🔍 Проверить")
        self.check_btn.clicked.connect(self._run_validation)
        self.check_btn.setEnabled(False)
        params_row.addWidget(self.check_btn)
        settings_layout.addLayout(params_row)
        
        layout.addWidget(settings_group)
        
        # Данные из Excel
        excel_group = QGroupBox("Загруженные данные из Excel")
        excel_layout = QVBoxLayout()
        excel_group.setLayout(excel_layout)
        
        # Информация о данных
        self.excel_info_label = QLabel("Данные не загружены")
        self.excel_info_label.setStyleSheet("color: #666; font-size: 9pt;")
        excel_layout.addWidget(self.excel_info_label)
        
        # Кнопка просмотра данных
        preview_row = QHBoxLayout()
        self.preview_btn = QPushButton("📄 Просмотр данных")
        self.preview_btn.clicked.connect(self._show_excel_preview)
        self.preview_btn.setEnabled(False)
        preview_row.addWidget(self.preview_btn)
        preview_row.addStretch()
        excel_layout.addLayout(preview_row)
        
        # Подсказка о сохранении форматов
        format_hint = QLabel("💡 Коды с ведущими нулями (007, 0123) сохраняются как текст и подсвечены желтым")
        format_hint.setStyleSheet("color: #666; font-size: 8pt; font-style: italic;")
        excel_layout.addWidget(format_hint)
        
        layout.addWidget(excel_group)
        
        # Результаты
        results_group = QGroupBox("Результаты проверки")
        results_layout = QVBoxLayout()
        results_group.setLayout(results_layout)
        
        self.stats_label = QLabel("Статистика: -")
        results_layout.addWidget(self.stats_label)
        
        self.errors_table = QTableWidget()
        self.errors_table.setAlternatingRowColors(True)
        self.errors_table.setColumnCount(5)
        self.errors_table.setHorizontalHeaderLabels([
            "№", "Текст из Excel", "Код из Excel", "D", "Лучшее совпадение в справочнике"
        ])
        header = self.errors_table.horizontalHeader()
        header.setStretchLastSection(True)
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.errors_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.errors_table.doubleClicked.connect(self._show_error_details)
        results_layout.addWidget(self.errors_table)
        
        layout.addWidget(results_group)
        
        # Кнопки
        buttons_row = QHBoxLayout()
        buttons_row.addStretch()
        
        self.export_btn = QPushButton("💾 Экспорт в Excel")
        self.export_btn.clicked.connect(self._export_results)
        self.export_btn.setEnabled(False)
        buttons_row.addWidget(self.export_btn)
        
        close_btn = QPushButton("Закрыть")
        close_btn.clicked.connect(self.accept)
        buttons_row.addWidget(close_btn)
        
        layout.addLayout(buttons_row)
        
    def _load_excel(self):
        """Загрузка Excel файла"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Выберите Excel файл", "", "Excel Files (*.xlsx *.xls)"
        )
        
        if not file_path:
            return
        
        self.excel_file_path = file_path
        self.header_row_spinbox.setValue(1)
        self.data_start_row_spinbox.setValue(2)
        
        # Автоматически определяем начало таблицы
        self._auto_detect_table_start()
        
        self._reload_excel_preview()
    
    def _on_header_row_changed(self, value):
        """При изменении строки заголовков - обновить минимум для начала данных"""
        # Данные не могут начинаться раньше заголовков
        if self.data_start_row_spinbox.value() < value:
            self.data_start_row_spinbox.setValue(value)
        self.data_start_row_spinbox.setMinimum(value)
    
    def _auto_detect_table_start(self):
        """Автоматическое определение начала таблицы"""
        if not self.excel_file_path:
            return
        
        try:
            # Читаем первые 30 строк файла для анализа
            df_sample = pd.read_excel(
                self.excel_file_path, 
                header=None,
                nrows=30
            )
            
            if df_sample.empty:
                return
            
            # Ищем строку с максимальным количеством непустых значений
            # Это вероятно заголовки
            max_filled = 0
            header_row_candidate = 1
            
            for idx, row in df_sample.iterrows():
                # Количество непустых ячеек
                filled_count = row.notna().sum()
                
                # Если больше половины столбцов заполнено и это максимум
                if filled_count > max_filled and filled_count >= len(df_sample.columns) * 0.5:
                    max_filled = filled_count
                    header_row_candidate = idx + 1  # +1 т.к. индексы с 0
            
            # Находим первую строку с данными после заголовков
            data_start_candidate = header_row_candidate + 1
            
            # Пропускаем пустые строки
            for idx in range(header_row_candidate, min(len(df_sample), header_row_candidate + 10)):
                row = df_sample.iloc[idx]
                filled_count = row.notna().sum()
                
                # Если строка имеет данные (хотя бы 20% заполнено)
                if filled_count >= len(df_sample.columns) * 0.2:
                    data_start_candidate = idx + 1
                    break
            
            # Устанавливаем значения
            self.header_row_spinbox.setValue(header_row_candidate)
            self.data_start_row_spinbox.setValue(data_start_candidate)
            
            self.auto_detect_btn.setEnabled(True)
            
            logger.info(f"Автоопределение: заголовки=строка {header_row_candidate}, данные=строка {data_start_candidate}")
            
        except Exception as e:
            logger.error(f"Ошибка автоопределения: {e}", exc_info=True)
            # В случае ошибки оставляем значения по умолчанию
            pass
    
    def _reload_excel_preview(self):
        """Перезагрузка Excel с учетом параметров"""
        if not self.excel_file_path:
            return
            
        try:
            header_row = self.header_row_spinbox.value()
            data_start_row = self.data_start_row_spinbox.value()
            
            # Пропускаем строки до заголовков
            skiprows = header_row - 1 if header_row > 1 else None
            
            # Читаем Excel как текст для сохранения ведущих нулей в кодах
            # dtype=str сохраняет "007", "0123", "01.05" как текст
            df = pd.read_excel(
                self.excel_file_path, 
                skiprows=skiprows, 
                header=0,
                dtype=str,
                keep_default_na=False
            )
            
            # Если данные начинаются не сразу после заголовков, удаляем промежуточные строки
            rows_between = data_start_row - header_row - 1
            if rows_between > 0:
                df = df.iloc[rows_between:].reset_index(drop=True)
            
            if df.empty:
                QMessageBox.warning(self, "Ошибка", "Excel файл пустой или все строки пропущены")
                return
                
            self.excel_data = df
            self.excel_path_label.setText(self.excel_file_path.split('\\')[-1])
            self.excel_path_label.setStyleSheet("color: #28a745; font-weight: bold;")
            
            # Заполняем комбобоксы
            columns = df.columns.tolist()
            self.text_column_combo.clear()
            self.text_column_combo.addItems(columns)
            
            self.code_column_combo.clear()
            self.code_column_combo.addItem("-- Без проверки кода --", None)
            self.code_column_combo.addItems(columns)
            
            # Активируем кнопки
            self.check_btn.setEnabled(True)
            self.reload_preview_btn.setEnabled(True)
            self.auto_detect_btn.setEnabled(True)
            self.preview_btn.setEnabled(True)
            
            # Обновляем информацию
            self._update_excel_info()
            
            logger.info(f"Загружен Excel: {self.excel_file_path}, header={header_row}, data_start={data_start_row}, строк: {len(df)}")
            
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось загрузить файл:\n{str(e)}")
            logger.error(f"Ошибка загрузки Excel: {e}", exc_info=True)
            
    def _show_excel_preview(self):
        """Показать предпросмотр данных Excel в отдельном окне"""
        if self.excel_data is None:
            QMessageBox.warning(self, "Ошибка", "Данные не загружены")
            return
        
        dialog = ExcelPreviewDialog(
            self.excel_data,
            self.header_row_spinbox.value(),
            self.data_start_row_spinbox.value(),
            self
        )
        dialog.exec_()
    
    def _update_excel_info(self):
        """Обновление информации о загруженных данных"""
        if self.excel_data is None:
            self.excel_info_label.setText("Данные не загружены")
            self.excel_info_label.setStyleSheet("color: #666; font-size: 9pt;")
            return
        
        header_row = self.header_row_spinbox.value()
        data_start_row = self.data_start_row_spinbox.value()
        total_rows = len(self.excel_data)
        
        info_text = (
            f"Всего строк данных: <b>{total_rows}</b> | "
            f"Заголовки: строка <b>{header_row}</b> | "
            f"Данные: с строки <b>{data_start_row}</b> | "
            f"Столбцов: <b>{len(self.excel_data.columns)}</b>"
        )
        
        rows_between = data_start_row - header_row - 1
        if rows_between > 0:
            info_text += f" | <span style='color: #ff8c00;'>Пропущено после заголовков: {rows_between}</span>"
        
        self.excel_info_label.setText(info_text)
        self.excel_info_label.setStyleSheet("color: #333; font-size: 9pt;")
    
    def _load_oktmo_list(self):
        """Загрузка списка ОКТМО из базы данных (8 разрядов)"""
        try:
            filter_date = self.date_edit.date().toString("yyyy-MM-dd")
            
            # Сохраняем текущий выбор
            selected_code = self.oktmo_combo.currentData() if self.oktmo_combo.count() > 0 else None
            
            # Очищаем и добавляем пункт "Все"
            self.oktmo_combo.clear()
            self.oktmo_combo.addItem("00000000 — Все муниципалитеты", "00000000")
            
            # Загружаем ОКТМО через DatabaseManager (8 разрядов, как в проектах)
            pairs = self.db_manager.load_oktmo_for_municipality(filter_date, code_length=8)
            pairs.sort(key=lambda p: (p[0] or "").lower())
            
            # Добавляем в комбобокс
            for code, name in pairs:
                display = f"{code} — {name}" if name else code
                self.oktmo_combo.addItem(display, code)
            
            # Восстанавливаем выбор
            if selected_code:
                idx = self.oktmo_combo.findData(selected_code)
                if idx >= 0:
                    self.oktmo_combo.setCurrentIndex(idx)
                    
            logger.info(f"Загружено ОКТМО: {len(pairs)} (дата: {filter_date})")
            
        except Exception as e:
            logger.error(f"Ошибка загрузки ОКТМО: {e}", exc_info=True)
    
    def _on_date_changed(self, qdate):
        """Обработчик изменения даты - обновляем список ОКТМО"""
        self._load_oktmo_list()
        
    def _run_validation(self):
        """Запуск проверки"""
        if self.excel_data is None:
            QMessageBox.warning(self, "Ошибка", "Сначала загрузите Excel файл")
            return
            
        text_column = self.text_column_combo.currentText()
        code_column = self.code_column_combo.currentText() if self.code_column_combo.currentData() else None
        
        reference_table = self.reference_combo.currentData()
        check_date = self.date_edit.date().toString("yyyy-MM-dd")
        oktmo = self.oktmo_combo.currentData()
        max_distance = self.distance_spinbox.value()
        
        # Извлекаем данные
        texts = self.excel_data[text_column].astype(str).str.strip().tolist()
        codes = self.excel_data[code_column].astype(str).str.strip().tolist() if code_column else None
        
        # Загружаем справочник
        try:
            conn = sqlite3.connect(self.db_path)
            
            # Формируем фильтр по ОКТМО
            if oktmo and oktmo != "00000000":
                ppocode_filter = ["00000000", oktmo]
            else:
                ppocode_filter = "00000000"
                
            reference_df = get_filtered_view(
                conn, reference_table,
                filter_date=check_date,
                filter_ppocode=ppocode_filter,
                deduplicate=True
            )
            conn.close()
            
            if reference_df.empty:
                QMessageBox.warning(self, "Ошибка", "Справочник пуст")
                return
                
            self.reference_data = reference_df
            self.reference_names = reference_df['name'].astype(str).str.strip().tolist()
            
            if codes:
                self.reference_codes = self._extract_reference_codes(reference_df, reference_table)
            else:
                self.reference_codes = None
                
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось загрузить справочник:\n{str(e)}")
            logger.error(f"Ошибка загрузки справочника: {e}", exc_info=True)
            return
            
        # Проверка
        progress = QProgressDialog("Проверка текстов...", "Отмена", 0, len(texts), self)
        progress.setWindowModality(Qt.WindowModal)
        progress.setMinimumDuration(0)
        
        try:
            def progress_callback(current, total):
                progress.setValue(current)
                return not progress.wasCanceled()
                
            all_results = find_errors(
                texts, self.reference_names,
                orig_codes=codes,
                reference_codes=self.reference_codes,
                max_distance=max_distance,
                progress_callback=progress_callback
            )
            
            progress.close()
            
            # Сохраняем все результаты
            self.all_items = sorted(all_results, key=lambda x: x.original_index)
            self.errors = [e for e in all_results if e.distance > 0]
            
            self._apply_filters()
            
            logger.info(f"Проверка завершена. Всего: {len(all_results)}, ошибок: {len(self.errors)}")
            
        except Exception as e:
            progress.close()
            QMessageBox.critical(self, "Ошибка", f"Ошибка при проверке:\n{str(e)}")
            logger.error(f"Ошибка проверки: {e}", exc_info=True)
            
    def _extract_reference_codes(self, df: pd.DataFrame, table_name: str) -> list:
        """Извлечение кодов из справочника"""
        try:
            if 'inc' in table_name.lower():
                codes = []
                for _, row in df.iterrows():
                    parts = []
                    for col in ['inctypecode', 'incsubtypecode', 'analyticalgroupcode']:
                        if col in row and pd.notna(row[col]):
                            parts.append(str(row[col]).strip())
                    codes.append(' '.join(parts))
                return codes
                
            elif 'cost' in table_name.lower():
                codes = []
                for _, row in df.iterrows():
                    parts = []
                    for col in ['grbscode', 'rzpr', 'kcsr', 'kvr']:
                        if col in row and pd.notna(row[col]):
                            parts.append(str(row[col]).strip())
                    codes.append(' '.join(parts))
                return codes
                
            elif 'source' in table_name.lower():
                if 'code' in df.columns:
                    return df['code'].astype(str).str.strip().tolist()
                    
            return []
            
        except Exception as e:
            logger.error(f"Ошибка извлечения кодов: {e}", exc_info=True)
            return []
            
    def _apply_filters(self):
        """Применение фильтров расстояния и показа всех"""
        if not self.all_items:
            return
            
        show_all = self.show_all_checkbox.isChecked()
        max_distance = self.distance_spinbox.value()
        
        # Выбираем источник
        source = self.all_items if show_all else self.errors
        
        # Фильтруем по расстоянию
        actual_max = max_distance + 1 if max_distance == self.distance_spinbox.maximum() else max_distance
        filtered = [e for e in source if e.distance <= actual_max]
        
        self._display_results(filtered)
        
    def _display_results(self, items):
        """Отображение результатов"""
        self.errors_table.setRowCount(len(items))
        
        # Статистика
        total = len(self.excel_data) if self.excel_data is not None else 0
        errors_count = len(self.errors)
        correct_count = len(self.all_items) - errors_count
        
        if self.show_all_checkbox.isChecked():
            self.stats_label.setText(
                f"Всего: {total} | "
                f"<span style='color: #dc3545;'>Ошибок: {errors_count}</span> | "
                f"<span style='color: #28a745;'>Корректных: {correct_count}</span> | "
                f"Отображено: {len(items)}"
            )
        else:
            self.stats_label.setText(
                f"Всего: {total} | Найдено ошибок: <span style='color: #dc3545;'>{len(items)}</span>"
            )
            
        # Таблица
        for row_idx, error in enumerate(items):
            # №
            num_item = QTableWidgetItem(str(error.original_index + 1))
            num_item.setTextAlignment(Qt.AlignCenter)
            self.errors_table.setItem(row_idx, 0, num_item)
            
            # Текст
            text_item = QTableWidgetItem(error.original_text[:100])
            self.errors_table.setItem(row_idx, 1, text_item)
            
            # Код
            code_text = ""
            if error.code_error:
                code_text = str(error.code_error.ref_code) if hasattr(error.code_error, 'ref_code') else ""
            code_item = QTableWidgetItem(code_text)
            self.errors_table.setItem(row_idx, 2, code_item)
            
            # Расстояние
            dist_item = QTableWidgetItem(str(error.distance))
            dist_item.setTextAlignment(Qt.AlignCenter)
            
            if error.distance == 0:
                dist_item.setBackground(QColor("#28a745"))
                dist_item.setForeground(QColor("white"))
            elif error.distance < 10:
                dist_item.setBackground(QColor("#90EE90"))
            elif error.distance < 50:
                dist_item.setBackground(QColor("#FFD700"))
            else:
                dist_item.setBackground(QColor("#FF6B6B"))
                dist_item.setForeground(QColor("white"))
                
            self.errors_table.setItem(row_idx, 3, dist_item)
            
            # Совпадение
            match_item = QTableWidgetItem(error.reference_text[:100])
            self.errors_table.setItem(row_idx, 4, match_item)
            
        self.errors_table.resizeColumnsToContents()
        self.export_btn.setEnabled(len(items) > 0)
        
    def _show_error_details(self, index):
        """Детали ошибки"""
        row = index.row()
        show_all = self.show_all_checkbox.isChecked()
        max_distance = self.distance_spinbox.value()
        actual_max = max_distance + 1 if max_distance == self.distance_spinbox.maximum() else max_distance
        
        source = self.all_items if show_all else self.errors
        filtered = [e for e in source if e.distance <= actual_max]
        
        if row < 0 or row >= len(filtered):
            return
            
        error = filtered[row]
        
        dialog = QDialog(self)
        dialog.setWindowTitle("Детали проверки")
        dialog.resize(900, 600)
        
        layout = QVBoxLayout(dialog)
        
        # Инфо
        info_widget = QWidget()
        info_widget.setObjectName("infoCard")
        info_layout = QHBoxLayout(info_widget)
        info_layout.setContentsMargins(6, 2, 6, 2)
        
        info_layout.addWidget(QLabel(f"<b>№:</b> {error.original_index + 1}"))
        info_layout.addWidget(QLabel(f"<b>D:</b> {error.distance}"))
        
        if error.code_error:
            status = "✓" if not error.code_error.has_error else "✗"
            color = "#28a745" if not error.code_error.has_error else "#dc3545"
            info_layout.addWidget(QLabel(f"<b>Код:</b> <span style='color: {color};'>{status}</span>"))
            
        info_layout.addStretch()
        layout.addWidget(info_widget)
        
        # Сравнение
        text_browser = QTextBrowser()
        
        orig_html = self._highlight_diff(error.original_text, error.diff_indices)
        ref_html = self._highlight_diff(error.reference_text, [c[0] for c in error.corrections])
        
        html = f"""
        <table width="100%" border="1" cellpadding="8" style="border-collapse: collapse;">
        <tr style="background-color: #f8f9fa;">
            <th width="50%">Текст из Excel</th>
            <th width="50%">Справочник</th>
        </tr>
        <tr>
            <td valign="top">{orig_html}</td>
            <td valign="top">{ref_html}</td>
        </tr>
        </table>
        """
        text_browser.setHtml(html)
        layout.addWidget(text_browser)
        
        close_btn = QPushButton("Закрыть")
        close_btn.clicked.connect(dialog.accept)
        layout.addWidget(close_btn, alignment=Qt.AlignRight)
        
        dialog.exec_()
        
    def _highlight_diff(self, text: str, indices: list) -> str:
        """Подсветка различий"""
        if not indices:
            return text
            
        html = ""
        for i, char in enumerate(text):
            if i in indices:
                html += f'<span style="background-color: #ffcccc;">{char}</span>'
            else:
                html += char
        return html
            
    def _export_results(self):
        """Экспорт в Excel"""
        show_all = self.show_all_checkbox.isChecked()
        max_distance = self.distance_spinbox.value()
        actual_max = max_distance + 1 if max_distance == self.distance_spinbox.maximum() else max_distance
        
        source = self.all_items if show_all else self.errors
        items = [e for e in source if e.distance <= actual_max]
        
        if not items:
            QMessageBox.warning(self, "Предупреждение", "Нет данных для экспорта")
            return
            
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Сохранить",
            f"validation_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            "Excel Files (*.xlsx)"
        )
        
        if not file_path:
            return
            
        try:
            data = []
            for e in items:
                data.append({
                    '№': e.original_index + 1,
                    'Текст из Excel': e.original_text,
                    'Код из Excel': e.code_error.ref_code if e.code_error else '',
                    'Расстояние': e.distance,
                    'Эталонный текст': e.reference_text,
                    'Статус': 'OK' if e.distance == 0 else 'Ошибка'
                })
                
            df = pd.DataFrame(data)
            df.to_excel(file_path, index=False, engine='openpyxl')
            
            QMessageBox.information(self, "Успех", f"Сохранено:\n{file_path}")
            logger.info(f"Экспорт: {file_path}")
            
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось экспортировать:\n{str(e)}")
            logger.error(f"Ошибка экспорта: {e}", exc_info=True)
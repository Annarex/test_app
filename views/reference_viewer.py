from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QTabWidget, 
                             QTableWidget, QTableWidgetItem, QLabel, QComboBox,
                             QHeaderView, QPushButton, QMessageBox, QTextEdit)
from PyQt5.QtCore import Qt
import pandas as pd

class ReferenceViewer(QWidget):
    """Виджет для просмотра справочников"""
    
    def __init__(self):
        super().__init__()
        self.references = {}
        self.init_ui()
    
    def init_ui(self):
        """Инициализация интерфейса"""
        layout = QVBoxLayout(self)
        
        # Панель управления
        control_layout = QHBoxLayout()
        
        control_layout.addWidget(QLabel("Справочник:"))
        self.ref_combo = QComboBox()
        self.ref_combo.currentTextChanged.connect(self.on_reference_changed)
        control_layout.addWidget(self.ref_combo)
        
        self.info_label = QLabel("Справочники не загружены")
        control_layout.addWidget(self.info_label)
        
        control_layout.addStretch()
        
        refresh_btn = QPushButton("Обновить")
        refresh_btn.clicked.connect(self.refresh_data)
        control_layout.addWidget(refresh_btn)
        
        layout.addLayout(control_layout)
        
        # Вкладки
        self.tabs = QTabWidget()
        
        # Вкладка с таблицей
        self.table_tab = QWidget()
        table_layout = QVBoxLayout(self.table_tab)
        
        self.ref_table = QTableWidget()
        self.ref_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        table_layout.addWidget(self.ref_table)
        
        self.tabs.addTab(self.table_tab, "Таблица")
        
        # Вкладка с информацией
        self.info_tab = QWidget()
        info_layout = QVBoxLayout(self.info_tab)
        
        self.ref_info = QTextEdit()
        self.ref_info.setReadOnly(True)
        info_layout.addWidget(self.ref_info)
        
        self.tabs.addTab(self.info_tab, "Информация")
        
        layout.addWidget(self.tabs)
    
    def load_references(self, references):
        """Загрузка справочников"""
        self.references = references
        self.ref_combo.clear()
        
        if references:
            self.ref_combo.addItems(list(references.keys()))
            self.info_label.setText(f"Загружено справочников: {len(references)}")
        else:
            self.info_label.setText("Справочники не загружены")
    
    def on_reference_changed(self, ref_type):
        """Обработка смены справочника"""
        if ref_type and ref_type in self.references:
            self.display_reference_data(ref_type)
    
    def display_reference_data(self, ref_type):
        """Отображение данных справочника"""
        df = self.references[ref_type]
        
        # Отображаем в таблице
        self.display_reference_table(df)
        
        # Отображаем информацию
        self.display_reference_info(df, ref_type)
    
    def display_reference_table(self, df):
        """Отображение справочника в таблице"""
        if df is None or df.empty:
            return
        
        # Настраиваем таблицу
        self.ref_table.setRowCount(len(df))
        self.ref_table.setColumnCount(len(df.columns))
        self.ref_table.setHorizontalHeaderLabels(df.columns.tolist())
        
        # Заполняем данные
        for row_idx, row in df.iterrows():
            for col_idx, value in enumerate(row):
                item = QTableWidgetItem(str(value) if pd.notna(value) else "")
                self.ref_table.setItem(row_idx, col_idx, item)
        
        # Настраиваем заголовки
        header = self.ref_table.horizontalHeader()
        for i in range(len(df.columns)):
            header.setSectionResizeMode(i, QHeaderView.ResizeToContents)
    
    def display_reference_info(self, df, ref_type):
        """Отображение информации о справочнике"""
        if df is None or df.empty:
            self.ref_info.setText("Нет данных")
            return
        
        info_text = f"<h2>Справочник: {ref_type}</h2>"
        info_text += f"<p><b>Количество записей:</b> {len(df)}</p>"
        info_text += f"<p><b>Колонки:</b> {', '.join(df.columns.tolist())}</p>"
        
        # Статистика по уровням
        if 'уровень_кода' in df.columns:
            level_stats = df['уровень_кода'].value_counts().sort_index()
            info_text += "<p><b>Распределение по уровням:</b></p><ul>"
            for level, count in level_stats.items():
                info_text += f"<li>Уровень {level}: {count} записей</li>"
            info_text += "</ul>"
        
        # Статистика по уникальным значениям
        info_text += "<p><b>Уникальные значения по колонкам:</b></p><ul>"
        for column in df.columns:
            unique_count = df[column].nunique()
            info_text += f"<li>{column}: {unique_count} уникальных значений</li>"
        info_text += "</ul>"
        
        # Примеры данных
        info_text += "<p><b>Примеры данных (первые 5 записей):</b></p>"
        info_text += "<table border='1' style='border-collapse: collapse;'>"
        info_text += "<tr>"
        for column in df.columns:
            info_text += f"<th style='padding: 5px;'>{column}</th>"
        info_text += "</tr>"
        
        for _, row in df.head().iterrows():
            info_text += "<tr>"
            for value in row:
                info_text += f"<td style='padding: 5px;'>{value}</td>"
            info_text += "</tr>"
        info_text += "</table>"
        
        self.ref_info.setHtml(info_text)
    
    def refresh_data(self):
        """Обновление данных"""
        if self.ref_combo.currentText():
            self.on_reference_changed(self.ref_combo.currentText())
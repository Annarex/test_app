from PyQt5.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QFormLayout, 
                             QLineEdit, QComboBox, QPushButton, QDialogButtonBox,
                             QLabel, QFileDialog, QMessageBox)
import os

class ReferenceDialog(QDialog):
    """Диалог загрузки справочника"""
    
    def __init__(self, parent=None, default_type: str = None):
        super().__init__(parent)
        self.setWindowTitle("Загрузка справочника")
        self.setModal(True)
        self.resize(500, 200)
        self.file_path = ""
        self.init_ui()

        # Если задан тип справочника по умолчанию - устанавливаем и фиксируем его
        if default_type in ("доходы", "источники"):
            index = self.ref_type_combo.findText(default_type)
            if index >= 0:
                self.ref_type_combo.setCurrentIndex(index)
                self.ref_type_combo.setEnabled(False)
    
    def init_ui(self):
        """Инициализация интерфейса"""
        layout = QVBoxLayout(self)
        
        # Форма ввода
        form_layout = QFormLayout()
        
        # Название справочника
        self.name_edit = QLineEdit()
        form_layout.addRow("Название справочника:", self.name_edit)
        
        # Тип справочника
        self.ref_type_combo = QComboBox()
        self.ref_type_combo.addItems(["доходы", "источники"])
        form_layout.addRow("Тип справочника:", self.ref_type_combo)
        
        # Файл справочника
        file_layout = QHBoxLayout()
        
        self.file_path_label = QLabel("Файл не выбран")
        file_layout.addWidget(self.file_path_label)
        
        browse_btn = QPushButton("Обзор...")
        browse_btn.clicked.connect(self.browse_file)
        file_layout.addWidget(browse_btn)
        
        form_layout.addRow("Файл справочника:", file_layout)
        
        layout.addLayout(form_layout)
        
        # Информация о необходимых колонках
        info_label = QLabel("Необходимые колонки:\n"
                           "• Для доходов: код_классификации_ДБ, уровень_кода\n"
                           "• Для источников: код_классификации_ИФДБ, уровень_кода")
        info_label.setWordWrap(True)
        layout.addWidget(info_label)
        
        # Кнопки
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)
    
    def browse_file(self):
        """Выбор файла"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Выберите файл справочника",
            "",
            "Excel files (*.xlsx *.xls);;All files (*.*)"
        )
        
        if file_path:
            self.file_path = file_path
            self.file_path_label.setText(os.path.basename(file_path))
            
            # Автоматически заполняем название, если оно пустое
            if not self.name_edit.text():
                self.name_edit.setText(os.path.basename(file_path).split('.')[0])
    
    def get_reference_data(self):
        """Получение данных справочника"""
        return {
            'name': self.name_edit.text(),
            'reference_type': self.ref_type_combo.currentText(),
            'file_path': self.file_path
        }
    
    def accept(self):
        """Проверка перед принятием"""
        if not self.name_edit.text():
            QMessageBox.warning(self, "Ошибка", "Введите название справочника")
            return
        
        if not self.file_path:
            QMessageBox.warning(self, "Ошибка", "Выберите файл справочника")
            return
        
        if not os.path.exists(self.file_path):
            QMessageBox.warning(self, "Ошибка", "Выбранный файл не существует")
            return
        
        super().accept()
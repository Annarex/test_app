"""
Диалог для формирования документов (заключения, письма)
"""
from PyQt5.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QFormLayout,
                             QPushButton, QLabel, QDateEdit, QLineEdit,
                             QDialogButtonBox, QGroupBox, QMessageBox)
from PyQt5.QtCore import Qt, QDate
from datetime import datetime
from logger import logger


class DocumentDialog(QDialog):
    """Диалог для формирования документов"""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Формирование документов")
        self.setMinimumWidth(500)
        
        self.result_paths = {}
        self.init_ui()
    
    def init_ui(self):
        """Инициализация UI"""
        layout = QVBoxLayout(self)
        
        # Группа для заключения
        conclusion_group = QGroupBox("Заключение")
        conclusion_layout = QFormLayout()
        
        self.protocol_date = QDateEdit()
        self.protocol_date.setDate(QDate.currentDate())
        self.protocol_date.setCalendarPopup(True)
        self.protocol_date.setDisplayFormat("dd.MM.yyyy")
        conclusion_layout.addRow("Дата протокола:", self.protocol_date)
        
        self.protocol_number = QLineEdit()
        conclusion_layout.addRow("Номер протокола:", self.protocol_number)
        
        self.letter_date = QDateEdit()
        self.letter_date.setDate(QDate.currentDate())
        self.letter_date.setCalendarPopup(True)
        self.letter_date.setDisplayFormat("dd.MM.yyyy")
        self.letter_date.setEnabled(False)
        conclusion_layout.addRow("Дата письма:", self.letter_date)
        
        self.letter_number = QLineEdit()
        self.letter_number.setEnabled(False)
        conclusion_layout.addRow("Номер письма:", self.letter_number)
        
        self.admin_date = QDateEdit()
        self.admin_date.setDate(QDate.currentDate())
        self.admin_date.setCalendarPopup(True)
        self.admin_date.setDisplayFormat("dd.MM.yyyy")
        self.admin_date.setEnabled(False)
        conclusion_layout.addRow("Дата постановления администрации:", self.admin_date)
        
        self.admin_number = QLineEdit()
        self.admin_number.setEnabled(False)
        conclusion_layout.addRow("Номер постановления администрации:", self.admin_number)
        
        conclusion_group.setLayout(conclusion_layout)
        layout.addWidget(conclusion_group)
        
        # Кнопки действий
        buttons_layout = QHBoxLayout()
        
        self.generate_conclusion_btn = QPushButton("Сформировать заключение")
        self.generate_conclusion_btn.clicked.connect(self.generate_conclusion)
        buttons_layout.addWidget(self.generate_conclusion_btn)
        
        self.generate_letters_btn = QPushButton("Сформировать письма")
        self.generate_letters_btn.clicked.connect(self.generate_letters)
        buttons_layout.addWidget(self.generate_letters_btn)
        
        layout.addLayout(buttons_layout)
        
        # Статус
        self.status_label = QLabel("")
        self.status_label.setWordWrap(True)
        layout.addWidget(self.status_label)
        
        # Кнопки диалога
        button_box = QDialogButtonBox(QDialogButtonBox.Close)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)
    
    def generate_conclusion(self):
        """Формирование заключения"""
        try:
            protocol_date = self.protocol_date.date().toPyDate()
            protocol_number = self.protocol_number.text().strip()
            
            if not protocol_number:
                QMessageBox.warning(self, "Ошибка", "Укажите номер протокола")
                return
            
            letter_date = None
            letter_number = None
            if self.letter_date.isEnabled():
                letter_date = self.letter_date.date().toPyDate()
                letter_number = self.letter_number.text().strip() or None
            
            admin_date = None
            admin_number = None
            if self.admin_date.isEnabled():
                admin_date = self.admin_date.date().toPyDate()
                admin_number = self.admin_number.text().strip() or None
            
            # Вызываем метод контроллера через родительское окно
            parent = self.parent()
            if hasattr(parent, 'controller'):
                result_path = parent.controller.generate_conclusion(
                    protocol_date=protocol_date,
                    protocol_number=protocol_number,
                    letter_date=letter_date,
                    letter_number=letter_number,
                    admin_date=admin_date,
                    admin_number=admin_number
                )
                
                if result_path:
                    self.result_paths['conclusion'] = result_path
                    self.status_label.setText(f"Заключение сформировано: {result_path}")
                    # Сохраняем путь в главном окне для возможности открытия
                    if hasattr(parent, 'last_exported_file'):
                        parent.last_exported_file = result_path
                    if hasattr(parent, 'open_last_file_action'):
                        parent.open_last_file_action.setEnabled(True)
                    if hasattr(parent, 'open_last_file_btn'):
                        parent.open_last_file_btn.setEnabled(True)
                    
                    # Предлагаем открыть файл
                    reply = QMessageBox.question(
                        self,
                        "Успех",
                        f"Заключение успешно сформировано:\n{result_path}\n\nОткрыть файл?",
                        QMessageBox.Yes | QMessageBox.No,
                        QMessageBox.Yes
                    )
                    if reply == QMessageBox.Yes and hasattr(parent, 'open_file'):
                        parent.open_file(result_path)
                else:
                    QMessageBox.warning(self, "Ошибка", "Не удалось сформировать заключение")
            else:
                QMessageBox.warning(self, "Ошибка", "Контроллер не найден")
                
        except Exception as e:
            logger.error(f"Ошибка формирования заключения: {e}", exc_info=True)
            QMessageBox.critical(self, "Ошибка", f"Ошибка формирования заключения:\n{str(e)}")
    
    def generate_letters(self):
        """Формирование писем"""
        try:
            protocol_date = self.protocol_date.date().toPyDate()
            protocol_number = self.protocol_number.text().strip()
            
            if not protocol_number:
                QMessageBox.warning(self, "Ошибка", "Укажите номер протокола")
                return
            
            # Вызываем метод контроллера через родительское окно
            parent = self.parent()
            if hasattr(parent, 'controller'):
                result = parent.controller.generate_letters(
                    protocol_date=protocol_date,
                    protocol_number=protocol_number
                )
                
                if result.get('admin') or result.get('council'):
                    self.result_paths.update(result)
                    status_text = []
                    if result.get('admin'):
                        status_text.append(f"Письмо администрации: {result['admin']}")
                    if result.get('council'):
                        status_text.append(f"Письмо совета: {result['council']}")
                    
                    self.status_label.setText("\n".join(status_text))
                    
                    # Сохраняем пути в главном окне для возможности открытия
                    if hasattr(parent, 'last_exported_file'):
                        # Сохраняем последний созданный файл (приоритет письму администрации)
                        parent.last_exported_file = result.get('admin') or result.get('council')
                    if hasattr(parent, 'open_last_file_action'):
                        parent.open_last_file_action.setEnabled(True)
                    if hasattr(parent, 'open_last_file_btn'):
                        parent.open_last_file_btn.setEnabled(True)
                    
                    # Предлагаем открыть файлы
                    files_to_open = []
                    if result.get('admin'):
                        files_to_open.append(('администрации', result['admin']))
                    if result.get('council'):
                        files_to_open.append(('совета', result['council']))
                    
                    if files_to_open:
                        files_text = "\n".join([f"Письмо {name}: {path}" for name, path in files_to_open])
                        reply = QMessageBox.question(
                            self,
                            "Успех",
                            f"Письма успешно сформированы:\n{files_text}\n\nОткрыть файлы?",
                            QMessageBox.Yes | QMessageBox.No,
                            QMessageBox.Yes
                        )
                        if reply == QMessageBox.Yes and hasattr(parent, 'open_file'):
                            for name, path in files_to_open:
                                parent.open_file(path)
                else:
                    QMessageBox.warning(self, "Ошибка", "Не удалось сформировать письма")
            else:
                QMessageBox.warning(self, "Ошибка", "Контроллер не найден")
                
        except Exception as e:
            logger.error(f"Ошибка формирования писем: {e}", exc_info=True)
            QMessageBox.critical(self, "Ошибка", f"Ошибка формирования писем:\n{str(e)}")


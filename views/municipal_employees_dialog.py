"""
Диалог для управления сотрудниками муниципальных образований
"""
from PyQt5.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QFormLayout,
                             QLineEdit, QDateEdit, QTextEdit, QPushButton,
                             QDialogButtonBox, QLabel, QComboBox, QGroupBox,
                             QMessageBox)
from PyQt5.QtCore import Qt, QDate
from datetime import datetime
from typing import Optional, Dict, Any
from logger import logger


class MunicipalEmployeeDialog(QDialog):
    """Диалог для добавления/редактирования сотрудников МО"""
    
    def __init__(self, db_manager, parent=None, employee_data: Optional[Dict[str, Any]] = None):
        super().__init__(parent)
        self.db_manager = db_manager
        self.employee_data = employee_data
        self.employee_id = employee_data.get('id') if employee_data else None
        
        self.setWindowTitle("Сотрудники муниципального образования")
        self.resize(800, 700)
        self.init_ui()
        
        if employee_data:
            self.load_employee_data(employee_data)
    
    def init_ui(self):
        """Инициализация интерфейса"""
        layout = QVBoxLayout(self)
        
        # Основная информация
        main_group = QGroupBox("Основная информация")
        main_layout = QFormLayout()
        
        # ОКТМО
        self.oktmo_combo = QComboBox()
        self.oktmo_combo.setEditable(True)
        self._load_oktmo_list()
        main_layout.addRow("Код ОКТМО (8 разрядов):", self.oktmo_combo)
        
        # Период действия
        period_layout = QHBoxLayout()
        self.startdate_edit = QDateEdit()
        self.startdate_edit.setCalendarPopup(True)
        self.startdate_edit.setDate(QDate.currentDate())
        self.startdate_edit.setDisplayFormat("dd.MM.yyyy")
        period_layout.addWidget(QLabel("С:"))
        period_layout.addWidget(self.startdate_edit)
        
        self.enddate_edit = QDateEdit()
        self.enddate_edit.setCalendarPopup(True)
        self.enddate_edit.setDate(QDate.currentDate().addYears(1))
        self.enddate_edit.setDisplayFormat("dd.MM.yyyy")
        self.enddate_edit.setSpecialValueText("Не задано")
        period_layout.addWidget(QLabel("По:"))
        period_layout.addWidget(self.enddate_edit)
        period_layout.addStretch()
        
        main_layout.addRow("Период действия:", period_layout)
        main_group.setLayout(main_layout)
        layout.addWidget(main_group)
        
        # Председатель совета
        council_group = QGroupBox("Председатель совета")
        council_layout = QFormLayout()
        
        self.council_position = QLineEdit()
        self.council_position.setPlaceholderText("Председатель совета депутатов")
        council_layout.addRow("Должность:", self.council_position)
        
        fio_layout = QHBoxLayout()
        self.council_surname = QLineEdit()
        self.council_surname.setPlaceholderText("Фамилия")
        fio_layout.addWidget(self.council_surname)
        
        self.council_first_name = QLineEdit()
        self.council_first_name.setPlaceholderText("Имя")
        fio_layout.addWidget(self.council_first_name)
        
        self.council_patronymic = QLineEdit()
        self.council_patronymic.setPlaceholderText("Отчество")
        fio_layout.addWidget(self.council_patronymic)
        
        council_layout.addRow("ФИО:", fio_layout)
        
        self.council_address = QTextEdit()
        self.council_address.setMaximumHeight(60)
        self.council_address.setPlaceholderText("Адрес совета")
        council_layout.addRow("Адрес:", self.council_address)
        
        self.council_email = QLineEdit()
        self.council_email.setPlaceholderText("email@example.com")
        council_layout.addRow("Email:", self.council_email)
        
        council_group.setLayout(council_layout)
        layout.addWidget(council_group)
        
        # Глава администрации
        admin_group = QGroupBox("Глава администрации")
        admin_layout = QFormLayout()
        
        self.admin_position = QLineEdit()
        self.admin_position.setPlaceholderText("Глава администрации")
        admin_layout.addRow("Должность:", self.admin_position)
        
        admin_fio_layout = QHBoxLayout()
        self.admin_surname = QLineEdit()
        self.admin_surname.setPlaceholderText("Фамилия")
        admin_fio_layout.addWidget(self.admin_surname)
        
        self.admin_first_name = QLineEdit()
        self.admin_first_name.setPlaceholderText("Имя")
        admin_fio_layout.addWidget(self.admin_first_name)
        
        self.admin_patronymic = QLineEdit()
        self.admin_patronymic.setPlaceholderText("Отчество")
        admin_fio_layout.addWidget(self.admin_patronymic)
        
        admin_layout.addRow("ФИО:", admin_fio_layout)
        
        self.admin_address = QTextEdit()
        self.admin_address.setMaximumHeight(60)
        self.admin_address.setPlaceholderText("Адрес администрации")
        admin_layout.addRow("Адрес:", self.admin_address)
        
        self.admin_email = QLineEdit()
        self.admin_email.setPlaceholderText("email@example.com")
        admin_layout.addRow("Email:", self.admin_email)
        
        admin_group.setLayout(admin_layout)
        layout.addWidget(admin_group)
        
        # Документы
        docs_group = QGroupBox("Документы")
        docs_layout = QFormLayout()
        
        self.agreement_date = QDateEdit()
        self.agreement_date.setCalendarPopup(True)
        self.agreement_date.setDisplayFormat("dd.MM.yyyy")
        self.agreement_date.setSpecialValueText("Не задано")
        docs_layout.addRow("Дата соглашения:", self.agreement_date)
        
        decision_layout = QHBoxLayout()
        self.decision_date = QDateEdit()
        self.decision_date.setCalendarPopup(True)
        self.decision_date.setDisplayFormat("dd.MM.yyyy")
        self.decision_date.setSpecialValueText("Не задано")
        decision_layout.addWidget(self.decision_date)
        
        self.decision_number = QLineEdit()
        self.decision_number.setPlaceholderText("Номер решения")
        decision_layout.addWidget(QLabel("№"))
        decision_layout.addWidget(self.decision_number)
        decision_layout.addStretch()
        
        docs_layout.addRow("Дата и номер решения:", decision_layout)
        
        docs_group.setLayout(docs_layout)
        layout.addWidget(docs_group)
        
        # Кнопки
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)
    
    def _load_oktmo_list(self):
        """Загрузка списка ОКТМО из БД с учетом даты действия"""
        try:
            # Загружаем ОКТМО через метод db_manager с учетом дат startdate/enddate
            filter_date = datetime.now().strftime('%Y-%m-%d')
            if self.employee_data and self.employee_data.get('startdate'):
                filter_date = self.employee_data.get('startdate')
            
            pairs = self.db_manager.load_oktmo_for_municipality(filter_date, code_length=8)
            pairs.sort(key=lambda p: (p[0] or "").lower())  # Сортируем по коду
            
            self.oktmo_combo.clear()
            for code, name in pairs:
                display = f"{code} — {name}" if name else code
                self.oktmo_combo.addItem(display, code)
                
            if self.oktmo_combo.count() == 0:
                self.oktmo_combo.addItem("Нет данных ОКТМО", "")
        except Exception as e:
            logger.error(f"Ошибка загрузки ОКТМО: {e}", exc_info=True)
            self.oktmo_combo.addItem("Ошибка загрузки", "")
    
    def load_employee_data(self, data: Dict[str, Any]):
        """Загрузка данных сотрудника в форму"""
        try:
            # ОКТМО
            oktmo_code = data.get('oktmo_code', '')
            if oktmo_code:
                index = self.oktmo_combo.findData(oktmo_code)
                if index >= 0:
                    self.oktmo_combo.setCurrentIndex(index)
                else:
                    self.oktmo_combo.setEditText(oktmo_code)
            
            # Период действия
            if data.get('startdate'):
                try:
                    date = datetime.strptime(data['startdate'], '%Y-%m-%d')
                    self.startdate_edit.setDate(QDate(date.year, date.month, date.day))
                except ValueError:
                    pass
            
            if data.get('enddate'):
                try:
                    date = datetime.strptime(data['enddate'], '%Y-%m-%d')
                    self.enddate_edit.setDate(QDate(date.year, date.month, date.day))
                except ValueError:
                    pass
            
            # Председатель совета
            self.council_position.setText(data.get('council_position', ''))
            self.council_surname.setText(data.get('council_surname', ''))
            self.council_first_name.setText(data.get('council_first_name', ''))
            self.council_patronymic.setText(data.get('council_patronymic', ''))
            self.council_address.setPlainText(data.get('council_address', ''))
            self.council_email.setText(data.get('council_email', ''))
            
            # Глава администрации
            self.admin_position.setText(data.get('administration_position', ''))
            self.admin_surname.setText(data.get('administration_surname', ''))
            self.admin_first_name.setText(data.get('administration_first_name', ''))
            self.admin_patronymic.setText(data.get('administration_patronymic', ''))
            self.admin_address.setPlainText(data.get('administration_address', ''))
            self.admin_email.setText(data.get('administration_email', ''))
            
            # Документы
            if data.get('agreement_date'):
                try:
                    date = datetime.strptime(data['agreement_date'], '%Y-%m-%d')
                    self.agreement_date.setDate(QDate(date.year, date.month, date.day))
                except ValueError:
                    pass
            
            if data.get('decision_date'):
                try:
                    date = datetime.strptime(data['decision_date'], '%Y-%m-%d')
                    self.decision_date.setDate(QDate(date.year, date.month, date.day))
                except ValueError:
                    pass
            
            self.decision_number.setText(data.get('decision_number', ''))
            
        except Exception as e:
            logger.error(f"Ошибка загрузки данных сотрудника: {e}", exc_info=True)
    
    def get_data(self) -> Dict[str, Any]:
        """Получение данных из полей формы"""
        # Получаем код ОКТМО
        oktmo_code = self.oktmo_combo.currentData()
        if not oktmo_code:
            oktmo_code = self.oktmo_combo.currentText().split('—')[0].strip()
        
        data = {
            'oktmo_code': oktmo_code,
            'startdate': self.startdate_edit.date().toString('yyyy-MM-dd'),
            'enddate': self.enddate_edit.date().toString('yyyy-MM-dd') if self.enddate_edit.date().isValid() else None,
            'council_position': self.council_position.text().strip(),
            'council_surname': self.council_surname.text().strip(),
            'council_first_name': self.council_first_name.text().strip(),
            'council_patronymic': self.council_patronymic.text().strip(),
            'council_address': self.council_address.toPlainText().strip(),
            'council_email': self.council_email.text().strip(),
            'administration_position': self.admin_position.text().strip(),
            'administration_surname': self.admin_surname.text().strip(),
            'administration_first_name': self.admin_first_name.text().strip(),
            'administration_patronymic': self.admin_patronymic.text().strip(),
            'administration_address': self.admin_address.toPlainText().strip(),
            'administration_email': self.admin_email.text().strip(),
            'agreement_date': self.agreement_date.date().toString('yyyy-MM-dd') if self.agreement_date.date().isValid() else None,
            'decision_date': self.decision_date.date().toString('yyyy-MM-dd') if self.decision_date.date().isValid() else None,
            'decision_number': self.decision_number.text().strip(),
        }
        
        if self.employee_id:
            data['id'] = self.employee_id
        
        return data
    
    def accept(self):
        """Валидация и сохранение данных"""
        data = self.get_data()
        
        # Валидация
        if not data['oktmo_code']:
            QMessageBox.warning(self, "Ошибка", "Необходимо указать код ОКТМО")
            return
        
        if len(data['oktmo_code']) != 8:
            QMessageBox.warning(self, "Ошибка", "Код ОКТМО должен содержать 8 разрядов")
            return
        
        super().accept()

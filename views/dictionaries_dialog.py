from PyQt5.QtWidgets import (
    QDialog,
    QVBoxLayout,
    QHBoxLayout,
    QTabWidget,
    QWidget,
    QTableWidget,
    QTableWidgetItem,
    QPushButton,
    QHeaderView,
    QLabel,
    QMessageBox,
)
from PyQt5.QtCore import Qt

from models.database import DatabaseManager
from models.base_models import YearRef, MunicipalityRef, FormTypeMeta, PeriodRef


class DictionariesDialog(QDialog):
    """
    Диалог редактирования справочников конфигурации:
    - Годы
    - Муниципальные образования (МО)
    - Типы форм
    - Периоды
    """

    def __init__(self, db_manager: DatabaseManager, parent=None):
        super().__init__(parent)
        self.db_manager = db_manager
        self.setWindowTitle("Справочники конфигурации")
        self.resize(900, 600)

        main_layout = QVBoxLayout(self)

        self.tabs = QTabWidget()
        main_layout.addWidget(self.tabs)

        # Вкладка годов
        self.years_tab = QWidget()
        self.years_table = QTableWidget()
        self._setup_years_tab()

        # Вкладка МО
        self.municip_tab = QWidget()
        self.municip_table = QTableWidget()
        self._setup_municip_tab()

        # Вкладка типов форм
        self.forms_tab = QWidget()
        self.forms_table = QTableWidget()
        self._setup_forms_tab()

        # Вкладка периодов
        self.periods_tab = QWidget()
        self.periods_table = QTableWidget()
        self._setup_periods_tab()

        # Кнопки управления
        buttons_layout = QHBoxLayout()
        buttons_layout.addStretch()

        self.reload_btn = QPushButton("Перезагрузить")
        self.reload_btn.clicked.connect(self.reload_all)
        buttons_layout.addWidget(self.reload_btn)

        self.save_btn = QPushButton("Сохранить")
        self.save_btn.clicked.connect(self.save_all)
        buttons_layout.addWidget(self.save_btn)

        self.close_btn = QPushButton("Закрыть")
        self.close_btn.clicked.connect(self.accept)
        buttons_layout.addWidget(self.close_btn)

        main_layout.addLayout(buttons_layout)

        # Инициализируем данными
        self.reload_all()

    # --- Настройка вкладок ---

    def _setup_years_tab(self):
        layout = QVBoxLayout(self.years_tab)
        layout.addWidget(QLabel("Годы (используются для структурирования проектов)"))
        self.years_table.setColumnCount(2)
        self.years_table.setHorizontalHeaderLabels(["Год", "Активен (1/0)"])
        self.years_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        layout.addWidget(self.years_table)
        self.tabs.addTab(self.years_tab, "Годы")

    def _setup_municip_tab(self):
        layout = QVBoxLayout(self.municip_tab)
        layout.addWidget(QLabel("Муниципальные образования"))
        self.municip_table.setColumnCount(3)
        self.municip_table.setHorizontalHeaderLabels(["Код", "Наименование", "Активен (1/0)"])
        self.municip_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        layout.addWidget(self.municip_table)
        self.tabs.addTab(self.municip_tab, "МО")

    def _setup_forms_tab(self):
        layout = QVBoxLayout(self.forms_tab)
        layout.addWidget(QLabel("Типы форм (например, 0503317, 0503314 и т.п.)"))
        # Добавляем явный ID типа формы, чтобы можно было задавать его вручную и сохранять
        self.forms_table.setColumnCount(5)
        self.forms_table.setHorizontalHeaderLabels(
            ["ID", "Код формы", "Наименование", "Периодичность", "Активен (1/0)"]
        )
        self.forms_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        layout.addWidget(self.forms_table)
        self.tabs.addTab(self.forms_tab, "Типы форм")

    def _setup_periods_tab(self):
        layout = QVBoxLayout(self.periods_tab)
        layout.addWidget(
            QLabel("Периоды (код: Y, Q1, Q2, H1 и т.п.; периодичность формы задаётся в типе формы)")
        )
        # Добавляем явный ID периода, чтобы можно было задавать его вручную и сохранять
        self.periods_table.setColumnCount(6)
        self.periods_table.setHorizontalHeaderLabels(
            ["ID", "Код периода", "Наименование", "Порядок сортировки", "Код формы (опц.)", "Активен (1/0)"]
        )
        self.periods_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        layout.addWidget(self.periods_table)
        self.tabs.addTab(self.periods_tab, "Периоды")

    # --- Загрузка данных ---

    def reload_all(self):
        self._load_years()
        self._load_municipalities()
        self._load_forms()
        self._load_periods()

    def _load_years(self):
        years = self.db_manager.load_years()
        self.years_table.setRowCount(len(years) + 1)  # +1 пустая строка для добавления
        for row_idx, y in enumerate(years):
            self.years_table.setItem(row_idx, 0, QTableWidgetItem(str(y.year)))
            self.years_table.setItem(row_idx, 1, QTableWidgetItem("1" if y.is_active else "0"))

    def _load_municipalities(self):
        municip = self.db_manager.load_municipalities()
        self.municip_table.setRowCount(len(municip) + 1)
        for row_idx, m in enumerate(municip):
            self.municip_table.setItem(row_idx, 0, QTableWidgetItem(m.code or ""))
            self.municip_table.setItem(row_idx, 1, QTableWidgetItem(m.name or ""))
            self.municip_table.setItem(row_idx, 2, QTableWidgetItem("1" if m.is_active else "0"))

    def _load_forms(self):
        forms = self.db_manager.load_form_types_meta()
        self.forms_table.setRowCount(len(forms) + 1)
        for row_idx, f in enumerate(forms):
            # ID типа формы (можно менять вручную, по умолчанию показываем существующий)
            self.forms_table.setItem(row_idx, 0, QTableWidgetItem(str(f.id) if f.id is not None else ""))
            self.forms_table.setItem(row_idx, 1, QTableWidgetItem(f.code or ""))
            self.forms_table.setItem(row_idx, 2, QTableWidgetItem(f.name or ""))
            self.forms_table.setItem(row_idx, 3, QTableWidgetItem(f.periodicity or ""))
            self.forms_table.setItem(row_idx, 4, QTableWidgetItem("1" if f.is_active else "0"))

    def _load_periods(self):
        periods = self.db_manager.load_periods()
        self.periods_table.setRowCount(len(periods) + 1)
        for row_idx, p in enumerate(periods):
            # ID периода (можно менять вручную, но по умолчанию выводим существующий)
            self.periods_table.setItem(row_idx, 0, QTableWidgetItem(str(p.id) if p.id is not None else ""))
            self.periods_table.setItem(row_idx, 1, QTableWidgetItem(p.code or ""))
            self.periods_table.setItem(row_idx, 2, QTableWidgetItem(p.name or ""))
            self.periods_table.setItem(row_idx, 3, QTableWidgetItem(str(p.sort_order)))
            self.periods_table.setItem(row_idx, 4, QTableWidgetItem(p.form_type_code or ""))
            self.periods_table.setItem(row_idx, 5, QTableWidgetItem("1" if p.is_active else "0"))

    # --- Сохранение данных ---

    def save_all(self):
        try:
            self._save_years()
            self._save_municipalities()
            self._save_forms()
            self._save_periods()
            QMessageBox.information(self, "Сохранено", "Справочники успешно сохранены.")
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Ошибка сохранения справочников: {e}")

    def _save_years(self):
        years: list[YearRef] = []
        for row in range(self.years_table.rowCount()):
            year_item = self.years_table.item(row, 0)
            active_item = self.years_table.item(row, 1)
            if not year_item or not year_item.text().strip():
                continue
            try:
                year_val = int(year_item.text())
            except ValueError:
                continue
            is_active = (active_item.text().strip() == "1") if active_item else True
            y = YearRef()
            y.year = year_val
            y.is_active = is_active
            years.append(y)
        self.db_manager.save_years_bulk(years)

    def _save_municipalities(self):
        municip_list: list[MunicipalityRef] = []
        for row in range(self.municip_table.rowCount()):
            code_item = self.municip_table.item(row, 0)
            name_item = self.municip_table.item(row, 1)
            active_item = self.municip_table.item(row, 2)
            if not name_item or not name_item.text().strip():
                continue
            m = MunicipalityRef()
            m.code = (code_item.text() if code_item else "").strip()
            m.name = name_item.text().strip()
            m.is_active = (active_item.text().strip() == "1") if active_item else True
            municip_list.append(m)
        self.db_manager.save_municipalities_bulk(municip_list)

    def _save_forms(self):
        forms_list: list[FormTypeMeta] = []
        for row in range(self.forms_table.rowCount()):
            id_item = self.forms_table.item(row, 0)
            code_item = self.forms_table.item(row, 1)
            name_item = self.forms_table.item(row, 2)
            periodicity_item = self.forms_table.item(row, 3)
            active_item = self.forms_table.item(row, 4)
            if not code_item or not code_item.text().strip():
                continue
            f = FormTypeMeta()
            # ID может задаваться вручную через интерфейс
            if id_item and id_item.text().strip():
                try:
                    f.id = int(id_item.text().strip())
                except ValueError:
                    f.id = None
            f.code = code_item.text().strip()
            f.name = (name_item.text() if name_item else f.code).strip()
            f.periodicity = (periodicity_item.text() if periodicity_item else "").strip()
            f.is_active = (active_item.text().strip() == "1") if active_item else True
            forms_list.append(f)
        self.db_manager.save_form_types_bulk(forms_list)

    def _save_periods(self):
        periods_list: list[PeriodRef] = []
        for row in range(self.periods_table.rowCount()):
            id_item = self.periods_table.item(row, 0)
            code_item = self.periods_table.item(row, 1)
            name_item = self.periods_table.item(row, 2)
            sort_item = self.periods_table.item(row, 3)
            form_code_item = self.periods_table.item(row, 4)
            active_item = self.periods_table.item(row, 5)
            if not code_item or not code_item.text().strip():
                continue
            p = PeriodRef()
            # ID может задаваться вручную через интерфейс
            if id_item and id_item.text().strip():
                try:
                    p.id = int(id_item.text().strip())
                except ValueError:
                    p.id = None
            p.code = code_item.text().strip()
            p.name = (name_item.text() if name_item else p.code).strip()
            try:
                p.sort_order = int(sort_item.text()) if sort_item and sort_item.text().strip() else 0
            except ValueError:
                p.sort_order = 0
            p.form_type_code = (form_code_item.text() if form_code_item else "").strip()
            p.is_active = (active_item.text().strip() == "1") if active_item else True
            periods_list.append(p)
        self.db_manager.save_periods_bulk(periods_list)



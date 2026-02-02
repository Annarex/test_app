from PyQt5.QtWidgets import (
    QDialog,
    QVBoxLayout,
    QFormLayout,
    QLineEdit,
    QComboBox,
    QDialogButtonBox,
    QLabel,
    QDateEdit,
)
from PyQt5.QtCore import QDate
from datetime import datetime, date

from models.base_models import FormType, Project
from models.database import DatabaseManager


class ProjectDialog(QDialog):
    """
    Диалог создания/редактирования проекта.
    
    В проекте задаются: название, год (ref_years), дата создания,
    муниципальное образование (из справочника ОКТМО, 8 разрядов, по дате создания).
    Тип формы, период и ревизии задаются при загрузке формы.
    """

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Создание проекта")
        self.setModal(True)
        self.resize(420, 300)

        self.db_manager: DatabaseManager = (
            getattr(getattr(parent, "controller", None), "db_manager", None)
            or DatabaseManager()
        )

        self._years_cache = []

        self.init_ui()
        self._load_years()
        filter_date = date.today().isoformat()
        self._load_oktmo_for_dialog(filter_date)

    def init_ui(self):
        """Инициализация интерфейса"""
        layout = QVBoxLayout(self)
        form_layout = QFormLayout()

        self.name_edit = QLineEdit()
        form_layout.addRow("Название проекта:", self.name_edit)

        self.year_combo = QComboBox()
        form_layout.addRow("Год:", self.year_combo)

        self.created_at_edit = QDateEdit()
        self.created_at_edit.setCalendarPopup(True)
        self.created_at_edit.setDate(QDate.currentDate())
        self.created_at_edit.setDisplayFormat("dd.MM.yyyy")
        form_layout.addRow("Дата создания:", self.created_at_edit)

        self.municipality_combo = QComboBox()
        form_layout.addRow("Муниципальное образование (ОКТМО):", self.municipality_combo)

        self.created_at_edit.dateChanged.connect(self._on_created_at_changed)

        layout.addLayout(form_layout)

        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)

    def _load_years(self):
        """Загрузка годов из справочника в комбобокс"""
        self.year_combo.clear()
        self._years_cache = self.db_manager.load_years()
        self._years_cache.sort(key=lambda y: y.year, reverse=True)
        for y in self._years_cache:
            self.year_combo.addItem(str(y.year), y.year)
        if self.year_combo.count() == 0:
            current_year = datetime.now().year
            self.year_combo.addItem(str(current_year), current_year)

    def _load_oktmo_for_dialog(self, filter_date: str):
        """Загрузка МО из ОКТМО (8 разрядов) по дате в комбобокс."""
        selected_code = self.municipality_combo.currentData() if self.municipality_combo.count() > 0 else None
        self.municipality_combo.clear()
        pairs = self.db_manager.load_oktmo_for_municipality(filter_date, code_length=8)
        pairs.sort(key=lambda p: (p[0] or "").lower())
        for code, name in pairs:
            display = f"{code} — {name}" if name else code
            self.municipality_combo.addItem(display, code)
        if self.municipality_combo.count() == 0:
            self.municipality_combo.addItem("Не задано", "")
        if selected_code:
            idx = self.municipality_combo.findData(selected_code)
            if idx >= 0:
                self.municipality_combo.setCurrentIndex(idx)

    def _on_created_at_changed(self, qdate: QDate):
        filter_date = f"{qdate.year():04d}-{qdate.month():02d}-{qdate.day():02d}"
        self._load_oktmo_for_dialog(filter_date)

    def set_project(self, project: Project):
        """Заполнить диалог данными существующего проекта (режим редактирования)."""
        if not project:
            return

        self.setWindowTitle("Редактирование проекта")

        self.name_edit.setText(project.name or "")

        year_val = None
        if project.year_id and self._years_cache:
            try:
                year_ref = next((y for y in self._years_cache if y.id == project.year_id), None)
                if year_ref:
                    year_val = year_ref.year
            except Exception:
                year_val = None
        if year_val:
            idx = self.year_combo.findData(year_val)
            if idx >= 0:
                self.year_combo.setCurrentIndex(idx)

        created_at = project.created_at or datetime.now()
        self.created_at_edit.setDate(QDate(created_at.year, created_at.month, created_at.day))

        if project.oktmo_code:
            idx = self.municipality_combo.findData((project.oktmo_code or "").strip())
            if idx >= 0:
                self.municipality_combo.setCurrentIndex(idx)

    def get_project_data(self):
        """Получение данных проекта."""
        name = self.name_edit.text().strip()

        year_id = None
        year_val_int = datetime.now().year
        if self.year_combo.count() > 0:
            year_val = self.year_combo.currentData()
            try:
                year_val_int = int(year_val)
                year_ref = next((y for y in self._years_cache if y.year == year_val_int), None)
                if year_ref:
                    year_id = year_ref.id
                else:
                    year_ref = self.db_manager.get_or_create_year(year_val_int)
                    year_id = year_ref.id
            except (TypeError, ValueError):
                pass

        oktmo_code = None
        if self.municipality_combo.count() > 0:
            oktmo_code = self.municipality_combo.currentData()
            if oktmo_code is not None:
                oktmo_code = (oktmo_code or "").strip() or None

        qd = self.created_at_edit.date()
        created_at = datetime(qd.year(), qd.month(), qd.day())

        return {
            "name": name,
            "year_id": year_id,
            "oktmo_code": oktmo_code,
            "created_at": created_at,
        }
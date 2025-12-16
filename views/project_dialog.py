from PyQt5.QtWidgets import (
    QDialog,
    QVBoxLayout,
    QFormLayout,
    QLineEdit,
    QComboBox,
    QDialogButtonBox,
    QLabel,
)
from PyQt5.QtCore import QDate
from datetime import datetime, date

from models.base_models import FormType, Project
from models.database import DatabaseManager


class ProjectDialog(QDialog):
    """
    Диалог создания/редактирования проекта.
    
    В проекте напрямую задаются только:
    - название проекта;
    - год (из справочника годов);
    - муниципальное образование (из справочника МО).
    
    Тип формы, период и ревизии задаются позже при загрузке формы
    и хранятся в связанной архитектуре project_forms / form_revisions.
    """

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Создание проекта")
        self.setModal(True)
        self.resize(420, 260)

        # Получаем db_manager от родителя, если он есть
        self.db_manager: DatabaseManager = (
            getattr(getattr(parent, "controller", None), "db_manager", None)
            or DatabaseManager()
        )

        self._years_cache = []
        self._municip_cache = []

        self.init_ui()
        self._load_years()
        self._load_municipalities()

    def init_ui(self):
        """Инициализация интерфейса"""
        layout = QVBoxLayout(self)

        form_layout = QFormLayout()

        # Название проекта
        self.name_edit = QLineEdit()
        form_layout.addRow("Название проекта:", self.name_edit)

        # Год проекта (из справочника ref_years)
        self.year_combo = QComboBox()
        form_layout.addRow("Год:", self.year_combo)

        # Муниципальное образование (из справочника ref_municipalities)
        self.municipality_combo = QComboBox()
        form_layout.addRow("Муниципальное образование:", self.municipality_combo)

        layout.addLayout(form_layout)

        # Кнопки
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)

    def _load_years(self):
        """Загрузка годов из справочника в комбобокс"""
        self.year_combo.clear()
        self._years_cache = self.db_manager.load_years()
        # Сортируем по убыванию года
        self._years_cache.sort(key=lambda y: y.year, reverse=True)
        for y in self._years_cache:
            self.year_combo.addItem(str(y.year), y.year)
        if self.year_combo.count() == 0:
            # Если справочник пуст — добавляем текущий год как fallback (только в UI)
            current_year = datetime.now().year
            self.year_combo.addItem(str(current_year), current_year)

    def _load_municipalities(self):
        """Загрузка МО из справочника в комбобокс"""
        self.municipality_combo.clear()
        self._municip_cache = self.db_manager.load_municipalities()
        # Сортируем по имени
        self._municip_cache.sort(key=lambda m: m.name.lower() if m.name else "")
        for m in self._municip_cache:
            display = f"{m.code} — {m.name}" if m.code else m.name
            self.municipality_combo.addItem(display, m.name)
        if self.municipality_combo.count() == 0:
            # Если справочник пуст — оставляем возможность ручного ввода через "по умолчанию"
            self.municipality_combo.addItem("Не задано", "")

    def set_project(self, project: Project):
        """
        Заполнить диалог данными существующего проекта (режим редактирования).
        """
        if not project:
            return

        self.setWindowTitle("Редактирование проекта")

        # Название
        self.name_edit.setText(project.name or "")

        # Год: по year_id (новая архитектура)
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

        # МО: ищем по municipality_id
        if project.municipality_id and self.municipality_combo.count() > 0 and self._municip_cache:
            try:
                municip_ref = next((m for m in self._municip_cache if m.id == project.municipality_id), None)
                if municip_ref:
                    idx = self.municipality_combo.findData(municip_ref.name)
                    if idx >= 0:
                        self.municipality_combo.setCurrentIndex(idx)
            except Exception:
                pass

    def get_project_data(self):
        """Получение данных проекта"""
        name = self.name_edit.text().strip()

        # Год — из справочника (получаем ID)
        year_id = None
        year_val_int = datetime.now().year
        if self.year_combo.count() > 0:
            year_val = self.year_combo.currentData()
            try:
                year_val_int = int(year_val)
                # Находим year_id из кэша
                year_ref = next((y for y in self._years_cache if y.year == year_val_int), None)
                if year_ref:
                    year_id = year_ref.id
                else:
                    # Если не найден в кэше, создаем/получаем через БД
                    year_ref = self.db_manager.get_or_create_year(year_val_int)
                    year_id = year_ref.id
            except (TypeError, ValueError):
                pass
        
        # Муниципальное образование — из справочника (получаем ID)
        municipality_id = None
        municipality_name = ""
        if self.municipality_combo.count() > 0:
            municipality_name = self.municipality_combo.currentData() or ""
            if municipality_name:
                # Находим municipality_id из кэша
                municip_ref = next((m for m in self._municip_cache if m.name == municipality_name), None)
                if municip_ref:
                    municipality_id = municip_ref.id
                else:
                    # Если не найден в кэше, создаем/получаем через БД
                    municip_ref = self.db_manager.get_or_create_municipality(municipality_name)
                    municipality_id = municip_ref.id

        return {
            "name": name,
            "year_id": year_id,
            "municipality_id": municipality_id,
        }
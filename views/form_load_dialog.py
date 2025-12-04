from PyQt5.QtWidgets import (
    QDialog,
    QVBoxLayout,
    QFormLayout,
    QComboBox,
    QLineEdit,
    QDialogButtonBox,
    QLabel,
)
from models.database import DatabaseManager


class FormLoadDialog(QDialog):
    """
    Диалог выбора типа формы, периода и ревизии при загрузке файла формы.

    Типы форм и периоды берутся из справочников ref_form_types и ref_periods.
    """

    def __init__(self, db_manager: DatabaseManager, parent=None):
        super().__init__(parent)
        self.db_manager = db_manager

        self.setWindowTitle("Параметры формы")
        self.resize(400, 220)

        self._form_types = []
        self._periods_by_form_code = {}

        self._init_ui()
        self._load_form_types()
        self._reload_periods()

    def _init_ui(self):
        layout = QVBoxLayout(self)

        form_layout = QFormLayout()

        # Тип формы
        self.form_type_combo = QComboBox()
        self.form_type_combo.currentIndexChanged.connect(self._reload_periods)
        form_layout.addRow("Тип формы:", self.form_type_combo)

        # Период
        self.period_combo = QComboBox()
        form_layout.addRow("Период:", self.period_combo)

        # Ревизия
        self.revision_edit = QLineEdit()
        self.revision_edit.setText("1.0")
        form_layout.addRow("Ревизия:", self.revision_edit)

        layout.addLayout(form_layout)

        # Подсказка
        layout.addWidget(QLabel("Выберите тип формы, период и ревизию перед загрузкой файла."))

        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(self._on_accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)

    def _load_form_types(self):
        """Заполняем список типов форм из справочника."""
        self.form_type_combo.clear()
        self._form_types = self.db_manager.load_form_types_meta()
        for ft in self._form_types:
            if not ft.is_active:
                continue
            label = f"{ft.code} — {ft.name}" if ft.name else ft.code
            self.form_type_combo.addItem(label, ft.code)

    def _reload_periods(self):
        """Обновляем список периодов в зависимости от выбранного типа формы."""
        self.period_combo.clear()
        if self.form_type_combo.count() == 0:
            return

        form_code = self.form_type_combo.currentData()
        periods = self.db_manager.load_periods(form_type_code=form_code) or self.db_manager.load_periods()

        self._periods_by_form_code[form_code] = periods
        for p in periods:
            if not p.is_active:
                continue
            label = p.name or p.code
            self.period_combo.addItem(label, p.code)

        if self.period_combo.count() == 0:
            # Fallback: добавляем "Год"
            self.period_combo.addItem("Год", "Y")

    def _on_accept(self):
        if self.form_type_combo.count() == 0:
            return
        if not self.revision_edit.text().strip():
            self.revision_edit.setText("1.0")
        self.accept()

    def get_form_params(self) -> dict:
        """
        Возвращает словарь:
        {
          "form_code": "0503317",
          "period_code": "Q1",
          "revision": "1.0",
        }
        """
        form_code = self.form_type_combo.currentData()
        period_code = self.period_combo.currentData() if self.period_combo.count() > 0 else "Y"
        revision = self.revision_edit.text().strip() or "1.0"
        return {
            "form_code": form_code,
            "period_code": period_code,
            "revision": revision,
        }



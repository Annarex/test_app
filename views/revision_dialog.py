from PyQt5.QtWidgets import (
    QDialog,
    QVBoxLayout,
    QHBoxLayout,
    QFormLayout,
    QLineEdit,
    QComboBox,
    QDialogButtonBox,
    QFileDialog,
    QPushButton,
)

from models.base_models import ProjectStatus, FormRevisionRecord
from models.database import DatabaseManager


class RevisionDialog(QDialog):
    """
    Диалог редактирования ревизии формы.
    
    Позволяет редактировать:
    - номер ревизии (revision)
    - статус ревизии (status)
    - путь к файлу (file_path)
    """

    def __init__(self, db_manager: DatabaseManager, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Редактирование ревизии")
        self.setModal(True)
        self.resize(500, 200)
        
        self.db_manager = db_manager
        self.revision_id = None
        self.project_id = None
        
        self.init_ui()

    def init_ui(self):
        """Инициализация интерфейса"""
        layout = QVBoxLayout(self)

        form_layout = QFormLayout()

        # Номер ревизии
        self.revision_edit = QLineEdit()
        self.revision_edit.setPlaceholderText("Например: 1.0, 2.0, 2.1")
        form_layout.addRow("Номер ревизии:", self.revision_edit)

        # Статус ревизии
        self.status_combo = QComboBox()
        for status in ProjectStatus:
            self.status_combo.addItem(status.value, status.value)
        form_layout.addRow("Статус:", self.status_combo)

        # Путь к файлу
        file_layout = QHBoxLayout()
        self.file_path_edit = QLineEdit()
        self.file_path_edit.setPlaceholderText("Путь к файлу формы")
        browse_button = QPushButton("Обзор...")
        browse_button.clicked.connect(self.browse_file)
        file_layout.addWidget(self.file_path_edit)
        file_layout.addWidget(browse_button)
        form_layout.addRow("Файл:", file_layout)

        layout.addLayout(form_layout)

        # Кнопки
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)

    def browse_file(self):
        """Открыть диалог выбора файла"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Выберите файл формы",
            "",
            "Excel files (*.xlsx *.xls);;All files (*.*)"
        )
        if file_path:
            self.file_path_edit.setText(file_path)

    def set_revision(self, revision: FormRevisionRecord, project_id: int):
        """
        Заполнить диалог данными существующей ревизии.
        """
        if not revision:
            return

        self.revision_id = revision.id
        self.project_id = project_id

        # Номер ревизии
        self.revision_edit.setText(revision.revision or "")

        # Статус
        status_value = revision.status.value if isinstance(revision.status, ProjectStatus) else str(revision.status)
        idx = self.status_combo.findData(status_value)
        if idx >= 0:
            self.status_combo.setCurrentIndex(idx)

        # Путь к файлу
        self.file_path_edit.setText(revision.file_path or "")

    def get_revision_data(self):
        """Получение данных ревизии"""
        revision = self.revision_edit.text().strip()
        status = self.status_combo.currentData() or "created"
        file_path = self.file_path_edit.text().strip()

        return {
            "revision": revision,
            "status": status,
            "file_path": file_path,
        }


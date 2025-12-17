"""
Диалог для загрузки и обработки решений о бюджете
"""
from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QFormLayout,
    QPushButton, QLabel, QFileDialog, QComboBox,
    QDialogButtonBox, QMessageBox, QTextEdit, QGroupBox
)
from PyQt5.QtCore import Qt
from pathlib import Path
from logger import logger


class SolutionLoadDialog(QDialog):
    """Диалог для загрузки и обработки решений о бюджете"""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Загрузка решения о бюджете")
        self.setMinimumWidth(600)
        self.setMinimumHeight(400)
        
        self.solution_file_path = None
        self.project_id = None
        self.solution_id = None
        
        self.init_ui()
        self.load_projects()
    
    def init_ui(self):
        """Инициализация UI"""
        layout = QVBoxLayout(self)
        
        # Группа для выбора проекта
        project_group = QGroupBox("Выбор проекта")
        project_layout = QFormLayout()
        
        self.project_combo = QComboBox()
        self.project_combo.currentIndexChanged.connect(self.update_process_button_state)
        project_layout.addRow("Проект:", self.project_combo)
        
        project_group.setLayout(project_layout)
        layout.addWidget(project_group)
        
        # Группа для выбора файла
        file_group = QGroupBox("Выбор файла решения")
        file_layout = QVBoxLayout()
        
        file_path_layout = QHBoxLayout()
        self.file_path_label = QLabel("Файл не выбран")
        self.file_path_label.setWordWrap(True)
        browse_btn = QPushButton("Обзор...")
        browse_btn.clicked.connect(self.browse_file)
        
        file_path_layout.addWidget(self.file_path_label)
        file_path_layout.addWidget(browse_btn)
        file_layout.addLayout(file_path_layout)
        
        file_group.setLayout(file_layout)
        layout.addWidget(file_group)
        
        # Кнопка обработки
        self.process_btn = QPushButton("Загрузить и обработать")
        self.process_btn.clicked.connect(self.process_solution)
        self.process_btn.setEnabled(False)
        layout.addWidget(self.process_btn)
        
        # Группа для результатов
        results_group = QGroupBox("Результаты обработки")
        results_layout = QVBoxLayout()
        
        self.results_text = QTextEdit()
        self.results_text.setReadOnly(True)
        self.results_text.setMaximumHeight(150)
        results_layout.addWidget(self.results_text)
        
        results_group.setLayout(results_layout)
        layout.addWidget(results_group)
        
        # Кнопки диалога
        button_box = QDialogButtonBox(QDialogButtonBox.Close)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)
    
    def load_projects(self):
        """Загрузка списка проектов"""
        try:
            parent = self.parent()
            if hasattr(parent, 'controller') and hasattr(parent.controller, 'db_manager'):
                db_manager = parent.controller.db_manager
                projects = db_manager.load_projects()
                
                self.project_combo.clear()
                for project in projects:
                    self.project_combo.addItem(project.name, project.id)
                
                # Если есть текущий проект, выбираем его
                if hasattr(parent, 'current_project_id') and parent.current_project_id:
                    idx = self.project_combo.findData(parent.current_project_id)
                    if idx >= 0:
                        self.project_combo.setCurrentIndex(idx)
        except Exception as e:
            logger.error(f"Ошибка загрузки проектов: {e}", exc_info=True)
            QMessageBox.warning(self, "Ошибка", f"Не удалось загрузить список проектов:\n{str(e)}")
    
    def browse_file(self):
        """Открыть диалог выбора файла"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Выберите файл решения",
            "",
            "Word Documents (*.docx);;All files (*.*)"
        )
        
        if file_path:
            self.solution_file_path = file_path
            self.file_path_label.setText(file_path)
            self.update_process_button_state()
    
    def update_process_button_state(self):
        """Обновление состояния кнопки обработки"""
        has_project = self.project_combo.count() > 0 and self.project_combo.currentData() is not None
        has_file = self.solution_file_path is not None
        
        self.process_btn.setEnabled(has_project and has_file)
    
    def process_solution(self):
        """Обработка решения"""
        try:
            self.project_id = self.project_combo.currentData()
            if not self.project_id:
                QMessageBox.warning(self, "Ошибка", "Выберите проект")
                return
            
            if not self.solution_file_path:
                QMessageBox.warning(self, "Ошибка", "Выберите файл решения")
                return
            
            # Вызываем метод контроллера через родительское окно
            parent = self.parent()
            if not hasattr(parent, 'controller'):
                QMessageBox.warning(self, "Ошибка", "Контроллер не найден")
                return
            
            # Проверяем наличие solution_controller
            if not hasattr(parent.controller, 'solution_controller'):
                QMessageBox.warning(self, "Ошибка", "Контроллер решений не найден")
                return
            
            solution_controller = parent.controller.solution_controller
            
            # Парсим документ
            self.results_text.clear()
            self.results_text.append("Обработка решения...")
            self.process_btn.setEnabled(False)
            
            result = solution_controller.parse_solution_document(
                file_path=self.solution_file_path,
                project_id=self.project_id
            )
            
            if result:
                # Сохраняем данные в БД
                self.solution_id = solution_controller.save_solution_data(
                    project_id=self.project_id,
                    solution_data=result,
                    file_path=self.solution_file_path
                )
                
                # Формируем отчет о результатах
                report_lines = [
                    "Решение успешно обработано!",
                    "",
                    f"Файл: {Path(self.solution_file_path).name}",
                    f"ID решения в БД: {self.solution_id}",
                    "",
                    "Статистика:",
                    f"  - Доходы (Приложение 1): {len(result.get('приложение1', []))} записей",
                    f"  - Расходы общие (Приложение 2): {len(result.get('приложение2', []))} записей",
                    f"  - Расходы по ГРБС (Приложение 3): {len(result.get('приложение3', []))} записей",
                ]
                
                self.results_text.clear()
                self.results_text.append("\n".join(report_lines))
                
                QMessageBox.information(
                    self,
                    "Успех",
                    f"Решение успешно обработано и сохранено в БД.\n\n"
                    f"Доходы: {len(result.get('приложение1', []))} записей\n"
                    f"Расходы общие: {len(result.get('приложение2', []))} записей\n"
                    f"Расходы по ГРБС: {len(result.get('приложение3', []))} записей"
                )
            else:
                self.results_text.clear()
                self.results_text.append("Ошибка: Не удалось обработать решение")
                QMessageBox.warning(self, "Ошибка", "Не удалось обработать решение")
            
        except Exception as e:
            logger.error(f"Ошибка обработки решения: {e}", exc_info=True)
            self.results_text.clear()
            self.results_text.append(f"Ошибка: {str(e)}")
            QMessageBox.critical(self, "Ошибка", f"Ошибка обработки решения:\n{str(e)}")
        finally:
            self.process_btn.setEnabled(True)
    
    def get_solution_id(self):
        """Получить ID сохраненного решения"""
        return self.solution_id

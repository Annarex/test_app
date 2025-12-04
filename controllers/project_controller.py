from PyQt5.QtCore import QObject, pyqtSignal
from typing import List, Optional, Dict, Any
from datetime import datetime

from models.database import DatabaseManager
from models.base_models import Project, ProjectStatus, FormType

class ProjectController(QObject):
    """Контроллер управления проектами"""
    
    # Сигналы
    projects_updated = pyqtSignal(list)
    project_loaded = pyqtSignal(Project)
    calculation_completed = pyqtSignal(dict)
    export_completed = pyqtSignal(str)
    error_occurred = pyqtSignal(str)
    
    def __init__(self, db_manager: DatabaseManager):
        super().__init__()
        self.db_manager = db_manager
        self.current_project = None
    
    def load_projects(self) -> List[Project]:
        """Загрузка всех проектов"""
        return self.db_manager.load_projects()
    
    def create_project(self, project_data: Dict[str, Any]) -> Optional[Project]:
        """Создание нового проекта"""
        try:
            project = Project()
            project.name = project_data.get('name', '')
            project.year_id = project_data.get('year_id')
            project.municipality_id = project_data.get('municipality_id')
            
            # Сохраняем в БД
            project_id = self.db_manager.save_project(project)
            project.id = project_id
            
            # Обновляем список проектов
            projects = self.load_projects()
            self.projects_updated.emit(projects)
            
            return project
            
        except Exception as e:
            self.error_occurred.emit(f"Ошибка создания проекта: {str(e)}")
            return None
    
    def load_project(self, project_id: int):
        """Загрузка проекта по ID"""
        try:
            projects = self.db_manager.load_projects()
            project = next((p for p in projects if p.id == project_id), None)
            
            if project:
                self.current_project = project
                self.project_loaded.emit(project)
            else:
                self.error_occurred.emit("Проект не найден")
                
        except Exception as e:
            self.error_occurred.emit(f"Ошибка загрузки проекта: {str(e)}")
    
    def update_project(self, project_data: Dict[str, Any]) -> bool:
        """Обновление проекта"""
        if not self.current_project:
            self.error_occurred.emit("Проект не выбран")
            return False
        
        try:
            self.current_project.name = project_data.get('name', self.current_project.name)
            if 'year_id' in project_data:
                self.current_project.year_id = project_data.get('year_id')
            if 'municipality_id' in project_data:
                self.current_project.municipality_id = project_data.get('municipality_id')
            
            self.db_manager.save_project(self.current_project)
            
            # Обновляем список проектов
            projects = self.load_projects()
            self.projects_updated.emit(projects)
            
            return True
            
        except Exception as e:
            self.error_occurred.emit(f"Ошибка обновления проекта: {str(e)}")
            return False
    
    def delete_project(self, project_id: int):
        """Удаление проекта"""
        try:
            self.db_manager.delete_project(project_id)
            
            # Обновляем список проектов
            projects = self.load_projects()
            self.projects_updated.emit(projects)
            
        except Exception as e:
            self.error_occurred.emit(f"Ошибка удаления проекта: {str(e)}")
from PyQt5.QtCore import QObject, pyqtSignal
from typing import List, Optional, Dict, Any
import os
from pathlib import Path
import pandas as pd
import sqlite3
import json

from models.database import DatabaseManager
from logger import logger
from models.base_models import (
    Project,
    Reference,
    ProjectStatus,
    FormType,
    FormTypeMeta,
    PeriodRef,
    ProjectForm,
    FormRevisionRecord,
)
from models.form_0503317 import Form0503317
from controllers.project_controller import ProjectController
from controllers.revision_controller import RevisionController
from controllers.form_controller import FormController
from controllers.reference_controller import ReferenceController
from controllers.calculation_controller import CalculationController
from controllers.tree_controller import TreeController

class MainController(QObject):
    """Главный контроллер приложения"""
    
    # Сигналы
    projects_updated = pyqtSignal(list)
    references_updated = pyqtSignal(list)
    project_loaded = pyqtSignal(Project)
    calculation_completed = pyqtSignal(dict)
    export_completed = pyqtSignal(str)
    error_occurred = pyqtSignal(str)
    
    def __init__(self):
        super().__init__()
        self.db_manager = DatabaseManager()
        
        # Основные контроллеры
        self.project_controller = ProjectController(self.db_manager)
        
        # Специализированные контроллеры
        self.revision_controller = RevisionController(self.db_manager)
        self.form_controller = FormController(self.db_manager)
        self.reference_controller = ReferenceController(self.db_manager)
        self.calculation_controller = CalculationController(self.db_manager)
        self.tree_controller = TreeController(self.db_manager, self.project_controller)
        
        # Текущий проект (синхронизируем с подконтроллерами)
        self.current_project = None
        self.current_form = None
        self.current_revision_id = None
        
        # Синхронизируем состояние между контроллерами
        self._sync_controller_state()
        
        # Подключаем сигналы
        self.project_controller.projects_updated.connect(self.projects_updated)
        self.project_controller.project_loaded.connect(self._on_project_loaded)
        self.project_controller.calculation_completed.connect(self.calculation_completed)
        self.project_controller.export_completed.connect(self.export_completed)
        self.project_controller.error_occurred.connect(self.error_occurred)
        
        # Сигналы специализированных контроллеров
        self.reference_controller.references_updated.connect(self.references_updated)
        self.reference_controller.error_occurred.connect(self.error_occurred)
        self.calculation_controller.calculation_completed.connect(self.calculation_completed)
        self.calculation_controller.export_completed.connect(self.export_completed)
        self.calculation_controller.error_occurred.connect(self.error_occurred)
    
    def _sync_controller_state(self):
        """Синхронизация состояния между контроллерами"""
        # Синхронизируем текущее состояние
        self.revision_controller.current_project = self.current_project
        self.revision_controller.current_form = self.current_form
        self.revision_controller.current_revision_id = self.current_revision_id
        
        self.form_controller.current_project = self.current_project
        self.form_controller.current_form = self.current_form
        self.form_controller.current_revision_id = self.current_revision_id
        self.form_controller.pending_form_type_code = self.revision_controller.pending_form_type_code
        self.form_controller.pending_revision = self.revision_controller.pending_revision
        
        self.calculation_controller.current_project = self.current_project
        self.calculation_controller.current_form = self.current_form
        self.calculation_controller.current_revision_id = self.current_revision_id

    # ------------------------------------------------------------------
    # Выбор формы/периода/ревизии пользователем (до загрузки файла)
    # ------------------------------------------------------------------

    def set_current_form_params(self, form_code: str, revision: str, period_code: Optional[str] = None) -> None:
        """Сохранить выбранные пользователем параметры формы для текущего проекта"""
        self.revision_controller.set_current_form_params(form_code, revision, period_code)
        self._sync_controller_state()
        # Переинициализируем форму под выбранный тип
        if self.current_project:
            self._initialize_form_for_project()

    def set_form_params_from_revision(self, revision_id: int):
        """Подтянуть параметры формы из существующей ревизии"""
        self.revision_controller.set_form_params_from_revision(revision_id)
        self._sync_controller_state()

    def get_pending_form_params(self) -> Dict[str, Optional[str]]:
        """Возвращает сохранённые параметры формы для префилла диалога."""
        return self.revision_controller.get_pending_form_params()
    
    def load_initial_data(self):
        """Загрузка начальных данных"""
        projects = self.project_controller.load_projects()
        references = self.reference_controller.load_references()
        self._sync_controller_state()
        
        # Сигнал по‑прежнему передаём список Project, но левая панель
        # теперь строится по новой архитектуре (год → проект → форма → период → ревизии)
        self.projects_updated.emit(projects)
        self.references_updated.emit(references)
    
    def refresh_references(self):
        """Обновление справочников (публичный метод)"""
        references = self.reference_controller.refresh_references()
        self._sync_controller_state()
        return references
    
    def create_project(self, project_data: Dict[str, Any]) -> Optional[Project]:
        """Создание нового проекта"""
        project = self.project_controller.create_project(project_data)
        if project:
            self.current_project = project
            self._sync_controller_state()
            self._initialize_form_for_project()
            # Ревизия создается только при загрузке формы, не при создании проекта
        return project
    
    def update_project(self, project_data: Dict[str, Any]) -> bool:
        """Обновление существующего проекта"""
        success = self.project_controller.update_project(project_data)
        if success:
            # Синхронизируем текущий проект
            self.current_project = self.project_controller.current_project
            self._sync_controller_state()
        return success
    
    def delete_form_revision(self, revision_id: int) -> None:
        """Удаление одной ревизии формы (новая архитектура)"""
        self.revision_controller.delete_form_revision(revision_id)
        self._sync_controller_state()
        # Обновляем список проектов после удаления
        projects = self.project_controller.load_projects()
        self.projects_updated.emit(projects)
    
    def update_form_revision(self, revision_id: int, revision_data: Dict[str, Any]) -> bool:
        """Обновление ревизии формы"""
        success = self.revision_controller.update_form_revision(revision_id, revision_data)
        self._sync_controller_state()
        if success:
            # Обновляем список проектов после обновления
            projects = self.project_controller.load_projects()
            self.projects_updated.emit(projects)
        return success
    
    def load_project(self, project_id: int):
        """Загрузка проекта"""
        self.project_controller.load_project(project_id)
    
    def load_revision(self, revision_id: int, project_id: int):
        """Загрузка конкретной ревизии проекта (делегирует к revision_controller)"""
        # Используем расширенный метод revision_controller для полной загрузки
        project = self.revision_controller.load_revision_with_form_initialization(
            revision_id,
            project_id,
            self.project_controller,
            self.form_controller
        )
        
        if project:
            # Обновляем состояние
            self.current_project = project
            self.current_revision_id = revision_id
            self.current_form = self.form_controller.current_form
            self._sync_controller_state()
            
            # Эмитируем сигнал загрузки проекта
            self.project_loaded.emit(project)
    
    def _on_project_loaded(self, project: Project):
        """Обработка загруженного проекта"""
        self.current_project = project
        self.current_revision_id = None  # Сбрасываем при загрузке проекта без указания ревизии
        self._sync_controller_state()
        self._initialize_form_for_project()

        # ОПТИМИЗАЦИЯ: Убран автоматический пересчет уровней при загрузке.
        # Теперь данные отображаются как есть в БД (быстрая загрузка).
        # Пересчет уровней выполняется только по кнопке "Пересчитать" в UI.
        
        # Инициализируем форму данными проекта, чтобы экспорт/проверка
        # работали сразу после загрузки без повторного парсинга файла.
        if self.current_form and self.current_project.data:
            self.current_form.load_saved_data(self.current_project.data)

        self._sync_controller_state()
        self.project_loaded.emit(project)
    
    def get_project_info(self, project: Project) -> Dict[str, Any]:
        """
        Получить информацию о проекте и текущей ревизии для отображения в UI
        
        Returns:
            Словарь с ключами: form_text, revision_text, status_text, 
            period_text, municipality_text, excel_path
        """
        rev_id = self.current_revision_id
        form_text = "—"
        revision_text = "—"
        status_text = "—"
        period_text = "—"
        municipality_text = "—"
        excel_path = None

        if rev_id:
            try:
                revision = self.db_manager.get_form_revision_by_id(rev_id)
                if revision:
                    # Ревизия и статус
                    revision_text = revision.revision or "—"
                    from models.base_models import ProjectStatus
                    if isinstance(revision.status, ProjectStatus):
                        status_text = revision.status.value
                    else:
                        status_text = str(revision.status or "—")

                    # Путь к файлу для Excel‑просмотра
                    excel_path = revision.file_path or None

                    # Находим связанную форму и её тип / период
                    project_forms = self.db_manager.load_project_forms(project.id)
                    pf = next((p for p in project_forms if p.id == revision.project_form_id), None)
                    if pf:
                        # Тип формы
                        form_types_meta = {ft.id: ft for ft in self.db_manager.load_form_types_meta()}
                        ft_meta = form_types_meta.get(pf.form_type_id)
                        if ft_meta:
                            # Показываем и код, и читаемое имя, если есть
                            if ft_meta.name:
                                form_text = f"{ft_meta.name} ({ft_meta.code})"
                            else:
                                form_text = ft_meta.code
                        # Период
                        if pf.period_id:
                            periods = self.db_manager.load_periods()
                            period_ref = next((p for p in periods if p.id == pf.period_id), None)
                            if period_ref:
                                period_text = period_ref.name or period_ref.code or period_text
                else:
                    # Если ревизия по ID не найдена — используем заглушки
                    revision_text = "—"
                    status_text = "—"
                    form_text = "—"
            except Exception as e:
                logger.error(f"Ошибка получения информации о ревизии: {e}", exc_info=True)
                # Fallback - используем заглушки
                revision_text = "—"
                status_text = "—"
                form_text = "—"
        else:
            # Проект без выбранной ревизии
            form_text = "—"
            revision_text = "—"
            status_text = "—"

        # МО — по коду ОКТМО и дате создания проекта
        try:
            if project.oktmo_code:
                filter_date = project.created_at.strftime("%Y-%m-%d") if project.created_at else None
                name = self.db_manager.get_oktmo_name_by_code(project.oktmo_code, filter_date)
                if name:
                    municipality_text = name
        except Exception as e:
            logger.warning(f"Ошибка получения МО для проекта {project.id}: {e}", exc_info=True)

        return {
            'form_text': form_text,
            'revision_text': revision_text,
            'status_text': status_text,
            'period_text': period_text,
            'municipality_text': municipality_text,
            'excel_path': excel_path
        }
    
    def _initialize_form_for_project(self, form_meta=None):
        """Инициализация формы для проекта"""
        self.form_controller.initialize_form_for_project(form_meta)
        self.current_form = self.form_controller.current_form
        self._sync_controller_state()
    
    def _copy_form_file_to_project(self, source_file_path: str, project_id: int) -> str:
        """Копирование файла формы в папку проекта с префиксом даты/времени"""
        return self.form_controller.copy_form_file_to_project(source_file_path, project_id)
    
    def load_form_file(self, file_path: str) -> bool:
        """Загрузка файла формы (делегирует к form_controller)"""
        if not self.current_project:
            self.error_occurred.emit("Проект не выбран")
            return False
        
        # Синхронизируем состояние перед загрузкой
        self._sync_controller_state()
        
        # Используем form_controller для полного цикла загрузки
        success = self.form_controller.load_and_register_form_file(
            file_path,
            self.revision_controller,
            self.db_manager
        )
        
        if not success:
            return False
        
        # Обновляем состояние после загрузки
        self.current_form = self.form_controller.current_form
        self.current_revision_id = self.revision_controller.current_revision_id
        
        # Обновляем данные проекта из формы
        if self.current_form:
            form_data = {
                'meta_info': self.current_form.meta_info,
                'income_data': self.current_form.income_data,
                'outcome_data': self.current_form.outcome_data,
                'source_financing_deficit_data': self.current_form.source_financing_deficit_data,
                'consolidated_calc_data': self.current_form.consolidated_calc_data
            }
            self.current_project.data = form_data
        
        # После успешного создания/обновления ревизии обновляем дерево проектов
        try:
            projects = self.project_controller.load_projects()
            self.projects_updated.emit(projects)
        except Exception as e:
            logger.error(f"Ошибка обновления списка проектов после сохранения ревизии: {e}", exc_info=True)

        logger.info(f"Форма успешно загружена. Данные: {len(self.current_form.income_data) if self.current_form else 0} доходов, "
              f"{len(self.current_form.outcome_data) if self.current_form else 0} расходов")
        
        self._sync_controller_state()
        return True
    
    def calculate_sums(self):
        """Расчет агрегированных сумм"""
        self.calculation_controller.calculate_sums()
        self._sync_controller_state()
    
    def export_validation(self, output_path: str) -> bool:
        """Экспорт формы с проверкой"""
        result = self.calculation_controller.export_validation(output_path)
        self._sync_controller_state()
        return result is not None

    # ------------------------------------------------------------------
    # Вспомогательная логика для новой архитектуры форм/ревизий
    # ------------------------------------------------------------------

    def _register_form_revision(self, project: Project, status: ProjectStatus, file_path: str, 
                                form_type_code: Optional[str] = None, period_code: Optional[str] = None,
                                revision: Optional[str] = None):
        """Зарегистрировать или обновить ревизию формы для указанного проекта"""
        # Синхронизируем состояние перед вызовом
        self._sync_controller_state()
        revision_record = self.revision_controller.register_form_revision(
            project, status, file_path, form_type_code, period_code, revision
        )
        if revision_record:
            self.current_revision_id = revision_record.id
            self._sync_controller_state()
        return revision_record

    # ------------------------------------------------------------------
    # Построение дерева проектов (Год → Проект → Форма → Период → Ревизии)
    # ------------------------------------------------------------------

    def build_project_tree(self) -> list:
        """Построение дерева проектов"""
        return self.tree_controller.build_project_tree()
    
    def load_reference_file(self, file_path: str, ref_type: str, name: str) -> bool:
        """Загрузка файла справочника"""
        success = self.reference_controller.load_reference_file(file_path, ref_type, name)
        self._sync_controller_state()
        return success
    
    def delete_project(self, project_id: int):
        """Удаление проекта"""
        self.project_controller.delete_project(project_id)
        
        # Если удален текущий проект, сбрасываем его
        if self.current_project and self.current_project.id == project_id:
            self.current_project = None
            self.current_form = None
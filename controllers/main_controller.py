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
from controllers.document_controller import DocumentController
from controllers.solution_controller import SolutionController
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
        self.document_controller = DocumentController(self.db_manager)
        self.solution_controller = SolutionController(self.db_manager)
        
        # Новые специализированные контроллеры
        self.revision_controller = RevisionController(self.db_manager)
        self.form_controller = FormController(self.db_manager)
        self.reference_controller = ReferenceController(self.db_manager)
        self.calculation_controller = CalculationController(self.db_manager)
        self.tree_controller = TreeController(self.db_manager, self.project_controller)
        
        # Текущий проект (синхронизируем с подконтроллерами)
        self.current_project = None
        self.current_form = None
        self.current_revision_id = None
        
        # Справочники (используем из reference_controller)
        self.references = self.reference_controller.references
        
        # Синхронизируем состояние между контроллерами
        self._sync_controller_state()
        
        # Подключаем сигналы
        self.project_controller.projects_updated.connect(self.projects_updated)
        self.project_controller.project_loaded.connect(self._on_project_loaded)
        self.project_controller.calculation_completed.connect(self.calculation_completed)
        self.project_controller.export_completed.connect(self.export_completed)
        self.project_controller.error_occurred.connect(self.error_occurred)
        
        # Сигналы контроллеров документов
        self.document_controller.document_generated.connect(self._on_document_generated)
        self.document_controller.error_occurred.connect(self.error_occurred)
        self.solution_controller.solution_parsed.connect(self._on_solution_parsed)
        self.solution_controller.error_occurred.connect(self.error_occurred)
        
        # Сигналы новых контроллеров
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
        self.revision_controller.references = self.references
        
        self.form_controller.current_project = self.current_project
        self.form_controller.current_form = self.current_form
        self.form_controller.current_revision_id = self.current_revision_id
        self.form_controller.references = self.references
        self.form_controller.pending_form_type_code = self.revision_controller.pending_form_type_code
        self.form_controller.pending_revision = self.revision_controller.pending_revision
        
        self.calculation_controller.current_project = self.current_project
        self.calculation_controller.current_form = self.current_form
        self.calculation_controller.current_revision_id = self.current_revision_id
        self.calculation_controller.references = self.references

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
        """Загрузка конкретной ревизии проекта"""
        try:
            # Сначала загружаем информацию о ревизии
            revision_record = self.db_manager.get_form_revision_by_id(revision_id)
            if not revision_record:
                self.error_occurred.emit("Ревизия не найдена")
                return
            
            # Загружаем проект
            self.project_controller.load_project(project_id)
            
            if not self.current_project:
                self.error_occurred.emit("Проект не найден")
                return
            
            # Определяем тип формы из project_form, связанного с ревизией
            # Оптимизация: создаем словарь для быстрого поиска вместо next()
            project_forms = self.db_manager.load_project_forms(project_id)
            project_forms_by_id = {pf.id: pf for pf in project_forms}
            project_form = project_forms_by_id.get(revision_record.project_form_id)
            
            if not project_form:
                self.error_occurred.emit(f"ProjectForm не найден для ревизии {revision_id}")
                return
            
            # Получаем метаданные типа формы
            form_types_meta = {ft.id: ft for ft in self.db_manager.load_form_types_meta()}
            form_meta = form_types_meta.get(project_form.form_type_id)
            
            if not form_meta:
                self.error_occurred.emit(f"Тип формы не найден для form_type_id={project_form.form_type_id}")
                return
            
            # Сохраняем ID текущей ревизии ДО инициализации формы
            # Это нужно, чтобы _initialize_form_for_project мог определить тип формы из ревизии
            self.current_revision_id = revision_id
            # Инициализируем форму заранее, чтобы при необходимости можно было спарсить файл
            self._initialize_form_for_project(form_meta=form_meta)
            
            # Загружаем данные ревизии
            revision_data = self.db_manager.load_revision_data(project_id, revision_id)
            need_sections = ['доходы_data', 'расходы_data', 'источники_финансирования_data', 'консолидируемые_расчеты_data']
            has_sections = revision_data and any(revision_data.get(k) for k in need_sections)

            if not has_sections:
                self.error_occurred.emit("Данные ревизии не найдены")
                return
            
            # Обновляем данные проекта данными ревизии (для отображения в UI)
            self.current_project.data = revision_data
            
            # Инициализируем форму с правильным типом (теперь current_revision_id уже установлен)
            # Передаём form_meta как fallback на случай, если определение из ревизии не сработает
            self._initialize_form_for_project(form_meta=form_meta)
            
            # Загружаем данные в форму
            if not self.current_form:
                self.error_occurred.emit(f"Форма типа '{form_meta.code}' не поддерживается")
                return
            
            self.current_form.load_saved_data(revision_data)
            # Пересчитываем дефицит/профицит при загрузке ревизии, чтобы
            # calculated_deficit_proficit всегда был доступен в project.data
            try:
                if hasattr(self.current_form, "_calculate_deficit_proficit"):
                    self.current_form._calculate_deficit_proficit()
                    if getattr(self.current_form, "calculated_deficit_proficit", None):
                        revision_data["calculated_deficit_proficit"] = (
                            self.current_form.calculated_deficit_proficit
                        )
                        self.current_project.data = revision_data
            except Exception as e:
                logger.error(
                    f"Ошибка пересчета дефицита/профицита при загрузке ревизии: {e}",
                    exc_info=True,
                )
            
            # Пересчитываем уровни и значения на основе справочников, если файл есть
            # Файл нужен только для покраски по уровням и отображения пересчитанных значений
            if revision_record.file_path and os.path.exists(revision_record.file_path):
                try:
                    reference_data_доходы = self.references.get('доходы')
                    reference_data_источники = self.references.get('источники')
                    
                    # Пересчитываем уровни и значения на основе справочников
                    if isinstance(self.current_form, Form0503317):
                        updated_data = self.current_form.recalculate_levels_with_references(
                            revision_data,
                            reference_data_доходы,
                            reference_data_источники
                        )
                        if updated_data:
                            # Обновляем данные проекта пересчитанными значениями
                            self.current_project.data = updated_data
                            # Перезагружаем данные в форму с пересчитанными значениями
                            self.current_form.load_saved_data(updated_data)
                            logger.info("Уровни и значения пересчитаны на основе справочников")
                except Exception as e:
                    logger.error(f"Ошибка пересчета уровней и значений: {e}", exc_info=True)
                    # Не блокируем загрузку ревизии из-за ошибки пересчета
            
            # current_revision_id уже установлен выше, перед инициализацией формы
            
            # Эмитируем сигнал загрузки проекта
            self.project_loaded.emit(self.current_project)
                
        except Exception as e:
            self.error_occurred.emit(f"Ошибка загрузки ревизии: {str(e)}")
            logger.error(f"Ошибка загрузки ревизии: {e}", exc_info=True)
    
    def _on_project_loaded(self, project: Project):
        """Обработка загруженного проекта"""
        self.current_project = project
        self.current_revision_id = None  # Сбрасываем при загрузке проекта без указания ревизии
        self._sync_controller_state()
        self._initialize_form_for_project()

        # При загрузке существующего проекта пересчитываем уровни строк
        # на основе актуальных справочников, если это поддерживаемая форма
        # НЕ пересчитываем, если загружена конкретная ревизия (чтобы не перезаписывать данные)
        if (
            self.current_form
            and isinstance(self.current_form, Form0503317)
            and self.current_project.data
            and self.current_revision_id is None  # Только если не загружена конкретная ревизия
        ):
            try:
                reference_data_доходы = self.references.get('доходы')
                reference_data_источники = self.references.get('источники')

                # Если справочники отсутствуют, явно предупреждаем пользователя
                missing_refs = []
                if reference_data_доходы is None:
                    missing_refs.append("доходов")
                if reference_data_источники is None:
                    missing_refs.append("источников финансирования")
                if missing_refs:
                    msg = (
                        "При загрузке проекта не найдены справочники: "
                        + ", ".join(missing_refs)
                        + ". Уровни строк для соответствующих разделов могут быть некорректны (0)."
                    )
                    self.error_occurred.emit(msg)

                # Проверяем, нужно ли пересчитывать уровни
                # Пересчитываем только если справочники доступны
                if reference_data_доходы is not None or reference_data_источники is not None:
                    updated_data = self.current_form.recalculate_levels_with_references(
                        self.current_project.data,
                        reference_data_доходы,
                        reference_data_источники
                    )
                    if updated_data:
                        self.current_project.data = updated_data
                        self.db_manager.save_project(self.current_project)
                        logger.info("Уровни строк пересчитаны на основе справочников")

                # Инициализируем форму данными проекта, чтобы экспорт/проверка
                # работали сразу после загрузки без повторного парсинга файла.
                self.current_form.load_saved_data(self.current_project.data)
            except Exception as e:
                error_msg = f"Ошибка пересчета уровней при загрузке проекта: {e}"
                logger.error(error_msg, exc_info=True)
                # Не блокируем загрузку проекта из-за ошибки пересчета

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
                    # Если ревизия по ID не найдена — fallback на старые поля проекта
                    revision_text = project.revision or "—"
                    status_text = getattr(project.status, "value", str(project.status)) if project.status else "—"
                    form_text = getattr(project.form_type, "value", str(project.form_type)) if project.form_type else "—"
            except Exception as e:
                logger.error(f"Ошибка получения информации о ревизии: {e}", exc_info=True)
                # Fallback на старые поля проекта
                revision_text = project.revision or "—"
                status_text = getattr(project.status, "value", str(project.status)) if project.status else "—"
                form_text = getattr(project.form_type, "value", str(project.form_type)) if project.form_type else "—"
        else:
            # Проект без выбранной ревизии (старые проекты или только что созданные)
            form_text = getattr(project.form_type, "value", str(project.form_type)) if project.form_type else "—"
            revision_text = project.revision or "—"
            status_text = getattr(project.status, "value", str(project.status)) if project.status else "—"

        # МО — берём из справочника по municipality_id проекта
        try:
            if hasattr(project, "municipality_id") and project.municipality_id:
                municip_list = self.db_manager.load_municipalities()
                municip_ref = next((m for m in municip_list if m.id == project.municipality_id), None)
                if municip_ref:
                    municipality_text = municip_ref.name or municipality_text
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
        """Загрузка файла формы"""
        if not self.current_project:
            self.error_occurred.emit("Проект не выбран")
            return False
        
        # Синхронизируем состояние перед загрузкой
        self._sync_controller_state()
        
        # Используем form_controller для загрузки и парсинга
        result = self.form_controller.load_form_file(file_path)
        if not result:
            return False
        
        form_data = result['form_data']
        copied_file_path = result['file_path']
        form_type_code = result['form_type_code']
        period_code = result['period_code']
        
        # Обновляем состояние
        self.current_form = self.form_controller.current_form
        self.current_project.data = form_data
        
        # Определяем период: в приоритете период, выбранный пользователем,
        # а затем (если есть) период из метаданных формы.
        if not period_code:
            period_code = (self.revision_controller.pending_period_code or "").strip() or None
            if form_data.get('meta_info'):
                period_str = form_data.get('meta_info', {}).get('period')
                if period_str:
                    period_code = str(period_str).strip()

        # Регистрируем/обновляем ревизию формы в новой архитектуре
        revision_record = None
        try:
            revision_record = self._register_form_revision(
                project=self.current_project,
                status=ProjectStatus.PARSED, 
                file_path=copied_file_path,
                form_type_code=form_type_code,
                period_code=period_code,
                # Используем выбранный пользователем номер ревизии (или "1.0" по умолчанию)
                revision=self.revision_controller.pending_revision or "1.0",
            )
        except Exception as e:
            # Не блокируем работу, если новая архитектура ревизий дала сбой
            logger.error(f"Ошибка регистрации ревизии формы: {e}", exc_info=True)
        
        # Сохраняем данные ревизии отдельно, если ревизия создана
        if revision_record and revision_record.id:
            self.current_revision_id = revision_record.id
            try:
                # Формируем полные данные ревизии, включая метаданные
                revision_data = {
                    'meta_info': form_data.get('meta_info', {}),
                    'доходы_data': form_data.get('доходы_data', []),
                    'расходы_data': form_data.get('расходы_data', []),
                    'источники_финансирования_data': form_data.get('источники_финансирования_data', []),
                    'консолидируемые_расчеты_data': form_data.get('консолидируемые_расчеты_data', [])
                }
                self.db_manager.save_revision_data(
                    self.current_project.id,
                    revision_record.id,
                    revision_data
                )
            except Exception as e:
                logger.error(f"Ошибка сохранения данных ревизии: {e}", exc_info=True)

            # После успешного создания/обновления ревизии обновляем дерево проектов
            try:
                projects = self.project_controller.load_projects()
                self.projects_updated.emit(projects)
            except Exception as e:
                logger.error(f"Ошибка обновления списка проектов после сохранения ревизии: {e}", exc_info=True)

        logger.info(f"Форма успешно загружена. Данные: {len(form_data.get('доходы_data', []))} доходов, "
              f"{len(form_data.get('расходы_data', []))} расходов")
        
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
    
    # ------------------------------------------------------------------
    # Работа с документами (заключения, письма, решения)
    # ------------------------------------------------------------------
    
    def generate_conclusion(
        self,
        protocol_date,
        protocol_number: str,
        letter_date=None,
        letter_number: str = None,
        admin_date=None,
        admin_number: str = None,
        output_path: str = None
    ) -> Optional[str]:
        """
        Формирование заключения на основе данных формы
        
        Args:
            protocol_date: Дата протокола
            protocol_number: Номер протокола
            letter_date: Дата письма (опционально)
            letter_number: Номер письма (опционально)
            admin_date: Дата постановления администрации (опционально)
            admin_number: Номер постановления администрации (опционально)
            output_path: Путь для сохранения файла (опционально)
        
        Returns:
            Путь к сгенерированному файлу или None при ошибке
        """
        if not self.current_project or not self.current_revision_id:
            self.error_occurred.emit("Не выбран проект или ревизия")
            return None
        
        return self.document_controller.generate_conclusion(
            project_id=self.current_project.id,
            revision_id=self.current_revision_id,
            protocol_date=protocol_date,
            protocol_number=protocol_number,
            letter_date=letter_date,
            letter_number=letter_number,
            admin_date=admin_date,
            admin_number=admin_number,
            output_path=output_path
        )
    
    def generate_letters(
        self,
        protocol_date,
        protocol_number: str,
        output_dir: str = None
    ) -> Dict[str, Optional[str]]:
        """
        Формирование писем администрации и совета
        
        Args:
            protocol_date: Дата протокола
            protocol_number: Номер протокола
            output_dir: Директория для сохранения файлов (опционально)
        
        Returns:
            Словарь с путями к файлам: {'admin': путь, 'council': путь}
        """
        if not self.current_project or not self.current_revision_id:
            self.error_occurred.emit("Не выбран проект или ревизия")
            return {'admin': None, 'council': None}
        
        return self.document_controller.generate_letters(
            project_id=self.current_project.id,
            revision_id=self.current_revision_id,
            protocol_date=protocol_date,
            protocol_number=protocol_number,
            output_dir=output_dir
        )
    
    def parse_solution_document(self, file_path: str) -> Optional[Dict[str, Any]]:
        """
        Парсинг Word документа с решением о бюджете
        
        Args:
            file_path: Путь к Word документу
        
        Returns:
            Словарь с распарсенными данными или None при ошибке
        """
        if not self.current_project:
            self.error_occurred.emit("Не выбран проект")
            return None
        
        return self.solution_controller.parse_solution_document(
            file_path=file_path,
            project_id=self.current_project.id
        )
    
    def _on_document_generated(self, file_path: str):
        """Обработчик сигнала о сгенерированном документе"""
        logger.info(f"Документ сгенерирован: {file_path}")
        # Можно добавить уведомление пользователя
    
    def _on_solution_parsed(self, result: Dict[str, Any]):
        """Обработчик сигнала о распарсенном решении"""
        logger.info(f"Решение распарсено: {len(result.get('приложение1', []))} доходов, "
                   f"{len(result.get('приложение2', []))} расходов")
        # Можно добавить сохранение данных в БД или отображение в UI
    
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
        self.project_controller = ProjectController(self.db_manager)
        
        # Текущий проект
        self.current_project = None
        self.current_form = None
        self.current_revision_id = None  # ID текущей загруженной ревизии
        # Параметры формы, выбранные пользователем до создания первой ревизии
        self.pending_form_type_code: Optional[str] = None
        self.pending_revision: str = "1.0"
        self.pending_period_code: Optional[str] = None
        
        # Справочники (храним как DataFrame)
        self.references = {}
        
        # Подключаем сигналы
        self.project_controller.projects_updated.connect(self.projects_updated)
        self.project_controller.project_loaded.connect(self._on_project_loaded)
        self.project_controller.calculation_completed.connect(self.calculation_completed)
        self.project_controller.export_completed.connect(self.export_completed)
        self.project_controller.error_occurred.connect(self.error_occurred)

    # ------------------------------------------------------------------
    # Выбор формы/периода/ревизии пользователем (до загрузки файла)
    # ------------------------------------------------------------------

    def set_current_form_params(self, form_code: str, revision: str, period_code: Optional[str] = None) -> None:
        """
        Сохранить выбранные пользователем параметры формы для текущего проекта
        до загрузки/создания первой ревизии.

        Эти параметры используются для:
        - инициализации Form0503317 и других моделей;
        - регистрации первой ревизии (form_revisions).
        """
        self.pending_form_type_code = (form_code or "").strip() if form_code else None
        self.pending_revision = (revision or "").strip() or "1.0"
        self.pending_period_code = (period_code or "").strip() if period_code else None

        # Пока ревизия ещё не создана
        self.current_revision_id = None

        # Переинициализируем форму под выбранный тип
        if self.current_project:
            self._initialize_form_for_project()

    def set_form_params_from_revision(self, revision_id: int):
        """
        Подтянуть параметры формы (тип, период, номер ревизии) из существующей ревизии.
        Используется как префилл диалога загрузки новой формы.
        """
        try:
            rev = self.db_manager.get_form_revision_by_id(revision_id)
            if not rev:
                return
            pf = self.db_manager.get_project_form_by_id(rev.project_form_id)
            form_type_code = None
            period_code = None
            if pf:
                ft = self.db_manager.get_form_type_meta_by_id(pf.form_type_id)
                if ft:
                    form_type_code = ft.code
                period = self.db_manager.get_period_by_id(pf.period_id) if pf.period_id else None
                if period:
                    period_code = period.code

            self.pending_form_type_code = form_type_code
            self.pending_period_code = period_code
            self.pending_revision = rev.revision or "1.0"
        except Exception as e:
            logger.warning(f"Не удалось подтянуть параметры из ревизии {revision_id}: {e}")

    def get_pending_form_params(self) -> Dict[str, Optional[str]]:
        """Возвращает сохранённые параметры формы для префилла диалога."""
        return {
            "form_code": self.pending_form_type_code,
            "period_code": self.pending_period_code,
            "revision": self.pending_revision,
        }
    
    def load_initial_data(self):
        """Загрузка начальных данных"""
        projects = self.project_controller.load_projects()
        references = self._load_references()
        
        # Сигнал по‑прежнему передаём список Project, но левая панель
        # теперь строится по новой архитектуре (год → проект → форма → период → ревизии)
        self.projects_updated.emit(projects)
        self.references_updated.emit(references)
    
    def refresh_references(self):
        """Обновление справочников (публичный метод)"""
        references = self._load_references()
        self.references_updated.emit(references)
        return references
    
    def _load_references(self) -> List[Reference]:
        """Загрузка справочников"""
        # Загружаем список справочников (метаданные)
        try:
            references = self.db_manager.load_references()
        except Exception as e:
            logger.error(f"Ошибка загрузки метаданных справочников: {e}", exc_info=True)
            references = []
        
        # Загружаем данные справочников исключительно из индивидуальных SQL-таблиц
        # Очищаем старые справочники перед загрузкой новых
        self.references.pop('доходы', None)
        self.references.pop('источники', None)
        
        try:
            income_df = self.db_manager.load_income_reference_df()
            if income_df is not None and not income_df.empty:
                self.references['доходы'] = income_df
                logger.info(f"Справочник доходов загружен: {income_df.shape}")
            else:
                logger.warning("Справочник доходов пуст или не найден")
        except Exception as e:
            logger.error(f"Ошибка загрузки справочника доходов из SQL: {e}", exc_info=True)

        try:
            sources_df = self.db_manager.load_sources_reference_df()
            if sources_df is not None and not sources_df.empty:
                self.references['источники'] = sources_df
                logger.info(f"Справочник источников загружен: {sources_df.shape}")
            else:
                logger.warning("Справочник источников пуст или не найден")
        except Exception as e:
            logger.error(f"Ошибка загрузки справочника источников из SQL: {e}", exc_info=True)
        
        return references
    
    def create_project(self, project_data: Dict[str, Any]) -> Optional[Project]:
        """Создание нового проекта"""
        project = self.project_controller.create_project(project_data)
        if project:
            self.current_project = project
            self._initialize_form_for_project()
            # Ревизия создается только при загрузке формы, не при создании проекта
        return project
    
    def update_project(self, project_data: Dict[str, Any]) -> bool:
        """Обновление существующего проекта"""
        success = self.project_controller.update_project(project_data)
        if success:
            # Синхронизируем текущий проект
            self.current_project = self.project_controller.current_project
        return success
    
    def delete_form_revision(self, revision_id: int) -> None:
        """Удаление одной ревизии формы (новая архитектура)"""
        try:
            self.db_manager.delete_form_revision(revision_id)
            # Сбрасываем текущие ссылки, если удалена активная ревизия
            if self.current_revision_id == revision_id:
                self.current_revision_id = None
                if self.current_project:
                    self.current_project.data = {}
                if self.current_form:
                    self.current_form = None
            # Обновляем список проектов после удаления
            projects = self.project_controller.load_projects()
            self.projects_updated.emit(projects)
        except Exception as e:
            self.error_occurred.emit(f"Ошибка удаления ревизии: {e}")
    
    def update_form_revision(self, revision_id: int, revision_data: Dict[str, Any]) -> bool:
        """Обновление ревизии формы"""
        try:
            from models.base_models import ProjectStatus
            
            revision = revision_data.get('revision', '').strip()
            status_str = revision_data.get('status', 'created')
            file_path = revision_data.get('file_path', '').strip()
            
            try:
                status = ProjectStatus(status_str)
            except ValueError:
                status = ProjectStatus.CREATED
            
            success = self.db_manager.update_form_revision(
                revision_id,
                revision,
                status,
                file_path
            )
            
            if success:
                # Обновляем список проектов после обновления
                projects = self.project_controller.load_projects()
                self.projects_updated.emit(projects)
            
            return success
        except Exception as e:
            self.error_occurred.emit(f"Ошибка обновления ревизии: {e}")
            return False
    
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

            # Если данных в *_values нет (например, ревизия только с legacy-данными), пробуем распарсить исходный файл и сохранить
            if not has_sections:
                source_file_path = revision_record.file_path
                if not source_file_path or not os.path.exists(source_file_path):
                    self.error_occurred.emit("Данные ревизии не найдены и отсутствует исходный файл для загрузки")
                    return
                try:
                    # Загружаем справочники при необходимости
                    reference_data_доходы = self.references.get('доходы')
                    if reference_data_доходы is None or reference_data_доходы.empty:
                        reference_data_доходы = self.db_manager.load_income_reference_df()
                        self.references['доходы'] = reference_data_доходы
                    reference_data_источники = self.references.get('источники')
                    if reference_data_источники is None or reference_data_источники.empty:
                        reference_data_источники = self.db_manager.load_sources_reference_df()
                        self.references['источники'] = reference_data_источники

                    if not self.current_form:
                        self.error_occurred.emit("Форма не инициализирована для парсинга исходного файла")
                        return

                    parsed_data = self.current_form.parse_excel(
                        source_file_path,
                        reference_data_доходы,
                        reference_data_источники,
                    )
                    if not parsed_data:
                        self.error_occurred.emit("Не удалось распарсить исходный файл ревизии")
                        return
                    # Сохраняем распарсенные данные в *_values и метаданные
                    self.db_manager.save_revision_data(project_id, revision_id, parsed_data)
                    revision_data = self.db_manager.load_revision_data(project_id, revision_id)
                    has_sections = revision_data and any(revision_data.get(k) for k in need_sections)
                except Exception as e:
                    self.error_occurred.emit(f"Не удалось восстановить данные ревизии: {e}")
                    return

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

        self.project_loaded.emit(project)
    
    def _initialize_form_for_project(self, form_meta=None):
        """
        Инициализация формы для проекта
        
        Args:
            form_meta: Опциональный FormTypeMeta для использования в качестве fallback
        """
        form_type_code = None
        revision_str = "1.0"

        # 1) Если есть текущая ревизия – берём тип формы и номер ревизии из неё
        if self.current_revision_id and self.current_project:
            try:
                revision_record = self.db_manager.get_form_revision_by_id(self.current_revision_id)
                if revision_record:
                    revision_str = revision_record.revision or revision_str
                    project_forms = self.db_manager.load_project_forms(self.current_project.id)
                    project_form = next((pf for pf in project_forms if pf.id == revision_record.project_form_id), None)
                    if project_form:
                        form_types_meta = {ft.id: ft for ft in self.db_manager.load_form_types_meta()}
                        form_meta_from_db = form_types_meta.get(project_form.form_type_id)
                        if form_meta_from_db:
                            form_type_code = form_meta_from_db.code
            except Exception as e:
                logger.warning(f"Ошибка определения типа формы из ревизии: {e}", exc_info=True)

        # 2) Fallback на переданный form_meta (если есть)
        if not form_type_code and form_meta:
            form_type_code = form_meta.code

        # 3) Если ревизии еще нет – используем параметры, выбранные пользователем
        if not form_type_code:
            form_type_code = self.pending_form_type_code
            revision_str = self.pending_revision or revision_str

        # Инициализируем форму по типу
        if form_type_code == "0503317" or form_type_code == FormType.FORM_0503317.value:
            # Пытаемся взять mapping колонок из справочника форм, если он задан
            column_mapping = None
            try:
                form_meta_from_db = self.db_manager.get_form_type_meta_by_code("0503317")
                if form_meta_from_db and form_meta_from_db.column_mapping:
                    column_mapping = form_meta_from_db.column_mapping
            except Exception:
                column_mapping = None

            self.current_form = Form0503317(revision_str, column_mapping=column_mapping)
            logger.info(f"Форма 0503317 инициализирована с ревизией {revision_str}")
        else:
            self.current_form = None
            if form_type_code:
                logger.warning(f"Форма типа '{form_type_code}' не поддерживается")
            else:
                logger.warning("Не удалось определить тип формы для инициализации")
    
    def _copy_form_file_to_project(self, source_file_path: str, project_id: int) -> str:
        """Копирование файла формы в папку проекта с префиксом даты/времени"""
        from datetime import datetime
        import shutil
        
        # Создаем папку для проекта
        project_dir = Path("data") / "projects" / str(project_id)
        project_dir.mkdir(parents=True, exist_ok=True)
        
        # Формируем имя файла с префиксом даты/времени
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        source_path = Path(source_file_path)
        file_name = f"{timestamp}_{source_path.name}"
        dest_file_path = project_dir / file_name
        
        # Копируем файл
        shutil.copy2(source_file_path, dest_file_path)
        
        return str(dest_file_path)
    
    def load_form_file(self, file_path: str) -> bool:
        """Загрузка файла формы"""
        if not self.current_project:
            self.error_occurred.emit("Проект не выбран")
            return False
        
        # Проверяем, что форма инициализирована
        if not self.current_form:
            self.error_occurred.emit("Форма не инициализирована. Убедитесь, что выбран правильный тип формы.")
            return False
        
        try:
            # Копируем файл в папку проекта с префиксом даты/времени
            copied_file_path = self._copy_form_file_to_project(file_path, self.current_project.id)
            
            # Получаем данные справочников как DataFrame
            reference_data_доходы = self.references.get('доходы')
            reference_data_источники = self.references.get('источники')
            
            # Явно предупреждаем, если справочники не загружены
            missing_refs = []
            if reference_data_доходы is None:
                missing_refs.append("доходов")
            if reference_data_источники is None:
                missing_refs.append("источников финансирования")
            if missing_refs:
                msg = (
                    "Не загружены справочники: "
                    + ", ".join(missing_refs)
                    + ". Уровни строк для соответствующих разделов будут установлены в 0."
                )
                self.error_occurred.emit(msg)
            
            logger.debug(
                f"Загрузка справочников: доходы={reference_data_доходы is not None}, "
                f"источники={reference_data_источники is not None}"
            )
            
            # Парсим форму из скопированного файла
            form_data = self.current_form.parse_excel(
                copied_file_path, 
                reference_data_доходы,  # DataFrame
                reference_data_источники  # DataFrame
            )

            # Сохраняем данные в проект (для отображения в UI)
            self.current_project.data = form_data

            # Определяем тип формы из текущей формы
            form_type_code = None
            if isinstance(self.current_form, Form0503317):
                form_type_code = "0503317"
            elif hasattr(self.current_form, 'form_type'):
                if isinstance(self.current_form.form_type, FormType):
                    form_type_code = self.current_form.form_type.value
                else:
                    form_type_code = str(self.current_form.form_type)
            
            # Определяем период: в приоритете период, выбранный пользователем,
            # а затем (если есть) период из метаданных формы.
            period_code = (self.pending_period_code or "").strip() or None
            if form_data.get('meta_info'):
                # Пытаемся извлечь период из метаданных и, если он задан, используем его
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
                    revision=self.pending_revision or "1.0",
                )
            except Exception as e:
                # Не блокируем работу, если новая архитектура ревизий дала сбой
                logger.error(f"Ошибка регистрации ревизии формы: {e}", exc_info=True)
            
            # Сохраняем данные ревизии отдельно, если ревизия создана
            # Включаем meta_info и результат_исполнения_data, которые относятся к этой ревизии
            if revision_record and revision_record.id:
                self.current_revision_id = revision_record.id
                try:
                    # Формируем полные данные ревизии, включая метаданные
                    revision_data = {
                        'meta_info': form_data.get('meta_info', {}),
                        'результат_исполнения_data': form_data.get('результат_исполнения_data'),
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
            
            return True
            
        except Exception as e:
            self.error_occurred.emit(f"Ошибка загрузки файла: {str(e)}")
            logger.error(f"Ошибка загрузки формы: {e}", exc_info=True)
            return False
    
    def calculate_sums(self):
        """Расчет агрегированных сумм"""
        if not self.current_project or not self.current_form:
            self.error_occurred.emit("Проект не выбран")
            return

        # Для формы 0503317 расчет для доходов и источников невозможен без справочников
        reference_data_доходы = self.references.get('доходы')
        reference_data_источники = self.references.get('источники')
        if isinstance(self.current_form, Form0503317):
            missing_refs = []
            if reference_data_доходы is None or reference_data_доходы.empty:
                missing_refs.append("доходов")
            if reference_data_источники is None or reference_data_источники.empty:
                missing_refs.append("источников финансирования")
            if missing_refs:
                self.error_occurred.emit(
                    "Невозможно выполнить расчет: не загружены справочники "
                    + ", ".join(missing_refs)
                    + ". Загрузка справочников обязательна для расчета доходов и источников."
                )
                return
        
        try:
            # Оптимизация: используем расчет напрямую из нормализованных данных БД
            # вместо преобразования в старый формат и обратно
            if self.current_revision_id:
                calculation_results = self.db_manager.calculate_sums_from_values(
                    project_id=self.current_project.id,
                    revision_id=self.current_revision_id,
                    reference_data_доходы=reference_data_доходы,
                    reference_data_источники=reference_data_источники,
                )
            else:
                # Fallback на старый метод, если ревизия не загружена
                calculation_results = self.current_form.calculate_sums()
            
            # Обновляем данные проекта (только разделы, meta_info и результат_исполнения_data остаются)
            self.current_project.data.update(calculation_results)
            
            # Обновляем статус ревизии напрямую (если ревизия существует)
            if self.current_revision_id:
                try:
                    revision_record = self.db_manager.get_form_revision_by_id(self.current_revision_id)
                    if revision_record:
                        # Обновляем статус ревизии
                        self.db_manager.update_form_revision(
                            revision_id=self.current_revision_id,
                            revision=revision_record.revision,
                            status=ProjectStatus.CALCULATED,
                            file_path=revision_record.file_path or ""
                        )
                        logger.info(f"Статус ревизии {self.current_revision_id} обновлен на CALCULATED")
                except Exception as e:
                    logger.error(f"Ошибка обновления статуса ревизии после расчёта: {e}", exc_info=True)
            
            # Сохраняем обновленные данные ревизии (включая meta_info и результат_исполнения_data)
            if self.current_revision_id:
                try:
                    # Оптимизация: обновляем только вычисленные значения напрямую в БД
                    # без полного пересохранения всех данных
                    self.db_manager.update_calculated_values(
                        self.current_project.id,
                        self.current_revision_id,
                        calculation_results
                    )
                    
                    # Также обновляем meta_info и результат_исполнения_data, если они изменились
                    meta_info = self.current_project.data.get('meta_info')
                    результат_исполнения_data = self.current_project.data.get('результат_исполнения_data')
                    if meta_info or результат_исполнения_data:
                        with sqlite3.connect(self.db_manager.db_path) as conn:
                            cursor = conn.cursor()
                            cursor.execute('DELETE FROM revision_metadata WHERE revision_id=?', (self.current_revision_id,))
                            meta_info_json = json.dumps(meta_info, ensure_ascii=False, default=str) if meta_info else None
                            результат_исполнения_json = json.dumps(результат_исполнения_data, ensure_ascii=False, default=str) if результат_исполнения_data else None
                            if meta_info_json or результат_исполнения_json:
                                cursor.execute(
                                    'INSERT INTO revision_metadata (revision_id, meta_info, результат_исполнения_data) VALUES (?, ?, ?)',
                                    (self.current_revision_id, meta_info_json, результат_исполнения_json)
                                )
                            conn.commit()

                    # После сохранения вычисленных значений перезагружаем ревизию из БД,
                    # чтобы в current_project.data были полные данные (оригинал + расчётные)
                    revision_data = self.db_manager.load_revision_data(
                        self.current_project.id,
                        self.current_revision_id,
                    )
                    if revision_data:
                        self.current_project.data = revision_data
                        # Обновляем форму с перезагруженными данными (включая расчетные значения)
                        if self.current_form:
                            self.current_form.load_saved_data(revision_data)
                        
                        # Отладочный вывод: проверяем наличие расчетных значений для консолидированных расчетов
                        cons_data = revision_data.get('консолидируемые_расчеты_data', [])
                        if cons_data:
                            sample_item = cons_data[0] if len(cons_data) > 0 else None
                            if sample_item:
                                has_calc = any(k.startswith('расчетный_поступления_') for k in sample_item.keys())
                                logger.debug(f"Консолидированные расчеты: {len(cons_data)} строк, расчетные значения: {'есть' if has_calc else 'отсутствуют'}")
                                if has_calc:
                                    calc_keys = [k for k in sample_item.keys() if k.startswith('расчетный_поступления_')]
                                    logger.debug(f"  Примеры ключей расчетных значений: {calc_keys[:3]}")
                except Exception as e:
                    logger.error(f"Ошибка сохранения данных ревизии после расчёта: {e}", exc_info=True)
            
            logger.info("Расчет агрегированных сумм завершен")
            self.calculation_completed.emit(calculation_results)
            
        except Exception as e:
            self.error_occurred.emit(f"Ошибка расчета: {str(e)}")
            logger.error(f"Ошибка расчета: {e}", exc_info=True)
    
    def export_validation(self, output_path: str) -> bool:
        """Экспорт формы с проверкой"""
        if not self.current_project or not self.current_form:
            self.error_occurred.emit("Проект не выбран")
            return False
        
        # Получаем путь к файлу из ревизии (новая архитектура) или из проекта (fallback)
        source_file_path = None
        if self.current_revision_id:
            try:
                revision_record = self.db_manager.get_form_revision_by_id(self.current_revision_id)
                if revision_record and revision_record.file_path:
                    source_file_path = revision_record.file_path
            except Exception as e:
                logger.warning(f"Ошибка получения пути к файлу из ревизии: {e}", exc_info=True)
        
        # Fallback на старый способ
        if not source_file_path:
            source_file_path = getattr(self.current_project, 'file_path', None)
        
        if not source_file_path or not os.path.exists(source_file_path):
            # Для экспорта рассчитываем по данным из БД, исходный файл не обязателен
            source_file_path = output_path  # будет перезаписан export_validation
        
        try:
            # Всегда используем данные из нормализованных таблиц *_values (level/source_row уже сохранены).
            # Если данных нет — возвращаем ошибку, парсинг Excel не выполняем.
            if not self.current_revision_id:
                self.error_occurred.emit("Ревизия не выбрана, данные для экспорта недоступны")
                return False

            values_data = self.db_manager.load_project_data_values(
                self.current_project.id,
                self.current_revision_id,
            )
            if not any(values_data.get(k) for k in (
                'доходы_data', 'расходы_data', 'источники_финансирования_data', 'консолидируемые_расчеты_data'
            )):
                self.error_occurred.emit("Нет данных в *_values для экспорта")
                return False

            # Обновляем данные формы из БД
            self.current_form.доходы_data = values_data.get('доходы_data', [])
            self.current_form.расходы_data = values_data.get('расходы_data', [])
            self.current_form.источники_финансирования_data = values_data.get('источники_финансирования_data', [])
            self.current_form.консолидируемые_расчеты_data = values_data.get('консолидируемые_расчеты_data', [])
            # Обновляем кэш проекта
            self.current_project.data.update(values_data)

            output_file = self.current_form.export_validation(
                source_file_path,
                output_path
            )
            
            # Обновляем статус ревизии напрямую (если ревизия существует)
            if self.current_revision_id:
                try:
                    revision_record = self.db_manager.get_form_revision_by_id(self.current_revision_id)
                    if revision_record:
                        # Обновляем статус ревизии и путь к файлу (может быть обновлен после экспорта)
                        self.db_manager.update_form_revision(
                            revision_id=self.current_revision_id,
                            revision=revision_record.revision,
                            status=ProjectStatus.EXPORTED,
                            file_path=output_file  # Обновляем путь на экспортированный файл
                        )
                        logger.info(f"Статус ревизии {self.current_revision_id} обновлен на EXPORTED")
                except Exception as e:
                    logger.error(f"Ошибка обновления статуса ревизии после экспорта: {e}", exc_info=True)
            
            # Сохраняем обновленные данные ревизии (включая meta_info и результат_исполнения_data)
            if self.current_revision_id:
                try:
                    # Формируем полные данные ревизии, сохраняя meta_info и результат_исполнения_data
                    revision_data = {
                        'meta_info': self.current_project.data.get('meta_info', {}),
                        'результат_исполнения_data': self.current_project.data.get('результат_исполнения_data'),
                        'доходы_data': self.current_project.data.get('доходы_data', []),
                        'расходы_data': self.current_project.data.get('расходы_data', []),
                        'источники_финансирования_data': self.current_project.data.get('источники_финансирования_data', []),
                        'консолидируемые_расчеты_data': self.current_project.data.get('консолидируемые_расчеты_data', [])
                    }
                    self.db_manager.save_revision_data(
                        self.current_project.id,
                        self.current_revision_id,
                        revision_data
                    )
                except Exception as e:
                    logger.error(f"Ошибка сохранения данных ревизии после экспорта: {e}", exc_info=True)
            
            self.export_completed.emit(output_file)
            return True
            
        except Exception as e:
            self.error_occurred.emit(f"Ошибка экспорта: {str(e)}")
            return False

    # ------------------------------------------------------------------
    # Вспомогательная логика для новой архитектуры форм/ревизий
    # ------------------------------------------------------------------

    def _register_form_revision(self, project: Project, status: ProjectStatus, file_path: str, 
                                form_type_code: Optional[str] = None, period_code: Optional[str] = None,
                                revision: Optional[str] = None):
        """
        Зарегистрировать или обновить ревизию формы для указанного проекта
        в новых таблицах project_forms / form_revisions.
        
        Args:
            project: Проект
            status: Статус ревизии
            file_path: Путь к файлу формы
            form_type_code: Код типа формы (например, '0503317'). Если не указан, определяется из project.form_type
            period_code: Код периода (например, 'Q1', 'Y'). Если не указан, берётся из
                         self.pending_period_code (выбран пользователем), а не из project.period
            revision: Номер ревизии. Если не указан, определяется из project.revision или используется "1.0"
        """
        if not project:
            return None

        # Определяем код формы
        if not form_type_code:
            # Fallback на старый способ
            form_type_enum = getattr(project, 'form_type', None)
            if isinstance(form_type_enum, FormType):
                form_type_code = form_type_enum.value
            elif form_type_enum:
                form_type_code = str(form_type_enum).strip()
        
        if not form_type_code:
            logger.warning("Не удалось определить тип формы для регистрации ревизии")
            return None

        # Мета‑информация по типу формы
        form_meta = self.db_manager.get_form_type_meta_by_code(form_type_code)
        if not form_meta:
            logger.warning(f"Тип формы '{form_type_code}' не найден в справочнике форм.")
            return None

        # Определяем период
        if not period_code:
            # Не опираемся больше на project.period (его нет в новой модели),
            # используем явно выбранный пользователем период, если он есть.
            period_code = (self.pending_period_code or "").strip() or None
        
        period_id = None
        if period_code:
            period = self.db_manager.get_period_by_code(
                code=period_code,
                form_type_code=form_meta.code,
            )
            if period:
                period_id = period.id
            else:
                logger.warning(f"Период '{period_code}' не найден для формы '{form_meta.code}'")

        # Создаём/находим ProjectForm для (проект, форма, период)
        project_form = self.db_manager.get_or_create_project_form(
            project_id=project.id,
            form_type_id=form_meta.id,
            period_id=period_id,
        )
        
        logger.debug(f"Создан/найден ProjectForm: id={project_form.id}, project_id={project.id}, "
              f"form_type_id={form_meta.id}, period_id={period_id}")

        # Определяем номер ревизии
        if not revision:
            revision = getattr(project, 'revision', None)
            if revision:
                revision = str(revision).strip()
            if not revision:
                revision = "1.0"
        else:
            revision = str(revision).strip() or "1.0"

        # Создаём или обновляем запись ревизии формы
        revision_record = self.db_manager.create_or_update_form_revision(
            project_form_id=project_form.id,
            revision=revision,
            status=status,
            file_path=file_path or "",
        )
        
        logger.debug(f"Создана/обновлена ревизия: id={revision_record.id}, revision={revision}, "
              f"project_form_id={project_form.id}")
        
        return revision_record

    # ------------------------------------------------------------------
    # Построение дерева проектов (Год → Проект → Форма → Период → Ревизии)
    # ------------------------------------------------------------------

    def build_project_tree(self) -> list:
        """
        Строит структуру для отображения в левой панели:
        [
          {
            "year": "2024",
            "projects": [
              {
                "id": ...,
                "name": "...",
                "municipality": "...",
                "forms": [
                  {
                    "form_code": "0503317",
                    "form_name": "...",
                    "periods": [
                      {
                        "period_code": "Q1",
                        "period_name": "...",
                        "revisions": [
                          {"revision": "1.0", "status": "parsed", "project_id": ..., "file_path": "..."},
                          ...
                        ]
                      }
                    ]
                  }
                ]
              }
            ]
          }
        ]
        """
        from collections import defaultdict
        import re

        # Загружаем все проекты (историческая таблица)
        projects = self.project_controller.load_projects()

        # Загружаем справочники формы и периодов для отображения (один раз, не в цикле)
        form_types_meta = {ft.id: ft for ft in self.db_manager.load_form_types_meta()}

        # Для периодов удобнее индексировать по id
        periods_all = self.db_manager.load_periods()
        periods_by_id = {p.id: p for p in periods_all if p.id is not None}

        # Загружаем справочники годов и МО один раз (оптимизация: не в цикле)
        years_all = self.db_manager.load_years()
        years_by_id = {y.id: y for y in years_all if y.id is not None}
        
        municipalities_all = self.db_manager.load_municipalities()
        municipalities_by_id = {m.id: m for m in municipalities_all if m.id is not None}

        # Год → { project_id → ... }
        years_map = defaultdict(lambda: defaultdict(lambda: defaultdict(lambda: defaultdict(list))))

        project_forms_map: Dict[int, list] = {}

        for project in projects:
            # Определяем год из year_id (новая архитектура) - используем предзагруженный словарь
            year = None
            if hasattr(project, 'year_id') and project.year_id:
                year_ref = years_by_id.get(project.year_id)
                if year_ref:
                    year = year_ref.year
            
            year_key = str(year) if year else "Без года"

            # Гарантируем наличие узла проекта в years_map (чтобы проект не пропадал при отсутствии ревизий)
            _ = years_map[year_key][project.id]

            # Загружаем project_forms и form_revisions для данного проекта
            project_forms = self.db_manager.load_project_forms(project.id)
            project_forms_map[project.id] = project_forms or []

            # Если у проекта нет ревизий (нет загруженных форм), 
            # добавляем проект в years_map с пустым forms_map
            # Это позволит показать проект в дереве без вложенности
            if not project_forms:
                # Инициализируем пустой forms_map для проекта
                if project.id not in years_map[year_key]:
                    years_map[year_key][project.id] = {}
                # Продолжаем - проект будет добавлен в итоговую структуру с пустым списком форм
                continue

            # Есть новые project_forms / form_revisions
            for pf in project_forms:
                ft_meta = form_types_meta.get(pf.form_type_id)
                # Код формы берём только из справочника типов форм
                form_code = ft_meta.code if ft_meta else "UNKNOWN"
                form_name = ft_meta.name if ft_meta else f"Форма {form_code}"

                # Код периода берём из ref_periods по period_id; если нет — используем 'Y' по умолчанию
                period_obj = periods_by_id.get(pf.period_id) if pf.period_id else None
                period_code = period_obj.code if period_obj else "Y"

                revisions = self.db_manager.load_form_revisions(pf.id)
                # Ревизии создаются только при загрузке формы, поэтому если их нет - пропускаем
                if not revisions:
                    continue
                    
                revisions_info = [{
                    "revision_id": r.id,
                    "revision": r.revision,
                    "status": r.status.value,
                    "project_id": project.id,
                    "file_path": r.file_path or "",
                } for r in revisions]

                years_map[year_key][project.id][form_code][period_code].extend(revisions_info)

        # Оптимизация: создаем словари для быстрого поиска (O(1) вместо O(n))
        projects_by_id = {p.id: p for p in projects}
        form_types_meta_by_code = {ft.code: ft for ft in form_types_meta.values()}
        periods_by_code = {p.code: p for p in periods_by_id.values() if p.code}
        
        # Преобразуем в удобную для UI структуру
        tree = []
        for year_key in sorted(years_map.keys(), reverse=True):
            year_entry = {"year": year_key, "projects": []}
            proj_map = years_map[year_key]
            for project_id, forms_map in proj_map.items():
                # Используем словарь для быстрого поиска проекта
                proj_obj = projects_by_id.get(project_id)
                if not proj_obj:
                    continue
                # Получаем название МО из справочника (новая архитектура) - используем предзагруженный словарь
                municipality_name = "-"  # Имя МО по умолчанию
                if hasattr(proj_obj, 'municipality_id') and proj_obj.municipality_id:
                    municip_ref = municipalities_by_id.get(proj_obj.municipality_id)
                    if municip_ref:
                        municipality_name = municip_ref.name
                
                proj_entry = {
                    "id": proj_obj.id,
                    "name": proj_obj.name,
                    "municipality": municipality_name,
                    "forms": [],
                }

                # Используем формы, связанные с проектом (даже если нет ревизий)
                project_forms = project_forms_map.get(project_id, [])
                if project_forms:
                    for pf in project_forms:
                        # Код/название формы
                        ft_meta = form_types_meta.get(pf.form_type_id)
                        form_code = ft_meta.code if ft_meta else "UNKNOWN"
                        form_name = ft_meta.name if ft_meta else f"Форма {form_code}"

                        form_entry = {
                            "form_code": form_code,
                            "form_name": form_name,
                            "periods": [],
                        }

                        # Код/название периода
                        period_obj = periods_by_id.get(pf.period_id) if pf.period_id else None
                        period_code = period_obj.code if period_obj else "Y"
                        period_name = period_obj.name if period_obj else period_code

                        # Ревизии берём из years_map (быстрее и учитывает только нужный проект/форму/период)
                        revisions = forms_map.get(form_code, {}).get(period_code, [])

                        period_entry = {
                            "period_code": period_code,
                            "period_name": period_name,
                            "revisions": sorted(
                                revisions,
                                key=lambda r: r["revision"],
                            ),
                        }

                        form_entry["periods"].append(period_entry)
                        proj_entry["forms"].append(form_entry)

                    # Сортируем формы/периоды
                    proj_entry["forms"].sort(key=lambda f: f["form_code"])
                    for f in proj_entry["forms"]:
                        f["periods"].sort(key=lambda p: p["period_code"])

                    # Сортируем формы по коду
                    proj_entry["forms"].sort(key=lambda f: f["form_code"])
                
                # Добавляем проект в дерево (даже если у него нет форм - это проект без загруженных ревизий)
                year_entry["projects"].append(proj_entry)

            # Сортируем проекты по имени
            year_entry["projects"].sort(key=lambda p: p["name"])
            tree.append(year_entry)

        return tree
    
    def load_reference_file(self, file_path: str, ref_type: str, name: str) -> bool:
        """Загрузка файла справочника"""
        try:
            # Читаем Excel файл в DataFrame
            df = pd.read_excel(file_path)
            logger.info(f"Справочник загружен: {df.shape}, колонки: {list(df.columns)}")

            # Нормализуем названия колонок (убираем пробелы по краям)
            df.columns = [str(c).strip() for c in df.columns]
            
            # Определяем колонку с кодом классификации и приводим её к единому формату
            code_column = None
            if ref_type == 'доходы' and 'код_классификации_ДБ' in df.columns:
                code_column = 'код_классификации_ДБ'
            elif ref_type == 'источники' and 'код_классификации_ИФДБ' in df.columns:
                code_column = 'код_классификации_ИФДБ'

            if code_column:
                # Приводим к строке, убираем пробелы/неразрывные пробелы и дополняем нулями до 20 символов
                df[code_column] = (
                    df[code_column]
                    .astype(str)
                    .str.strip()
                    .str.replace(' ', '', regex=False)
                    .str.replace('\u00A0', '', regex=False)
                    .str.zfill(20)
                )
                logger.debug(f"Колонка '{code_column}' нормализована для справочника '{name}'")
            
            # Проверяем необходимые колонки
            required_columns = []
            if ref_type == 'доходы':
                required_columns = ['код_классификации_ДБ', 'уровень_кода']
            elif ref_type == 'источники':
                required_columns = ['код_классификации_ИФДБ', 'уровень_кода']
            
            missing_columns = [col for col in required_columns if col not in df.columns]
            if missing_columns:
                self.error_occurred.emit(f"В справочнике отсутствуют колонки: {missing_columns}")
                return False
            
            # Создаем объект справочника
            reference = Reference()
            reference.name = name
            reference.reference_type = ref_type
            reference.file_path = file_path
            
            # Сохраняем в БД (метаданные)
            self.db_manager.save_reference(reference)
            
            # Дополнительно сохраняем строки справочника в отдельные SQL-таблицы
            reference_data = df.to_dict('records')
            self.db_manager.save_reference_records(ref_type, reference_data)
            
            # Обновляем кэш как DataFrame из SQL-таблиц
            if ref_type == 'доходы':
                self.references['доходы'] = self.db_manager.load_income_reference_df()
            elif ref_type == 'источники':
                self.references['источники'] = self.db_manager.load_sources_reference_df()
            
            # Обновляем список справочников (метаданные)
            references = self.db_manager.load_references()
            self.references_updated.emit(references)
            
            logger.info(f"Справочник '{name}' успешно загружен. Уровни: {df['уровень_кода'].unique()}")
            return True
            
        except Exception as e:
            self.error_occurred.emit(f"Ошибка загрузки справочника: {str(e)}")
            logger.error(f"Ошибка загрузки справочника: {e}", exc_info=True)
            return False
    
    def delete_project(self, project_id: int):
        """Удаление проекта"""
        self.project_controller.delete_project(project_id)
        
        # Если удален текущий проект, сбрасываем его
        if self.current_project and self.current_project.id == project_id:
            self.current_project = None
            self.current_form = None
    
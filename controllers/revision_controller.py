from typing import Dict, Any, Optional, List
import os

from PyQt5.QtCore import QObject, pyqtSignal

from logger import logger
from models.base_models import Project, ProjectStatus, FormTypeMeta
from models.form_0503317 import Form0503317
from models.database import DatabaseManager


class RevisionController(QObject):
    """
    Контроллер, отвечающий за операции с ревизиями форм:
    - загрузка конкретной ревизии;
    - обновление и удаление ревизий;
    - регистрация ревизии в новых таблицах project_forms / form_revisions;
    - управление выбранными пользователем параметрами формы (тип, период, номер ревизии).

    Этот контроллер инкапсулирует тяжелую бизнес-логику, которая ранее жила
    внутри MainController, чтобы упростить последний.
    """

    # Сигналы проксируются наружу, чтобы MainController мог их использовать
    project_loaded = pyqtSignal(Project)
    error_occurred = pyqtSignal(str)

    def __init__(self, db_manager: DatabaseManager, parent: Optional[QObject] = None):
        super().__init__(parent)
        self.db_manager = db_manager

        # Текущее состояние (устанавливается/используется MainController)
        self.current_project: Optional[Project] = None
        self.current_form = None
        self.current_revision_id: Optional[int] = None

        # Параметры формы, выбранные пользователем до создания первой ревизии
        self.pending_form_type_code: Optional[str] = None
        self.pending_revision: str = "1.0"
        self.pending_period_code: Optional[str] = None

        # Кэш справочников (как в MainController) – передаётся снаружи
        self.references: Dict[str, Any] = {}

    # ------------------------------------------------------------------
    # Параметры формы/периода/ревизии, выбранные пользователем
    # ------------------------------------------------------------------

    def set_current_form_params(self, form_code: str, revision: str, period_code: Optional[str] = None) -> None:
        """
        Сохранить выбранные пользователем параметры формы для текущего проекта
        до загрузки/создания первой ревизии.
        """
        self.pending_form_type_code = (form_code or "").strip() if form_code else None
        self.pending_revision = (revision or "").strip() or "1.0"
        self.pending_period_code = (period_code or "").strip() if period_code else None

        # Пока ревизия ещё не создана
        self.current_revision_id = None

    def set_form_params_from_revision(self, revision_id: int) -> None:
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

    # ------------------------------------------------------------------
    # Операции с ревизиями
    # ------------------------------------------------------------------

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
        except Exception as e:
            self.error_occurred.emit(f"Ошибка удаления ревизии: {e}")

    def update_form_revision(self, revision_id: int, revision_data: Dict[str, Any]) -> bool:
        """Обновление ревизии формы"""
        try:
            revision = revision_data.get("revision", "").strip()
            status_str = revision_data.get("status", "created")
            file_path = revision_data.get("file_path", "").strip()

            try:
                status = ProjectStatus(status_str)
            except ValueError:
                status = ProjectStatus.CREATED

            success = self.db_manager.update_form_revision(
                revision_id,
                revision,
                status,
                file_path,
            )
            return success
        except Exception as e:
            self.error_occurred.emit(f"Ошибка обновления ревизии: {e}")
            return False

    def load_revision(self, revision_id: int, project_id: int, form_meta: Optional[FormTypeMeta] = None) -> Optional[Project]:
        """
        Загрузка конкретной ревизии проекта.

        Часть логики по инициализации формы остаётся в MainController,
        поэтому здесь мы фокусируемся на загрузке данных ревизии и проекта.
        """
        try:
            # Сначала загружаем информацию о ревизии
            revision_record = self.db_manager.get_form_revision_by_id(revision_id)
            if not revision_record:
                self.error_occurred.emit("Ревизия не найдена")
                return None

            # Загружаем проект (MainController должен будет обновить current_project)
            # Здесь предполагается, что вызывающая сторона сама дернёт ProjectController.load_project(...)
            # и передаст актуальный Project в self.current_project.

            # Загружаем данные ревизии
            revision_data = self.db_manager.load_revision_data(project_id, revision_id)
            need_sections = [
                "доходы_data",
                "расходы_data",
                "источники_финансирования_data",
                "консолидируемые_расчеты_data",
            ]
            has_sections = revision_data and any(revision_data.get(k) for k in need_sections)

            if not has_sections:
                self.error_occurred.emit("Данные ревизии не найдены")
                return None

            # Обновляем состояние
            self.current_revision_id = revision_id
            if self.current_project:
                self.current_project.data = revision_data

            return self.current_project
        except Exception as e:
            self.error_occurred.emit(f"Ошибка загрузки ревизии: {str(e)}")
            logger.error(f"Ошибка загрузки ревизии: {e}", exc_info=True)
            return None
    
    def load_revision_with_form_initialization(
        self, 
        revision_id: int, 
        project_id: int,
        project_controller,
        form_controller
    ) -> Optional[Project]:
        """
        Полная загрузка ревизии с инициализацией формы
        
        Args:
            revision_id: ID ревизии
            project_id: ID проекта
            project_controller: Контроллер проектов для загрузки проекта
            form_controller: Контроллер форм для инициализации формы
        
        Returns:
            Загруженный проект или None
        """
        try:
            # Загружаем информацию о ревизии
            revision_record = self.db_manager.get_form_revision_by_id(revision_id)
            if not revision_record:
                self.error_occurred.emit("Ревизия не найдена")
                return None
            
            # Загружаем проект
            project_controller.load_project(project_id)
            project = project_controller.current_project
            
            if not project:
                self.error_occurred.emit("Проект не найден")
                return None
            
            # Определяем тип формы из project_form
            project_forms = self.db_manager.load_project_forms(project_id)
            project_forms_by_id = {pf.id: pf for pf in project_forms}
            project_form = project_forms_by_id.get(revision_record.project_form_id)
            
            if not project_form:
                self.error_occurred.emit(f"ProjectForm не найден для ревизии {revision_id}")
                return None
            
            # Получаем метаданные типа формы
            form_types_meta = {ft.id: ft for ft in self.db_manager.load_form_types_meta()}
            form_meta = form_types_meta.get(project_form.form_type_id)
            
            if not form_meta:
                self.error_occurred.emit(f"Тип формы не найден для form_type_id={project_form.form_type_id}")
                return None
            
            # Сохраняем ID текущей ревизии ДО инициализации формы
            self.current_revision_id = revision_id
            self.current_project = project
            
            # Инициализируем форму
            form_controller.current_project = project
            form_controller.current_revision_id = revision_id
            form_controller.initialize_form_for_project(form_meta=form_meta)
            
            # Загружаем данные ревизии
            revision_data = self.db_manager.load_revision_data(project_id, revision_id)
            need_sections = ['доходы_data', 'расходы_data', 'источники_финансирования_data', 'консолидируемые_расчеты_data']
            has_sections = revision_data and any(revision_data.get(k) for k in need_sections)

            if not has_sections:
                self.error_occurred.emit("Данные ревизии не найдены")
                return None
            
            # Обновляем данные проекта данными ревизии
            project.data = revision_data
            
            # Загружаем данные в форму
            if not form_controller.current_form:
                self.error_occurred.emit(f"Форма типа '{form_meta.code}' не поддерживается")
                return None
            
            form_controller.current_form.load_saved_data(revision_data)
            
            # Пересчитываем дефицит/профицит
            try:
                if hasattr(form_controller.current_form, "_calculate_deficit_proficit"):
                    form_controller.current_form._calculate_deficit_proficit()
                    if getattr(form_controller.current_form, "calculated_deficit_proficit", None):
                        revision_data["calculated_deficit_proficit"] = (
                            form_controller.current_form.calculated_deficit_proficit
                        )
                        project.data = revision_data
            except Exception as e:
                logger.error(
                    f"Ошибка пересчета дефицита/профицита при загрузке ревизии: {e}",
                    exc_info=True,
                )
            
            # Пересчитываем уровни и значения на основе справочников, если файл есть
            if revision_record.file_path and os.path.exists(revision_record.file_path):
                try:
                    reference_data_доходы = self.references.get('доходы')
                    reference_data_источники = self.references.get('источники')
                    
                    if isinstance(form_controller.current_form, Form0503317):
                        updated_data = form_controller.current_form.recalculate_levels_with_references(
                            revision_data,
                            reference_data_доходы,
                            reference_data_источники
                        )
                        if updated_data:
                            project.data = updated_data
                            form_controller.current_form.load_saved_data(updated_data)
                            logger.info("Уровни и значения пересчитаны на основе справочников")
                except Exception as e:
                    logger.error(f"Ошибка пересчета уровней и значений: {e}", exc_info=True)
            
            return project
                
        except Exception as e:
            self.error_occurred.emit(f"Ошибка загрузки ревизии: {str(e)}")
            logger.error(f"Ошибка загрузки ревизии: {e}", exc_info=True)
            return None

    # ------------------------------------------------------------------
    # Регистрация ревизии (новая архитектура)
    # ------------------------------------------------------------------

    def register_form_revision(
        self,
        project: Project,
        status: ProjectStatus,
        file_path: str,
        form_type_code: Optional[str] = None,
        period_code: Optional[str] = None,
        revision: Optional[str] = None,
    ):
        """
        Зарегистрировать или обновить ревизию формы для указанного проекта
        в новых таблицах project_forms / form_revisions.
        """
        if not project:
            return None

        # Определяем код формы
        if not form_type_code:
            form_type_enum = getattr(project, "form_type", None)
            if form_type_enum:
                form_type_code = getattr(form_type_enum, "value", str(form_type_enum)).strip()

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

        logger.debug(
            f"Создан/найден ProjectForm: id={project_form.id}, project_id={project.id}, "
            f"form_type_id={form_meta.id}, period_id={period_id}"
        )

        # Определяем номер ревизии
        if not revision:
            revision = getattr(project, "revision", None)
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

        logger.debug(
            f"Создана/обновлена ревизия: id={revision_record.id}, revision={revision}, "
            f"project_form_id={project_form.id}"
        )

        self.current_revision_id = revision_record.id
        return revision_record


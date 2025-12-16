from typing import Optional, Dict, Any
from pathlib import Path
from datetime import datetime
import shutil
import os

from PyQt5.QtCore import QObject, pyqtSignal

from logger import logger
from models.base_models import Project, FormTypeMeta, FormType, ProjectStatus
from models.form_0503317 import Form0503317
from models.database import DatabaseManager


class FormController(QObject):
    """
    Контроллер, отвечающий за работу с формами:
    - инициализация формы для проекта;
    - загрузка файла формы;
    - копирование файлов в папку проекта.
    """

    error_occurred = pyqtSignal(str)

    def __init__(self, db_manager: DatabaseManager, parent: Optional[QObject] = None):
        super().__init__(parent)
        self.db_manager = db_manager

        # Текущее состояние (устанавливается MainController)
        self.current_project: Optional[Project] = None
        self.current_form = None
        self.current_revision_id: Optional[int] = None

        # Параметры формы (из RevisionController)
        self.pending_form_type_code: Optional[str] = None
        self.pending_revision: str = "1.0"

        # Кэш справочников (передаётся снаружи)
        self.references: Dict[str, Any] = {}

    def initialize_form_for_project(self, form_meta: Optional[FormTypeMeta] = None) -> None:
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

    def copy_form_file_to_project(self, source_file_path: str, project_id: int) -> str:
        """Копирование файла формы в папку проекта с префиксом даты/времени"""
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

    def load_form_file(self, file_path: str) -> Dict[str, Any]:
        """
        Загрузка и парсинг файла формы
        
        Returns:
            Словарь с данными формы или None при ошибке
        """
        if not self.current_project:
            self.error_occurred.emit("Проект не выбран")
            return None
        
        # Проверяем, что форма инициализирована
        if not self.current_form:
            self.error_occurred.emit("Форма не инициализирована. Убедитесь, что выбран правильный тип формы.")
            return None
        
        try:
            # Копируем файл в папку проекта с префиксом даты/времени
            copied_file_path = self.copy_form_file_to_project(file_path, self.current_project.id)
            
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
            period_code = None
            if form_data.get('meta_info'):
                # Пытаемся извлечь период из метаданных и, если он задан, используем его
                period_str = form_data.get('meta_info', {}).get('period')
                if period_str:
                    period_code = str(period_str).strip()

            # Возвращаем данные формы и метаинформацию
            return {
                'form_data': form_data,
                'file_path': copied_file_path,
                'form_type_code': form_type_code,
                'period_code': period_code,
            }
            
        except Exception as e:
            self.error_occurred.emit(f"Ошибка загрузки файла: {str(e)}")
            logger.error(f"Ошибка загрузки формы: {e}", exc_info=True)
            return None

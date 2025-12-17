from typing import Dict, Any, Optional
import os

from PyQt5.QtCore import QObject, pyqtSignal

from logger import logger
from models.base_models import Project, ProjectStatus
from models.form_0503317 import Form0503317
from models.database import DatabaseManager


class CalculationController(QObject):
    """
    Контроллер, отвечающий за расчеты и экспорт:
    - расчет агрегированных сумм;
    - экспорт формы с проверкой.
    """

    calculation_completed = pyqtSignal(dict)
    export_completed = pyqtSignal(str)
    error_occurred = pyqtSignal(str)

    def __init__(self, db_manager: DatabaseManager, parent: Optional[QObject] = None):
        super().__init__(parent)
        self.db_manager = db_manager

        # Текущее состояние (устанавливается MainController)
        self.current_project: Optional[Project] = None
        self.current_form = None
        self.current_revision_id: Optional[int] = None

        # Кэш справочников (передаётся снаружи)
        self.references: Dict[str, Any] = {}

    def calculate_sums(self) -> Optional[Dict[str, Any]]:
        """Расчет агрегированных сумм"""
        if not self.current_project or not self.current_form:
            self.error_occurred.emit("Проект не выбран")
            return None

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
                return None
        
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
            
            # Обновляем данные проекта (только разделы, meta_info и calculated_deficit_proficit остаются)
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
            
            # Сохраняем обновленные данные ревизии (через *_values и, при наличии, meta_info)
            if self.current_revision_id:
                try:
                    # Оптимизация: обновляем только вычисленные значения напрямую в БД
                    # без полного пересохранения всех данных
                    self.db_manager.update_calculated_values(
                        self.current_project.id,
                        self.current_revision_id,
                        calculation_results
                    )
                    
                    # Также при необходимости обновляем meta_info (шапку формы)
                    meta_info = self.current_project.data.get('meta_info')
                    if meta_info:
                        try:
                            self.db_manager.save_revision_data(
                                project_id=self.current_project.id,
                                revision_id=self.current_revision_id,
                                data={'meta_info': meta_info},
                            )
                        except Exception as e:
                            logger.error(
                                f"Ошибка обновления revision_metadata после расчёта: {e}",
                                exc_info=True,
                            )

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
                            # Пересчитываем дефицит/профицит по актуальным данным формы
                            try:
                                if hasattr(self.current_form, "_calculate_deficit_proficit"):
                                    self.current_form._calculate_deficit_proficit()
                                    if getattr(self.current_form, "calculated_deficit_proficit", None):
                                        self.current_project.data[
                                            "calculated_deficit_proficit"
                                        ] = self.current_form.calculated_deficit_proficit
                            except Exception as e:
                                logger.error(
                                    f"Ошибка пересчета дефицита/профицита после загрузки ревизии: {e}",
                                    exc_info=True,
                                )
                        
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
            return calculation_results
            
        except Exception as e:
            self.error_occurred.emit(f"Ошибка расчета: {str(e)}")
            logger.error(f"Ошибка расчета: {e}", exc_info=True)
            return None

    def export_validation(self, output_path: str) -> Optional[str]:
        """Экспорт формы с проверкой"""
        if not self.current_project or not self.current_form:
            self.error_occurred.emit("Проект не выбран")
            return None
        
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
        
        # Проверяем существование исходного файла перед экспортом
        if not source_file_path or not os.path.exists(source_file_path):
            error_msg = (
                f"Исходный файл не найден: {source_file_path}. "
                f"Для экспорта с валидацией необходим исходный Excel файл."
            )
            logger.error(error_msg)
            self.error_occurred.emit(error_msg)
            return None
        
        try:
            # Всегда используем данные из нормализованных таблиц *_values (level/source_row уже сохранены).
            # Если данных нет — возвращаем ошибку, парсинг Excel не выполняем.
            if not self.current_revision_id:
                self.error_occurred.emit("Ревизия не выбрана, данные для экспорта недоступны")
                return None

            values_data = self.db_manager.load_project_data_values(
                self.current_project.id,
                self.current_revision_id,
            )
            if not any(values_data.get(k) for k in (
                'доходы_data', 'расходы_data', 'источники_финансирования_data', 'консолидируемые_расчеты_data'
            )):
                self.error_occurred.emit("Нет данных в *_values для экспорта")
                return None

            # Выполняем расчет сумм из нормализованных данных, чтобы получить расчетные значения
            # для подсветки ошибок в Excel (аналогично calculate_sums).
            reference_data_доходы = self.references.get('доходы')
            reference_data_источники = self.references.get('источники')
            calculation_results = self.db_manager.calculate_sums_from_values(
                project_id=self.current_project.id,
                revision_id=self.current_revision_id,
                reference_data_доходы=reference_data_доходы,
                reference_data_источники=reference_data_источники,
            )

            # Обновляем данные формы расчетными значениями
            self.current_form.доходы_data = calculation_results.get('доходы_data', [])
            self.current_form.расходы_data = calculation_results.get('расходы_data', [])
            self.current_form.источники_финансирования_data = calculation_results.get('источники_финансирования_data', [])
            self.current_form.консолидируемые_расчеты_data = calculation_results.get('консолидируемые_расчеты_data', [])
            self.current_form.calculated_deficit_proficit = calculation_results.get('calculated_deficit_proficit')

            # Обновляем кэш проекта (оригинальные + расчетные данные)
            self.current_project.data.update(calculation_results)

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
            
            # Сохраняем обновленные данные ревизии (включая meta_info)
            if self.current_revision_id:
                try:
                    # Формируем полные данные ревизии, сохраняя meta_info
                    revision_data = {
                        'meta_info': self.current_project.data.get('meta_info', {}),
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
            return output_file
            
        except Exception as e:
            self.error_occurred.emit(f"Ошибка экспорта: {str(e)}")
            logger.error(f"Ошибка экспорта: {e}", exc_info=True)
            return None

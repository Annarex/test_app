from typing import List, Dict, Any, Optional
import pandas as pd

from PyQt5.QtCore import QObject, pyqtSignal

from logger import logger
from models.base_models import Reference
from models.database import DatabaseManager


class ReferenceController(QObject):
    """
    Контроллер, отвечающий за работу со справочниками:
    - загрузка справочников из БД;
    - загрузка справочников из файлов;
    - обновление кэша справочников.
    """

    references_updated = pyqtSignal(list)
    error_occurred = pyqtSignal(str)

    def __init__(self, db_manager: DatabaseManager, parent: Optional[QObject] = None):
        super().__init__(parent)
        self.db_manager = db_manager
        
        # Справочники (храним как DataFrame)
        self.references: Dict[str, Any] = {}

    def load_references(self) -> List[Reference]:
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

    def refresh_references(self) -> List[Reference]:
        """Обновление справочников (публичный метод)"""
        references = self.load_references()
        self.references_updated.emit(references)
        return references

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

    def get_reference(self, ref_type: str):
        """Получить справочник по типу"""
        return self.references.get(ref_type)

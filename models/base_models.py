from abc import ABC, abstractmethod
from datetime import datetime
from enum import Enum
import sqlite3
from pathlib import Path
from typing import Dict, List, Any, Optional
import pandas as pd

class FormType(Enum):
    FORM_0503317 = "0503317"

class ProjectStatus(Enum):
    CREATED = "created"
    PARSED = "parsed"
    CALCULATED = "calculated"
    EXPORTED = "exported"

class BaseFormModel(ABC):
    """Базовый класс для всех форм"""
    
    def __init__(self, form_type: FormType, revision: str):
        self.form_type = form_type
        self.revision = revision
        self.meta_info = {}
        self.sections = {}
        
    @abstractmethod
    def parse_excel(self, file_path: str, reference_data_доходы: pd.DataFrame = None, reference_data_источники: pd.DataFrame = None) -> Dict[str, Any]:
        """Парсинг Excel файла формы"""
        pass
    
    @abstractmethod
    def calculate_sums(self) -> Dict[str, Any]:
        """Расчет агрегированных сумм"""
        pass
    
    @abstractmethod
    def validate_data(self) -> List[Dict[str, Any]]:
        """Валидация данных"""
        pass
    
    @abstractmethod
    def get_form_constants(self) -> Any:
        """Получение констант формы"""
        pass

class Project:
    """Проект работы с формой"""
    
    def __init__(self):
        self.id = None
        self.name = ""
        # Поля для связи со справочниками
        self.year_id = None
        self.municipality_id = None
        self.created_at = datetime.now()
        self.data = {}
        
    def to_dict(self) -> Dict[str, Any]:
        return {
            'id': self.id,
            'name': self.name,
            'year_id': self.year_id,
            'municipality_id': self.municipality_id,
            'created_at': self.created_at.isoformat(),
            'data': self.data
        }
    
    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'Project':
        project = cls()
        project.id = data.get('id')
        project.name = data.get('name', '')
        project.year_id = data.get('year_id')
        project.municipality_id = data.get('municipality_id')
        created_at = data.get('created_at')
        project.created_at = datetime.fromisoformat(created_at) if created_at else datetime.now()
        project.data = data.get('data', {})
        return project

class Reference:
    """Справочник"""
    
    def __init__(self):
        self.id = None
        self.name = ""
        self.reference_type = ""  # 'доходы', 'источники'
        self.file_path = ""
        self.loaded_at = datetime.now()
        # Поле data удалено - данные справочников хранятся в отдельных SQL-таблицах
        
    def to_dict(self) -> Dict[str, Any]:
        return {
            'id': self.id,
            'name': self.name,
            'reference_type': self.reference_type,
            'file_path': self.file_path,
            'loaded_at': self.loaded_at.isoformat()
        }
    
    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'Reference':
        ref = cls()
        ref.id = data.get('id')
        ref.name = data.get('name', '')
        ref.reference_type = data.get('reference_type', '')
        ref.file_path = data.get('file_path', '')
        loaded_at = data.get('loaded_at')
        ref.loaded_at = datetime.fromisoformat(loaded_at) if loaded_at else datetime.now()
        return ref


# --------------------------------------------------------------------
# Новые модельные классы для расширенной архитектуры
# --------------------------------------------------------------------

class YearRef:
    """Справочник лет (для привязки проектов к году)"""

    def __init__(self):
        self.id: Optional[int] = None
        self.year: int = 0
        self.is_active: bool = True

    @classmethod
    def from_row(cls, row: Dict[str, Any]) -> "YearRef":
        y = cls()
        y.id = row.get("id")
        y.year = int(row.get("year", 0) or 0)
        y.is_active = bool(row.get("is_active", 1))
        return y


class MunicipalityRef:
    """Справочник муниципальных образований"""

    def __init__(self):
        self.id: Optional[int] = None
        self.code: str = ""
        self.name: str = ""
        self.is_active: bool = True

    @classmethod
    def from_row(cls, row: Dict[str, Any]) -> "MunicipalityRef":
        m = cls()
        m.id = row.get("id")
        m.code = (row.get("code") or "").strip()
        m.name = (row.get("name") or "").strip()
        m.is_active = bool(row.get("is_active", 1))
        return m


class FormTypeMeta:
    """
    Мета‑информация о типе формы (0503317, 0503314 и т.п.).
    Код должен совпадать со значением enum FormType, если он есть.
    """

    def __init__(self):
        self.id: Optional[int] = None
        self.code: str = ""        # '0503317', '0503314', ...
        self.name: str = ""        # Читаемое название
        self.periodicity: str = "" # 'yearly', 'quarterly', 'half_year', ...
        self.column_mapping: Optional[Dict[str, Any]] = None  # Mapping колонок для экспорта/валидации
        self.is_active: bool = True

    @classmethod
    def from_row(cls, row: Dict[str, Any]) -> "FormTypeMeta":
        import json
        f = cls()
        f.id = row.get("id")
        f.code = (row.get("code") or "").strip()
        f.name = (row.get("name") or "").strip()
        f.periodicity = (row.get("periodicity") or "").strip()
        column_mapping_str = row.get("column_mapping")
        if column_mapping_str:
            try:
                f.column_mapping = json.loads(column_mapping_str)
            except (json.JSONDecodeError, TypeError):
                f.column_mapping = None
        else:
            f.column_mapping = None
        f.is_active = bool(row.get("is_active", 1))
        return f


class PeriodRef:
    """Справочник периодов (год, квартал, полугодие и т.п.)"""

    def __init__(self):
        self.id: Optional[int] = None
        self.code: str = ""        # Y, Q1, Q2, Q3, Q4, H1, H2, ...
        self.name: str = ""
        self.sort_order: int = 0
        self.form_type_code: str = ""  # опциональная привязка к форме
        self.is_active: bool = True

    @classmethod
    def from_row(cls, row: Dict[str, Any]) -> "PeriodRef":
        p = cls()
        p.id = row.get("id")
        p.code = (row.get("code") or "").strip()
        p.name = (row.get("name") or "").strip()
        p.sort_order = int(row.get("sort_order", 0) or 0)
        p.form_type_code = (row.get("form_type_code") or "").strip()
        p.is_active = bool(row.get("is_active", 1))
        return p


class ProjectForm:
    """
    Связка Проект ↔ Форма ↔ Период.
    Один проект может содержать несколько форм и периодов.
    """

    def __init__(self):
        self.id: Optional[int] = None
        self.project_id: Optional[int] = None
        self.form_type_id: Optional[int] = None
        self.period_id: Optional[int] = None

    @classmethod
    def from_row(cls, row: Dict[str, Any]) -> "ProjectForm":
        pf = cls()
        pf.id = row.get("id")
        pf.project_id = row.get("project_id")
        pf.form_type_id = row.get("form_type_id")
        pf.period_id = row.get("period_id")
        return pf


class FormRevisionRecord:
    """
    Ревизия конкретной формы в рамках ProjectForm.
    Здесь будет привязка к файлу, статусу и данным формы.
    """

    def __init__(self):
        self.id: Optional[int] = None
        self.project_form_id: Optional[int] = None
        self.revision: str = ""               # '1.0', '2.0', '2.1'
        self.status: ProjectStatus = ProjectStatus.CREATED
        self.file_path: str = ""
        self.created_at: datetime = datetime.now()

    @classmethod
    def from_row(cls, row: Dict[str, Any]) -> "FormRevisionRecord":
        fr = cls()
        fr.id = row.get("id")
        fr.project_form_id = row.get("project_form_id")
        fr.revision = (row.get("revision") or "").strip()
        status = row.get("status") or ProjectStatus.CREATED.value
        fr.status = ProjectStatus(status)
        fr.file_path = (row.get("file_path") or "").strip()
        created_at = row.get("created_at")
        fr.created_at = datetime.fromisoformat(created_at) if created_at else datetime.now()
        return fr
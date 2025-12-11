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


# --------------------------------------------------------------------
# Модели для справочников классификаций
# --------------------------------------------------------------------

class IncomeCode:
    """Модель кода дохода из справочника"""
    
    def __init__(self):
        self.код: str = ""
        self.название: str = ""
        self.уровень: int = 0
        self.наименование: str = ""
    
    @classmethod
    def from_row(cls, row: Dict[str, Any]) -> "IncomeCode":
        code = cls()
        code.код = str(row.get("код", "")).strip()
        code.название = str(row.get("название", "")).strip()
        code.уровень = int(row.get("уровень", 0) or 0)
        code.наименование = str(row.get("наименование", "")).strip()
        return code


class ExpenseCode:
    """Модель кода расхода из справочника"""
    
    def __init__(self):
        self.код: str = ""
        self.название: str = ""
        self.уровень: int = 0
        self.код_Р: str = ""
        self.код_ПР: str = ""
        self.код_ЦС: str = ""
        self.код_ВР: str = ""
        self.наименование: str = ""
    
    @classmethod
    def from_row(cls, row: Dict[str, Any]) -> "ExpenseCode":
        code = cls()
        code.код = str(row.get("код", "")).strip()
        code.название = str(row.get("название", "")).strip()
        code.уровень = int(row.get("уровень", 0) or 0)
        code.код_Р = str(row.get("код_Р", "")).strip()
        code.код_ПР = str(row.get("код_ПР", "")).strip()
        code.код_ЦС = str(row.get("код_ЦС", "")).strip()
        code.код_ВР = str(row.get("код_ВР", "")).strip()
        code.наименование = str(row.get("наименование", "")).strip()
        return code


class MunicipalityTypeRef:
    """Справочник видов муниципальных образований"""
    
    def __init__(self):
        self.код_вида_МО: str = ""
        self.наименование: str = ""
    
    @classmethod
    def from_row(cls, row: Dict[str, Any]) -> "MunicipalityTypeRef":
        m = cls()
        m.код_вида_МО = str(row.get("код_вида_МО", "")).strip()
        m.наименование = str(row.get("наименование", "")).strip()
        return m


class ExtendedMunicipalityRef(MunicipalityRef):
    """Расширенная модель муниципального образования с полными данными"""
    
    def __init__(self):
        super().__init__()
        self.код_вида_МО: Optional[str] = None
        self.адрес_совет: str = ""
        self.адрес_администрация: str = ""
        self.совет_почта: str = ""
        self.администрация_почта: str = ""
        self.должность_совет: str = ""
        self.фамилия_совет: str = ""
        self.имя_совет: str = ""
        self.отчество_совет: str = ""
        self.должность_администрация: str = ""
        self.фамилия_администрация: str = ""
        self.имя_администрация: str = ""
        self.отчество_администрация: str = ""
        self.родительный_падеж: str = ""
        self.дата_соглашения: Optional[datetime] = None
        self.дата_решения: Optional[datetime] = None
        self.номер_решения: str = ""
        self.начальная_доходы: float = 0.0
        self.начальная_расходы: float = 0.0
        self.начальная_дефицит: float = 0.0
    
    @classmethod
    def from_row(cls, row: Dict[str, Any]) -> "ExtendedMunicipalityRef":
        m = cls()
        m.id = row.get("id")
        m.code = str(row.get("code", "")).strip()
        m.name = str(row.get("name", "")).strip()
        m.код_вида_МО = row.get("код_вида_МО")
        m.адрес_совет = str(row.get("адрес_совет", "")).strip()
        m.адрес_администрация = str(row.get("адрес_администрация", "")).strip()
        m.совет_почта = str(row.get("совет_почта", "")).strip()
        m.администрация_почта = str(row.get("администрация_почта", "")).strip()
        m.должность_совет = str(row.get("должность_совет", "")).strip()
        m.фамилия_совет = str(row.get("фамилия_совет", "")).strip()
        m.имя_совет = str(row.get("имя_совет", "")).strip()
        m.отчество_совет = str(row.get("отчество_совет", "")).strip()
        m.должность_администрация = str(row.get("должность_администрация", "")).strip()
        m.фамилия_администрация = str(row.get("фамилия_администрация", "")).strip()
        m.имя_администрация = str(row.get("имя_администрация", "")).strip()
        m.отчество_администрация = str(row.get("отчество_администрация", "")).strip()
        m.родительный_падеж = str(row.get("родительный_падеж", "")).strip()
        
        # Даты
        дата_соглашения = row.get("дата_соглашения")
        if дата_соглашения:
            if isinstance(дата_соглашения, str):
                try:
                    m.дата_соглашения = datetime.fromisoformat(дата_соглашения)
                except:
                    m.дата_соглашения = None
            elif isinstance(дата_соглашения, datetime):
                m.дата_соглашения = дата_соглашения
        
        дата_решения = row.get("дата_решения")
        if дата_решения:
            if isinstance(дата_решения, str):
                try:
                    m.дата_решения = datetime.fromisoformat(дата_решения)
                except:
                    m.дата_решения = None
            elif isinstance(дата_решения, datetime):
                m.дата_решения = дата_решения
        
        m.номер_решения = str(row.get("номер_решения", "")).strip()
        m.начальная_доходы = float(row.get("начальная_доходы", 0) or 0)
        m.начальная_расходы = float(row.get("начальная_расходы", 0) or 0)
        m.начальная_дефицит = float(row.get("начальная_дефицит", 0) or 0)
        m.is_active = bool(row.get("is_active", 1))
        return m


class GRBSRef:
    """Справочник ГРБС (Главных распорядителей бюджетных средств)"""
    
    def __init__(self):
        self.код_ГРБС: str = ""
        self.наименование: str = ""
    
    @classmethod
    def from_row(cls, row: Dict[str, Any]) -> "GRBSRef":
        g = cls()
        g.код_ГРБС = str(row.get("код_ГРБС", "")).strip()
        g.наименование = str(row.get("наименование", "")).strip()
        return g


class ExpenseSectionRef:
    """Справочник разделов/подразделов классификации расходов"""
    
    def __init__(self):
        self.код_РП: str = ""
        self.наименование: str = ""
        self.утверждающий_документ: Optional[str] = None
    
    @classmethod
    def from_row(cls, row: Dict[str, Any]) -> "ExpenseSectionRef":
        e = cls()
        e.код_РП = str(row.get("код_РП", "")).strip()
        e.наименование = str(row.get("наименование", "")).strip()
        e.утверждающий_документ = row.get("утверждающий_документ")
        return e


class TargetExpenseRef:
    """Справочник целевых статей расходов"""
    
    def __init__(self):
        self.код_ЦСР: str = ""
        self.наименование: str = ""
    
    @classmethod
    def from_row(cls, row: Dict[str, Any]) -> "TargetExpenseRef":
        t = cls()
        t.код_ЦСР = str(row.get("код_ЦСР", "")).strip()
        t.наименование = str(row.get("наименование", "")).strip()
        return t


class ExpenseTypeRef:
    """Справочник видов статей расходов"""
    
    def __init__(self):
        self.код_вида_СР: str = ""
        self.наименование: str = ""
    
    @classmethod
    def from_row(cls, row: Dict[str, Any]) -> "ExpenseTypeRef":
        e = cls()
        e.код_вида_СР = str(row.get("код_вида_СР", "")).strip()
        e.наименование = str(row.get("наименование", "")).strip()
        return e


class ProgramNonProgramRef:
    """Справочник программных/непрограммных статей"""
    
    def __init__(self):
        self.код_ПНС: str = ""
        self.наименование: str = ""
    
    @classmethod
    def from_row(cls, row: Dict[str, Any]) -> "ProgramNonProgramRef":
        p = cls()
        p.код_ПНС = str(row.get("код_ПНС", "")).strip()
        p.наименование = str(row.get("наименование", "")).strip()
        return p


class ExpenseKindRef:
    """Справочник видов расходов"""
    
    def __init__(self):
        self.код_ВР: str = ""
        self.наименование: str = ""
        self.утверждающий_документ: Optional[str] = None
    
    @classmethod
    def from_row(cls, row: Dict[str, Any]) -> "ExpenseKindRef":
        e = cls()
        e.код_ВР = str(row.get("код_ВР", "")).strip()
        e.наименование = str(row.get("наименование", "")).strip()
        e.утверждающий_документ = row.get("утверждающий_документ")
        return e


class NationalProjectRef:
    """Справочник национальных проектов целевой статьи расходов"""
    
    def __init__(self):
        self.код_НПЦСР: str = ""
        self.наименование: str = ""
        self.утверждающий_документ: Optional[str] = None
    
    @classmethod
    def from_row(cls, row: Dict[str, Any]) -> "NationalProjectRef":
        n = cls()
        n.код_НПЦСР = str(row.get("код_НПЦСР", "")).strip()
        n.наименование = str(row.get("наименование", "")).strip()
        n.утверждающий_документ = row.get("утверждающий_документ")
        return n


class GADBRef:
    """Справочник ГАДБ (Главных администраторов доходов бюджета)"""
    
    def __init__(self):
        self.код_ГАДБ: str = ""
        self.наименование: str = ""
    
    @classmethod
    def from_row(cls, row: Dict[str, Any]) -> "GADBRef":
        g = cls()
        g.код_ГАДБ = str(row.get("код_ГАДБ", "")).strip()
        g.наименование = str(row.get("наименование", "")).strip()
        return g


class IncomeGroupRef:
    """Справочник групп доходов бюджетов"""
    
    def __init__(self):
        self.код_группы_ДБ: str = ""
        self.наименование: str = ""
    
    @classmethod
    def from_row(cls, row: Dict[str, Any]) -> "IncomeGroupRef":
        i = cls()
        i.код_группы_ДБ = str(row.get("код_группы_ДБ", "")).strip()
        i.наименование = str(row.get("наименование", "")).strip()
        return i


class IncomeSubgroupRef:
    """Справочник подгрупп доходов бюджетов"""
    
    def __init__(self):
        self.код_подгруппы_ДБ: str = ""
        self.наименование: str = ""
    
    @classmethod
    def from_row(cls, row: Dict[str, Any]) -> "IncomeSubgroupRef":
        i = cls()
        i.код_подгруппы_ДБ = str(row.get("код_подгруппы_ДБ", "")).strip()
        i.наименование = str(row.get("наименование", "")).strip()
        return i


class IncomeArticleRef:
    """Справочник статей/подстатей доходов бюджетов"""
    
    def __init__(self):
        self.код_статьи_подстатьи_ДБ: str = ""
        self.наименование: str = ""
    
    @classmethod
    def from_row(cls, row: Dict[str, Any]) -> "IncomeArticleRef":
        i = cls()
        i.код_статьи_подстатьи_ДБ = str(row.get("код_статьи_подстатьи_ДБ", "")).strip()
        i.наименование = str(row.get("наименование", "")).strip()
        return i


class IncomeElementRef:
    """Справочник элементов доходов бюджетов"""
    
    def __init__(self):
        self.код_элемента_ДБ: str = ""
        self.наименование: str = ""
    
    @classmethod
    def from_row(cls, row: Dict[str, Any]) -> "IncomeElementRef":
        i = cls()
        i.код_элемента_ДБ = str(row.get("код_элемента_ДБ", "")).strip()
        i.наименование = str(row.get("наименование", "")).strip()
        return i


class IncomeSubtypeGroupRef:
    """Справочник групп подвидов доходов бюджетов"""
    
    def __init__(self):
        self.код_группы_ПДБ: str = ""
        self.наименование: str = ""
    
    @classmethod
    def from_row(cls, row: Dict[str, Any]) -> "IncomeSubtypeGroupRef":
        i = cls()
        i.код_группы_ПДБ = str(row.get("код_группы_ПДБ", "")).strip()
        i.наименование = str(row.get("наименование", "")).strip()
        return i


class IncomeAnalyticGroupRef:
    """Справочник аналитических групп подвидов доходов бюджетов"""
    
    def __init__(self):
        self.код_группы_АПДБ: str = ""
        self.наименование: str = ""
    
    @classmethod
    def from_row(cls, row: Dict[str, Any]) -> "IncomeAnalyticGroupRef":
        i = cls()
        i.код_группы_АПДБ = str(row.get("код_группы_АПДБ", "")).strip()
        i.наименование = str(row.get("наименование", "")).strip()
        return i


class IncomeLevelRef:
    """Справочник уровней доходов"""
    
    def __init__(self):
        self.код_уровня: str = ""
        self.наименование: str = ""
        self.цвет: Optional[str] = None
    
    @classmethod
    def from_row(cls, row: Dict[str, Any]) -> "IncomeLevelRef":
        i = cls()
        i.код_уровня = str(row.get("код_уровня", "")).strip()
        i.наименование = str(row.get("наименование", "")).strip()
        i.цвет = row.get("цвет")
        return i
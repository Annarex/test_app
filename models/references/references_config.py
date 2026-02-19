"""
Единая конфигурация справочников системы

Этот модуль содержит полную конфигурацию всех справочников:
- Типы справочников (REFERENCE_TYPES)
- Маппинг полей (TABLE_FIELD_MAPPINGS)
- Функции для работы с конфигурацией

Использование:
    from models.references.references_config import (
        REFERENCE_TYPES,
        get_reference_config,
        get_table_config,
        get_russian_name,
        get_display_columns,
        get_search_columns
    )
"""

# ============================================================================
# ЧАСТЬ 1: МАППИНГ ПОЛЕЙ (латинские названия → русские заголовки)
# ============================================================================

# Общие поля для всех таблиц
COMMON_FIELDS = {
    'id': 'ID',
    'code': 'код',
    'name': 'наименование',
    'guid': 'GUID',
    'created_at': 'Дата_создания',
    'loaddate': 'Дата_загрузки',
    'startdate': 'Дата_начала',
    'enddate': 'Дата_окончания',
    'status': 'Статус',
    'level': 'уровень',
    'doc': 'документ',
    'ppocode': 'ОКТМО',
    'pponame': 'Наименование_МО',
    'budgetname': 'Бюджет',
    'stagename': 'Стадия',
    'year': 'Год',
}

# Базовые маппинги для групп таблиц (используются для ФУ/МО/merged версий)
BASE_FIELD_MAPPINGS = {
    # Группа: классификация доходов
    'income_classification': {
        'inctypecode': 'Тип_дохода',
        'incsubtypecode': 'Подтип_дохода',
        'analyticalgroupcode': 'Аналит_группа',
    },
    # Группа: расходы
    'costs': {
        'rzpr': 'РзПР',
        'kcsr': 'КЦСР',
        'kvr': 'КВР',
        'grbscode': 'Код_ГРБС',
        'id_code': 'ID_код',
    },
    # Группа: источники финансирования
    'sources': {
        'gaifcode': 'Код_ГАИФ',
    },
    # Группа: КЦСР
    'kcsr': {
        'dbkkcsr_id': 'Справочник_БК_КЦСР',
        'parentcode': 'Код_родителя',
        'id_code': 'ID_кода',
        'idparent': 'ID_родителя',
        'signsistem': 'Признак_системной_записи',
        'levelkcsr': 'Уровень_КЦСР',
    },
}


def _apply_base_mappings(base_key: str, table_patterns: list, additional: dict = None) -> dict:
    """
    Применяет базовый маппинг к группе таблиц.
    
    Args:
        base_key: Ключ в BASE_FIELD_MAPPINGS
        table_patterns: Список имён таблиц для применения
        additional: Дополнительные поля для конкретных таблиц
    
    Returns:
        Словарь {table_name: mapping}
    """
    result = {}
    base_mapping = BASE_FIELD_MAPPINGS.get(base_key, {})
    
    for table in table_patterns:
        result[table] = base_mapping.copy()
        if additional and table in additional:
            result[table].update(additional[table])
    
    return result


# Специфичные маппинги для уникальных таблиц
_UNIQUE_FIELD_MAPPINGS = {
    'oktmo': {
        'code': 'Код_ОКТМО',
        'name': 'Наименование_МО',
        'section': 'Раздел',
        'regioncode': 'Код_региона',
        'areacode': 'Код_района',
        'citycode': 'Код_города',
        'centrename': 'Центр',
        'clarification': 'Уточнение',
    },
    'ref_municipal_employees': {
        'oktmo_code': 'ОКТМО',
        'council_position': 'Совет_Должность',
        'council_surname': 'Совет_Фамилия',
        'council_first_name': 'Совет_Имя',
        'council_patronymic': 'Совет_Отчество',
        'council_address': 'Совет_Адрес',
        'council_email': 'Совет_Email',
        'administration_position': 'Администрация_Должность',
        'administration_surname': 'Администрация_Фамилия',
        'administration_first_name': 'Администрация_Имя',
        'administration_patronymic': 'Администрация_Отчество',
        'administration_address': 'Администрация_Адрес',
        'administration_email': 'Администрация_Email',
        'agreement_date': 'Дата_соглашения',
        'decision_date': 'Дата_решения',
        'decision_number': 'Номер_решения',
    },
    'npa': {
        'numdoc': 'Номер',
        'approvaldate': 'Дата',
        'kindname': 'Тип_документа',
    },
}

# Собираем TABLE_FIELD_MAPPINGS из групп
TABLE_FIELD_MAPPINGS = {}
TABLE_FIELD_MAPPINGS.update(_UNIQUE_FIELD_MAPPINGS)

# Доходы: ФУ, МО, объединенные
TABLE_FIELD_MAPPINGS.update(_apply_base_mappings(
    'income_classification',
    ['budgetclastypeinc', 'budgetclassubtypincmo'],
))
# Доходы объединенные с дополнительным полем
TABLE_FIELD_MAPPINGS['v_budgetclastypeinc_merged'] = {
    **BASE_FIELD_MAPPINGS['income_classification'],
    'concatenated_code': 'код',
}

# Расходы: ФУ, МО, объединенные
TABLE_FIELD_MAPPINGS.update(_apply_base_mappings(
    'costs',
    ['budgetclascosts', 'budgetclascostsmo', 'v_budgetclascosts_merged'],
))

# Источники: ФУ, МО, объединенные
TABLE_FIELD_MAPPINGS.update(_apply_base_mappings(
    'sources',
    ['budgetclassources', 'budgetclassourcesmo', 'v_budgetclassources_merged'],
))

# КЦСР: ФУ, МО, объединенные
TABLE_FIELD_MAPPINGS.update(_apply_base_mappings(
    'kcsr',
    ['budgetclaskcsr', 'budgetclaskcsrmo', 'v_budgetclaskcsr_merged'],
))

# Распорядители МО с дополнительным полем
TABLE_FIELD_MAPPINGS['budgetclasgrbsmo'] = {'codereestr': 'Код_реестра'}


# Паттерны для колонок поиска
_SEARCH_PATTERNS = {
    'default': ['code', 'name'],
    'with_ppo': ['code', 'name', 'pponame'],
    'with_level_ppo': ['code', 'name', 'level', 'pponame'],
}

# Специфичные колонки поиска для уникальных таблиц
_UNIQUE_SEARCH_COLUMNS = {
    'oktmo': ['code', 'name', 'centrename'],
    'ref_municipal_employees': ['oktmo_code', 'council_surname', 'council_first_name', 
                                'administration_surname', 'administration_first_name'],
    'v_budgetclastypeinc_merged': ['inctypecode', 'incsubtypecode', 'analyticalgroupcode', 
                                   'name', 'level', 'concatenated_code'],
    'budgetclastypeinc': ['inctypecode', 'incsubtypecode', 'analyticalgroupcode', 
                         'name', 'level', 'pponame'],
    'budgetclassubtypincmo': ['inctypecode', 'incsubtypecode', 'analyticalgroupcode', 
                             'name', 'level', 'pponame'],
    'source_reference_records': ['code', 'name', 'level', 'doc'],
    'budgetclascosts': ['name', 'rzpr', 'kcsr', 'kvr', 'grbscode', 'pponame'],
    'budgetclascostsmo': ['name', 'rzpr', 'kcsr', 'kvr', 'grbscode', 'pponame'],
    'v_budgetclascosts_merged': ['name', 'rzpr', 'kcsr', 'kvr', 'grbscode', 'pponame'],
    'budgetclasgrbsmo': ['code', 'name', 'pponame', 'codereestr'],
    'budgetclassources': ['code', 'name', 'level', 'gaifcode', 'pponame'],
    'budgetclassourcesmo': ['code', 'name', 'level', 'gaifcode', 'pponame'],
    'v_budgetclassources_merged': ['code', 'name', 'level', 'gaifcode', 'pponame'],
    'budgetclaskcsr': ['code', 'name', 'level', 'parentcode', 'levelkcsr', 'pponame'],
    'budgetclaskcsrmo': ['code', 'name', 'level', 'parentcode', 'levelkcsr', 'pponame'],
    'v_budgetclaskcsr_merged': ['code', 'name', 'level', 'parentcode', 'levelkcsr', 'pponame'],
    'npa': ['numdoc', 'name', 'kindname'],
}

# Таблицы с паттерном 'with_ppo'
_TABLES_WITH_PPO = [
    'budgetclasgrbs', 'v_budgetclasgrbs_merged',
    'budgetclasgabs', 'budgetclasgabsmo', 'v_budgetclasgabs_merged',
    'budgetclasgaiffb', 'budgetclasgaifmo', 'v_budgetclasgaif_merged',
]

# Таблицы с паттерном 'with_level_ppo'
_TABLES_WITH_LEVEL_PPO = [
    'budgetclaskvr', 'budgetclasrzpr',
]

# Собираем SEARCH_COLUMNS
SEARCH_COLUMNS = _UNIQUE_SEARCH_COLUMNS.copy()

for table in _TABLES_WITH_PPO:
    SEARCH_COLUMNS[table] = _SEARCH_PATTERNS['with_ppo']

for table in _TABLES_WITH_LEVEL_PPO:
    SEARCH_COLUMNS[table] = _SEARCH_PATTERNS['with_level_ppo']


# ============================================================================
# ЧАСТЬ 2: ТИПЫ СПРАВОЧНИКОВ (сгруппированы по категориям)
# ============================================================================

# Загружаемые справочники (из Excel)
LOADABLE_REFERENCES = {
    'Коды доходов': {
        'table': 'v_budgetclastypeinc_merged',
        'load_method': 'load_income_sources_reference',
        'load_type': 'доходы',
        'is_view': True,
        'columns': ['concatenated_code', 'name', 'level', 'inctypecode', 'incsubtypecode', 'analyticalgroupcode'],
    },
    'Коды источников': {
        'table': 'source_reference_records',
        'load_method': 'load_income_sources_reference',
        'load_type': 'источники',
        'columns': ['code', 'name', 'level', 'doc'],
    },
}

# Конфигурационные справочники (редактируются в системе)
CONFIG_REFERENCES = {
    'Годы': {
        'table': 'ref_years',
        'columns': ['year', 'is_active'],
        'is_config': True,
        'load_func': '_load_years',
        'save_func': '_save_years',
    },
    'Типы форм': {
        'table': 'ref_form_types',
        'columns': ['id', 'code', 'name', 'periodicity', 'is_active'],
        'is_config': True,
        'load_func': '_load_forms',
        'save_func': '_save_forms',
    },
    'Периоды': {
        'table': 'ref_periods',
        'columns': ['id', 'code', 'name', 'sort_order', 'form_type_code', 'is_active'],
        'is_config': True,
        'load_func': '_load_periods',
        'save_func': '_save_periods',
    },
    'Сотрудники МО': {
        'table': 'ref_municipal_employees',
        'columns': ['id', 'oktmo_code', 'startdate', 'enddate', 'council_position', 'council_surname', 
                   'council_first_name', 'council_patronymic', 'council_address', 'council_email',
                   'administration_position', 'administration_surname', 'administration_first_name', 
                   'administration_patronymic', 'administration_address', 'administration_email',
                   'agreement_date', 'decision_date', 'decision_number'],
        'editable': True,
        'dialog_class': 'MunicipalEmployeeDialog',
    },
}

# Онлайн справочники (из бюджетной системы)
ONLINE_REFERENCES = {
    'ОКТМО': {'table': 'oktmo'},
    'Классификаторы доходов бюджета ФУ': {'table': 'budgetclastypeinc', 'has_npa': True},
    'Классификаторы доходов бюджета МО': {'table': 'budgetclassubtypincmo', 'has_npa': True},
    'Администраторы бюджета ФУ': {'table': 'budgetclasgabs', 'has_npa': True},
    'Администраторы бюджета МО': {'table': 'budgetclasgabsmo', 'has_npa': True},
    'Распорядители бюджета ФУ': {'table': 'budgetclasgrbs', 'has_npa': True},
    'Распорядители бюджета МО': {'table': 'budgetclasgrbsmo', 'has_npa': True},
    'Классификаторы расходов бюджета ФУ': {'table': 'budgetclascosts', 'has_npa': True},
    'Классификаторы расходов бюджета МО': {'table': 'budgetclascostsmo', 'has_npa': True},
    'Источники финансирования дефицита ФУ': {'table': 'budgetclasgaiffb', 'has_npa': True},
    'Источники финансирования дефицита МО': {'table': 'budgetclasgaifmo', 'has_npa': True},
    'Классификаторы источников финансирования ФУ': {'table': 'budgetclassources', 'has_npa': True},
    'Классификаторы источников финансирования МО': {'table': 'budgetclassourcesmo', 'has_npa': True},
    'Коды видов расходов (КВР)': {'table': 'budgetclaskvr', 'has_npa': True},
    'Коды разделов и подразделов (РЗПР)': {'table': 'budgetclasrzpr', 'has_npa': True},
    'Коды целевых статей расходов ФУ (КЦСР)': {'table': 'budgetclaskcsr', 'has_npa': True},
    'Коды целевых статей расходов МО (КЦСР)': {'table': 'budgetclaskcsrmo', 'has_npa': True},
}

# Объединенные представления (VIEW)
VIEW_REFERENCES = {
    'Классификаторы доходов (объединенные)': {'table': 'v_budgetclastypeinc_merged', 'has_npa': True},
    'Классификаторы расходов (объединенные)': {'table': 'v_budgetclascosts_merged', 'has_npa': True},
    'Распорядители бюджета (объединенные)': {'table': 'v_budgetclasgrbs_merged', 'has_npa': True},
    'Администраторы бюджета (объединенные)': {'table': 'v_budgetclasgabs_merged', 'has_npa': True},
    'Источники финансирования дефицита (объединенные)': {'table': 'v_budgetclasgaiffb_merged', 'has_npa': True},
    'Классификаторы источников финансирования (объединенные)': {'table': 'v_budgetclassources_merged', 'has_npa': True},
    'Коды целевых статей расходов (объединенные)': {'table': 'v_budgetclaskcsr_merged', 'has_npa': True},
}


def _normalize_reference_config(config: dict, defaults: dict) -> dict:
    """Добавляет значения по умолчанию в конфигурацию справочника"""
    normalized = defaults.copy()
    normalized.update(config)
    return normalized


def _build_reference_types() -> dict:
    """
    Собирает REFERENCE_TYPES из сгруппированных справочников.
    Добавляет значения по умолчанию для каждой группы.
    """
    result = {}
    
    # Загружаемые справочники
    loadable_defaults = {
        'load_method': None,
        'columns': None,
        'display_columns': None,
        'search_columns': None,
        'has_npa': False,
        'is_view': False,
    }
    for name, config in LOADABLE_REFERENCES.items():
        result[name] = _normalize_reference_config(config, loadable_defaults)
    
    # Конфигурационные справочники
    config_defaults = {
        'load_method': None,
        'display_columns': None,
        'search_columns': None,
        'has_npa': False,
        'is_config': False,
        'editable': False,
    }
    for name, config in CONFIG_REFERENCES.items():
        result[name] = _normalize_reference_config(config, config_defaults)
    
    # Онлайн справочники
    result['─── Онлайн справочники ───'] = {'table': None, 'is_separator': True}
    
    online_defaults = {
        'load_method': None,
        'columns': None,
        'display_columns': None,
        'search_columns': None,
        'is_online': True,
        'has_npa': False,
    }
    for name, config in ONLINE_REFERENCES.items():
        result[name] = _normalize_reference_config(config, online_defaults)
    
    # Объединенные представления
    result['─── Объединенные представления ───'] = {'table': None, 'is_separator': True}
    
    view_defaults = {
        'load_method': None,
        'columns': None,
        'display_columns': None,
        'search_columns': None,
        'is_view': True,
        'has_npa': False,
    }
    for name, config in VIEW_REFERENCES.items():
        result[name] = _normalize_reference_config(config, view_defaults)
    
    return result


# Собираем единый словарь REFERENCE_TYPES
REFERENCE_TYPES = _build_reference_types()


# ============================================================================
# ЧАСТЬ 3: ФУНКЦИИ ДЛЯ РАБОТЫ С КОНФИГУРАЦИЕЙ
# ============================================================================

def get_display_columns(table_name: str, columns: list) -> list:
    mapping = TABLE_FIELD_MAPPINGS.get(table_name, {})
    result = []
    
    for col in columns:
        if col in mapping:
            result.append(f"{col} AS {mapping[col]}")
        elif col in COMMON_FIELDS:
            result.append(f"{col} AS {COMMON_FIELDS[col]}")
        else:
            result.append(col)
    
    return result


def get_russian_name(table_name: str, field_name: str) -> str:
    mapping = TABLE_FIELD_MAPPINGS.get(table_name, {})
    
    if field_name in mapping:
        return mapping[field_name]
    elif field_name in COMMON_FIELDS:
        return COMMON_FIELDS[field_name]
    else:
        return field_name


def get_search_columns(table_name: str) -> list:
    return SEARCH_COLUMNS.get(table_name, ['code', 'name'])


def get_reference_types():
    return REFERENCE_TYPES


def get_reference_config(reference_name: str) -> dict:
    return REFERENCE_TYPES.get(reference_name)


def get_table_config(table_name: str) -> dict:
    for ref_name, config in REFERENCE_TYPES.items():
        if config.get('table') == table_name:
            return config
    return None

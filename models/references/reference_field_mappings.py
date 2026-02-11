"""
Маппинг полей справочных таблиц: латинские названия → русские заголовки

Этот модуль централизует конфигурацию отображения полей для всех справочных таблиц.
Вместо дублирования маппинга в каждом месте кода, все названия хранятся здесь.

Использование:
    1. Добавьте таблицу в TABLE_FIELD_MAPPINGS с маппингом полей
    2. В REFERENCE_TYPES установите display_columns=None и search_columns=None
    3. Модуль автоматически сгенерирует SQL выражения с алиасами

Пример:
    TABLE_FIELD_MAPPINGS = {
        'my_table': {
            'field_name': 'Русское_название',
            'other_field': 'Другое_поле'
        }
    }
    
    # В references_management_dialog.py:
    'Мой справочник': {
        'table': 'my_table',
        'columns': ['field_name', 'other_field'],
        'display_columns': None,  # Генерируется автоматически
        'search_columns': None,   # Генерируется автоматически
    }

Функции:
    - get_display_columns(table, cols) → список "field AS Русское_название"
    - get_russian_name(table, field) → русское название поля
    - get_search_columns(table) → список полей для поиска
"""

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

# Специфичные маппинги для каждой таблицы (только уникальные поля, не входящие в COMMON_FIELDS)
TABLE_FIELD_MAPPINGS = {
    # Таблица ОКТМО
    'oktmo': {
        'code': 'Код_ОКТМО',  # Переопределяем для более точного названия
        'name': 'Наименование_МО',  # Переопределяем для более точного названия
        'section': 'Раздел',
        'regioncode': 'Код_региона',
        'areacode': 'Код_района',
        'citycode': 'Код_города',
        'centrename': 'Центр',
        'clarification': 'Уточнение',
    },
    
    # Сотрудники МО
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
    
    # Классификация доходов (объединенная)
    'v_budgetclastypeinc_merged': {
        'inctypecode': 'Тип_дохода',
        'incsubtypecode': 'Подтип_дохода',
        'analyticalgroupcode': 'Аналит_группа',
        'concatenated_code': 'код',
    },
    
    # Доходы ФУ и МО
    'budgetclastypeinc': {
        'inctypecode': 'Тип_дохода',
        'incsubtypecode': 'Подтип_дохода',
        'analyticalgroupcode': 'Аналит_группа',
    },
    
    'budgetclassubtypincmo': {
        'inctypecode': 'Тип_дохода',
        'incsubtypecode': 'Подтип_дохода',
        'analyticalgroupcode': 'Аналит_группа',
    },
    
    # Классификация источников
    'source_reference_records': {
        # Все поля из COMMON_FIELDS
    },
    
    # Расходы ФУ, МО и объединенные
    'budgetclascosts': {
        'rzpr': 'РзПР',
        'kcsr': 'КЦСР',
        'kvr': 'КВР',
        'grbscode': 'Код_ГРБС',
        'id_code': 'ID_код',
    },
    
    'budgetclascostsmo': {
        'rzpr': 'РзПР',
        'kcsr': 'КЦСР',
        'kvr': 'КВР',
        'grbscode': 'Код_ГРБС',
        'id_code': 'ID_код',
    },
    
    'v_budgetclascosts_merged': {
        'rzpr': 'РзПР',
        'kcsr': 'КЦСР',
        'kvr': 'КВР',
        'grbscode': 'Код_ГРБС',
        'id_code': 'ID_код',
    },
    
    # Распорядители ФУ, МО и объединенные
    'budgetclasgrbs': {
        # Все общие поля
    },
    
    'budgetclasgrbsmo': {
        'codereestr': 'Код_реестра',
    },
    
    'v_budgetclasgrbs_merged': {
        # Все общие поля
    },
    
    # Администраторы ФУ, МО и объединенные
    'budgetclasgabs': {
        # Все общие поля
    },
    
    'budgetclasgabsmo': {
        # Все общие поля
    },
    
    'v_budgetclasgabs_merged': {
        # Все общие поля
    },
    
    # Источники дефицита ФУ, МО и объединенные
    'budgetclasgaiffb': {
        # Все общие поля
    },
    
    'budgetclasgaifmo': {
        # Все общие поля
    },
    
    'v_budgetclasgaif_merged': {
        # Все общие поля
    },
    
    # Классификаторы источников финансирования ФУ, МО и объединенные
    'budgetclassources': {
        'gaifcode': 'Код_ГАИФ',
    },
    
    'budgetclassourcesmo': {
        'gaifcode': 'Код_ГАИФ',
    },
    
    'v_budgetclassources_merged': {
        'gaifcode': 'Код_ГАИФ',
    },
    
    # НПА (нормативно-правовые акты)
    'npa': {
        'numdoc': 'Номер',
        'approvaldate': 'Дата',
        'kindname': 'Тип_документа',
        # name из COMMON_FIELDS
    },
}


def get_display_columns(table_name: str, columns: list) -> list:
    """
    Генерирует SQL выражения с алиасами для отображения
    
    Args:
        table_name: Название таблицы
        columns: Список полей для отображения
    
    Returns:
        Список строк вида "field AS Русское_название"
    """
    mapping = TABLE_FIELD_MAPPINGS.get(table_name, {})
    result = []
    
    for col in columns:
        if col in mapping:
            result.append(f"{col} AS {mapping[col]}")
        elif col in COMMON_FIELDS:
            result.append(f"{col} AS {COMMON_FIELDS[col]}")
        else:
            # Если маппинга нет, оставляем как есть
            result.append(col)
    
    return result


def get_russian_name(table_name: str, field_name: str) -> str:
    """
    Получить русское название для поля
    
    Args:
        table_name: Название таблицы
        field_name: Название поля
    
    Returns:
        Русское название или исходное поле
    """
    mapping = TABLE_FIELD_MAPPINGS.get(table_name, {})
    
    if field_name in mapping:
        return mapping[field_name]
    elif field_name in COMMON_FIELDS:
        return COMMON_FIELDS[field_name]
    else:
        return field_name


def get_search_columns(table_name: str) -> list:
    """
    Получить список полей для поиска по таблице
    
    Args:
        table_name: Название таблицы
    
    Returns:
        Список полей для поиска
    """
    # Определяем ключевые поля для поиска по каждой таблице
    search_mappings = {
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
        'budgetclasgrbs': ['code', 'name', 'pponame'],
        'budgetclasgrbsmo': ['code', 'name', 'pponame', 'codereestr'],
        'v_budgetclasgrbs_merged': ['code', 'name', 'pponame'],
        'budgetclasgabs': ['code', 'name', 'pponame'],
        'budgetclasgabsmo': ['code', 'name', 'pponame'],
        'v_budgetclasgabs_merged': ['code', 'name', 'pponame'],
        'budgetclasgaiffb': ['code', 'name', 'pponame'],
        'budgetclasgaifmo': ['code', 'name', 'pponame'],
        'v_budgetclasgaif_merged': ['code', 'name', 'pponame'],
        'budgetclassources': ['code', 'name', 'level', 'gaifcode', 'pponame'],
        'budgetclassourcesmo': ['code', 'name', 'level', 'gaifcode', 'pponame'],
        'v_budgetclassources_merged': ['code', 'name', 'level', 'gaifcode', 'pponame'],
        'npa': ['numdoc', 'name', 'kindname'],
    }
    
    return search_mappings.get(table_name, ['code', 'name'])

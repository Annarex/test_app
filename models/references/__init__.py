"""
Модуль для работы с бюджетными справочниками из API бюджетной системы.

Единая конфигурация справочников:
- references_config: объединенная конфигурация типов справочников и маппинга полей
"""

from .references_config import (
    REFERENCE_TYPES,
    get_reference_types,
    get_reference_config,
    get_table_config,
    get_display_columns,
    get_russian_name,
    get_search_columns,
    TABLE_FIELD_MAPPINGS,
    COMMON_FIELDS,
    SEARCH_COLUMNS
)

__all__ = [
    'REFERENCE_TYPES',
    'get_reference_types',
    'get_reference_config',
    'get_table_config',
    'get_display_columns',
    'get_russian_name',
    'get_search_columns',
    'TABLE_FIELD_MAPPINGS',
    'COMMON_FIELDS',
    'SEARCH_COLUMNS'
]

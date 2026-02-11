"""
Модуль для работы с бюджетными справочниками из API бюджетной системы.

Включает:
- reference_field_mappings: маппинг латинских названий полей на русские заголовки
"""

from .reference_field_mappings import (
    get_display_columns,
    get_russian_name,
    get_search_columns,
    TABLE_FIELD_MAPPINGS,
    COMMON_FIELDS
)

__all__ = [
    'get_display_columns',
    'get_russian_name',
    'get_search_columns',
    'TABLE_FIELD_MAPPINGS',
    'COMMON_FIELDS'
]

"""
Пакет text_diff_tool для сравнения текстов и создания отчетов об ошибках.

Основные модули:
- text_comparator: сравнение строк и поиск различий
- text_highlighter: создание подсвеченного текста для Excel
- error_finder: поиск ошибок в данных
- excel_exporter: экспорт результатов в Excel
"""

from .text_comparator import find_differences
from .text_highlighter import create_highlighted_text, BLACK_FONT, RED_FONT, GREEN_FONT
from .error_finder import find_errors, find_best_match, ErrorInfo, CodeErrorInfo
from .excel_exporter import create_error_report, save_workbook
from .code_processor import normalize_classification_code, compare_codes

__version__ = "1.0.0"
__all__ = [
    "find_differences",
    "create_highlighted_text",
    "BLACK_FONT",
    "RED_FONT",
    "GREEN_FONT",
    "find_errors",
    "find_best_match",
    "ErrorInfo",
    "CodeErrorInfo",
    "create_error_report",
    "save_workbook",
    "normalize_classification_code",
    "compare_codes",
]


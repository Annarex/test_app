"""Конфигуратор заголовков дерева"""
from models.constants.form_0503317_constants import Form0503317Constants


class TreeHeaderConfigurator:
    """Класс для конфигурации заголовков дерева"""
    
    def configure_headers(self, section_name: str) -> dict:
        """Конфигурация заголовков дерева под выбранный раздел
        
        Args:
            section_name: Название раздела ("Доходы", "Расходы", и т.д.)
        
        Returns:
            Словарь с ключами:
            - headers: список заголовков для отображения
            - tooltips: список подсказок для заголовков
            - mapping: словарь с метаданными о колонках
        """
        base_headers = ["Наименование", "Код строки", "Код классификации", "Уровень"]
        display_headers = base_headers[:]
        tooltip_headers = base_headers[:]
        mapping = {
            "type": "base",
            "base_count": len(base_headers)
        }

        if section_name in ["Доходы", "Расходы", "Источники финансирования"]:
            mapping.update(self._build_budget_headers(display_headers, tooltip_headers))
        elif section_name == "Консолидируемые расчеты":
            mapping.update(self._build_consolidated_headers(display_headers, tooltip_headers))

        return {
            "headers": display_headers,
            "tooltips": tooltip_headers,
            "mapping": mapping
        }
    
    def _build_budget_headers(self, display_headers: list, tooltip_headers: list) -> dict:
        """Построение заголовков для бюджетных разделов
        
        Args:
            display_headers: Список заголовков для отображения (изменяется in-place)
            tooltip_headers: Список подсказок (изменяется in-place)
        
        Returns:
            Словарь с метаданными для бюджетных разделов
        """
        budget_cols = Form0503317Constants.BUDGET_COLUMNS
        mapping = {
            "type": "budget",
            "budget_columns": budget_cols,
            "approved_start": len(display_headers),
            "executed_start": len(display_headers) + len(budget_cols)
        }

        for col in budget_cols:
            display_headers.append(f"У. {col}")
            tooltip_headers.append(f"Утвержденный — {col}")
        for col in budget_cols:
            display_headers.append(f"И. {col}")
            tooltip_headers.append(f"Исполненный — {col}")

        return mapping
    
    def _build_consolidated_headers(self, display_headers: list, tooltip_headers: list) -> dict:
        """Построение заголовков для консолидированных расчетов
        
        Args:
            display_headers: Список заголовков для отображения (изменяется in-place)
            tooltip_headers: Список подсказок (изменяется in-place)
        
        Returns:
            Словарь с метаданными для консолидированных расчетов
        """
        cons_cols = Form0503317Constants.CONSOLIDATED_COLUMNS
        mapping = {
            "type": "consolidated",
            "value_start": len(display_headers),
            "columns": cons_cols
        }
        for col in cons_cols:
            display_headers.append(col)
            tooltip_headers.append(col)

        return mapping

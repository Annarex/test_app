"""
Модуль для работы с подсветкой текста в Excel (rich text formatting).
Содержит функции для создания подсвеченного текста с ошибками и исправлениями.
"""
from openpyxl.cell.text import InlineFont
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.xml.functions import Element, whitespace as original_whitespace
import openpyxl.xml.functions as xml_funcs

# Monkey patch для TextBlock.to_tree - устанавливаем xml:space="preserve" для пробелов
_original_textblock_to_tree = TextBlock.to_tree
def textblock_to_tree_preserve(self):
    """Патч для TextBlock.to_tree с принудительным xml:space="preserve" для пробелов"""
    el = Element("r")
    el.append(self.font.to_tree(tagname="rPr"))
    t = Element("t")
    t.text = self.text
    # Всегда устанавливаем xml:space="preserve" если есть пробелы
    if self.text and (' ' in self.text or '\t' in self.text or '\n' in self.text):
        t.set("{%s}space" % xml_funcs.XML_NS, "preserve")
    else:
        # Вызываем оригинальную whitespace для других случаев
        original_whitespace(t)
    el.append(t)
    return el
TextBlock.to_tree = textblock_to_tree_preserve


# Определение шрифтов для подсветки
BLACK_FONT = InlineFont(color="000000", sz=11)
RED_FONT = InlineFont(color="FF0000", b=True, u="single", sz=11)   # Красный с подчеркиванием для ошибок
GREEN_FONT = InlineFont(color="00AA00", b=True, u="single", sz=11)  # Зеленый с подчеркиванием для правильного текста


def create_highlighted_text(
    orig: str,
    diff_idx: list[int],
    corrections: list[tuple[int, int, str, str]]
) -> CellRichText:
    """
    Создает подсвеченный текст с ошибками и исправлениями.
    
    Args:
        orig: Исходная строка
        diff_idx: Список индексов позиций с ошибками
        corrections: Список кортежей (start_pos, end_pos, correct_text, kind),
                     где kind ∈ {"replace", "delete", "insert"}
    
    Returns:
        CellRichText объект с подсветкой
    """
    parts = []
    
    if not orig:
        return CellRichText([])
    
    # Разделяем исправления на:
    # - corrections_map: для replace/delete (привязаны к диапазону ошибок)
    # - insert_map: для insert (отсутствует в orig, есть только в ref)
    corrections_map: dict[int, str] = {}
    insert_map: dict[int, list[str]] = {}
    for start, end, correct_text, kind in corrections:
        if kind in ("replace", "delete"):
            # Привязываем правильный текст к началу ошибочного диапазона
            corrections_map[start] = correct_text
        elif kind == "insert":
            # Вставка: позиция между символами 0..len(orig)
            if start not in insert_map:
                insert_map[start] = []
            insert_map[start].append(correct_text)

    # Сначала обрабатываем все insert в начале строки (позиция 0)
    if 0 in insert_map:
        for txt in insert_map[0]:
            parts.append(TextBlock(GREEN_FONT, txt))
    
    i = 0
    while i < len(orig):
        if i in diff_idx:
            # Нашли ошибку (replace/delete): показываем правильный текст (зеленым),
            # затем неправильный фрагмент (красным).
            error_start = i
            while i < len(orig) and i in diff_idx:
                i += 1
            error_end = i
            error_text = orig[error_start:error_end]

            # Правильный текст для этой ошибки (если есть)
            if error_start in corrections_map:
                correct_text = corrections_map[error_start]
                if correct_text:
                    parts.append(TextBlock(GREEN_FONT, correct_text))

            # Неправильный текст подсвечиваем красным
            parts.append(TextBlock(RED_FONT, error_text))
        else:
            # Правильный текст - черным
            start_ok = i
            while i < len(orig) and i not in diff_idx:
                # Проверяем, есть ли insert после текущего символа
                next_pos = i + 1
                if next_pos in insert_map:
                    # Добавляем текущий символ
                    if start_ok <= i:
                        text_ok = orig[start_ok:i+1]
                        if text_ok:
                            parts.append(TextBlock(BLACK_FONT, text_ok))
                    # Добавляем insert после этого символа
                    for txt in insert_map[next_pos]:
                        parts.append(TextBlock(GREEN_FONT, txt))
                    start_ok = i + 1
                i += 1
            text_ok = orig[start_ok:i]
            if text_ok:
                parts.append(TextBlock(BLACK_FONT, text_ok))
    
    # Обрабатываем вставки в самом конце строки (позиция len(orig))
    if len(orig) in insert_map:
        for txt in insert_map[len(orig)]:
            parts.append(TextBlock(GREEN_FONT, txt))

    return CellRichText(parts)
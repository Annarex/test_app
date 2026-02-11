"""
Модуль для сравнения строк и поиска различий.
Содержит функции для нахождения оптимального выравнивания и различий между строками.
"""
import difflib


def find_differences(orig: str, ref: str) -> tuple[list[int], list[tuple[int, int, str, str]]]:
    """
    Находит индексы реальных различий между строками, используя оптимальное выравнивание.
    
    Args:
        orig: Исходная строка
        ref: Эталонная строка для сравнения
    
    Returns:
        Кортеж из:
        - список индексов позиций в orig, которые являются ошибками
        - список кортежей (start_pos, end_pos, correct_text, kind),
          где kind ∈ {"replace", "delete", "insert"}
    
    Примеры:
        - если в orig есть лишнее "на" в начале, подсветит только "на",
          а не весь последующий текст как неправильный;
        - если наоборот "на" есть только в ref, подсветится ближайшая позиция в orig,
          чтобы строка всё равно считалась ошибочной.
    """
    if not orig:
        return [], []
    if not ref:
        return list(range(len(orig))), []

    # Используем SequenceMatcher для оптимального выравнивания
    matcher = difflib.SequenceMatcher(None, orig, ref, autojunk=False)

    # Собираем индексы позиций в orig, которые являются различиями
    diff_positions: set[int] = set()
    # Собираем информацию о правильном тексте для каждой ошибки
    corrections: list[tuple[int, int, str, str]] = []  # (start_pos, end_pos, correct_text, kind)

    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "replace":
            # Замена: подсвечиваем все символы в orig, которые заменены
            diff_positions.update(range(i1, i2))
            # Сохраняем правильный текст из ref
            correct_text = ref[j1:j2]
            if correct_text:
                corrections.append((i1, i2, correct_text, "replace"))
        elif tag == "delete":
            # Удаление: подсвечиваем удаленные символы в orig
            diff_positions.update(range(i1, i2))
            # В ref на этом месте ничего нет (или есть что-то другое), но показываем пустоту
            corrections.append((i1, i2, "", "delete"))
        elif tag == "insert":
            # Вставка символов только в ref.
            # Для отсутствия слова мы хотим показать только зеленый текст без красной подсветки,
            # поэтому НЕ добавляем позицию в diff_positions.
            # i1 указывает на позицию в orig, где должна быть вставка (между символами)
            # Позиция может быть от 0 (перед первым символом) до len(orig) (после последнего)
            insert_pos = i1  # i1 уже корректная позиция для вставки
            # Показываем, что должно быть вставлено
            correct_text = ref[j1:j2]
            if correct_text:
                corrections.append((insert_pos, insert_pos, correct_text, "insert"))
        # tag == "equal" - совпадающие части не подсвечиваем

    return sorted(diff_positions), corrections


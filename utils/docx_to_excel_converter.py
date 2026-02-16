"""
Модуль для конвертации документов DOCX в Excel.

Модуль извлекает текст и таблицы из документов Word (docx)
и экспортирует их в формат Excel (xlsx).
"""
import os
from docx import Document
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from typing import Optional


class DocxToExcelConverter:
    """Конвертер документов DOCX в Excel."""
    
    def __init__(self, remove_empty_rows: bool = False, max_merge_rows: int = 5, max_merge_cols: int = 20):
        """
        Инициализация конвертера.
        
        Args:
            remove_empty_rows: Если True, не добавлять пустые строки между таблицами
            max_merge_rows: Максимальное количество строк для проверки объединений (по умолчанию 5)
            max_merge_cols: Максимальное количество столбцов для проверки объединений (по умолчанию 20)
        """
        self.workbook = None
        self.current_row = 1
        self.remove_empty_rows = remove_empty_rows
        self.max_merge_rows = max_merge_rows
        self.max_merge_cols = max_merge_cols
        self.last_element_type = None  # Отслеживание последнего добавленного элемента ('paragraph', 'table', None)
        
    def convert(self, docx_path: str, excel_path: str) -> bool:
        """
        Конвертирует docx файл в excel.
        
        Args:
            docx_path: Путь к исходному docx файлу
            excel_path: Путь для сохранения excel файла
            
        Returns:
            True если конвертация успешна, False в противном случае
        """
        try:
            # Проверка существования файла
            if not os.path.exists(docx_path):
                raise FileNotFoundError(f"Файл не найден: {docx_path}")
            
            # Загрузка документа Word
            doc = Document(docx_path)
            
            # Создание новой книги Excel
            self.workbook = Workbook()
            sheet = self.workbook.active
            sheet.title = "Содержимое документа"
            
            # Сброс счетчика строк и отслеживания элементов
            self.current_row = 1
            self.last_element_type = None
            
            # Обработка параграфов и таблиц
            for element in doc.element.body:
                # Обработка параграфов
                if element.tag.endswith('p'):
                    paragraph = None
                    for para in doc.paragraphs:
                        if para._element == element:
                            paragraph = para
                            break
                    
                    if paragraph and paragraph.text.strip():
                        self._add_paragraph_to_sheet(sheet, paragraph)
                
                # Обработка таблиц
                elif element.tag.endswith('tbl'):
                    table = None
                    for tbl in doc.tables:
                        if tbl._element == element:
                            table = tbl
                            break
                    
                    if table:
                        self._add_table_to_sheet(sheet, table)
            
            # Автоматическая настройка ширины столбцов
            self._adjust_column_widths(sheet)
            
            # Сохранение файла Excel
            self.workbook.save(excel_path)
            
            return True
            
        except Exception as e:
            raise Exception(f"Ошибка при конвертации: {str(e)}")
    
    def _add_paragraph_to_sheet(self, sheet, paragraph):
        """
        Добавляет параграф в лист Excel.
        
        Args:
            sheet: Активный лист Excel
            paragraph: Параграф из документа Word
        """
        # Если параграф пустой и включена опция удаления пустых строк, пропускаем
        if self.remove_empty_rows and not paragraph.text.strip():
            return
        
        cell = sheet.cell(row=self.current_row, column=1)
        cell.value = paragraph.text
        
        # Стилизация для заголовков (жирный текст)
        if paragraph.runs and paragraph.runs[0].bold:
            cell.font = Font(bold=True, size=12)
            cell.fill = PatternFill(start_color="E7E6E6", end_color="E7E6E6", fill_type="solid")
        else:
            cell.font = Font(size=11)
        
        cell.alignment = Alignment(wrap_text=True, vertical='top')
        
        self.current_row += 1
        self.last_element_type = 'paragraph'  # Запоминаем, что добавили параграф
    
    def _parse_cell_value(self, text: str):
        """
        Пытается преобразовать текст в числовое значение.
        
        Args:
            text: Текст из ячейки
            
        Returns:
            Число (int или float) если преобразование успешно, иначе исходный текст
        """
        if not text:
            return text
        
        # Убираем пробелы
        text_clean = text.strip()
        if not text_clean:
            return text
        
        # Проверка на ведущие нули (коды, номера) - не конвертируем
        # Например: "007", "0123" должны остаться текстом
        if len(text_clean) > 1 and text_clean[0] == '0' and text_clean[1].isdigit():
            return text
        
        # Заменяем запятую на точку для русской локали
        text_normalized = text_clean.replace(',', '.').replace(' ', '')
        
        try:
            # Пытаемся преобразовать в число
            if '.' in text_normalized:
                # Вещественное число
                return float(text_normalized)
            else:
                # Целое число
                return int(text_normalized)
        except (ValueError, AttributeError):
            # Если не число, возвращаем исходный текст
            return text
    
    def _add_table_to_sheet(self, sheet, table):
        """
        Добавляет таблицу из Word в лист Excel с поддержкой вертикального объединения ячеек.
        
        Args:
            sheet: Активный лист Excel
            table: Таблица из документа Word
        """
        # Добавляем пустую строку перед таблицей
        # Всегда добавляем, если перед таблицей был текст (чтобы не склеивать)
        # Или если опция remove_empty_rows выключена
        if self.last_element_type == 'paragraph' or not self.remove_empty_rows:
            self.current_row += 1
        
        start_row = self.current_row
        
        # Стили для границ
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        # Копируем таблицу построчно, пропуская полностью пустые строки если опция включена
        actual_row_idx = 0
        for row_idx, row in enumerate(table.rows):
            # Проверяем, пустая ли строка (все ячейки пустые)
            if self.remove_empty_rows:
                is_empty_row = all(not cell.text.strip() for cell in row.cells)
                if is_empty_row:
                    continue  # Пропускаем пустую строку
            
            for col_idx, cell in enumerate(row.cells):
                excel_row = start_row + actual_row_idx
                excel_col = col_idx + 1
                
                try:
                    excel_cell = sheet.cell(row=excel_row, column=excel_col)
                    # Пытаемся преобразовать в число, иначе оставляем текст
                    excel_cell.value = self._parse_cell_value(cell.text.strip())
                    excel_cell.border = thin_border
                    excel_cell.alignment = Alignment(wrap_text=True, vertical='top', horizontal='left')
                    
                    # Стилизация для заголовков (первая строка)
                    if actual_row_idx == 0:
                        excel_cell.font = Font(bold=True, size=11)
                        excel_cell.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
                        excel_cell.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center')
                    else:
                        excel_cell.font = Font(size=10)
                except Exception as e:
                    pass
            
            actual_row_idx += 1
        
        # Ищем вертикальные объединения используя простой метод - сравнивая _tc элементы
        num_rows = len(table.rows)
        num_cols = len(table.columns) if table.rows else 0
        
        # Ограничиваем проверку объединений для ускорения
        check_rows = min(num_rows, self.max_merge_rows)
        check_cols = min(num_cols, self.max_merge_cols)
        
        for col_idx in range(check_cols):
            merge_start_row = None
            prev_tc_id = None
            
            for row_idx in range(check_rows):
                try:
                    cell = table.cell(row_idx, col_idx)
                    current_tc_id = id(cell._tc)
                    
                    # Если это та же _tc что и выше - продолжаем объединение
                    if row_idx > 0:
                        prev_cell = table.cell(row_idx - 1, col_idx)
                        if id(prev_cell._tc) == current_tc_id:
                            # Продолжается объединение
                            if merge_start_row is None:
                                merge_start_row = row_idx - 1
                        else:
                            # Закончилось объединение
                            if merge_start_row is not None:
                                try:
                                    sheet.merge_cells(
                                        start_row=start_row + merge_start_row,
                                        start_column=col_idx + 1,
                                        end_row=start_row + row_idx - 1,
                                        end_column=col_idx + 1
                                    )
                                    # Центрируем текст
                                    merged_cell = sheet.cell(row=start_row + merge_start_row, column=col_idx + 1)
                                    merged_cell.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center')
                                except Exception as e:
                                    pass
                            merge_start_row = None
                
                except Exception as e:
                    pass
            
            # Завершаем объединение в конце столбца
            if merge_start_row is not None:
                try:
                    sheet.merge_cells(
                        start_row=start_row + merge_start_row,
                        start_column=col_idx + 1,
                        end_row=start_row + check_rows - 1,
                        end_column=col_idx + 1
                    )
                    # Центрируем текст
                    merged_cell = sheet.cell(row=start_row + merge_start_row, column=col_idx + 1)
                    merged_cell.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center')
                except Exception as e:
                    pass
        
        # Ищем горизонтальные объединения (по строкам)
        for row_idx in range(check_rows):
            merge_start_col = None
            
            for col_idx in range(check_cols):
                try:
                    cell = table.cell(row_idx, col_idx)
                    current_tc_id = id(cell._tc)
                    
                    # Если это та же _tc что и слева - продолжаем объединение
                    if col_idx > 0:
                        prev_cell = table.cell(row_idx, col_idx - 1)
                        if id(prev_cell._tc) == current_tc_id:
                            # Продолжается объединение
                            if merge_start_col is None:
                                merge_start_col = col_idx - 1
                        else:
                            # Закончилось объединение
                            if merge_start_col is not None:
                                try:
                                    sheet.merge_cells(
                                        start_row=start_row + row_idx,
                                        start_column=merge_start_col + 1,
                                        end_row=start_row + row_idx,
                                        end_column=col_idx
                                    )
                                    # Центрируем текст
                                    merged_cell = sheet.cell(row=start_row + row_idx, column=merge_start_col + 1)
                                    merged_cell.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center')
                                except Exception as e:
                                    pass
                            merge_start_col = None
                
                except Exception as e:
                    pass
            
            # Завершаем объединение в конце строки
            if merge_start_col is not None:
                try:
                    sheet.merge_cells(
                        start_row=start_row + row_idx,
                        start_column=merge_start_col + 1,
                        end_row=start_row + row_idx,
                        end_column=check_cols
                    )
                    # Центрируем текст
                    merged_cell = sheet.cell(row=start_row + row_idx, column=merge_start_col + 1)
                    merged_cell.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center')
                except Exception as e:
                    pass
        
        # Обновляем текущую строку (используем actual_row_idx для учета пропущенных строк)
        self.current_row += actual_row_idx
        
        # Добавляем пустую строку после таблицы (если опция не включена)
        if not self.remove_empty_rows:
            self.current_row += 1
        
        self.last_element_type = 'table'  # Запоминаем, что добавили таблицу
    
    def _adjust_column_widths(self, sheet):
        """
        Автоматически настраивает ширину столбцов.
        
        Args:
            sheet: Активный лист Excel
        """
        for column in sheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            
            for cell in column:
                try:
                    if cell.value:
                        cell_length = len(str(cell.value))
                        if cell_length > max_length:
                            max_length = cell_length
                except:
                    pass
            
            # Устанавливаем ширину с учетом максимальной длины
            adjusted_width = min(max_length + 2, 100)  # Максимум 100 символов
            sheet.column_dimensions[column_letter].width = adjusted_width


def convert_docx_to_excel(docx_path: str, excel_path: Optional[str] = None) -> str:
    """
    Утилита для быстрой конвертации docx в excel.
    
    Args:
        docx_path: Путь к исходному docx файлу
        excel_path: Путь для сохранения excel файла (опционально)
        
    Returns:
        Путь к созданному excel файлу
        
    Raises:
        Exception: Если произошла ошибка при конвертации
    """
    if excel_path is None:
        # Автоматическое создание имени для excel файла
        base_name = os.path.splitext(docx_path)[0]
        excel_path = f"{base_name}.xlsx"
    
    converter = DocxToExcelConverter(remove_empty_rows=False)
    converter.convert(docx_path, excel_path)
    
    return excel_path

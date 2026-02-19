"""
Виджет для проверки текстов классификации
Находит ош ибки и подсвечивает различия между текстом в проекте и справочником
"""
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
                             QLabel, QTableWidget, QTableWidgetItem, QHeaderView,
                             QComboBox, QMessageBox, QProgressDialog, QFileDialog,
                             QApplication, QStyledItemDelegate, QStyle, QSpinBox,
                             QDialog, QTextBrowser, QSplitter, QStyleOptionViewItem,
                             QTabWidget, QSizePolicy, QCheckBox)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QRectF, QSize, QTimer
from PyQt5.QtGui import QFont, QTextDocument, QAbstractTextDocumentLayout, QPalette, QColor
from datetime import datetime
import sqlite3
import pandas as pd
from logger import logger
from utils.db_utils import get_filtered_view
from utils.text_validation.display_helpers import richtext_to_html, find_alternative_variants


class HtmlDelegate(QStyledItemDelegate):
    """Делегат для отображения HTML в ячейках таблицы"""
    
    def paint(self, painter, option, index):
        """Отрисовка HTML содержимого ячейки"""
        options = option
        self.initStyleOption(options, index)
        
        painter.save()
        
        # Создаем QTextDocument для рендеринга HTML
        doc = QTextDocument()
        doc.setHtml(options.text)
        doc.setTextWidth(options.rect.width())
        
        # Очищаем текст из options, чтобы стандартная отрисовка не конфликтовала
        options.text = ""
        
        # Отрисовываем фон и границы ячейки
        options.widget.style().drawControl(QStyle.CE_ItemViewItem, options, painter)
        
        # Сдвигаем painter на позицию ячейки
        painter.translate(options.rect.left(), options.rect.top())
        
        # Создаем rect для содержимого с отступами (QRectF для совместимости)
        clip = QRectF(0, 0, options.rect.width(), options.rect.height())
        
        # Отрисовываем HTML
        ctx = QAbstractTextDocumentLayout.PaintContext()
        ctx.clip = clip
        doc.documentLayout().draw(painter, ctx)
        
        painter.restore()
    
    def sizeHint(self, option, index):
        """Вычисление размера ячейки с учетом HTML"""
        options = option
        self.initStyleOption(options, index)
        
        doc = QTextDocument()
        doc.setHtml(options.text)
        doc.setTextWidth(options.rect.width())
        
        return doc.size().toSize()


class WordWrapDelegate(QStyledItemDelegate):
    """Делегат для автоматического переноса текста в ячейках"""
    
    def paint(self, painter, option, index):
        """Отрисовка с переносом текста"""
        opt = QStyleOptionViewItem(option)
        self.initStyleOption(opt, index)
        
        painter.save()
        
        # Создаем QTextDocument для переноса текста
        doc = QTextDocument()
        doc.setDefaultFont(opt.font)
        doc.setPlainText(opt.text)
        doc.setTextWidth(opt.rect.width())
        
        # Очищаем текст из options
        opt.text = ""
        
        # Отрисовываем фон и границы
        opt.widget.style().drawControl(QStyle.CE_ItemViewItem, opt, painter)
        
        # Сдвигаем painter
        painter.translate(opt.rect.left(), opt.rect.top())
        
        # Отрисовываем текст
        clip = QRectF(0, 0, opt.rect.width(), opt.rect.height())
        ctx = QAbstractTextDocumentLayout.PaintContext()
        ctx.clip = clip
        doc.documentLayout().draw(painter, ctx)
        
        painter.restore()
    
    def sizeHint(self, option, index):
        """Вычисление высоты с учетом переноса"""
        opt = QStyleOptionViewItem(option)
        self.initStyleOption(opt, index)
        
        doc = QTextDocument()
        doc.setDefaultFont(opt.font)
        doc.setPlainText(opt.text)
        doc.setTextWidth(option.rect.width() if option.rect.width() > 0 else 200)
        
        return QSize(int(doc.idealWidth()), int(doc.size().height()))


class TextValidationWorker(QThread):
    """Фоновый поток для проверки текстов"""
    progress = pyqtSignal(int, int)  # current, total
    finished = pyqtSignal(list, list, list, list, list, list)  # errors, all_items, orig_codes, reference_codes, reference_names, reference_data
    error = pyqtSignal(str)  # сообщение об ошибке
    
    # Маппинг разделов на таблицы справочников
    SECTION_TABLE_MAP = {
        'Доходы': 'v_budgetclastypeinc_merged',
        'Расходы': 'v_budgetclascosts_merged',
        'Источники финансирования': 'v_budgetclassources_merged',
    }
    
    def __init__(self, section, project_data, ppocode, filter_date, db_path):
        super().__init__()
        self.section = section
        self.project_data = project_data
        self.ppocode = ppocode
        self.filter_date = filter_date
        self.db_path = db_path
    
    def run(self):
        """Выполнение проверки в фоновом потоке"""
        try:
            logger.info(f"Запуск проверки текстов для раздела '{self.section}'")
            from utils.text_validation import find_errors, ErrorInfo
            
            # Определяем таблицу справочника по разделу
            table_name = self.SECTION_TABLE_MAP.get(self.section)
            if not table_name:
                logger.error(f"Неподдерживаемый раздел: {self.section}")
                self.error.emit(f"Неподдерживаемый раздел: {self.section}")
                return
            
            logger.info(f"Загрузка справочника из таблицы '{table_name}' с фильтрами: дата={self.filter_date}, ppocode={self.ppocode}")
            
            # Загружаем справочник из БД
            conn = sqlite3.connect(self.db_path)
            
            # Фильтруем по дате и ppocode
            reference_df = get_filtered_view(
                conn, 
                table_name, 
                filter_date=self.filter_date,
                filter_ppocode=self.ppocode
            )
            conn.close()
            
            logger.info(f"Загружено {len(reference_df)} записей из справочника")
            
            if reference_df.empty:
                logger.warning(f"Справочник '{self.section}' пуст для указанной даты и ppocode")
                self.error.emit(f"Справочник '{self.section}' пуст для указанной даты и ppocode")
                return
            
            # Получаем наименования из справочника
            reference_names = reference_df['name'].astype(str).str.strip().tolist()
            
            # Получаем коды из справочника через единый метод
            reference_codes = TextValidationWidget._extract_codes_from_reference(reference_df, self.section)
            
            # Получаем данные из проекта
            orig_names = []
            orig_codes = []
            
            for row in self.project_data:
                name = row.get('наименование_показателя', '')
                
                if name and name.strip():
                    orig_names.append(name.strip())
            
            # Извлекаем коды через единый метод
            orig_codes = TextValidationWidget._extract_codes_from_project_data(self.project_data)
            
            if not orig_names:
                logger.warning(f"Нет данных для проверки в разделе '{self.section}'")
                self.error.emit(f"Нет данных для проверки в разделе '{self.section}'")
                return
            
            logger.info(f"Начало проверки: {len(orig_names)} записей в проекте, {len(reference_names)} в справочнике")
            
            # Выполняем проверку
            def progress_callback(current, total):
                self.progress.emit(current, total)
            
            errors = find_errors(
                orig_names=orig_names,
                reference_names=reference_names,
                orig_codes=orig_codes,
                reference_codes=reference_codes,
                max_distance=300,
                progress_callback=progress_callback
            )
            
            logger.info(f"Проверка завершена: найдено {len(errors)} ошибок")
            
            # Создаем список всех проверенных элементов (включая корректные)
            # Строим множество индексов ошибок для быстрого поиска
            error_indices = {err.original_index for err in errors}
            
            all_items = list(errors)  # Начинаем с ошибок
            
            # Добавляем корректные элементы (которых нет в ошибках)
            for idx, orig_name in enumerate(orig_names):
                if idx not in error_indices:
                    # Создаем "корректный" ErrorInfo с distance = 0
                    # Находим точное совпадение в справочнике
                    ref_idx = None
                    if orig_name in reference_names:
                        ref_idx = reference_names.index(orig_name)
                    
                    correct_item = ErrorInfo(
                        original_text=orig_name,
                        reference_text=orig_name,  # Точное совпадение
                        distance=0,
                        diff_indices=[],
                        corrections=[],
                        original_index=idx,
                        reference_index=ref_idx,
                        code_error=None
                    )
                    all_items.append(correct_item)
            
            # Сортируем all_items по original_index для сохранения порядка
            all_items.sort(key=lambda x: x.original_index)
            
            # Преобразуем DataFrame в список словарей для передачи
            reference_data = reference_df.to_dict('records')
            
            self.finished.emit(errors, all_items, orig_codes, reference_codes, reference_names, reference_data)
            
        except Exception as e:
            logger.error(f"Ошибка при проверке текстов: {e}", exc_info=True)
            self.error.emit(f"Ошибка: {str(e)}")


class TextValidationWidget(QWidget):
    """Виджет для проверки текстов классификации"""
    
    # Маппинг разделов на ключи в project.data
    SECTION_DATA_KEY_MAP = {
        'Доходы': 'income_data',
        'Расходы': 'outcome_data',
        'Источники финансирования': 'source_financing_deficit_data',
    }
    
    def __init__(self, main_window):
        super().__init__()
        self.main_window = main_window
        self.controller = main_window.controller
        self.errors = []
        self.all_items = []  # Все проверенные элементы (включая корректные)
        self.orig_codes = []
        self.reference_codes = []
        self._reference_names = []  # Для поиска альтернатив
        self._reference_data = []  # Полные данные справочника
        self.worker = None
        self._last_loaded_key = None  # Ключ последней загруженной комбинации (project_id, revision_id, section)
        
        self.init_ui()
    
    def init_ui(self):
        """Инициализация интерфейса"""
        layout = QVBoxLayout(self)
        layout.setContentsMargins(5, 5, 5, 5)
        
        # Заголовок и фильтры
        header_layout = QHBoxLayout()
        
        info_label = QLabel("Проверка соответствия текстов классификации справочнику:")
        info_label.setFont(QFont("Arial", 10, QFont.Bold))
        header_layout.addWidget(info_label)
        
        header_layout.addStretch()
        
        # Фильтр по разделу
        header_layout.addWidget(QLabel("Раздел:"))
        self.section_combo = QComboBox()
        self.section_combo.addItems(["Доходы", "Расходы", "Источники финансирования"])
        self.section_combo.currentTextChanged.connect(self._on_section_changed)
        header_layout.addWidget(self.section_combo)
        
        # Фильтр по предельной ошибке
        header_layout.addWidget(QLabel("Дистанция:"))
        self.distance_threshold = QSpinBox()
        self.distance_threshold.setMinimum(1)
        self.distance_threshold.setMaximum(301)
        self.distance_threshold.setValue(200)  # Значение по умолчанию
        self.distance_threshold.setToolTip("Показывать только ошибки с distance не более указанного значения")
        self.distance_threshold.valueChanged.connect(self._apply_distance_filter)
        header_layout.addWidget(self.distance_threshold)
        
        # Checkbox для отображения всех текстов
        self.show_all_checkbox = QCheckBox("Показать все тексты")
        self.show_all_checkbox.setToolTip("Показывать все проверенные тексты, включая корректные")
        self.show_all_checkbox.stateChanged.connect(self._apply_distance_filter)
        header_layout.addWidget(self.show_all_checkbox)
        
        # Кнопка проверки
        self.check_btn = QPushButton("Проверить")
        self.check_btn.clicked.connect(self.check_texts)
        header_layout.addWidget(self.check_btn)
        
        # Кнопка загрузки сохраненных результатов
        self.load_btn = QPushButton("Загрузить сохраненные")
        self.load_btn.clicked.connect(self.load_saved_errors)
        header_layout.addWidget(self.load_btn)
        
        # Кнопка экспорта
        self.export_btn = QPushButton("Экспорт...")
        self.export_btn.clicked.connect(self.export_errors)
        self.export_btn.setEnabled(False)
        header_layout.addWidget(self.export_btn)
        
        layout.addLayout(header_layout)
        
        # Таблица ошибок
        self.errors_table = QTableWidget()
        self.errors_table.setColumnCount(5)
        self.errors_table.setHorizontalHeaderLabels([
            "Текст в проекте",
            "Эталон из справочника",
            "Distance",
            "Код в проекте",
            "Код справочника"
        ])
        
        # Настройка таблицы
        header = self.errors_table.horizontalHeader()
        header.setStretchLastSection(False)
        for i in range(5):
            header.setSectionResizeMode(i, QHeaderView.Interactive)
        
        header.resizeSection(0, 350)  # Текст в проекте
        header.resizeSection(1, 350)  # Эталон
        header.resizeSection(2, 80)   # Distance
        header.resizeSection(3, 200)  # Код в проекте
        header.resizeSection(4, 200)  # Код справочника
        
        self.errors_table.setAlternatingRowColors(True)
        self.errors_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.errors_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.errors_table.setWordWrap(True)
        
        # Обработчик двойного клика
        self.errors_table.cellDoubleClicked.connect(self._show_error_details)
        
        # Устанавливаем делегат для отображения HTML во всех текстовых столбцах
        html_delegate = HtmlDelegate(self.errors_table)
        self.errors_table.setItemDelegateForColumn(0, html_delegate)  # Текст в проекте
        self.errors_table.setItemDelegateForColumn(1, html_delegate)  # Эталон из справочника
        self.errors_table.setItemDelegateForColumn(3, html_delegate)  # Код в проекте
        self.errors_table.setItemDelegateForColumn(4, html_delegate)  # Код справочника
        
        layout.addWidget(self.errors_table)
        
        # Статистика
        self.stats_label = QLabel("Нет данных для проверки")
        self.stats_label.setFont(QFont("Arial", 9))
        layout.addWidget(self.stats_label)
    
    def _get_section_key(self, section: str) -> str:
        """Получение ключа раздела в project.data по его названию"""
        return self.SECTION_DATA_KEY_MAP.get(section, '')
    
    def _get_section_data(self, section: str):
        """Получение данных раздела из текущего проекта"""
        if not self.controller.current_project:
            raise ValueError("Проект не выбран")
        
        if not self.controller.current_project.data:
            raise ValueError("Проект не содержит данных")
        
        section_key = self._get_section_key(section)
        if not section_key:
            raise ValueError(f"Неизвестный раздел: {section}")
        
        section_data = self.controller.current_project.data.get(section_key, [])
        if not section_data:
            raise ValueError(f"Раздел '{section}' не содержит данных")
        
        return section_data
    
    @staticmethod
    def _extract_codes_from_reference(reference_df, section: str):
        """
        Извлечение кодов классификации из DataFrame справочника.
        Возвращает список склеенных кодов в зависимости от раздела.
        """
        reference_codes = []
        
        if section == 'Доходы':
            # inctypecode + incsubtypecode + analyticalgroupcode
            code_df = reference_df[['inctypecode', 'incsubtypecode', 'analyticalgroupcode']].fillna('').astype(str)
            reference_codes = (
                code_df['inctypecode'] + 
                code_df['incsubtypecode'] + 
                code_df['analyticalgroupcode']
            ).str.replace(' ', '').tolist()
        elif section == 'Расходы':
            # grbscode + rzpr + kcsr + kvr
            code_df = reference_df[['grbscode', 'rzpr', 'kcsr', 'kvr']].fillna('').astype(str)
            reference_codes = (
                code_df['grbscode'] + 
                code_df['rzpr'] + 
                code_df['kcsr'] + 
                code_df['kvr']
            ).str.replace(' ', '').tolist()
        elif section == 'Источники финансирования':
            # code
            reference_codes = reference_df['code'].fillna('').astype(str).str.replace(' ', '').tolist()
        
        return reference_codes
    
    @staticmethod
    def _extract_codes_from_project_data(section_data):
        """
        Извлечение кодов классификации из данных проекта.
        Для кода 'x' возвращает пустую строку (проверка только по наименованию).
        """
        codes = []
        for row in section_data:
            code = row.get('код_классификации', '')
            code_str = str(code).replace(' ', '').strip() if code else ''
            if code_str.lower() == 'x':
                code_str = ''  # При коде 'x' проверяем только наименование
            codes.append(code_str)
        return codes
    
    def showEvent(self, event):
        """Событие при показе виджета - автоматически загружаем сохраненные данные"""
        super().showEvent(event)
        # Загружаем данные только если есть активный проект и ревизия
        # Проверку на пустоту таблицы убираем, чтобы данные обновлялись при переключении проектов
        if self.controller.current_project and self.controller.current_revision_id:
            self._auto_load_saved_errors()
    
    def _on_section_changed(self):
        """Обработчик изменения раздела - автоматически загружаем данные для нового раздела"""
        if self.controller.current_project and self.controller.current_revision_id:
            self._auto_load_saved_errors()
    
    def _apply_distance_filter(self):
        """Применение фильтра по distance - обновление таблицы"""
        # Проверяем наличие данных
        show_all = self.show_all_checkbox.isChecked()
        items_to_check = self.all_items if show_all else self.errors
        
        if items_to_check:
            mode = "все тексты" if show_all else "только ошибки"
            logger.info(f"Фильтр изменен: distance={self.distance_threshold.value()}, режим={mode}")
            self._update_errors_table()
    
    def _auto_load_saved_errors(self):
        """Автоматическая загрузка сохраненных ошибок без диалогов"""
        section = self.section_combo.currentText()
        
        # Проверяем, не загружали ли мы уже эту комбинацию
        current_key = (
            self.controller.current_project.id,
            self.controller.current_revision_id,
            section
        )
        
        if current_key == self._last_loaded_key:
            # Уже загружено, пропускаем
            return
        
        try:
            # Загружаем ошибки из БД
            error_dicts = self.controller.db_manager.get_text_validation_errors(
                project_id=self.controller.current_project.id,
                revision_id=self.controller.current_revision_id,
                section=section
            )
            
            if not error_dicts:
                # Нет сохраненных данных - очищаем таблицу
                self.errors = []
                self.all_items = []
                self.orig_codes = []
                self.reference_codes = []
                self.errors_table.setRowCount(0)
                self.stats_label.setText("Нет сохраненных данных. Нажмите 'Проверить' для проверки.")
                self.export_btn.setEnabled(False)
                self.show_all_checkbox.setChecked(False)
                self.show_all_checkbox.setEnabled(False)
                self._last_loaded_key = current_key  # Запоминаем, что загрузили (пустые данные)
                return
            
            # Используем общий метод загрузки
            self._do_load_errors(error_dicts)
            self._last_loaded_key = current_key  # Запоминаем успешную загрузку
            
        except Exception as e:
            logger.error(f"Ошибка при автоматической загрузке сохраненных ошибок: {e}", exc_info=True)
            # Не показываем диалог ошибки при автоматической загрузке
    
    def _do_load_errors(self, error_dicts):
        """Общая логика загрузки ошибок из БД (используется и для автоматической, и для ручной загрузки)"""
        section = self.section_combo.currentText()
        
        # Получаем данные раздела через единый метод
        section_data = self._get_section_data(section)
        
        # Извлекаем коды из данных проекта
        self.orig_codes = []
        self.reference_codes = []
        
        for row in section_data:
            code = row.get('код_классификации', '')
            code_str = str(code).replace(' ', '').strip() if code else ''
            if code_str.lower() == 'x':
                code_str = ''
            self.orig_codes.append(code_str)
        
        # Загружаем справочник для альтернативных вариантов
        try:
            project = self.controller.current_project
            ppocode_from_meta = project.oktmo_code if project else None
            if ppocode_from_meta and ppocode_from_meta != "00000000":
                ppocode_filter = ["00000000", ppocode_from_meta]
            else:
                ppocode_filter = "00000000"
            
            filter_date = self.main_window.reference_date_edit.text().strip()
            if not filter_date:
                from datetime import datetime
                filter_date = datetime.now().strftime("%Y-%m-%d")
            
            # Определяем таблицу справочника
            section_table_map = {
                'Доходы': 'v_budgetclastypeinc_merged',
                'Расходы': 'v_budgetclascosts_merged',
                'Источники финансирования': 'v_budgetclassources_merged',
            }
            table_name = section_table_map.get(section)
            
            if table_name:
                conn = sqlite3.connect(self.controller.db_manager.db_path)
                # Загружаем с дедупликацией для основной проверки (актуальные записи)
                reference_df = get_filtered_view(conn, table_name, filter_date=filter_date, filter_ppocode=ppocode_filter, deduplicate=True)
                conn.close()
                
                if not reference_df.empty:
                    self._reference_names = reference_df['name'].astype(str).str.strip().tolist()
                    self.reference_codes = self._extract_codes_from_reference(reference_df, section)
                    self._reference_data = reference_df.to_dict('records')  # Сохраняем полные данные
                else:
                    self._reference_names = []
                    self.reference_codes = []
                    self._reference_data = []
            else:
                self._reference_names = []
                self.reference_codes = []
                self._reference_data = []
        except Exception as e:
            logger.error(f"Ошибка при загрузке справочника: {e}", exc_info=True)
            self._reference_names = []
            self.reference_codes = []
            self._reference_data = []
        
        # Преобразуем словари в объекты ErrorInfo
        # Преобразуем словари в объекты ErrorInfo
        from utils.text_validation import ErrorInfo, CodeErrorInfo
        
        self.errors = []
        for err_dict in error_dicts:
            # Восстанавливаем объект code_error
            code_error = None
            if err_dict.get('code_error'):
                ce = err_dict['code_error']
                code_error = CodeErrorInfo(
                    has_error=ce.get('has_error', False),
                    ref_code=ce.get('ref_code', ''),
                    distance=ce.get('distance', 0),
                    diff_indices=ce.get('diff_indices', []),
                    corrections=ce.get('corrections', []),
                    code_ref_idx=ce.get('code_ref_idx')
                )
            
            error_info = ErrorInfo(
                original_text=err_dict['original_text'],
                reference_text=err_dict['reference_text'],
                distance=err_dict['distance'],
                diff_indices=err_dict['diff_indices'],
                corrections=err_dict['corrections'],
                original_index=err_dict['original_index'],
                reference_index=err_dict['reference_index'],
                code_error=code_error
            )
            self.errors.append(error_info)
        
        # При загрузке из БД нет информации о корректных элементах
        self.all_items = list(self.errors)  # Копируем только ошибки
        self.show_all_checkbox.setChecked(False)  # Снимаем галочку
        self.show_all_checkbox.setEnabled(False)  # Отключаем checkbox (нет данных о корректных элементах)
        
        # Обновляем таблицу
        self._update_errors_table()
        self.stats_label.setText(f"Загружено из БД: {len(self.errors)} ошибок (сохранено: {error_dicts[0]['created_at']})")
        self.export_btn.setEnabled(len(self.errors) > 0)
    
    def check_texts(self):
        """Запуск проверки текстов классификации"""
        logger.info("Вызван метод check_texts()")
        # Получаем текущий проект и ревизию
        project = self.controller.current_project
        if not project:
            logger.warning("Попытка проверки текстов без выбранного проекта")
            QMessageBox.warning(self, "Предупреждение", "Не выбран проект")
            return
        
        revision_id = self.controller.current_revision_id
        if not revision_id:
            logger.warning("Попытка проверки текстов без выбранной ревизии")
            QMessageBox.warning(self, "Предупреждение", "Не выбрана ревизия")
            return
        
        logger.info(f"Проверка текстов для проекта '{project.name}' (ID: {project.id}), ревизия ID: {revision_id}")
        
        # Получаем ppocode из проекта
        ppocode_from_meta = project.oktmo_code
        if ppocode_from_meta and ppocode_from_meta != "00000000":
            ppocode_filter = ["00000000", ppocode_from_meta]  # ФУ + ОКТМО из проекта
        else:
            ppocode_filter = "00000000"  # только ФУ
        
        # Получаем дату актуальности
        filter_date = self.main_window.reference_date_edit.text().strip()
        if not filter_date:
            filter_date = datetime.now().strftime("%Y-%m-%d")
        
        # Получаем данные раздела через единый метод
        section = self.section_combo.currentText()
        
        try:
            section_data = self._get_section_data(section)
            logger.info(f"Раздел '{section}' содержит {len(section_data)} записей")
        except ValueError as e:
            logger.warning(f"Ошибка получения данных раздела: {e}")
            QMessageBox.warning(self, "Предупреждение", str(e))
            return
        
        # Создаем диалог прогресса
        progress_dialog = QProgressDialog("Проверка текстов...", "Отмена", 0, 100, self)
        progress_dialog.setWindowTitle("Проверка")
        progress_dialog.setWindowModality(Qt.WindowModal)
        progress_dialog.setMinimumDuration(0)
        progress_dialog.setValue(0)
        
        # Создаем и запускаем фоновый поток
        self.worker = TextValidationWorker(
            section=section,
            project_data=section_data,
            ppocode=ppocode_filter,
            filter_date=filter_date,
            db_path=self.controller.db_manager.db_path
        )
        
        self.worker.progress.connect(
            lambda current, total: progress_dialog.setValue(int(current * 100 / total) if total > 0 else 0)
        )
        self.worker.finished.connect(
            lambda errors, all_items, orig_codes, ref_codes, ref_names, ref_data: self._on_check_finished(errors, all_items, orig_codes, ref_codes, ref_names, ref_data, progress_dialog)
        )
        self.worker.error.connect(lambda msg: self._on_check_error(msg, progress_dialog))
        
        progress_dialog.canceled.connect(self.worker.terminate)
        
        self.worker.start()
    
    def _on_check_finished(self, errors, all_items, orig_codes, ref_codes, ref_names, ref_data, progress_dialog):
        """Обработка завершения проверки"""
        logger.info(f"Проверка завершена: получено {len(errors)} ошибок из {len(all_items)} всего")
        progress_dialog.close()
        
        self.errors = errors
        self.all_items = all_items  # Сохраняем все проверенные элементы
        self.orig_codes = orig_codes
        self.reference_codes = ref_codes
        self._reference_names = ref_names  # Сохраняем для поиска альтернатив
        self._reference_data = ref_data  # Сохраняем полные данные справочника
        
        # Включаем checkbox для отображения всех текстов (теперь есть данные)
        self.show_all_checkbox.setEnabled(True)
        
        try:
            logger.info("Обновление таблицы ошибок...")
            self._update_errors_table()
            logger.info("Таблица ошибок обновлена успешно")
            
            # Сохраняем результаты в БД
            if self.controller.current_project and self.controller.current_revision_id:
                section = self.section_combo.currentText()
                try:
                    logger.info(f"Сохранение {len(errors)} ошибок текстов в БД...")
                    self.controller.db_manager.save_text_validation_errors(
                        project_id=self.controller.current_project.id,
                        revision_id=self.controller.current_revision_id,
                        section=section,
                        errors=errors
                    )
                    logger.info("Ошибки текстов сохранены в БД успешно")
                    
                    # Обновляем ключ последней загрузки, так как данные изменились
                    self._last_loaded_key = (
                        self.controller.current_project.id,
                        self.controller.current_revision_id,
                        section
                    )
                except Exception as e:
                    logger.error(f"Ошибка при сохранении ошибок текстов в БД: {e}", exc_info=True)
                    # Не блокируем работу, если сохранение не удалось
            
        except Exception as e:
            logger.error(f"Ошибка при обновлении таблицы: {e}", exc_info=True)
            QMessageBox.critical(self, "Ошибка", f"Ошибка при отображении результатов: {str(e)}")
            return
        
        if not errors:
            QMessageBox.information(self, "Результат", "Ошибок не найдено!")
            self.stats_label.setText(f"Ошибок не найдено (проверено: {len(all_items)} текстов)")
        else:
            self.stats_label.setText(f"Найдено ошибок: {len(errors)} из {len(all_items)} проверенных")
        
        self.export_btn.setEnabled(len(errors) > 0)
    
    def _on_check_error(self, error_msg, progress_dialog):
        """Обработка ошибки проверки"""
        progress_dialog.close()
        QMessageBox.critical(self, "Ошибка", error_msg)
    
    def load_saved_errors(self):
        """Загрузка ранее сохраненных ошибок текстов из БД (ручной вызов с диалогами)"""
        if not self.controller.current_project:
            QMessageBox.warning(self, "Предупреждение", "Сначала выберите проект")
            return
        
        if not self.controller.current_revision_id:
            QMessageBox.warning(self, "Предупреждение", "Сначала загрузите ревизию формы")
            return
        
        # Используем внутренний метод для загрузки
        section = self.section_combo.currentText()
        
        try:
            # Загружаем ошибки из БД
            logger.info(f"Ручная загрузка сохраненных ошибок текстов для раздела '{section}'...")
            error_dicts = self.controller.db_manager.get_text_validation_errors(
                project_id=self.controller.current_project.id,
                revision_id=self.controller.current_revision_id,
                section=section
            )
            
            if not error_dicts:
                QMessageBox.information(self, "Информация", "Сохраненных ошибок для данного раздела не найдено")
                self.errors = []
                self.all_items = []
                self.orig_codes = []
                self.reference_codes = []
                self._update_errors_table()
                self.stats_label.setText("Сохраненных ошибок не найдено")
                self.export_btn.setEnabled(False)
                self.show_all_checkbox.setChecked(False)
                self.show_all_checkbox.setEnabled(False)
                return
            
            # Выполняем загрузку
            self._do_load_errors(error_dicts)
            
            # Обновляем ключ последней загрузки
            self._last_loaded_key = (
                self.controller.current_project.id,
                self.controller.current_revision_id,
                section
            )
            
            # Показываем сообщение об успехе
            QMessageBox.information(
                self, 
                "Успех", 
                f"Загружено {len(self.errors)} ошибок из БД\nДата сохранения: {error_dicts[0]['created_at']}"
            )
            
        except Exception as e:
            logger.error(f"Ошибка при загрузке сохраненных ошибок: {e}", exc_info=True)
            QMessageBox.critical(self, "Ошибка", f"Ошибка при загрузке сохраненных ошибок:\n{str(e)}")
    
    def _richtext_to_html(self, orig_text, diff_idx, corrections, invert_colors=False):
        """Преобразование подсвеченного текста в HTML (делегирует в общую функцию)"""
        return richtext_to_html(orig_text, diff_idx, corrections, invert_colors)
    
    def _update_errors_table(self):
        """Обновление таблицы ошибок"""
        from utils.text_validation import create_highlighted_text, normalize_classification_code
        
        # Выбираем источник данных в зависимости от состояния checkbox
        show_all = self.show_all_checkbox.isChecked()
        items_to_display = self.all_items if show_all else self.errors
        
        logger.info(f"Обновление таблицы: {'все элементы' if show_all else 'только ошибки'} ({len(items_to_display)} записей)")
        
        if not items_to_display:
            self.errors_table.setRowCount(0)
            logger.info("Нет элементов для отображения")
            return
        
        # Фильтруем по порогу distance
        max_distance = self.distance_threshold.value()
        # Если выбрана максимальная дистанция, увеличиваем на 1 для отображения всех записей
        if max_distance == self.distance_threshold.maximum():
            max_distance += 1
        filtered_items = [item for item in items_to_display if item.distance <= max_distance]
        
        logger.info(f"После фильтрации по distance <= {max_distance}: {len(filtered_items)} записей")
        
        if not filtered_items:
            self.errors_table.setRowCount(0)
            logger.info("Нет элементов после фильтрации")
            return
        
        self.errors_table.setRowCount(len(filtered_items))
        
        for row_idx, error in enumerate(filtered_items):
            try:
                # Текст в проекте - HTML с подсветкой для Qt таблицы
                html_text = self._richtext_to_html(
                    error.original_text,
                    error.diff_indices,
                    error.corrections
                )
                item = QTableWidgetItem()
                item.setData(Qt.DisplayRole, html_text)
                item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
                self.errors_table.setItem(row_idx, 0, item)
                
                # Эталон из справочника (оборачиваем в HTML для единообразного отображения)
                ref_text = error.reference_text if error.reference_text else ""
                ref_item = QTableWidgetItem()
                ref_item.setData(Qt.DisplayRole, ref_text)  # Обычный текст, без HTML
                ref_item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
                self.errors_table.setItem(row_idx, 1, ref_item)
                
                # Distance
                self.errors_table.setItem(row_idx, 2, QTableWidgetItem(str(error.distance)))
                
                # Код в проекте (с подсветкой, если есть ошибка в коде)
                orig_code = ""
                if error.original_index < len(self.orig_codes):
                    orig_code = self.orig_codes[error.original_index]
                
                if error.code_error:
                    # Нормализуем код и добавляем HTML подсветку
                    normalized_code = normalize_classification_code(orig_code)
                    code_html = self._richtext_to_html(
                        normalized_code,
                        error.code_error.diff_indices,
                        error.code_error.corrections
                    )
                    item = QTableWidgetItem()
                    item.setData(Qt.DisplayRole, code_html)
                    item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
                    self.errors_table.setItem(row_idx, 3, item)
                else:
                    self.errors_table.setItem(row_idx, 3, QTableWidgetItem(orig_code))
                
                # Код справочника (оборачиваем для единообразного отображения)
                ref_code = ""
                if error.code_error and error.code_error.ref_code:
                    ref_code = error.code_error.ref_code
                elif error.reference_index is not None and error.reference_index < len(self.reference_codes):
                    ref_code = self.reference_codes[error.reference_index]
                
                ref_code_item = QTableWidgetItem()
                ref_code_item.setData(Qt.DisplayRole, ref_code)  # Обычный текст, без HTML
                ref_code_item.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
                self.errors_table.setItem(row_idx, 4, ref_code_item)
            
            except Exception as e:
                logger.error(f"Ошибка при обработке строки {row_idx}: {e}", exc_info=True)
                # Заполняем строку пустыми значениями, чтобы не сломать всю таблицу
                for col in range(5):
                    if self.errors_table.item(row_idx, col) is None:
                        self.errors_table.setItem(row_idx, col, QTableWidgetItem(""))
        
        # Подгоняем высоту строк
        logger.info("Подгонка высоты строк...")
        self.errors_table.resizeRowsToContents()
        logger.info("Таблица заполнена успешно")
    
    def _show_error_details(self, row: int, column: int):
        """Показ детальной информации об ошибке текста при двойном клике с альтернативными вариантами"""
        if row < 0:
            return
        
        # Получаем отфильтрованный список элементов (в зависимости от режима отображения)
        show_all = self.show_all_checkbox.isChecked()
        items_to_display = self.all_items if show_all else self.errors
        
        if not items_to_display:
            return
        
        max_distance = self.distance_threshold.value()
        filtered_items = [item for item in items_to_display if item.distance <= max_distance]
        
        if row >= len(filtered_items):
            return
        
        error = filtered_items[row]
        section = self.section_combo.currentText()
        
        # Получаем коды
        orig_code = ""
        if error.original_index < len(self.orig_codes):
            orig_code = self.orig_codes[error.original_index]
        
        ref_code = ""
        if error.code_error and error.code_error.ref_code:
            ref_code = error.code_error.ref_code
        elif error.reference_index is not None and error.reference_index < len(self.reference_codes):
            ref_code = self.reference_codes[error.reference_index]
        
        # Информация о коде
        if error.code_error and error.code_error.has_error:
            code_status = "<span style='color: red;'>❌ Несоответствие</span>"
            code_distance = error.code_error.distance
        elif orig_code and ref_code:
            code_status = "<span style='color: green;'>✓ Совпадает</span>"
            code_distance = 0
        else:
            code_status = "<span style='color: gray;'>— Не проверялся</span>"
            code_distance = "—"
        
        # Получаем startdate для текущего эталона
        ref_startdate = ""
        if error.reference_index is not None and hasattr(self, '_reference_data') and error.reference_index < len(self._reference_data):
            ref_startdate = self._reference_data[error.reference_index].get('startdate', '')
            if ref_startdate:
                try:
                    from datetime import datetime
                    date_obj = datetime.strptime(ref_startdate.split()[0], '%Y-%m-%d')
                    ref_startdate = date_obj.strftime('%d.%m.%Y')
                except:
                    pass
        
        # Создаем расширенный диалог
        dialog = QDialog(self)
        dialog.setWindowTitle("🔍 Детали ошибки текста с альтернативами")
        dialog.setMinimumSize(1200, 650)
        
        main_layout = QVBoxLayout(dialog)
        main_layout.setSpacing(8)
        main_layout.setContentsMargins(10, 10, 10, 10)
        
        # === КОМПАКТНАЯ КАРТОЧКА С ОСНОВНОЙ ИНФОРМАЦИЕЙ ===
        info_card = QWidget()
        info_card.setObjectName("infoCard")
        info_card.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Minimum)
        info_card.setMaximumHeight(30)
        info_layout = QHBoxLayout(info_card)
        info_layout.setSpacing(8)
        info_layout.setContentsMargins(6, 2, 6, 2)
        
        # Вся информация в одной строке
        section_label = QLabel(f"<b>📂 Раздел:</b> {section}")
        info_layout.addWidget(section_label)
        
        # Distance для текста
        text_distance_color = "#28a745" if error.distance < 10 else ("#ffc107" if error.distance < 50 else "#dc3545")
        text_status = QLabel(f"<b>📝 Текст:</b> <span style='color: {text_distance_color}; font-weight: bold;'>D:{error.distance}</span>")
        info_layout.addWidget(text_status)
        
        # Статус кода
        code_status_label = QLabel(f"<b>🔢 Код:</b> {code_status}")
        info_layout.addWidget(code_status_label)
        
        # Дата
        if ref_startdate:
            date_label = QLabel(f"<b>📅 Дата:</b> {ref_startdate}")
            info_layout.addWidget(date_label)
        
        # info_layout.addStretch()
        
        main_layout.addWidget(info_card)
        
        # === ОСНОВНОЙ SPLITTER ===
        splitter = QSplitter(Qt.Horizontal)
        
        # === ЛЕВАЯ ПАНЕЛЬ: Сравнение текстов ===
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(0, 0, 0, 0)
        
        left_label = QLabel("<b>🔄 Сравнение текстов:</b>")
        left_label.setObjectName("sectionHeader")
        left_layout.addWidget(left_label)
        
        details = f"""
        <div style='padding: 10px;'>
        <table cellpadding='8' style='border: 1px solid #dee2e6; border-radius: 4px; width: 100%;'>
        <tr style='background-color: #fff3cd;'>
            <td style='width: 120px;'><b>📄 В проекте:</b></td>
            <td>{error.original_text}</td>
        </tr>
        <tr style='background-color: #d1ecf1;'>
            <td><b>✅ Эталон:</b></td>
            <td>{error.reference_text if error.reference_text else '—'}</td>
        </tr>
        <tr>
            <td><b>🔢 Код (проект):</b></td>
            <td><code>{orig_code if orig_code else '—'}</code></td>
        </tr>
        <tr>
            <td><b>🔢 Код (справочник):</b></td>
            <td><code>{ref_code if ref_code else '—'}</code></td>
        </tr>
        </table>
        <p style='margin-top: 10px; color: #6c757d; font-size: 9pt;'>
        💡 <i>Красным выделены различия, зеленым — исправления</i>
        </p>
        </div>
        """
        
        text_browser = QTextBrowser()
        text_browser.setHtml(details)
        text_browser.setOpenExternalLinks(False)
        left_layout.addWidget(text_browser)
        
        splitter.addWidget(left_widget)
        
        # === ПРАВАЯ ПАНЕЛЬ: ВКЛАДКИ С АЛЬТЕРНАТИВАМИ ===
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setContentsMargins(0, 0, 0, 0)
        
        right_label = QLabel("<b>🔎 Альтернативные варианты:</b>")
        right_label.setObjectName("sectionHeader")
        right_layout.addWidget(right_label)
        
        # Вкладки для двух режимов
        tabs = QTabWidget()
        tabs.setObjectName("alternativesTabs")
        
        # ВКЛАДКА 1: Актуальные (дедублированные)
        tab1 = QWidget()
        tab1_layout = QVBoxLayout(tab1)
        alternatives_table_dedup = QTableWidget()
        alternatives_table_dedup.setColumnCount(2)
        alternatives_table_dedup.setHorizontalHeaderLabels(["D / Код / Дата / ОКТМО", "Текст"])
        alternatives_table_dedup.horizontalHeader().setStretchLastSection(False)
        alternatives_table_dedup.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        alternatives_table_dedup.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        alternatives_table_dedup.setSelectionBehavior(QTableWidget.SelectRows)
        alternatives_table_dedup.setSelectionMode(QTableWidget.SingleSelection)
        alternatives_table_dedup.setEditTriggers(QTableWidget.NoEditTriggers)
        alternatives_table_dedup.setWordWrap(True)
        alternatives_table_dedup.setAlternatingRowColors(True)
        alternatives_table_dedup.setObjectName("alternativesTable")
        
        # Устанавливаем делегаты
        alternatives_table_dedup.setItemDelegateForColumn(0, WordWrapDelegate(alternatives_table_dedup))
        alternatives_table_dedup.setItemDelegateForColumn(1, HtmlDelegate(alternatives_table_dedup))
        
        # Находим альтернативные варианты с дедупликацией
        alternatives = self._find_alternative_variants(error.original_text, error.original_index, deduplicate=True)
        alternatives_table_dedup.setRowCount(len(alternatives))
        
        # Импортируем функцию для поиска различий
        from utils.text_validation.text_comparator import find_differences
        
        for idx, (distance, text, code, ref_idx, startdate, ppocode, pponame) in enumerate(alternatives):
            # Объединяем Distance, код, startdate и ОКТМО в одну ячейку
            distance_code_text = f"D:{distance}\n{code}"
            if startdate:
                # Форматируем дату для отображения
                try:
                    from datetime import datetime
                    date_obj = datetime.strptime(startdate.split()[0], '%Y-%m-%d')
                    formatted_date = date_obj.strftime('%d.%m.%Y')
                    distance_code_text += f"\n{formatted_date}"
                except:
                    distance_code_text += f"\n{startdate}"
            # Добавляем ОКТМО
            if ppocode:
                distance_code_text += f"\n{ppocode}"
            distance_code_item = QTableWidgetItem(distance_code_text)
            
            # Генерируем HTML с подсветкой различий между исходным текстом и альтернативным
            # Сравниваем альтернативный текст (показываем) с исходным текстом из проекта (эталон)
            diff_indices, corrections = find_differences(text, error.original_text)
            # Для альтернатив инвертируем цвета: отличия = правильные варианты (зеленый)
            html_text = self._richtext_to_html(text, diff_indices, corrections, invert_colors=True)
            
            text_item = QTableWidgetItem()
            text_item.setData(Qt.DisplayRole, html_text)
            text_item.setData(Qt.UserRole, text)  # Сохраняем чистый текст для замены
            
            # Проверяем совпадение кодов
            from utils.text_validation.code_processor import normalize_classification_code
            normalized_orig_code = normalize_classification_code(orig_code) if orig_code else ""
            normalized_alt_code = normalize_classification_code(code) if code else ""
            codes_match = normalized_orig_code and normalized_alt_code and normalized_orig_code == normalized_alt_code
            
            # Цветовое кодирование по distance или совпадению кодов
            if codes_match:
                # Если коды совпадают - зеленый фон с более ярким оттенком
                distance_code_item.setBackground(QColor("#28a745"))  # Яркий зеленый
                distance_code_item.setForeground(Qt.white)
            elif distance < 10:
                distance_code_item.setBackground(Qt.green)
                distance_code_item.setForeground(Qt.white)
            elif distance < 50:
                distance_code_item.setBackground(Qt.yellow)
            else:
                distance_code_item.setBackground(Qt.red)
                distance_code_item.setForeground(Qt.white)
            
            # Выравнивание по верхнему левому углу
            distance_code_item.setTextAlignment(Qt.AlignTop | Qt.AlignLeft)
            text_item.setTextAlignment(Qt.AlignTop | Qt.AlignLeft)
            
            alternatives_table_dedup.setItem(idx, 0, distance_code_item)
            alternatives_table_dedup.setItem(idx, 1, text_item)
            
            # Сохраняем индекс и другие данные в userData
            distance_code_item.setData(Qt.UserRole, ref_idx)
            distance_code_item.setData(Qt.UserRole + 1, code)
            distance_code_item.setData(Qt.UserRole + 2, distance)
            distance_code_item.setData(Qt.UserRole + 3, ppocode)
            distance_code_item.setData(Qt.UserRole + 4, pponame)
        
        # Подгоняем высоту строк под содержимое
        alternatives_table_dedup.resizeRowsToContents()
        
        # Обработчик двойного клика для просмотра деталей
        alternatives_table_dedup.cellDoubleClicked.connect(
            lambda r, c: self._show_alternative_details(alternatives_table_dedup, r, section)
        )
        
        tab1_layout.addWidget(alternatives_table_dedup)
        
        # Кнопка применения выбранного варианта
        apply_btn_dedup = QPushButton("✓ Применить выбранный вариант")
        apply_btn_dedup.setObjectName("buttonApplyAlternative")
        apply_btn_dedup.clicked.connect(
            lambda: self._apply_selected_variant(dialog, alternatives_table_dedup, error, row)
        )
        tab1_layout.addWidget(apply_btn_dedup)
        
        tabs.addTab(tab1, f"✨ Актуальные ({len(alternatives)})")
        
        # ВКЛАДКА 2: Все варианты (с историческими)
        tab2 = QWidget()
        tab2_layout = QVBoxLayout(tab2)
        
        alternatives_table_all = QTableWidget()
        alternatives_table_all.setColumnCount(2)
        alternatives_table_all.setHorizontalHeaderLabels(["D / Код / Дата / ОКТМО", "Текст"])
        alternatives_table_all.horizontalHeader().setStretchLastSection(False)
        alternatives_table_all.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        alternatives_table_all.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        alternatives_table_all.setSelectionBehavior(QTableWidget.SelectRows)
        alternatives_table_all.setSelectionMode(QTableWidget.SingleSelection)
        alternatives_table_all.setEditTriggers(QTableWidget.NoEditTriggers)
        alternatives_table_all.setWordWrap(True)
        alternatives_table_all.setAlternatingRowColors(True)
        alternatives_table_all.setObjectName("alternativesTable")
        
        # Устанавливаем делегаты
        alternatives_table_all.setItemDelegateForColumn(0, WordWrapDelegate(alternatives_table_all))
        alternatives_table_all.setItemDelegateForColumn(1, HtmlDelegate(alternatives_table_all))
        
        # Находим альтернативные варианты без дедупликации
        alternatives_all = self._find_alternative_variants(error.original_text, error.original_index, deduplicate=False)
        alternatives_table_all.setRowCount(len(alternatives_all))
        
        for idx, (distance, text, code, ref_idx, startdate, ppocode, pponame) in enumerate(alternatives_all):
            # Объединяем Distance, код, startdate и ОКТМО в одну ячейку
            distance_code_text = f"D:{distance}\n{code}"
            if startdate:
                try:
                    from datetime import datetime
                    date_obj = datetime.strptime(startdate.split()[0], '%Y-%m-%d')
                    formatted_date = date_obj.strftime('%d.%m.%Y')
                    distance_code_text += f"\n{formatted_date}"
                except:
                    distance_code_text += f"\n{startdate}"
            # Добавляем ОКТМО
            if ppocode:
                distance_code_text += f"\n{ppocode}"
            distance_code_item = QTableWidgetItem(distance_code_text)
            
            # Генерируем HTML с подсветкой различий
            diff_indices, corrections = find_differences(text, error.original_text)
            html_text = self._richtext_to_html(text, diff_indices, corrections, invert_colors=True)
            
            text_item = QTableWidgetItem()
            text_item.setData(Qt.DisplayRole, html_text)
            text_item.setData(Qt.UserRole, text)
            
            # Проверяем совпадение кодов
            from utils.text_validation.code_processor import normalize_classification_code
            normalized_orig_code = normalize_classification_code(orig_code) if orig_code else ""
            normalized_alt_code = normalize_classification_code(code) if code else ""
            codes_match = normalized_orig_code and normalized_alt_code and normalized_orig_code == normalized_alt_code
            
            # Цветовое кодирование по distance или совпадению кодов
            if codes_match:
                # Если коды совпадают - зеленый фон с более ярким оттенком
                distance_code_item.setBackground(QColor("#28a745"))  # Яркий зеленый
                distance_code_item.setForeground(Qt.white)
            elif distance < 10:
                distance_code_item.setBackground(Qt.green)
                distance_code_item.setForeground(Qt.white)
            elif distance < 50:
                distance_code_item.setBackground(Qt.yellow)
            else:
                distance_code_item.setBackground(Qt.red)
                distance_code_item.setForeground(Qt.white)
            
            distance_code_item.setTextAlignment(Qt.AlignTop | Qt.AlignLeft)
            text_item.setTextAlignment(Qt.AlignTop | Qt.AlignLeft)
            
            alternatives_table_all.setItem(idx, 0, distance_code_item)
            alternatives_table_all.setItem(idx, 1, text_item)
            
            distance_code_item.setData(Qt.UserRole, ref_idx)
            distance_code_item.setData(Qt.UserRole + 1, code)
            distance_code_item.setData(Qt.UserRole + 2, distance)
            distance_code_item.setData(Qt.UserRole + 3, ppocode)
            distance_code_item.setData(Qt.UserRole + 4, pponame)
        
        alternatives_table_all.resizeRowsToContents()
        
        alternatives_table_all.cellDoubleClicked.connect(
            lambda r, c: self._show_alternative_details(alternatives_table_all, r, section)
        )
        
        tab2_layout.addWidget(alternatives_table_all)
        
        apply_btn_all = QPushButton("✓ Применить выбранный вариант")
        apply_btn_all.setObjectName("buttonApplyAlternative")
        apply_btn_all.clicked.connect(
            lambda: self._apply_selected_variant(dialog, alternatives_table_all, error, row)
        )
        tab2_layout.addWidget(apply_btn_all)
        
        tabs.addTab(tab2, f"📚 Все варианты ({len(alternatives_all)})")
        
        right_layout.addWidget(tabs)
        
        splitter.addWidget(right_widget)
        
        # Устанавливаем пропорции: 55% левая, 45% правая
        splitter.setStretchFactor(0, 55)
        splitter.setStretchFactor(1, 45)
        
        main_layout.addWidget(splitter)
        
        # === ИНФОРМАЦИОННАЯ ПАНЕЛЬ С ПОДСКАЗКАМИ ===
        info_footer = QLabel()
        info_footer.setObjectName("infoFooter")
        info_footer_text = "<b>💡 Подсказка:</b> <b>⭐ Актуальные</b> - топ-5 из текущих версий справочника (без дублей по коду) | <br><b>📚 Все варианты</b> - топ-5 из всех версий (с историческими) | <br><b>Цвета:</b> <span style='background-color: #28a745; color: white; padding: 1px 4px;'>D&lt;10</span> <span style='background-color: #ffc107; padding: 1px 4px;'>10≤D&lt;50</span> <span style='background-color: #dc3545; color: white; padding: 1px 4px;'>D≥50</span>"
        info_footer.setText(info_footer_text)
        info_footer.setWordWrap(True)
        info_footer.setMaximumHeight(60)
        main_layout.addWidget(info_footer)
        
        # Кнопка закрытия
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        close_btn = QPushButton("✖ Закрыть")
        close_btn.setObjectName("buttonClose")
        close_btn.clicked.connect(dialog.accept)
        button_layout.addWidget(close_btn)
        main_layout.addLayout(button_layout)
        
        # Отложенный вызов для правильного расчета высоты строк после отображения диалога
        QTimer.singleShot(100, lambda: (alternatives_table_dedup.resizeRowsToContents(), alternatives_table_all.resizeRowsToContents()))
        
        dialog.exec_()
    
    def _find_alternative_variants(self, original_text: str, original_index: int, deduplicate: bool = True):
        """
        Поиск альтернативных вариантов из справочника (топ-5 по distance).
        
        Args:
            original_text: Исходный текст для поиска
            original_index: Индекс в исходных данных
            deduplicate: Если True, загружает справочник с дедупликацией; если False - все записи
        
        Returns:
            List[(distance, text, code, ref_index, startdate, ppocode, pponame)]
        """
        from Levenshtein import distance as levenshtein_distance
        
        # Загружаем справочник с нужным режимом дедупликации
        try:
            section = self.section_combo.currentText()
            section_table_map = {
                'Доходы': 'v_budgetclastypeinc_merged',
                'Расходы': 'v_budgetclascosts_merged',
                'Источники финансирования': 'v_budgetclassources_merged',
            }
            table_name = section_table_map.get(section)
            
            if not table_name:
                return []
            
            # Получаем параметры фильтрации
            ppocode_from_meta = getattr(self.main_window, 'current_ppocode', None)
            if ppocode_from_meta and ppocode_from_meta.strip():
                ppocode_filter = ["00000000", ppocode_from_meta]
            else:
                ppocode_filter = "00000000"
            
            filter_date = self.main_window.reference_date_edit.text().strip()
            if not filter_date:
                from datetime import datetime
                filter_date = datetime.now().strftime("%Y-%m-%d")
            
            # Загружаем справочник
            import sqlite3
            from utils.db_utils import get_filtered_view
            conn = sqlite3.connect(self.controller.db_manager.db_path)
            reference_df = get_filtered_view(conn, table_name, filter_date=filter_date, filter_ppocode=ppocode_filter, deduplicate=deduplicate)
            conn.close()
            
            if reference_df.empty:
                return []
            
            reference_names = reference_df['name'].astype(str).str.strip().tolist()
            reference_codes = self._extract_codes_from_reference(reference_df, section)
            reference_data = reference_df.to_dict('records')
            
        except Exception as e:
            logger.error(f"Ошибка загрузки альтернатив: {e}", exc_info=True)
            return []
        
        # Вычисляем distance для всех записей справочника
        variants = []
        for idx, ref_name in enumerate(reference_names):
            dist = levenshtein_distance(original_text, ref_name)
            ref_code = reference_codes[idx] if idx < len(reference_codes) else ""
            startdate = reference_data[idx].get('startdate', '') if idx < len(reference_data) else ""
            ppocode = reference_data[idx].get('ppocode', '') if idx < len(reference_data) else ""
            pponame = reference_data[idx].get('pponame', '') if idx < len(reference_data) else ""
            variants.append((dist, ref_name, ref_code, idx, startdate, ppocode, pponame))
        
        # Сортируем по distance и возвращаем топ-5 (СТАРАЯ РЕАЛИЗАЦИЯ - УДАЛИТЬ)
        variants.sort(key=lambda x: x[0])
        return variants[:5]
    
    def _find_alternative_variants(self, original_text: str, original_index: int, deduplicate: bool = True):
        """Поиск альтернативных вариантов из справочника (делегирует в общую функцию)"""
        section = self.section_combo.currentText()
        section_table_map = {
            'Доходы': 'v_budgetclastypeinc_merged',
            'Расходы': 'v_budgetclascosts_merged',
            'Источники финансирования': 'v_budgetclassources_merged',
        }
        table_name = section_table_map.get(section)
        
        if not table_name:
            return []
        
        # Получаем параметры фильтрации
        ppocode_from_meta = getattr(self.main_window, 'current_ppocode', None)
        if ppocode_from_meta and ppocode_from_meta.strip():
            ppocode_filter = ["00000000", ppocode_from_meta]
        else:
            ppocode_filter = "00000000"
        
        filter_date = self.main_window.reference_date_edit.text().strip()
        if not filter_date:
            from datetime import datetime
            filter_date = datetime.now().strftime("%Y-%m-%d")
        
        # Создаем функцию-экстрактор кодов
        def code_extractor(df):
            return self._extract_codes_from_reference(df, section)
        
        return find_alternative_variants(
            original_text=original_text,
            db_path=self.controller.db_manager.db_path,
            table_name=table_name,
            filter_date=filter_date,
            ppocode_filter=ppocode_filter,
            code_extractor=code_extractor,
            deduplicate=deduplicate,
            top_n=5
        )
    
    def _show_alternative_details(self, table: QTableWidget, row: int, section: str):
        """Показ детальной информации об альтернативном варианте"""
        if row < 0 or row >= table.rowCount():
            return
        
        distance_code_item = table.item(row, 0)
        text_item = table.item(row, 1)
        
        if not all([distance_code_item, text_item]):
            return
        
        # Извлекаем данные из UserRole
        ref_idx = distance_code_item.data(Qt.UserRole)
        distance = distance_code_item.data(Qt.UserRole + 2)
        code = distance_code_item.data(Qt.UserRole + 1)
        text = text_item.data(Qt.UserRole)  # Чистый текст
        
        # Получаем полную запись из справочника
        ref_record = None
        if ref_idx is not None and hasattr(self, '_reference_data') and ref_idx < len(self._reference_data):
            ref_record = self._reference_data[ref_idx]
        
        # Формируем детальное описание
        details = f"""<h3>Детали альтернативного варианта</h3>
        <table cellpadding='5' style='border: 1px solid #ccc;'>
        <tr><td><b>Раздел:</b></td><td>{section}</td></tr>
        <tr style='background-color: #e6f7ff;'><td><b>Текст из справочника:</b></td><td>{text}</td></tr>
        <tr><td><b>Код классификации:</b></td><td>{code if code else '—'}</td></tr>
        <tr><td><b>Distance от текста в проекте:</b></td><td><b>{distance}</b></td></tr>
        """
        
        # Добавляем все поля из справочника, если есть
        if ref_record:
            details += "<tr><td colspan='2'><hr/><b>Полная информация из справочника:</b></td></tr>"
            
            # Определяем читаемые названия полей
            field_names = {
                'startdate': 'Дата начала действия',
                'enddate': 'Дата окончания действия',
                'level': 'Уровень',
                'stagename': 'Этап',
                'budgetname': 'Наименование бюджета',
                'pponame': 'ОКТМО (наименование)',
                'ppocode': 'ОКТМО (код)',
                'year': 'Год',
                'inctypecode': 'Код типа дохода',
                'incsubtypecode': 'Код подтипа дохода',
                'analyticalgroupcode': 'Код аналитической группы',
                'concatenated_code': 'Полный код',
                'rzpr': 'Раздел/Подраздел',
                'kcsr': 'КЦСР',
                'kvr': 'КВР',
                'grbscode': 'Код ГРБС',
                'grbsname': 'Наименование ГРБС',
                'gabscode': 'Код администратора',
                'gabsname': 'Наименование администратора',
                'sourcecode': 'Код источника',
                'npa_id': 'Код НПА',
                'created_at': 'Дата создания',
            }
            
            for key, value in ref_record.items():
                if key != 'name' and key != 'id':  # name уже выведено, id не нужен
                    # Форматируем ключ (делаем читаемым)
                    display_key = field_names.get(key, key.replace('_', ' ').title())
                    display_value = str(value) if value is not None else '—'
                    details += f"<tr><td><i>{display_key}:</i></td><td>{display_value}</td></tr>"
        
        details += """</table>
        <p style='margin-top: 10px;'><i>ℹ️ Это один из возможных вариантов из справочника</i></p>
        """
        
        # Создаем диалог для просмотра
        detail_dialog = QDialog(self)
        detail_dialog.setWindowTitle("Детали альтернативного варианта")
        detail_dialog.setMinimumSize(500, 300)
        
        layout = QVBoxLayout(detail_dialog)
        
        text_browser = QTextBrowser()
        text_browser.setHtml(details)
        layout.addWidget(text_browser)
        
        # Кнопка закрытия
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        close_btn = QPushButton("Закрыть")
        close_btn.clicked.connect(detail_dialog.accept)
        button_layout.addWidget(close_btn)
        layout.addLayout(button_layout)
        
        detail_dialog.exec_()
    
    def _apply_selected_variant(self, parent_dialog: QDialog, table: QTableWidget, error, error_row: int):
        """Применение выбранного альтернативного варианта"""
        selected_rows = table.selectionModel().selectedRows()
        if not selected_rows:
            QMessageBox.warning(parent_dialog, "Предупреждение", "Выберите вариант из списка")
            return
        
        row = selected_rows[0].row()
        distance_code_item = table.item(row, 0)
        text_item = table.item(row, 1)
        
        if not all([distance_code_item, text_item]):
            return
        
        # Получаем чистый текст (без HTML подсветки) для замены
        new_text = text_item.data(Qt.UserRole)
        new_code = distance_code_item.data(Qt.UserRole + 1)
        new_distance = distance_code_item.data(Qt.UserRole + 2)
        new_ref_index = distance_code_item.data(Qt.UserRole)
        
        # Подтверждение
        msg = QMessageBox(parent_dialog)
        msg.setIcon(QMessageBox.Question)
        msg.setWindowTitle("Подтверждение")
        msg.setText(f"Заменить текущий эталон на:\n\n\"{new_text}\"\n\nDistance: {new_distance}")
        msg.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
        msg.setDefaultButton(QMessageBox.No)
        
        if msg.exec_() != QMessageBox.Yes:
            return
        
        # Обновляем ошибку в памяти (не сохраняем альтернативы в БД, только применяем выбранную)
        from utils.text_validation import ErrorInfo, CodeErrorInfo
        from utils.text_validation.text_comparator import find_differences
        from utils.text_validation.code_processor import compare_codes, normalize_classification_code
        
        # Пересчитываем различия для нового эталонного текста
        new_diff_indices, new_corrections = find_differences(error.original_text, new_text)
        
        # Пересчитываем code_error, если есть информация о коде в проекте и альтернативе
        new_code_error = None
        if error.original_index < len(self.orig_codes):
            orig_code = self.orig_codes[error.original_index]
            normalized_orig_code = normalize_classification_code(orig_code)
            
            # Если есть оригинальный код И код в альтернативе, сравниваем их
            if normalized_orig_code and new_code:
                # Сравниваем оригинальный код с новым кодом из альтернативы
                has_code_error, ref_code, code_dist, code_diff_idx, code_corrections, code_ref_idx = compare_codes(
                    normalized_orig_code, [new_code], max_distance=200
                )
                
                if has_code_error:
                    # Проверяем процент несовпадения
                    code_max_len = max(len(normalized_orig_code), len(ref_code)) if ref_code else len(normalized_orig_code)
                    code_mismatch_percent = (code_dist / code_max_len * 100) if code_max_len > 0 else 0
                    
                    if code_mismatch_percent > 50:
                        code_corrections = []
                    
                    new_code_error = CodeErrorInfo(
                        has_error=has_code_error,
                        ref_code=ref_code,
                        distance=code_dist,
                        diff_indices=code_diff_idx,
                        corrections=code_corrections,
                        code_ref_idx=0  # Индекс 0, так как мы передали список из одного элемента
                    )
        
        # Создаем обновленную ошибку с новым reference_text и пересчитанными различиями
        updated_error = ErrorInfo(
            original_text=error.original_text,
            reference_text=new_text,
            distance=new_distance,
            diff_indices=new_diff_indices,
            corrections=new_corrections,
            original_index=error.original_index,
            reference_index=new_ref_index,
            code_error=new_code_error
        )
        
        # Обновляем в списке ошибок
        # Находим позицию в полном списке ошибок
        for i, err in enumerate(self.errors):
            if err.original_index == error.original_index:
                self.errors[i] = updated_error
                break
        
        # Обновляем таблицу
        self._update_errors_table()
        
        # Сохраняем изменения в БД
        if self.controller.current_project and self.controller.current_revision_id:
            self._save_errors_to_db()
        
        QMessageBox.information(parent_dialog, "Успешно", "Эталонный текст обновлен")
        parent_dialog.accept()
    
    def _save_errors_to_db(self):
        """Сохранение текущих ошибок в БД"""
        if not self.controller.current_project or not self.controller.current_revision_id:
            return
        
        section = self.section_combo.currentText()
        try:
            logger.info(f"Сохранение {len(self.errors)} ошибок текстов в БД...")
            self.controller.db_manager.save_text_validation_errors(
                project_id=self.controller.current_project.id,
                revision_id=self.controller.current_revision_id,
                section=section,
                errors=self.errors
            )
            logger.info("Ошибки текстов сохранены в БД успешно")
            
            # Обновляем ключ последней загрузки
            self._last_loaded_key = (
                self.controller.current_project.id,
                self.controller.current_revision_id,
                section
            )
        except Exception as e:
            logger.error(f"Ошибка при сохранении ошибок текстов в БД: {e}", exc_info=True)
    
    def export_errors(self):
        """Экспорт ошибок в Excel"""
        if not self.errors:
            QMessageBox.warning(self, "Предупреждение", "Нет ошибок для экспорта")
            return
        
        # Диалог сохранения файла
        file_name, _ = QFileDialog.getSaveFileName(
            self,
            "Сохранить отчет",
            f"Ошибки_текстов_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            "Excel Files (*.xlsx)"
        )
        
        if not file_name:
            return
        
        try:
            from utils.text_validation import create_error_report, save_workbook
            
            # Получаем данные проекта для наименований справочника
            project = self.controller.current_project
            if not project or not project.data:
                QMessageBox.warning(self, "Предупреждение", "Нет данных проекта")
                return
            
            section = self.section_combo.currentText()
            section_data = []
            
            if section == "Доходы":
                section_data = project.data.get('income_data', [])
            elif section == "Расходы":
                section_data = project.data.get('outcome_data', [])
            elif section == "Источники финансирования":
                section_data = project.data.get('source_financing_deficit_data', [])
            
            reference_names = [row.get('наименование_показателя', '') for row in section_data]
            
            # Создаем Excel отчет
            wb = create_error_report(
                errors=self.errors,
                klass_codes=self.orig_codes,
                reference_codes=self.reference_codes,
                reference_names=reference_names
            )
            
            save_workbook(wb, file_name)
            
            QMessageBox.information(
                self,
                "Успех",
                f"Отчет сохранен: {file_name}"
            )
        except Exception as e:
            logger.error(f"Ошибка экспорта: {e}", exc_info=True)
            QMessageBox.critical(self, "Ошибка", f"Не удалось экспортировать: {str(e)}")

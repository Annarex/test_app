"""
Диалог для обновления онлайн справочников из API бюджетной системы
"""
from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QTableWidget, QTableWidgetItem,
    QPushButton, QLabel, QHeaderView, QMessageBox, QProgressBar,
    QAbstractItemView, QWidget, QStackedWidget, QMenu
)
from PyQt5.QtCore import Qt, QEvent, QObject
from PyQt5.QtWidgets import QApplication
from PyQt5.QtGui import QFont
from datetime import datetime
from typing import Optional, Dict
from logger import logger

from models.database import DatabaseManager
from views.column_visibility_dialog import ColumnVisibilityDialog
from services.budget_references_service import (
    BudgetReferencesService,
    URL_TO_TABLE,
    get_default_filters
)


# Названия справочников для отображения
REFERENCE_NAMES = {
    'oktmo': 'ОКТМО',
    'budgetclastypeinc': 'Классификаторы доходов бюджета ФУ',
    'budgetclassubtypincmo': 'Классификаторы доходов бюджета МО',
    'budgetclasgabs': 'Администраторы бюджета ФУ',
    'budgetclasgabsmo': 'Администраторы бюджета МО',
    'budgetclasgrbs': 'Распорядители бюджета ФУ',
    'budgetclasgrbsmo': 'Распорядители бюджета МО',
    'budgetclascosts': 'Классификаторы расходов бюджета ФУ',
    'budgetclascostsmo': 'Классификаторы расходов бюджета МО',
    'budgetclasgaiffb': 'Источники финансирования дефицита ФУ',
    'budgetclasgaifmo': 'Источники финансирования дефицита МО',
    'budgetclassources': 'Классификаторы источников финансирования ФУ',
    'budgetclassourcesmo': 'Классификаторы источников финансирования МО',
}

# Сопоставление таблиц с URL
TABLE_TO_URL = {v: k for k, v in URL_TO_TABLE.items()}



class BudgetReferencesUpdateDialog(QDialog):
    """Диалог для обновления онлайн справочников"""
    
    def __init__(self, db_manager: DatabaseManager, parent=None):
        super().__init__(parent)
        self.db_manager = db_manager
        self.service = BudgetReferencesService(db_manager.db_path)
        self.update_status = {}  # table_name -> 'none', 'success', 'error', 'cancelled'
        self.updated_tables = set()  # Множество обновленных таблиц в текущей сессии
        self._is_cancelled = {}  # table_name -> bool (для отмены обновления)
        self.update_stats = {}  # table_name -> {'success': bool, 'api_count': int, 'db_count': int, 'diff': int}
        self.is_mass_update = False  # Флаг массового обновления
        
        self.setWindowTitle("Обновление онлайн справочников")
        self.setMinimumWidth(800)
        self.setMinimumHeight(500)
        
        self.init_ui()
        self.load_references_info()
    
    def _find_table_row(self, table_name: str) -> Optional[int]:
        """Находит номер строки в таблице по имени справочника"""
        for i in range(self.table.rowCount()):
            item = self.table.item(i, 0)
            if item and item.data(Qt.UserRole) == table_name:
                return i
        return None
    
    def _get_db_count(self, table_name: str) -> int:
        """Получает количество записей из БД"""
        try:
            import sqlite3
            conn = sqlite3.connect(self.db_manager.db_path)
            cursor = conn.cursor()
            cursor.execute(f"SELECT COUNT(*) FROM {table_name}")
            count = cursor.fetchone()[0]
            conn.close()
            return count
        except Exception as e:
            logger.warning(f"Не удалось получить количество записей из БД для {table_name}: {e}")
            return 0
    
    def _show_confirmation_dialog(self, table_name: str, api_count: int, db_count: int, error: str = None) -> bool:
        """Показывает диалог подтверждения обновления"""
        name = REFERENCE_NAMES.get(table_name, table_name)
        if error:
            text = f"Справочник: {name}\n\nОшибка при проверке количества записей: {error}\nПродолжить обновление?"
        elif api_count == 0:
            text = f"Справочник: {name}\n\nЗаписей в API: {api_count} (возможна ошибка получения)\nЗаписей в БД: {db_count}\n\nПродолжить обновление?"
        else:
            text = f"Справочник: {name}\n\nЗаписей в API: {api_count}\nЗаписей в БД: {db_count}\n\nПродолжить обновление?"
        
        return QMessageBox.question(self, "Подтверждение обновления", text, QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes) == QMessageBox.Yes
    
    def _update_button_ui(self, table_name: str, status: str, enabled: bool = True):
        """Обновляет UI кнопки обновления"""
        row = self._find_table_row(table_name)
        if row is None:
            return
        
        stacked_widget = self.table.cellWidget(row, 4)
        if not stacked_widget:
            return
        
        stacked_widget.setCurrentIndex(0)
        update_btn = stacked_widget.currentWidget()
        if update_btn:
            update_btn.setEnabled(enabled)
            self.update_status[table_name] = status
            update_btn.setProperty('update_status', status)
            
            status_to_object_name = {
                'success': 'buttonSuccess',
                'error': 'buttonDanger',
                'cancelled': 'buttonWarning',
                'loading': 'buttonLoading',
                'none': 'buttonPrimary'
            }
            update_btn.setObjectName(status_to_object_name.get(status, 'buttonPrimary'))
            update_btn.style().unpolish(update_btn)
            update_btn.style().polish(update_btn)
    
    def init_ui(self):
        """Инициализация интерфейса"""
        layout = QVBoxLayout(self)
        
        # Заголовок
        title_label = QLabel("Обновление справочников из API Электронного бюджета")
        title_font = QFont()
        title_font.setPointSize(12)
        title_font.setBold(True)
        title_label.setFont(title_font)
        layout.addWidget(title_label)
        
        # Таблица справочников
        self.table = QTableWidget()
        self.table.setColumnCount(5)
        self.table.setHorizontalHeaderLabels([
            "Наименование справочника",
            "Дата загрузки",
            "Дата последнего обновления",
            "Количество записей",
            "Действие"
        ])
        header = self.table.horizontalHeader()
        header.setStretchLastSection(False)
        header.setSectionResizeMode(0, QHeaderView.Stretch)
        for col in range(1, 5):
            header.setSectionResizeMode(col, QHeaderView.ResizeToContents)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.verticalHeader().setDefaultSectionSize(36)
        layout.addWidget(self.table)
        
        # Кнопки управления
        buttons_layout = QHBoxLayout()
        
        self.update_all_btn = QPushButton("Обновить все справочники")
        self.update_all_btn.setObjectName("buttonSuccessLarge")
        self.update_all_btn.clicked.connect(self.update_all_references)
        buttons_layout.addWidget(self.update_all_btn)
        
        buttons_layout.addStretch()
        
        close_btn = QPushButton("Закрыть")
        close_btn.clicked.connect(self.accept)
        buttons_layout.addWidget(close_btn)
        
        layout.addLayout(buttons_layout)
    
    def load_references_info(self):
        """Загрузка информации о справочниках"""
        # Получаем информацию из БД
        references_info = self.db_manager.get_all_budget_references_info()
        info_dict = {item['table_name']: item for item in references_info}
        
        # Заполняем таблицу
        self.table.setRowCount(len(REFERENCE_NAMES))
        
        for row, (table_name, display_name) in enumerate(REFERENCE_NAMES.items()):
            info = info_dict.get(table_name, {})
            name_item = QTableWidgetItem(display_name)
            name_item.setData(Qt.UserRole, table_name)
            self.table.setItem(row, 0, name_item)
            self.table.setItem(row, 1, QTableWidgetItem(info.get('created_at', 'Никогда')))
            self.table.setItem(row, 2, QTableWidgetItem(info.get('last_update_date', 'Никогда')))
            count = info.get('records_count', 0) or self._get_db_count(table_name)
            self.table.setItem(row, 3, QTableWidgetItem(str(count)))
            
            # Используем QStackedWidget для переключения между кнопкой, прогресс-баром и кнопкой остановки
            stacked_widget = QStackedWidget()
            
            update_btn = QPushButton("Обновить")
            update_btn.setObjectName("buttonPrimary")
            update_btn.clicked.connect(lambda checked, tn=table_name: self.update_reference(tn))
            stacked_widget.addWidget(update_btn)
            
            progress_bar = QProgressBar()
            progress_bar.setMinimum(0)
            progress_bar.setMaximum(100)
            progress_bar.setTextVisible(True)
            progress_bar.setFormat("%p% (%v / %m)")
            progress_bar.setObjectName("progressBarLoading")
            stacked_widget.addWidget(progress_bar)
            
            cancel_btn = QPushButton("Остановить")
            cancel_btn.setObjectName("buttonDanger")
            cancel_btn.clicked.connect(lambda checked, tn=table_name: self.cancel_update(tn))
            stacked_widget.addWidget(cancel_btn)
            
            # Устанавливаем event filter для переключения между прогресс-баром и кнопкой остановки
            class StackedWidgetEventFilter(QObject):
                def __init__(self, stacked_widget_ref):
                    super().__init__()
                    self.stacked_widget = stacked_widget_ref
                
                def eventFilter(self, obj, event):
                    if event.type() == QEvent.Enter:
                        # При наведении на прогресс-бар переключаемся на кнопку остановки
                        if self.stacked_widget.currentIndex() == 1:
                            self.stacked_widget.setCurrentIndex(2)
                        return True
                    elif event.type() == QEvent.Leave:
                        # При потере фокуса возвращаемся на прогресс-бар
                        if self.stacked_widget.currentIndex() == 2:
                            self.stacked_widget.setCurrentIndex(1)
                        return True
                    return super().eventFilter(obj, event)
            
            event_filter = StackedWidgetEventFilter(stacked_widget)
            progress_bar.installEventFilter(event_filter)
            cancel_btn.installEventFilter(event_filter)
            
            stacked_widget.setProperty('table_name', table_name)
            self.table.setCellWidget(row, 4, stacked_widget)
            self.update_status[table_name] = 'none'
    
    def cancel_update(self, table_name: str):
        """Отмена обновления справочника"""
        if table_name in self._is_cancelled:
            self._is_cancelled[table_name] = True
            logger.info(f"Отмена обновления справочника {table_name}")
    
    def update_reference(self, table_name: str):
        """Обновление одного справочника"""
        if table_name in self._is_cancelled:
            QMessageBox.warning(self, "Предупреждение", f"Справочник '{REFERENCE_NAMES.get(table_name, table_name)}' уже обновляется")
            return
        
        # Сбрасываем статус только при начале нового обновления
        # После завершения статус останется результирующим (success/error)
        if table_name not in self._is_cancelled:
            self.updated_tables.clear()
        
        url = TABLE_TO_URL.get(table_name)
        if not url:
            QMessageBox.warning(self, "Ошибка", f"Не найден URL для справочника '{REFERENCE_NAMES.get(table_name, table_name)}'")
            return
        
        filters = get_default_filters("21").get(table_name)
        self._check_and_start_update(table_name, url, filters)
    
    def _check_and_start_update(self, table_name: str, url: str, filters: Optional[Dict], skip_confirmation: bool = False):
        """Проверка количества записей и запуск обновления синхронно"""
        try:
            # Получаем количество записей из API и БД
            api_count = self.service.get_record_count(url, filters)
            db_count = self._get_db_count(table_name)
            
            logger.info(f"[{table_name}] Проверка количества: API={api_count}, БД={db_count}")
            
            # Если skip_confirmation=True, пропускаем диалог и сразу запускаем обновление
            if skip_confirmation:
                logger.info(f"[{table_name}] Запуск обновления без подтверждения: API={api_count}, БД={db_count}")
                self._update_reference(table_name, url, filters, api_count=api_count, db_count=db_count)
                return
            
            # Показываем диалог подтверждения
            if self._show_confirmation_dialog(table_name, api_count, db_count):
                self._update_reference(table_name, url, filters, api_count=api_count, db_count=db_count)
                
        except Exception as e:
            logger.error(f"Ошибка при проверке количества записей для {table_name}: {e}")
            
            # При ошибке проверки запрашиваем подтверждение или запускаем без него
            if skip_confirmation or self._show_confirmation_dialog(table_name, 0, 0, error=str(e)):
                logger.warning(f"[{table_name}] Запуск обновления после ошибки проверки")
                self._update_reference(table_name, url, filters, api_count=0, db_count=0)
    
    def on_update_progress(self, table_name: str, current: int, total: int):
        """Обработчик обновления прогресса"""
        row = self._find_table_row(table_name)
        if row is None:
            return
        
        stacked_widget = self.table.cellWidget(row, 4)
        if not stacked_widget:
            return
        
        # Убеждаемся, что показывается прогресс-бар (страница 1) или кнопка остановки (страница 2)
        current_index = stacked_widget.currentIndex()
        if current_index not in [1, 2]:
            stacked_widget.setCurrentIndex(1)
        
        # Обновляем прогресс-бар (он всегда на странице 1)
        progress_bar = stacked_widget.widget(1)
        if not progress_bar:
            return
        
        # Если total неизвестен или равен 0, показываем индикатор загрузки
        if total and total > 0:
            percentage = int((current / total) * 100)
            progress_bar.setMaximum(total)
            progress_bar.setValue(current)
            progress_bar.setFormat(f"{percentage}% ({current} / {total})")
        else:
            # Показываем индикатор загрузки без точного прогресса
            progress_bar.setMaximum(0)  # Неопределенный режим
            progress_bar.setValue(0)
            progress_bar.setFormat(f"Загрузка... ({current} записей)")
    
    def on_update_finished(self, table_name: str, success: bool, message: str):
        """Обработчик завершения обновления"""
        self._update_button_ui(table_name, 'success' if success else 'cancelled')
        
        # Обновляем статистику
        if table_name in self.update_stats:
            self.update_stats[table_name]['success'] = success
            if success:
                # Вычисляем разницу для успешных обновлений
                api_count = self.update_stats[table_name].get('api_count', 0)
                db_count = self.update_stats[table_name].get('db_count', 0)
                self.update_stats[table_name]['diff'] = api_count - db_count
                
                # Сохраняем количество реально загруженных записей
                try:
                    loaded_count = int(message.split(": ")[-1])
                except:
                    loaded_count = 0
                self.update_stats[table_name]['loaded_count'] = loaded_count
        
        if success:
            row = self._find_table_row(table_name)
            if row is not None:
                try:
                    count = int(message.split(": ")[-1])
                except:
                    count = 0
                
                if count == 0:
                    count = self._get_db_count(table_name)
                
                last_info = self.db_manager.get_budget_reference_last_info(table_name)
                if last_info:
                    self.table.setItem(row, 1, QTableWidgetItem(last_info.get('created_at', 'Никогда')))
                    self.table.setItem(row, 2, QTableWidgetItem(last_info.get('update_at', datetime.now().strftime('%d.%m.%Y %H:%M:%S'))))
                else:
                    current_time = datetime.now().strftime('%d.%m.%Y %H:%M:%S')
                    self.table.setItem(row, 1, QTableWidgetItem(current_time))
                    self.table.setItem(row, 2, QTableWidgetItem(current_time))
                
                self.table.setItem(row, 3, QTableWidgetItem(str(count)))
        
        logger.info(f"Обновление справочника '{REFERENCE_NAMES.get(table_name, table_name)}' завершено: {message}")
        # При одиночном обновлении проверяем сразу, но не очищаем множество
        self._check_all_updates_completed(is_mass_update=False)
        # Показываем статистику для одиночного обновления
        if not self.is_mass_update:
            self._show_update_summary(table_name)
    
    def on_update_error(self, table_name: str, error_message: str):
        """Обработчик ошибки обновления"""
        self._update_button_ui(table_name, 'error')
        
        # Обновляем статистику для неудачных обновлений
        if table_name in self.update_stats:
            self.update_stats[table_name]['success'] = False
        
        logger.error(f"Ошибка при обновлении справочника '{REFERENCE_NAMES.get(table_name, table_name)}': {error_message}")
        self._check_all_updates_completed(is_mass_update=False)
        # Показываем статистику для одиночного обновления
        if not self.is_mass_update:
            self._show_update_summary(table_name)
    
    def update_all_references(self):
        """Обновление всех справочников последовательно (синхронно)"""
        reply = QMessageBox.question(
            self, 
            "Подтверждение", 
            "Вы уверены, что хотите обновить все справочники?\nВсе обновления будут выполнены последовательно.",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if reply != QMessageBox.Yes:
            return
        
        # Очищаем множество обновленных таблиц и статистику для новой сессии
        self.updated_tables.clear()
        self.update_stats.clear()
        self.is_mass_update = True
        
        # Отключаем кнопку обновления всех
        self.update_all_btn.setEnabled(False)
        self.update_all_btn.setText("Обновление...")
        
        # Получаем фильтры по умолчанию
        default_filters = get_default_filters("21")
        
        # Обновляем все справочники последовательно
        for table_name in REFERENCE_NAMES.keys():
            url = TABLE_TO_URL.get(table_name)
            if url:
                # Получаем количество записей из API и БД перед обновлением
                try:
                    api_count = self.service.get_record_count(url, default_filters.get(table_name))
                    db_count = self._get_db_count(table_name)
                    # Сохраняем статистику перед обновлением
                    self.update_stats[table_name] = {
                        'success': False,
                        'api_count': api_count,
                        'db_count': db_count,
                        'diff': 0
                    }
                except Exception as e:
                    logger.warning(f"Не удалось получить количество записей для {table_name}: {e}")
                    self.update_stats[table_name] = {
                        'success': False,
                        'api_count': 0,
                        'db_count': 0,
                        'diff': 0
                    }
                
                self._check_and_start_update(table_name, url, default_filters.get(table_name), skip_confirmation=True)
                QApplication.processEvents()
        
        # Включаем кнопку обратно после завершения всех обновлений
        self.update_all_btn.setEnabled(True)
        self.update_all_btn.setText("Обновить все справочники")
        
        # Включаем все кнопки обновления (но сохраняем их результирующий статус)
        for table_name in REFERENCE_NAMES.keys():
            row = self._find_table_row(table_name)
            if row is not None:
                stacked_widget = self.table.cellWidget(row, 4)
                if stacked_widget:
                    stacked_widget.setCurrentIndex(0)
                    update_btn = stacked_widget.currentWidget()
                    if update_btn:
                        update_btn.setEnabled(True)
        
        # Проверяем завершение всех обновлений и показываем итоговое сообщение
        self._check_all_updates_completed(is_mass_update=True)
        self._show_update_summary()
        self.is_mass_update = False
    
    def _update_reference(self, table_name: str, url: str, filters: Optional[Dict], api_count: Optional[int] = None, db_count: Optional[int] = None):
        """Синхронное обновление справочника"""
        row = self._find_table_row(table_name)
        if row is None:
            return
        
        # Сохраняем статистику перед обновлением, если еще не сохранена
        # При массовом обновлении значения уже сохранены в update_all_references
        if table_name not in self.update_stats:
            self.update_stats[table_name] = {
                'success': False,
                'api_count': api_count or 0,
                'db_count': db_count or 0,
                'diff': 0
            }
        # При массовом обновлении не перезаписываем значения, чтобы использовать те, что были получены до обновления
        # Для одиночного обновления обновляем значения, если они переданы
        elif not self.is_mass_update:
            if api_count is not None:
                self.update_stats[table_name]['api_count'] = api_count
            if db_count is not None:
                self.update_stats[table_name]['db_count'] = db_count
        
        self._is_cancelled[table_name] = False
        
        stacked_widget = self.table.cellWidget(row, 4)
        if stacked_widget:
            # Устанавливаем светло-зеленый стиль для кнопки в процессе загрузки
            update_btn = stacked_widget.widget(0)
            if update_btn:
                self.update_status[table_name] = 'loading'
                update_btn.setProperty('update_status', 'loading')
                update_btn.setObjectName("buttonLoading")
                update_btn.style().unpolish(update_btn)
                update_btn.style().polish(update_btn)
            stacked_widget.setCurrentIndex(1)  # Переключаемся на прогресс-бар
            progress_bar = stacked_widget.widget(1)
            if progress_bar:
                progress_bar.setMaximum(0)  # Неопределенный режим для начала
                progress_bar.setValue(0)
                progress_bar.setFormat("Загрузка...")
            QApplication.processEvents()  # Обновляем UI сразу
        
        logger.info(f"[{table_name}] Начало синхронного обновления: API={api_count}, БД={db_count}")
        
        try:
            def progress_callback(current, total):
                QApplication.processEvents()  # Обновляем UI
                self.on_update_progress(table_name, current, total)
            
            count = self.service.fill_table_from_url(
                url=url,
                table_name=table_name,
                filters=filters,
                clear_existing=False,
                api_count=api_count,
                db_count=db_count,
                progress_callback=progress_callback,
                cancel_check=lambda: self._is_cancelled.get(table_name, False)
            )
            
            if self._is_cancelled.get(table_name, False):
                self.on_update_finished(table_name, False, "Загрузка прервана пользователем")
            else:
                self.on_update_finished(table_name, True, f"Обновлено записей: {count}")
                # Добавляем в updated_tables только если данные были реально перезаписаны (count > 0)
                if count > 0:
                    self.updated_tables.add(table_name)
                
        except Exception as e:
            error_msg = str(e)
            logger.error(f"Ошибка при обновлении {table_name}: {e}", exc_info=True)
            self.on_update_error(table_name, error_msg)
        
        finally:
            # Удаляем флаг отмены
            if table_name in self._is_cancelled:
                del self._is_cancelled[table_name]
    
    def _check_all_updates_completed(self, is_mass_update: bool = False):
        """Проверка завершения всех обновлений"""
        # Обновляем уровни только если была реально перезаписана таблица budgetclassubtypincmo
        # Проверяем, что таблица была обновлена И данные были перезаписаны (loaded_count > 0)
        if 'budgetclassubtypincmo' in self.updated_tables:
            # Проверяем статистику, чтобы убедиться, что данные были реально перезаписаны
            stats = self.update_stats.get('budgetclassubtypincmo', {})
            was_replaced = False
            
            # Проверяем, была ли перезапись данных
            if stats.get('success', False):
                # Если loaded_count > 0, значит данные были реально загружены и перезаписаны
                loaded_count = stats.get('loaded_count', 0)
                if loaded_count > 0:
                    was_replaced = True
            
            if was_replaced:
                logger.info(f"Таблица budgetclassubtypincmo была перезаписана (загружено {stats.get('loaded_count', 0)} записей). Запускаем обновление уровней.")
                self._update_levels_after_all_updates()
            else:
                logger.info("Таблица budgetclassubtypincmo не была перезаписана (данные не были загружены). Пропускаем обновление уровней.")
        
        # Очищаем множество только после массового обновления всех справочников
        if is_mass_update:
            self.updated_tables.clear()
    
    def _update_levels_after_all_updates(self):
        """Обновление уровней после завершения всех обновлений справочников"""
        try:
            fu_count = self._get_db_count('budgetclastypeinc')
            if fu_count == 0:
                logger.info("В таблице budgetclastypeinc нет записей, пропускаем обновление уровней")
                return
            
            logger.info(f"Найдено {fu_count} записей в budgetclastypeinc. Запускаем пересчет уровней.")
            from services.budget_level_processor import BudgetLevelProcessor
            BudgetLevelProcessor(self.db_manager.db_path).process_and_update_levels()
            logger.info("Пересчет уровней в budgetclassubtypincmo завершен")
        except Exception as e:
            logger.error(f"Ошибка при обновлении уровней в budgetclassubtypincmo: {e}", exc_info=True)
    
    def _show_update_summary(self, single_table_name: Optional[str] = None):
        """Показывает итоговое сообщение после обновления справочников (одиночного или массового)
        
        Args:
            single_table_name: Имя таблицы для одиночного обновления (если None, то массовое обновление)
        """
        if not self.update_stats:
            return
        
        # Для одиночного обновления показываем статистику только для одной таблицы
        if single_table_name and not self.is_mass_update:
            stats_to_process = {single_table_name: self.update_stats.get(single_table_name, {})}
        else:
            stats_to_process = self.update_stats
        
        # Подсчитываем статистику
        total_success_diff = 0  # Сумма разниц для успешных обновлений
        total_success_api_count = 0  # Сумма api_count для успешных обновлений
        total_failed_api_count = 0  # Сумма api_count для неудачных обновлений
        success_count = 0  # Количество успешных обновлений
        
        for table_name, stats in stats_to_process.items():
            if not stats:
                continue
            if stats.get('success', False):
                # Для успешных: добавляем разницу (api_count - db_count) и api_count
                diff = stats.get('diff', 0)
                api_count = stats.get('api_count', 0)
                total_success_diff += diff
                total_success_api_count += api_count
                success_count += 1
            else:
                # Для неудачных: добавляем api_count
                api_count = stats.get('api_count', 0)
                total_failed_api_count += api_count
        
        # Формируем сообщение
        message_parts = []
        
        # Если разница равна 0, значит обновление не требовалось
        if total_success_diff == 0:
            if total_failed_api_count == 0:
                # Нет изменений и нет ошибок
                message = "Обновление таблиц не требуется."
            else:
                # Нет изменений, но есть ошибки
                message_parts.append("Обновление таблиц не требуется.")
                message_parts.append(f"Неудачно обновлено: {total_failed_api_count}")
                message = "\n".join(message_parts)
        else:
            # Есть изменения - показываем успешные обновления
            if success_count > 0:
                message_parts.append(f"Успешно обновлено записей: {total_success_api_count} ({total_success_diff:+d})")
            
            # Показываем неудачные обновления, если они были
            if total_failed_api_count > 0:
                message_parts.append(f"Неудачно обновлено: {total_failed_api_count}")
            
            if not message_parts:
                message = "Обновление завершено. Нет изменений."
            else:
                message = "\n".join(message_parts)
        
        QMessageBox.information(
            self,
            "Результаты обновления",
            message
        )
        
        # Очищаем статистику после показа сообщения
        if self.is_mass_update:
            # При массовом обновлении очищаем всю статистику
            self.update_stats.clear()
        elif single_table_name:
            # При одиночном обновлении очищаем статистику только для текущей таблицы
            if single_table_name in self.update_stats:
                del self.update_stats[single_table_name]
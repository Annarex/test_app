"""
Диалог для конвертации документов DOCX в Excel.
"""
from PyQt5.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel, 
                             QPushButton, QFileDialog, QLineEdit, QMessageBox,
                             QGroupBox, QApplication, QListWidget, QListWidgetItem,
                             QCheckBox, QProgressBar)
from PyQt5.QtCore import Qt
import os
from utils.docx_to_excel_converter import DocxToExcelConverter


class DocxToExcelDialog(QDialog):
    """Диалог для конвертации DOCX в Excel."""
    
    def __init__(self, parent=None):
        """
        Args:
            parent: Родительское окно
        """
        super().__init__(parent)
        self.setWindowTitle("Конвертация DOCX в Excel")
        self.setMinimumWidth(700)
        self.setMinimumHeight(500)
        self.setModal(True)
        
        self.docx_files = []  # Список файлов для конвертации
        self.output_folder = None
        
        self._init_ui()
    
    def _init_ui(self):
        """Инициализация интерфейса."""
        layout = QVBoxLayout()
        
        # Группа выбора исходных файлов
        source_group = QGroupBox("Исходные файлы (DOCX)")
        source_layout = QVBoxLayout()
        source_group.setLayout(source_layout)
        
        # Кнопки управления файлами
        buttons_layout = QHBoxLayout()
        
        self.add_files_btn = QPushButton("➕ Добавить файлы")
        self.add_files_btn.clicked.connect(self._add_files)
        buttons_layout.addWidget(self.add_files_btn)
        
        self.remove_file_btn = QPushButton("✖ Удалить выбранный")
        self.remove_file_btn.clicked.connect(self._remove_selected_file)
        self.remove_file_btn.setEnabled(False)
        buttons_layout.addWidget(self.remove_file_btn)
        
        self.clear_all_btn = QPushButton("🗑 Очистить все")
        self.clear_all_btn.clicked.connect(self._clear_all_files)
        self.clear_all_btn.setEnabled(False)
        buttons_layout.addWidget(self.clear_all_btn)
        
        buttons_layout.addStretch()
        source_layout.addLayout(buttons_layout)
        
        # Список файлов
        self.files_list = QListWidget()
        self.files_list.setAlternatingRowColors(True)
        self.files_list.setSelectionMode(QListWidget.ExtendedSelection)
        self.files_list.itemSelectionChanged.connect(self._on_selection_changed)
        source_layout.addWidget(self.files_list)
        
        layout.addWidget(source_group)
        
        # Группа выбора папки назначения
        target_group = QGroupBox("Папка назначения (Excel)")
        target_layout = QVBoxLayout()
        target_group.setLayout(target_layout)
        
        # Поле пути к папке
        folder_layout = QHBoxLayout()
        self.output_folder_edit = QLineEdit()
        self.output_folder_edit.setPlaceholderText("Папка для сохранения Excel файлов (по умолчанию - рядом с исходными)...")
        self.output_folder_edit.setReadOnly(True)
        folder_layout.addWidget(self.output_folder_edit)
        
        self.browse_folder_btn = QPushButton("📁 Обзор...")
        self.browse_folder_btn.clicked.connect(self._browse_output_folder)
        folder_layout.addWidget(self.browse_folder_btn)
        
        target_layout.addLayout(folder_layout)
        target_group.setLayout(target_layout)
        layout.addWidget(target_group)
        
        # Группа настроек конвертации
        options_group = QGroupBox("Параметры конвертации")
        options_layout = QVBoxLayout()
        options_group.setLayout(options_layout)
        
        self.remove_empty_rows_checkbox = QCheckBox("Убрать пустые строки между таблицами")
        self.remove_empty_rows_checkbox.setChecked(False)
        self.remove_empty_rows_checkbox.setToolTip(
            "Если включено, пустые строки до и после таблиц не будут добавляться в Excel"
        )
        options_layout.addWidget(self.remove_empty_rows_checkbox)
        
        self.use_converted_folder_checkbox = QCheckBox("Сохранять в папку data/converted")
        self.use_converted_folder_checkbox.setChecked(False)
        self.use_converted_folder_checkbox.setToolTip(
            "Если включено, все файлы будут сохраняться в папку data/converted\n"
            "Если выключено - сохранение рядом с исходным файлом или в выбранную папку"
        )
        self.use_converted_folder_checkbox.stateChanged.connect(self._on_converted_folder_changed)
        options_layout.addWidget(self.use_converted_folder_checkbox)
        
        layout.addWidget(options_group)
        
        # Информация
        info_label = QLabel(
            "ℹ️ Конвертер извлечет текст и таблицы из документов Word\n"
            "и сохранит их в формате Excel с сохранением структуры.\n"
            "Можно добавить несколько файлов для пакетной конвертации."
        )
        info_label.setWordWrap(True)
        info_label.setStyleSheet("color: #555; padding: 10px; background-color: #f0f0f0; border-radius: 5px;")
        layout.addWidget(info_label)
        
        # Прогресс-бар
        self.progress_label = QLabel("")
        self.progress_label.setVisible(False)
        layout.addWidget(self.progress_label)
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setTextVisible(True)
        layout.addWidget(self.progress_bar)
        
        # Кнопки действий
        buttons_layout = QHBoxLayout()
        buttons_layout.addStretch()
        
        self.convert_btn = QPushButton("🔄 Конвертировать все")
        self.convert_btn.setEnabled(False)
        self.convert_btn.clicked.connect(self._start_conversion)
        self.convert_btn.setDefault(True)
        self.convert_btn.setStyleSheet("font-weight: bold; padding: 8px 16px;")
        buttons_layout.addWidget(self.convert_btn)
        
        self.close_btn = QPushButton("Закрыть")
        self.close_btn.clicked.connect(self.close)
        buttons_layout.addWidget(self.close_btn)
        
        layout.addLayout(buttons_layout)
        
        self.setLayout(layout)
    
    def _add_files(self):
        """Добавление файлов DOCX."""
        file_paths, _ = QFileDialog.getOpenFileNames(
            self,
            "Выберите файлы DOCX",
            "",
            "Документы Word (*.docx);;Документы Word (*.doc);;Все файлы (*.*)"
        )
        
        if file_paths:
            for file_path in file_paths:
                # Проверяем, не добавлен ли уже этот файл
                if file_path not in self.docx_files:
                    self.docx_files.append(file_path)
                    # Добавляем только имя файла в список для отображения
                    item = QListWidgetItem(os.path.basename(file_path))
                    item.setToolTip(file_path)  # Полный путь в подсказке
                    self.files_list.addItem(item)
            
            self._update_buttons_state()
    
    def _remove_selected_file(self):
        """Удаление выбранных файлов из списка."""
        selected_items = self.files_list.selectedItems()
        if not selected_items:
            return
        
        for item in selected_items:
            file_path = item.toolTip()
            if file_path in self.docx_files:
                self.docx_files.remove(file_path)
            self.files_list.takeItem(self.files_list.row(item))
        
        self._update_buttons_state()
    
    def _clear_all_files(self):
        """Очистка всего списка файлов."""
        reply = QMessageBox.question(
            self,
            "Подтверждение",
            "Очистить весь список файлов?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            self.docx_files.clear()
            self.files_list.clear()
            self._update_buttons_state()
    
    def _on_selection_changed(self):
        """Обработка изменения выбора в списке."""
        has_selection = len(self.files_list.selectedItems()) > 0
        self.remove_file_btn.setEnabled(has_selection)
    
    def _on_converted_folder_changed(self, state):
        """Обработчик изменения состояния checkbox для папки data/converted."""
        if state == Qt.Checked:
            # Устанавливаем путь к data/converted
            converted_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'data', 'converted')
            converted_path = os.path.normpath(converted_path)
            self.output_folder = converted_path
            self.output_folder_edit.setText(converted_path)
            # Блокируем кнопку обзора
            self.browse_folder_btn.setEnabled(False)
        else:
            # Разблокируем кнопку обзора и очищаем путь
            self.browse_folder_btn.setEnabled(True)
            self.output_folder = None
            self.output_folder_edit.clear()
    
    def _browse_output_folder(self):
        """Выбор папки для сохранения Excel файлов."""
        folder_path = QFileDialog.getExistingDirectory(
            self,
            "Выберите папку для сохранения Excel файлов",
            ""
        )
        
        if folder_path:
            self.output_folder = folder_path
            self.output_folder_edit.setText(folder_path)
    
    def _update_buttons_state(self):
        """Обновление состояния кнопок."""
        has_files = len(self.docx_files) > 0
        self.convert_btn.setEnabled(has_files)
        self.clear_all_btn.setEnabled(has_files)
        self.remove_file_btn.setEnabled(len(self.files_list.selectedItems()) > 0)
    
    def _start_conversion(self):
        """Запуск процесса конвертации."""
        if not self.docx_files:
            QMessageBox.warning(
                self,
                "Предупреждение",
                "Добавьте хотя бы один файл DOCX."
            )
            return
        
        converted_files = []
        failed_files = []
        
        # Получаем параметры конвертации
        remove_empty_rows = self.remove_empty_rows_checkbox.isChecked()
        
        # Блокируем кнопки во время конвертации
        self.convert_btn.setEnabled(False)
        self.add_files_btn.setEnabled(False)
        self.remove_file_btn.setEnabled(False)
        self.clear_all_btn.setEnabled(False)
        self.browse_folder_btn.setEnabled(False)
        
        # Показываем прогресс-бар
        self.progress_bar.setVisible(True)
        self.progress_label.setVisible(True)
        self.progress_bar.setMaximum(len(self.docx_files))
        self.progress_bar.setValue(0)
        self.progress_bar.repaint()
        self.progress_label.setText("Начало конвертации...")
        self.progress_label.repaint()
        QApplication.processEvents()
        
        try:
            converter = DocxToExcelConverter(remove_empty_rows=remove_empty_rows)
            
            for idx, docx_path in enumerate(self.docx_files, 1):
                try:
                    # Обновляем прогресс
                    file_name = os.path.basename(docx_path)
                    self.progress_label.setText(f"Обработка {idx} из {len(self.docx_files)}: {file_name}...")
                    self.progress_label.repaint()
                    self.progress_bar.setValue(idx - 1)
                    self.progress_bar.repaint()
                    QApplication.processEvents()  # Обновляем интерфейс
                    
                    # Определяем путь для Excel файла
                    if self.output_folder:
                        # Если указана выходная папка, сохраняем туда
                        # Создаем папку если её нет
                        os.makedirs(self.output_folder, exist_ok=True)
                        base_name = os.path.splitext(os.path.basename(docx_path))[0]
                        excel_path = os.path.join(self.output_folder, f"{base_name}.xlsx")
                    else:
                        # Иначе сохраняем рядом с исходным файлом
                        base_name = os.path.splitext(docx_path)[0]
                        excel_path = f"{base_name}.xlsx"
                    
                    # Конвертация
                    converter.convert(docx_path, excel_path)
                    converted_files.append((os.path.basename(docx_path), excel_path))
                    
                    # Обновляем прогресс после завершения
                    self.progress_bar.setValue(idx)
                    self.progress_bar.repaint()
                    self.progress_label.setText(f"Завершено {idx} из {len(self.docx_files)}")
                    self.progress_label.repaint()
                    QApplication.processEvents()  # Обновляем интерфейс
                    
                except Exception as e:
                    failed_files.append((os.path.basename(docx_path), str(e)))
            
            # Разблокируем кнопки
            self.convert_btn.setEnabled(True)
            self.add_files_btn.setEnabled(True)
            self.clear_all_btn.setEnabled(len(self.docx_files) > 0)
            self.browse_folder_btn.setEnabled(True)
            
            # Скрываем прогресс-бар
            self.progress_bar.setVisible(False)
            self.progress_label.setVisible(False)
            
            # Формируем сообщение о результатах
            if converted_files and not failed_files:
                # Все файлы успешно сконвертированы
                message = f"Успешно сконвертировано файлов: {len(converted_files)}\n\n"
                if len(converted_files) <= 5:
                    for file_name, excel_path in converted_files:
                        message += f"✓ {file_name}\n"
                else:
                    message += f"Все файлы успешно сконвертированы."
                
                QMessageBox.information(self, "Успех", message)
                
                # Предложение открыть папку с результатами
                if self.output_folder:
                    reply = QMessageBox.question(
                        self,
                        "Открыть папку?",
                        "Хотите открыть папку с результатами?",
                        QMessageBox.Yes | QMessageBox.No,
                        QMessageBox.Yes
                    )
                    
                    if reply == QMessageBox.Yes:
                        os.startfile(self.output_folder)
                
            elif converted_files and failed_files:
                # Частичный успех
                message = f"Сконвертировано: {len(converted_files)} из {len(self.docx_files)}\n\n"
                message += "Ошибки:\n"
                for file_name, error in failed_files:
                    message += f"✗ {file_name}: {error}\n"
                
                QMessageBox.warning(self, "Частичный успех", message)
                
            else:
                # Все файлы с ошибками
                message = "Не удалось сконвертировать ни один файл:\n\n"
                for file_name, error in failed_files:
                    message += f"✗ {file_name}: {error}\n"
                
                QMessageBox.critical(self, "Ошибка", message)
                
        except Exception as e:
            # Разблокируем кнопки
            self.convert_btn.setEnabled(True)
            self.add_files_btn.setEnabled(True)
            self.clear_all_btn.setEnabled(len(self.docx_files) > 0)
            self.browse_folder_btn.setEnabled(True)
            
            # Скрываем прогресс-бар
            self.progress_bar.setVisible(False)
            self.progress_label.setVisible(False)
            
            # Показываем общую ошибку
            QMessageBox.critical(
                self,
                "Ошибка",
                f"Критическая ошибка конвертации:\n{str(e)}"
            )

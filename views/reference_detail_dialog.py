"""
Диалог для отображения детальной информации о записи справочника
"""
from PyQt5.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QFormLayout, QTextEdit, QScrollArea, QWidget, QFrame, QSizePolicy
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont, QPalette


class ReferenceDetailDialog(QDialog):
    """Диалог для отображения детальной информации о записи справочника"""
    
    def __init__(self, row_data: dict, column_headers: list, parent=None):
        super().__init__(parent)
        self.row_data = row_data
        self.column_headers = column_headers
        
        self.setWindowTitle("Детальная информация о записи")
        self.setMinimumSize(900, 600)
        self.setModal(True)
        
        # Стили применяются глобально через QApplication
        
        self.init_ui()
    
    def init_ui(self):
        """Инициализация интерфейса"""
        layout = QVBoxLayout(self)
        layout.setSpacing(15)
        layout.setContentsMargins(20, 20, 20, 20)
        
        # Заголовок с рамкой
        header_frame = QFrame()
        header_frame.setObjectName("headerFrame")
        header_layout = QVBoxLayout(header_frame)
        header_layout.setContentsMargins(15, 10, 15, 10)
        
        title_label = QLabel("📋 Детальная информация о записи")
        title_font = QFont()
        title_font.setPointSize(14)
        title_font.setBold(True)
        title_label.setFont(title_font)
        title_label.setWordWrap(True)
        header_layout.addWidget(title_label)
        
        layout.addWidget(header_frame)
        
        # Область прокрутки для формы
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        # Стили для QScrollArea уже применены через общий QSS
        
        # Виджет с формой в рамке
        form_frame = QFrame()
        form_frame.setObjectName("formFrame")
        
        form_widget = QWidget()
        form_layout = QFormLayout(form_widget)
        form_layout.setSpacing(12)
        form_layout.setContentsMargins(20, 20, 20, 20)
        form_layout.setLabelAlignment(Qt.AlignRight | Qt.AlignTop)
        # Устанавливаем политику роста полей - поля не растягиваются автоматически
        form_layout.setFieldGrowthPolicy(QFormLayout.FieldsStayAtSizeHint)
        
        # Добавляем поля формы
        for idx, header in enumerate(self.column_headers):
            value = self.row_data.get(header, "")
            value_str = str(value) if value else ""
            
            # Создаем метку для названия поля с ограничением ширины и многострочным режимом
            label = QLabel(f"{header}:")
            label_font = QFont()
            label_font.setPointSize(9)
            label_font.setBold(True)
            label.setFont(label_font)
            # Ограничиваем ширину метки и включаем многострочный режим
            label.setMaximumWidth(350)  # Максимальная ширина метки
            label.setWordWrap(True)  # Многострочный режим
            label.setAlignment(Qt.AlignRight | Qt.AlignTop)  # Выравнивание по правому краю и сверху
            # Устанавливаем размерную политику - не растягиваемся по вертикали
            label.setSizePolicy(QSizePolicy.Preferred, QSizePolicy.Preferred)
            # Явно убираем обводку
            label.setStyleSheet("border: none; background-color: transparent;")
            
            # Создаем виджет для значения
            # Используем QTextEdit для длинных значений или если есть переносы строк
            if len(value_str) > 80 or '\n' in value_str or len(header) > 30:
                # Для длинных значений используем QTextEdit
                text_edit = QTextEdit()
                text_edit.setPlainText(value_str)
                text_edit.setReadOnly(True)
                # Устанавливаем минимальную ширину для QTextEdit
                text_edit.setMinimumWidth(600)
                # Устанавливаем размерную политику - растягиваемся по горизонтали
                text_edit.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
                # Вычисляем оптимальную высоту на основе содержимого
                doc = text_edit.document()
                doc.setTextWidth(600)  # Устанавливаем ширину документа для правильного расчета
                ideal_height = int(doc.size().height()) + 20  # Добавляем небольшой отступ
                # Ограничиваем максимальную высоту
                max_height = 150 if len(value_str) > 200 else 100
                text_edit.setMaximumHeight(min(ideal_height, max_height))
                # Устанавливаем минимальную высоту, чтобы не было слишком маленьким
                text_edit.setMinimumHeight(min(ideal_height, 50))
                # Устанавливаем objectName для применения специального стиля из QSS
                text_edit.setObjectName("formTextEdit")
                form_layout.addRow(label, text_edit)
            else:
                # Для коротких значений используем QLabel с многострочным режимом
                value_label = QLabel(value_str if value_str else "(пусто)")
                value_label.setWordWrap(True)  # Многострочный режим
                value_label.setTextInteractionFlags(Qt.TextSelectableByMouse | Qt.TextSelectableByKeyboard)
                # Устанавливаем минимальную ширину для лучшего отображения
                value_label.setMinimumWidth(600)
                # Устанавливаем размерную политику - растягиваемся по горизонтали
                value_label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
                # Явно убираем обводку
                value_label.setStyleSheet("border: none; background-color: transparent;")
                form_layout.addRow(label, value_label)
        
        # Разделитель перед кнопками
        separator = QFrame()
        separator.setFrameShape(QFrame.HLine)
        separator.setFrameShadow(QFrame.Sunken)
        separator.setObjectName("separator")
        form_layout.addRow(separator)
        
        form_frame_layout = QVBoxLayout(form_frame)
        form_frame_layout.setContentsMargins(0, 0, 0, 0)
        form_frame_layout.addWidget(form_widget)
        
        scroll_area.setWidget(form_frame)
        layout.addWidget(scroll_area)
        
        # Кнопки
        buttons_frame = QFrame()
        buttons_layout = QHBoxLayout(buttons_frame)
        buttons_layout.setContentsMargins(0, 0, 0, 0)
        buttons_layout.addStretch()
        
        close_btn = QPushButton("✕ Закрыть")
        close_btn.clicked.connect(self.accept)
        close_btn.setDefault(True)
        buttons_layout.addWidget(close_btn)
        
        layout.addWidget(buttons_frame)

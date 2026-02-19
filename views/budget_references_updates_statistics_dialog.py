"""
Диалог статистики обновления онлайн справочников
"""
from collections import defaultdict
from datetime import datetime
from typing import Dict, List, Tuple

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QColor, QPainter, QPen
from PyQt5.QtWidgets import (
    QComboBox,
    QDialog,
    QHBoxLayout,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QPushButton,
    QSplitter,
    QVBoxLayout,
    QWidget,
)

from models.database import DatabaseManager
from views.budget_references_update_dialog import REFERENCE_NAMES


class TimeSeriesChartWidget(QWidget):
    def __init__(self, title: str, y_title: str, parent=None):
        super().__init__(parent)
        self.title = title
        self.y_title = y_title
        self.series: Dict[str, List[Tuple[datetime, float]]] = {}
        self.setMinimumHeight(240)

    def set_series(self, series: Dict[str, List[Tuple[datetime, float]]]):
        self.series = series
        self.update()

    def paintEvent(self, event):
        super().paintEvent(event)
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)

        width = self.width()
        height = self.height()

        left = 60
        top = 34
        right = width - 20
        bottom = height - 38

        painter.setPen(QPen(QColor("#2D3748"), 1))
        painter.drawText(12, 22, self.title)

        all_points = [point for points in self.series.values() for point in points]
        if not all_points:
            painter.setPen(QPen(QColor("#718096"), 1))
            painter.drawRect(left, top, right - left, bottom - top)
            painter.drawText(left + 14, top + 26, "Нет данных за выбранный период")
            return

        min_x = min(point[0] for point in all_points)
        max_x = max(point[0] for point in all_points)
        min_y = min(point[1] for point in all_points)
        max_y = max(point[1] for point in all_points)

        if min_x == max_x:
            max_x = max_x.replace(second=(max_x.second + 1) % 60)
        if min_y == max_y:
            delta = 1 if min_y == 0 else abs(min_y) * 0.05
            min_y -= delta
            max_y += delta

        painter.setPen(QPen(QColor("#A0AEC0"), 1))
        painter.drawLine(left, bottom, right, bottom)
        painter.drawLine(left, top, left, bottom)

        painter.setPen(QPen(QColor("#4A5568"), 1))
        painter.drawText(8, top + (bottom - top) // 2, self.y_title)

        y_ticks = 5
        x_ticks = 5

        for i in range(y_ticks + 1):
            ratio = i / y_ticks
            y_val = min_y + (max_y - min_y) * (1 - ratio)
            y_pos = top + int((bottom - top) * ratio)
            painter.setPen(QPen(QColor("#E2E8F0"), 1))
            painter.drawLine(left, y_pos, right, y_pos)
            painter.setPen(QPen(QColor("#4A5568"), 1))
            painter.drawText(6, y_pos + 4, f"{y_val:.0f}")

        total_seconds = max((max_x - min_x).total_seconds(), 1)
        for i in range(x_ticks + 1):
            ratio = i / x_ticks
            x_pos = left + int((right - left) * ratio)
            tick_seconds = total_seconds * ratio
            tick_dt = min_x.timestamp() + tick_seconds
            tick_date = datetime.fromtimestamp(tick_dt)
            painter.setPen(QPen(QColor("#E2E8F0"), 1))
            painter.drawLine(x_pos, top, x_pos, bottom)
            painter.setPen(QPen(QColor("#4A5568"), 1))
            painter.drawText(x_pos - 26, bottom + 18, tick_date.strftime("%d.%m"))

        colors = [
            QColor("#3182CE"), QColor("#38A169"), QColor("#D69E2E"), QColor("#E53E3E"),
            QColor("#805AD5"), QColor("#319795"), QColor("#DD6B20"), QColor("#2B6CB0"),
        ]

        legend_x = right - 170
        legend_y = top + 8
        for idx, (name, points) in enumerate(self.series.items()):
            color = colors[idx % len(colors)]
            pen = QPen(color, 2)
            painter.setPen(pen)

            path_points = []
            for dt_value, y_value in points:
                x_ratio = (dt_value - min_x).total_seconds() / total_seconds
                y_ratio = (y_value - min_y) / (max_y - min_y)
                x = left + int((right - left) * x_ratio)
                y = bottom - int((bottom - top) * y_ratio)
                path_points.append((x, y))

            if len(path_points) == 1:
                x, y = path_points[0]
                painter.drawEllipse(x - 2, y - 2, 4, 4)
            else:
                for point_index in range(1, len(path_points)):
                    painter.drawLine(
                        path_points[point_index - 1][0],
                        path_points[point_index - 1][1],
                        path_points[point_index][0],
                        path_points[point_index][1],
                    )

            if idx < 6:
                painter.setPen(QPen(color, 2))
                painter.drawLine(legend_x, legend_y + idx * 16, legend_x + 14, legend_y + idx * 16)
                painter.setPen(QPen(QColor("#2D3748"), 1))
                painter.drawText(legend_x + 18, legend_y + 4 + idx * 16, name[:24])


class BudgetReferencesUpdatesStatisticsDialog(QDialog):
    PERIOD_OPTIONS = [
        ("7 дней", 7),
        ("30 дней", 30),
        ("90 дней", 90),
        ("180 дней", 180),
        ("365 дней", 365),
        ("За всё время", None),
    ]

    def __init__(self, db_manager: DatabaseManager, parent=None):
        super().__init__(parent)
        self.db_manager = db_manager

        self.setWindowTitle("Статистика обновления онлайн справочников")
        self.resize(1200, 760)

        self.tables_list: QListWidget = None
        self.period_combo: QComboBox = None
        self.records_chart: TimeSeriesChartWidget = None
        self.updates_chart: TimeSeriesChartWidget = None
        self.summary_label: QLabel = None

        self._init_ui()
        self._load_tables()
        self.refresh_charts()

    def _init_ui(self):
        root = QVBoxLayout(self)

        title = QLabel("Статистика обновления онлайн справочников")
        title.setStyleSheet("font-size: 15px; font-weight: 600;")
        root.addWidget(title)

        splitter = QSplitter(Qt.Horizontal)
        root.addWidget(splitter)

        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        left_layout.addWidget(QLabel("Таблицы справочников"))

        self.tables_list = QListWidget()
        self.tables_list.itemChanged.connect(self.refresh_charts)
        left_layout.addWidget(self.tables_list)

        select_buttons = QHBoxLayout()
        select_all_btn = QPushButton("Выбрать все")
        clear_btn = QPushButton("Снять выбор")
        select_all_btn.clicked.connect(self._select_all_tables)
        clear_btn.clicked.connect(self._clear_all_tables)
        select_buttons.addWidget(select_all_btn)
        select_buttons.addWidget(clear_btn)
        left_layout.addLayout(select_buttons)

        splitter.addWidget(left_panel)

        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)

        controls = QHBoxLayout()
        controls.addWidget(QLabel("Период:"))

        self.period_combo = QComboBox()
        for text, value in self.PERIOD_OPTIONS:
            self.period_combo.addItem(text, value)
        self.period_combo.setCurrentIndex(1)
        self.period_combo.currentIndexChanged.connect(self.refresh_charts)
        controls.addWidget(self.period_combo)

        controls.addStretch()

        refresh_btn = QPushButton("Обновить")
        refresh_btn.clicked.connect(self.refresh_charts)
        controls.addWidget(refresh_btn)

        right_layout.addLayout(controls)

        self.records_chart = TimeSeriesChartWidget("Количество записей по обновлениям", "Записей")
        self.updates_chart = TimeSeriesChartWidget("Количество обновлений по дням", "Обновлений")

        right_layout.addWidget(self.records_chart)
        right_layout.addWidget(self.updates_chart)

        self.summary_label = QLabel()
        right_layout.addWidget(self.summary_label)

        splitter.addWidget(right_panel)
        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)
        splitter.setSizes([280, 900])

        close_btn = QPushButton("Закрыть")
        close_btn.clicked.connect(self.accept)
        root.addWidget(close_btn, alignment=Qt.AlignRight)

    def _load_tables(self):
        self.tables_list.blockSignals(True)
        self.tables_list.clear()

        # Получаем количество записей обновлений для каждой таблицы
        import sqlite3
        updates_count = {}
        with sqlite3.connect(self.db_manager.db_path) as conn:
            cursor = conn.cursor()
            cursor.execute("""
                SELECT table_name, COUNT(*) as count
                FROM budget_references_updates
                GROUP BY table_name
            """)
            for row in cursor.fetchall():
                updates_count[row[0]] = row[1]

        for table_name, display_name in REFERENCE_NAMES.items():
            count = updates_count.get(table_name, 0)
            label = f"{display_name} ({count})"
            item = QListWidgetItem(label)
            item.setData(Qt.UserRole, table_name)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            item.setCheckState(Qt.Checked)
            self.tables_list.addItem(item)

        self.tables_list.blockSignals(False)

    def _selected_tables(self) -> List[str]:
        selected = []
        for index in range(self.tables_list.count()):
            item = self.tables_list.item(index)
            if item.checkState() == Qt.Checked:
                selected.append(item.data(Qt.UserRole))
        return selected

    def _select_all_tables(self):
        self.tables_list.blockSignals(True)
        for index in range(self.tables_list.count()):
            self.tables_list.item(index).setCheckState(Qt.Checked)
        self.tables_list.blockSignals(False)
        self.refresh_charts()

    def _clear_all_tables(self):
        self.tables_list.blockSignals(True)
        for index in range(self.tables_list.count()):
            self.tables_list.item(index).setCheckState(Qt.Unchecked)
        self.tables_list.blockSignals(False)
        self.refresh_charts()

    def refresh_charts(self):
        selected_tables = self._selected_tables()
        period_days = self.period_combo.currentData()

        if not selected_tables:
            self.records_chart.set_series({})
            self.updates_chart.set_series({})
            self.summary_label.setText("Выберите хотя бы одну таблицу справочника")
            return

        rows = self.db_manager.get_budget_references_updates_history(
            table_names=selected_tables,
            period_days=period_days,
        )

        records_series = defaultdict(list)
        updates_by_day = defaultdict(int)

        for row in rows:
            try:
                dt_value = datetime.strptime(row['update_at'], "%d.%m.%Y %H:%M:%S")
            except (TypeError, ValueError):
                continue

            table_name = row['table_name']
            table_label = REFERENCE_NAMES.get(table_name, table_name)
            records_series[table_label].append((dt_value, float(row['records_count'] or 0)))
            updates_by_day[dt_value.strftime("%Y-%m-%d")] += 1

        updates_series = {
            "Все выбранные таблицы": [
                (datetime.strptime(day_key, "%Y-%m-%d"), float(count))
                for day_key, count in sorted(updates_by_day.items(), key=lambda item: item[0])
            ]
        }

        filtered_records_series = {
            table_label: points
            for table_label, points in records_series.items()
            if points
        }

        self.records_chart.set_series(filtered_records_series)
        self.updates_chart.set_series(updates_series if updates_by_day else {})

        total_points = sum(len(points) for points in filtered_records_series.values())
        self.summary_label.setText(
            f"Выбрано таблиц: {len(selected_tables)} | Записей истории: {total_points}"
        )

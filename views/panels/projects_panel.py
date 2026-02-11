"""Панель проектов"""
from collections import defaultdict

from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
                             QLabel, QTreeWidget, QTreeWidgetItem, QMenu,
                             QMessageBox)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont
from logger import logger

# Роли данных в узле дерева
ROLE_PROJECT_ID = Qt.UserRole
ROLE_REVISION_ID = Qt.UserRole + 1
ROLE_IS_REVISION = Qt.UserRole + 2

DEFAULT_STRUCTURE = "Год>Проект>Форма>Период>Ревизия"

# Пресеты: id -> (подпись, строка структуры). В строке: > вложенность, + склейка подписи через пробел
TREE_PRESETS = {
    "full": ("Год → Проект → Форма → Период → Ревизия", "Год>Проект>Форма>Период>Ревизия"),
    "compact": ("Год → Проект → Ревизия", "Год>Проект>Ревизия"),
    "year_period_form_rev": ("Год → Период → Форма+Ревизия", "Год>Период>Форма+Ревизия"),
    "year_period_mo_form_rev": ("Год → Период → МО → Форма+Ревизия", "Год>Период>МО>Форма+Ревизия"),
    "year_period_mo_form_rev2": ("Год → Период → МО → Форма → Ревизия", "Год>Период>МО>Форма>Ревизия"),
    "year_mo_period_form_rev": ("Год → МО → Период → Форма+Ревизия", "Год>МО>Период>Форма+Ревизия"),
    "year_mo_period_proj_form_rev": ("Год → МО → Период → Проект+Форма+Ревизия", "Год>МО>Период>Проект+Форма+Ревизия"),
    "year_mo_period_proj_form_rev_1": ("Год → МО → Период → Проект → Форма+Ревизия", "Год>МО>Период>Проект>Форма+Ревизия"),
}
CONFIG_KEY_TREE_PRESET = "projects_tree_preset"


def parse_structure(s: str):
    """
    Разбор строки структуры: > вложенность, + склейка подписи через пробел.
    Возвращает список уровней, каждый уровень — список имён для подписи.
    """
    s = (s or "").strip()
    if not s:
        return parse_structure(DEFAULT_STRUCTURE)
    levels = []
    for part in s.split(">"):
        names = [n.strip() for n in part.split("+") if n.strip()]
        if names:
            levels.append(names)
    return levels if levels else parse_structure(DEFAULT_STRUCTURE)

def _record_from_base(base, form=None, period=None, rev=None):
    """Собрать одну запись для плоского списка (форма/период/рев опциональны)."""
    rec = {**base, "form_code": None, "form_name": None, "period_code": None, "period_name": None, "revision_id": None, "revision": None, "status": None}
    if form:
        rec["form_code"], rec["form_name"] = form.get("form_code"), form.get("form_name")
    if period:
        rec["period_code"] = period.get("period_code")
        rec["period_name"] = period.get("period_name") or period.get("period_code") or "—"
    if rev:
        rec["revision_id"], rec["revision"] = rev.get("revision_id"), rev.get("revision") or ""
        rec["status"] = rev.get("status")
    return rec

def flatten_tree_data(tree_data):
    """
    Иерархия год→проекты→формы→периоды→ревизии в плоский список записей.
    Проекты без форм дают одну запись с пустыми form/period/revision.
    """
    records = []
    for year_entry in tree_data:
        year = year_entry.get("year", "")
        for proj in year_entry.get("projects") or []:
            base = {
                "year": year,
                "project_id": proj.get("id"),
                "project_name": proj.get("name") or "—",
                "municipality": proj.get("municipality") or "—",
            }
            if not proj.get("forms"):
                records.append(_record_from_base(base))
                continue
            for form in proj["forms"]:
                for period in form.get("periods") or []:
                    revs = period.get("revisions") or []
                    if not revs:
                        records.append(_record_from_base(base, form, period))
                        continue
                    for rev in revs:
                        records.append(_record_from_base(base, form, period, rev))
    return records

def _key_for_name(record, name):
    n = (name or "").strip()
    if n == "Год":
        return record.get("year")
    if n == "Проект":
        return record.get("project_id")
    if n == "МО":
        return record.get("municipality")
    if n == "Форма":
        return (record.get("form_code"), record.get("form_name"))
    if n == "Период":
        return record.get("period_code")
    if n == "Ревизия":
        return record.get("revision_id")
    return record.get(n)

def _label_for_name(record, name):
    n = (name or "").strip()
    if n == "Год":
        return f"Год {record.get('year', '')}"
    if n == "Проект":
        return record.get("project_name") or "—"
    if n == "МО":
        return record.get("municipality") or "—"
    if n == "Форма":
        fc, fn = record.get("form_code"), record.get("form_name")
        return f"{fn}" if fn else f"{fc}"
    if n == "Период":
        return record.get("period_name") or record.get("period_code") or "—"
    if n == "Ревизия":
        if record.get("revision_id") is None and not record.get("revision"):
            return "Нет ревизий"
        icon = "✅" if record.get("status") == "calculated" else "📝"
        return f"{icon} рев. {record.get('revision') or ''}"
    return str(record.get(n, ""))

def _record_key(record, level_names):
    """Ключ группировки записи по уровню."""
    return tuple(_key_for_name(record, n) for n in level_names)

def _record_label(record, level_names):
    """Подпись узла: одно поле или склейка через пробел (+)."""
    return " ".join(_label_for_name(record, n) for n in level_names if _label_for_name(record, n))

def _group_by(records, level_names):
    """Сгруппировать записи по ключу уровня."""
    groups = defaultdict(list)
    for rec in records:
        groups[_record_key(rec, level_names)].append(rec)
    return dict(groups)


def _expand_tree_recursive(item):
    """Развернуть узел и всех потомков."""
    item.setExpanded(True)
    for i in range(item.childCount()):
        _expand_tree_recursive(item.child(i))


class ProjectsPanel:
    """Класс для управления панелью проектов"""
    
    def __init__(self, main_window):
        """
        Args:
            main_window: Ссылка на главное окно для доступа к контроллеру и обработчикам
        """
        self.main_window = main_window
        self.controller = main_window.controller
    
    def create_projects_panel(self) -> QWidget:
        """Создание панели проектов"""
        # Основная панель с содержимым
        inner_panel = QWidget()
        layout = QVBoxLayout(inner_panel)
        layout.setContentsMargins(6, 6, 2, 6)
        
        # Заголовок
        title_label = QLabel("Проекты")
        title_label.setFont(QFont("Arial", 12, QFont.Bold))
        layout.addWidget(title_label)
        
        # Кнопки управления проектами
        buttons_layout = QHBoxLayout()
        
        new_project_btn = QPushButton("Новый")
        new_project_btn.clicked.connect(self.main_window.show_new_project_dialog)
        buttons_layout.addWidget(new_project_btn)
        
        refresh_btn = QPushButton("Обновить")
        refresh_btn.clicked.connect(self.main_window.refresh_projects)
        buttons_layout.addWidget(refresh_btn)
        layout.addLayout(buttons_layout)

        # Дерево проектов
        self.projects_tree = QTreeWidget()
        self.projects_tree.setIndentation(10)
        self.projects_tree.setHeaderHidden(True)
        self.projects_tree.itemDoubleClicked.connect(self.on_project_tree_double_clicked)
        self.projects_tree.setContextMenuPolicy(Qt.CustomContextMenu)
        self.projects_tree.customContextMenuRequested.connect(self.show_project_context_menu)
        layout.addWidget(self.projects_tree)
        
        # Информация о проекте
        self.project_info_label = QLabel("Выберите проект")
        self.project_info_label.setWordWrap(True)
        layout.addWidget(self.project_info_label)
        
        # Сохраняем ссылку на дерево в главном окне
        self.main_window.projects_tree = self.projects_tree
        self.main_window.project_info_label = self.project_info_label
        
        # Контейнер, в котором слева основная панель, справа узкая кнопка-свертка
        container = QWidget()
        container_layout = QHBoxLayout(container)
        container_layout.setContentsMargins(0, 0, 0, 0)
        container_layout.setSpacing(0)

        container_layout.addWidget(inner_panel)

        # Узкая вертикальная кнопка на правом краю панели
        toggle_button = QPushButton("◀")
        toggle_button.setFixedWidth(14)
        toggle_button.setFlat(True)
        toggle_button.setFocusPolicy(Qt.NoFocus)
        toggle_button.setToolTip("Свернуть/развернуть панель проектов")
        toggle_button.clicked.connect(self.main_window.on_projects_side_button_clicked)
        container_layout.addWidget(toggle_button)

        self.projects_inner_panel = inner_panel
        self.projects_toggle_button = toggle_button
        self.main_window.projects_inner_panel = inner_panel
        self.main_window.projects_toggle_button = toggle_button

        return container

    def _db(self):
        """Доступ к менеджеру БД для конфига (опционально)."""
        return getattr(self.controller, "db_manager", None)

    def _load_tree_preset(self):
        """Загрузить сохранённый пресет дерева из конфига."""
        db = self._db()
        if db and hasattr(db, "load_config"):
            return db.load_config(CONFIG_KEY_TREE_PRESET) or "full"
        return "full"

    def _save_tree_preset(self, preset_id):
        """Сохранить выбранный пресет в конфиг."""
        db = self._db()
        if db and hasattr(db, "save_config"):
            db.save_config(CONFIG_KEY_TREE_PRESET, preset_id)

    def get_tree_preset_list(self):
        """Список пресетов дерева для меню: [(preset_id, label), ...]."""
        return [(pid, label) for pid, (label, _) in TREE_PRESETS.items()]

    def get_current_tree_preset(self):
        """Текущий пресет дерева (из конфига)."""
        return self._load_tree_preset()

    def set_tree_preset(self, preset_id):
        """Установить пресет дерева и обновить список."""
        if preset_id and preset_id in TREE_PRESETS:
            self._save_tree_preset(preset_id)
            self.update_projects_list(None)

    def _get_current_preset(self):
        """Текущий пресет (из конфига)."""
        return self._load_tree_preset()

    def update_projects_list(self, _projects):
        """Обновление дерева по данным контроллера и выбранному пресету."""
        self.projects_tree.clear()
        tree_data = self.controller.build_project_tree()
        preset_id = self._get_current_preset()
        self._build_tree_from_data(tree_data, preset_id)
        for i in range(self.projects_tree.topLevelItemCount()):
            _expand_tree_recursive(self.projects_tree.topLevelItem(i))

    def _build_tree_from_data(self, tree_data, preset_id):
        """Построить дерево по строке структуры пресета: > вложенность, + склейка подписи через пробел."""
        preset = TREE_PRESETS.get(preset_id, TREE_PRESETS["full"])
        structure_str = preset[1] if len(preset) > 1 else DEFAULT_STRUCTURE
        levels = parse_structure(structure_str)
        records = flatten_tree_data(tree_data)
        self._add_level_from_records(records, levels, 0, None)

    def _add_level_from_records(self, records, levels, level_index, parent_item):
        """
        Рекурсивно построить уровень: сгруппировать records по levels[level_index], создать узел на каждый ключ.
        levels — список уровней, каждый уровень — список имён (подпись может быть из нескольких полей через +).
        """
        if not records or level_index >= len(levels):
            return
        level_names = levels[level_index]
        groups = _group_by(records, level_names)

        if not groups:
            if parent_item:
                placeholder = "Нет данных"
                if level_index > 0 and level_names and (level_names[0] or "").strip() == "Ревизия":
                    placeholder = "Нет ревизий"
                parent_item.addChild(QTreeWidgetItem([placeholder]))
            return

        is_leaf = level_index == len(levels) - 1
        for key, group_records in sorted(groups.items(), key=lambda x: str(x[0])):
            rec = group_records[0]
            label = _record_label(rec, level_names)
            item = QTreeWidgetItem([label])
            pid = rec.get("project_id")
            rid = rec.get("revision_id")
            if pid is not None:
                item.setData(0, ROLE_PROJECT_ID, pid)
            if rid is not None:
                item.setData(0, ROLE_REVISION_ID, rid)
            if rec.get("revision_id") is not None:
                item.setData(0, ROLE_IS_REVISION, 1)

            if parent_item is None:
                self.projects_tree.addTopLevelItem(item)
            else:
                parent_item.addChild(item)

            if not is_leaf:
                self._add_level_from_records(group_records, levels, level_index + 1, item)

    def _resolve_ids(self, item):
        """Поднимаясь по дереву, возвращает (project_id, revision_id)."""
        proj_id = rev_id = None
        cur = item
        while cur:
            if proj_id is None:
                proj_id = cur.data(0, ROLE_PROJECT_ID)
            if rev_id is None:
                rev_id = cur.data(0, ROLE_REVISION_ID)
            if proj_id is not None and rev_id is not None:
                break
            cur = cur.parent()
        return proj_id, rev_id

    def _is_revision_item(self, item):
        """True, если узел — ревизия (установлено при построении дерева)."""
        return item.data(0, ROLE_IS_REVISION) == 1

    def on_project_tree_double_clicked(self, item, column):
        """Обработка двойного клика по дереву проектов"""
        project_id, revision_id = self._resolve_ids(item)
        if not project_id:
            return

        if self._is_revision_item(item):
            # Подтягиваем параметры формы из ревизии для последующей загрузки файлов
            self.controller.set_form_params_from_revision(revision_id)
            # Загружаем конкретную ревизию
            logger.info(f"Загрузка ревизии {revision_id} для проекта {project_id}")
            self.controller.load_revision(revision_id, project_id)
        else:
            # Клик по проекту/форме/периоду/заглушке — выбираем проект, чтобы можно было загрузить новую форму
            if project_id:
                self.controller.project_controller.load_project(project_id)
            else:
                logger.warning("Проект не определён для выбранного узла")

    def show_project_context_menu(self, position):
        """Контекстное меню для дерева проектов"""
        item = self.projects_tree.itemAt(position)
        if not item:
            return
        project_id, revision_id = self._resolve_ids(item)
        if not project_id:
            return

        is_revision = self._is_revision_item(item)
        menu = QMenu()
        edit_action = None
        edit_rev_action = None
        delete_rev_action = None
        delete_project_action = None

        # Если это узел ревизии
        if is_revision:
            # Для ревизии нужен revision_id для редактирования/удаления
            if revision_id is not None:
                edit_rev_action = menu.addAction("Редактировать ревизию")
                delete_rev_action = menu.addAction("Удалить ревизию")
            # Если revision_id не установлен (виртуальная ревизия из старой модели),
            # действия редактирования/удаления недоступны
        else:
            # Для узла проекта (не ревизии) показываем действия проекта
            edit_action = menu.addAction("Редактировать проект")
            delete_project_action = menu.addAction("Удалить проект")

        action = menu.exec_(self.projects_tree.mapToGlobal(position))

        if action == edit_action:
            self.main_window.edit_project(project_id)
        elif edit_rev_action is not None and action == edit_rev_action and revision_id:
            self.main_window.edit_revision(revision_id, project_id)
        elif delete_rev_action is not None and action == delete_rev_action and revision_id:
            reply = QMessageBox.question(
                self.main_window,
                "Подтверждение",
                "Вы уверены, что хотите удалить выбранную ревизию?",
                QMessageBox.Yes | QMessageBox.No,
            )
            if reply == QMessageBox.Yes:
                self.controller.delete_form_revision(revision_id)
                # После удаления ревизии обновляем дерево
                self.update_projects_list(None)
        elif action == delete_project_action:
            reply = QMessageBox.question(
                self.main_window,
                "Подтверждение",
                "Вы уверены, что хотите удалить проект (все ревизии)?",
                QMessageBox.Yes | QMessageBox.No,
            )
            if reply == QMessageBox.Yes:
                self.controller.delete_project(project_id)

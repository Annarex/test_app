"""Панель проектов"""
from PyQt5.QtWidgets import (QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
                             QLabel, QTreeWidget, QTreeWidgetItem, QMenu,
                             QMessageBox)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont
from logger import logger


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
        
        # Дерево проектов: Год -> Проект -> Форма -> Ревизия
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
    
    def update_projects_list(self, _projects):
        """Обновление дерева проектов по новой архитектуре MainController.build_project_tree"""
        self.projects_tree.clear()

        # Получаем структурированные данные от контроллера
        tree_data = self.controller.build_project_tree()

        for year_entry in tree_data:
            year_label = f"Год {year_entry['year']}"
            year_item = QTreeWidgetItem([year_label])
            self.projects_tree.addTopLevelItem(year_item)

            for proj in year_entry["projects"]:
                proj_item = QTreeWidgetItem([proj["name"]])
                # Сохраняем ID проекта на уровне узла проекта
                proj_item.setData(0, Qt.UserRole, proj["id"])
                year_item.addChild(proj_item)

                # Формы/периоды/ревизии (показываем даже пустые, с заглушками)
                if proj.get("forms"):
                    for form in proj["forms"]:
                        form_label = f"{form['form_name']} ({form['form_code']})"
                        form_item = QTreeWidgetItem([form_label])
                        proj_item.addChild(form_item)

                        periods = form.get("periods") or []
                        if not periods:
                            form_item.addChild(QTreeWidgetItem(["Нет периодов"]))
                            continue

                        for period in periods:
                            period_label = period.get("period_name") or period.get("period_code") or "—"
                            period_item = QTreeWidgetItem([period_label])
                            form_item.addChild(period_item)

                            revisions = period.get("revisions") or []
                            if revisions:
                                for rev in revisions:
                                    status_icon = "✅" if rev["status"] == "calculated" else "📝"
                                    rev_text = f"{status_icon} рев. {rev['revision']}"
                                    rev_item = QTreeWidgetItem([rev_text])
                                    rev_item.setData(0, Qt.UserRole, rev.get("project_id"))
                                    revision_id = rev.get("revision_id")
                                    rev_item.setData(0, Qt.UserRole + 1, revision_id)
                                    if revision_id:
                                        logger.debug(
                                            f"Сохранена ревизия в дереве: "
                                            f"revision_id={revision_id}, project_id={rev.get('project_id')}, revision={rev.get('revision')}"
                                        )
                                    period_item.addChild(rev_item)
                            else:
                                period_item.addChild(QTreeWidgetItem(["Нет ревизий"]))
                else:
                    # Совсем нет форм — заглушка
                    placeholder = QTreeWidgetItem(["Нет ревизий"])
                    proj_item.addChild(placeholder)

        # Разворачиваем верхние уровни (год, проект, форма, период)
        # Ревизии остаются свернутыми по умолчанию
        for i in range(self.projects_tree.topLevelItemCount()):
            year_item = self.projects_tree.topLevelItem(i)
            year_item.setExpanded(True)
            for j in range(year_item.childCount()):
                proj_item = year_item.child(j)
                proj_item.setExpanded(True)
                for k in range(proj_item.childCount()):
                    form_item = proj_item.child(k)
                    form_item.setExpanded(True)
                    for m in range(form_item.childCount()):
                        period_item = form_item.child(m)
                        period_item.setExpanded(True)

    def on_project_tree_double_clicked(self, item, column):
        """Обработка двойного клика по дереву проектов"""
        # Поднимаемся по дереву, чтобы найти project_id/revision_id даже при клике на заглушки
        def _resolve_ids(it):
            proj_id = None
            rev_id = None
            cur = it
            while cur:
                if proj_id is None:
                    proj_id = cur.data(0, Qt.UserRole)
                if rev_id is None:
                    rev_id = cur.data(0, Qt.UserRole + 1)
                if proj_id is not None and rev_id is not None:
                    break
                cur = cur.parent()
            return proj_id, rev_id

        project_id, revision_id = _resolve_ids(item)
        
        if not project_id:
            return
        
        # Определяем, является ли узел ревизией (ревизия имеет revision_id и является дочерним элементом периода)
        is_revision = False
        if revision_id is not None and revision_id != 0:
            # Проверяем структуру дерева: ревизия является дочерним элементом периода
            parent = item.parent()
            if parent and item.childCount() == 0:
                # Период является дочерним элементом формы
                grandparent = parent.parent() if parent else None
                if grandparent:
                    grandparent_text = grandparent.text(0).lower()
                    if "форма" in grandparent_text or "(" in grandparent_text:
                        is_revision = True
        
        if is_revision:
            # Подтягиваем параметры формы из ревизии для последующей загрузки файлов
            self.controller.set_form_params_from_revision(revision_id)
            # Загружаем конкретную ревизию
            logger.info(f"Загрузка ревизии {revision_id} для проекта {project_id}")
            self.controller.load_revision(revision_id, project_id)
        else:
            # Клик по проекту/форме/периоду/заглушке — выбираем проект, чтобы можно было загрузить новую форму
            if project_id:
                logger.debug(f"Выбор проекта {project_id}")
                self.controller.project_controller.load_project(project_id)
            else:
                logger.warning("Проект не определён для выбранного узла")

    def show_project_context_menu(self, position):
        """Контекстное меню для дерева проектов"""
        item = self.projects_tree.itemAt(position)
        if not item:
            return
        project_id = item.data(0, Qt.UserRole)
        revision_id = item.data(0, Qt.UserRole + 1)

        # Если нет ID проекта — контекстное меню не показываем
        if not project_id:
            return

        # Определяем, является ли узел ревизией
        # Структура дерева: Год -> Проект -> Форма -> Период -> Ревизия
        # Ревизия - это узел, который является дочерним элементом периода
        # и не имеет дочерних элементов
        is_revision = False
        
        # Проверяем структуру дерева: ревизия является дочерним элементом периода
        parent = item.parent()
        if parent and item.childCount() == 0:
            # Период является дочерним элементом формы
            grandparent = parent.parent() if parent else None
            if grandparent:
                # Проверяем, что дедушка - это форма (содержит "форма" или "(")
                grandparent_text = grandparent.text(0).lower()
                if "форма" in grandparent_text or "(" in grandparent_text:
                    # Родитель - период, значит текущий узел - ревизия
                    is_revision = True

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

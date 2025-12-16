from typing import List, Dict, Any
from collections import defaultdict

from PyQt5.QtCore import QObject

from logger import logger
from models.database import DatabaseManager
from controllers.project_controller import ProjectController


class TreeController(QObject):
    """
    Контроллер, отвечающий за построение дерева проектов:
    - построение иерархической структуры (Год → Проект → Форма → Период → Ревизии)
    """

    def __init__(self, db_manager: DatabaseManager, project_controller: ProjectController, parent=None):
        super().__init__(parent)
        self.db_manager = db_manager
        self.project_controller = project_controller

    def build_project_tree(self) -> List[Dict[str, Any]]:
        """
        Строит структуру для отображения в левой панели:
        [
          {
            "year": "2024",
            "projects": [
              {
                "id": ...,
                "name": "...",
                "municipality": "...",
                "forms": [
                  {
                    "form_code": "0503317",
                    "form_name": "...",
                    "periods": [
                      {
                        "period_code": "Q1",
                        "period_name": "...",
                        "revisions": [
                          {"revision": "1.0", "status": "parsed", "project_id": ..., "file_path": "..."},
                          ...
                        ]
                      }
                    ]
                  }
                ]
              }
            ]
          }
        ]
        """
        # Загружаем все проекты (историческая таблица)
        projects = self.project_controller.load_projects()

        # Загружаем справочники формы и периодов для отображения (один раз, не в цикле)
        form_types_meta = {ft.id: ft for ft in self.db_manager.load_form_types_meta()}

        # Для периодов удобнее индексировать по id
        periods_all = self.db_manager.load_periods()
        periods_by_id = {p.id: p for p in periods_all if p.id is not None}

        # Загружаем справочники годов и МО один раз (оптимизация: не в цикле)
        years_all = self.db_manager.load_years()
        years_by_id = {y.id: y for y in years_all if y.id is not None}
        
        municipalities_all = self.db_manager.load_municipalities()
        municipalities_by_id = {m.id: m for m in municipalities_all if m.id is not None}

        # Год → { project_id → ... }
        years_map = defaultdict(lambda: defaultdict(lambda: defaultdict(lambda: defaultdict(list))))

        project_forms_map: Dict[int, list] = {}

        for project in projects:
            # Определяем год из year_id (новая архитектура) - используем предзагруженный словарь
            year = None
            if hasattr(project, 'year_id') and project.year_id:
                year_ref = years_by_id.get(project.year_id)
                if year_ref:
                    year = year_ref.year
            
            year_key = str(year) if year else "Без года"

            # Гарантируем наличие узла проекта в years_map (чтобы проект не пропадал при отсутствии ревизий)
            _ = years_map[year_key][project.id]

            # Загружаем project_forms и form_revisions для данного проекта
            project_forms = self.db_manager.load_project_forms(project.id)
            project_forms_map[project.id] = project_forms or []

            # Если у проекта нет ревизий (нет загруженных форм), 
            # добавляем проект в years_map с пустым forms_map
            # Это позволит показать проект в дереве без вложенности
            if not project_forms:
                # Инициализируем пустой forms_map для проекта
                if project.id not in years_map[year_key]:
                    years_map[year_key][project.id] = {}
                # Продолжаем - проект будет добавлен в итоговую структуру с пустым списком форм
                continue

            # Есть новые project_forms / form_revisions
            for pf in project_forms:
                ft_meta = form_types_meta.get(pf.form_type_id)
                # Код формы берём только из справочника типов форм
                form_code = ft_meta.code if ft_meta else "UNKNOWN"
                form_name = ft_meta.name if ft_meta else f"Форма {form_code}"

                # Код периода берём из ref_periods по period_id; если нет — используем 'Y' по умолчанию
                period_obj = periods_by_id.get(pf.period_id) if pf.period_id else None
                period_code = period_obj.code if period_obj else "Y"

                revisions = self.db_manager.load_form_revisions(pf.id)
                # Ревизии создаются только при загрузке формы, поэтому если их нет - пропускаем
                if not revisions:
                    continue
                    
                revisions_info = [{
                    "revision_id": r.id,
                    "revision": r.revision,
                    "status": r.status.value,
                    "project_id": project.id,
                    "file_path": r.file_path or "",
                } for r in revisions]

                years_map[year_key][project.id][form_code][period_code].extend(revisions_info)

        # Оптимизация: создаем словари для быстрого поиска (O(1) вместо O(n))
        projects_by_id = {p.id: p for p in projects}
        form_types_meta_by_code = {ft.code: ft for ft in form_types_meta.values()}
        periods_by_code = {p.code: p for p in periods_by_id.values() if p.code}
        
        # Преобразуем в удобную для UI структуру
        tree = []
        for year_key in sorted(years_map.keys(), reverse=True):
            year_entry = {"year": year_key, "projects": []}
            proj_map = years_map[year_key]
            for project_id, forms_map in proj_map.items():
                # Используем словарь для быстрого поиска проекта
                proj_obj = projects_by_id.get(project_id)
                if not proj_obj:
                    continue
                # Получаем название МО из справочника (новая архитектура) - используем предзагруженный словарь
                municipality_name = "-"  # Имя МО по умолчанию
                if hasattr(proj_obj, 'municipality_id') and proj_obj.municipality_id:
                    municip_ref = municipalities_by_id.get(proj_obj.municipality_id)
                    if municip_ref:
                        municipality_name = municip_ref.name
                
                proj_entry = {
                    "id": proj_obj.id,
                    "name": proj_obj.name,
                    "municipality": municipality_name,
                    "forms": [],
                }

                # Используем формы, связанные с проектом (даже если нет ревизий)
                project_forms = project_forms_map.get(project_id, [])
                if project_forms:
                    for pf in project_forms:
                        # Код/название формы
                        ft_meta = form_types_meta.get(pf.form_type_id)
                        form_code = ft_meta.code if ft_meta else "UNKNOWN"
                        form_name = ft_meta.name if ft_meta else f"Форма {form_code}"

                        form_entry = {
                            "form_code": form_code,
                            "form_name": form_name,
                            "periods": [],
                        }

                        # Код/название периода
                        period_obj = periods_by_id.get(pf.period_id) if pf.period_id else None
                        period_code = period_obj.code if period_obj else "Y"
                        period_name = period_obj.name if period_obj else period_code

                        # Ревизии берём из years_map (быстрее и учитывает только нужный проект/форму/период)
                        revisions = forms_map.get(form_code, {}).get(period_code, [])

                        period_entry = {
                            "period_code": period_code,
                            "period_name": period_name,
                            "revisions": sorted(
                                revisions,
                                key=lambda r: r["revision"],
                            ),
                        }

                        form_entry["periods"].append(period_entry)
                        proj_entry["forms"].append(form_entry)

                    # Сортируем формы/периоды
                    proj_entry["forms"].sort(key=lambda f: f["form_code"])
                    for f in proj_entry["forms"]:
                        f["periods"].sort(key=lambda p: p["period_code"])

                    # Сортируем формы по коду
                    proj_entry["forms"].sort(key=lambda f: f["form_code"])
                
                # Добавляем проект в дерево (даже если у него нет форм - это проект без загруженных ревизий)
                year_entry["projects"].append(proj_entry)

            # Сортируем проекты по имени
            year_entry["projects"].sort(key=lambda p: p["name"])
            tree.append(year_entry)

        return tree

# План рефакторинга main_controller и main_window

## Текущее состояние

- `main_controller.py`: 1306 строк
- `main_window.py`: 3728 строк

Оба файла содержат слишком много ответственностей и требуют разбиения на более мелкие, логически связанные модули.

---

## Предлагаемая структура

### 1. Рефакторинг MainController

#### 1.1. Создать `controllers/revision_controller.py`
**Ответственность:** Управление ревизиями форм

**Методы:**
- `delete_form_revision(revision_id)`
- `update_form_revision(revision_id, revision_data)`
- `load_revision(revision_id, project_id)`
- `_register_form_revision(...)`
- `set_current_form_params(...)`
- `set_form_params_from_revision(...)`
- `get_pending_form_params()`

#### 1.2. Создать `controllers/form_controller.py`
**Ответственность:** Работа с формами (инициализация, загрузка, парсинг)

**Методы:**
- `_initialize_form_for_project(form_meta)`
- `load_form_file(file_path)`
- `_copy_form_file_to_project(source_file_path, project_id)`

#### 1.3. Создать `controllers/reference_controller.py`
**Ответственность:** Управление справочниками

**Методы:**
- `_load_references()`
- `refresh_references()`
- `load_reference_file(file_path, ref_type, name)`

#### 1.4. Создать `controllers/calculation_controller.py`
**Ответственность:** Расчеты и экспорт

**Методы:**
- `calculate_sums()`
- `export_validation(output_path)`

#### 1.5. Создать `controllers/tree_controller.py`
**Ответственность:** Построение дерева проектов

**Методы:**
- `build_project_tree()`

#### 1.6. Обновить `controllers/main_controller.py`
**Остается как координатор:**
- Инициализация всех подконтроллеров
- Делегирование вызовов подконтроллерам
- Управление текущим состоянием (current_project, current_form, current_revision_id)
- Сигналы и их проксирование

---

### 2. Рефакторинг MainWindow

#### 2.1. Создать `views/widgets/` папку с кастомными виджетами

**`views/widgets/custom_headers.py`**
- `WrapHeaderView` - кастомный заголовок с переносом текста

**`views/widgets/custom_delegates.py`**
- `WordWrapItemDelegate` - делегат для переноса текста в ячейках

**`views/widgets/detached_tab_window.py`**
- `DetachedTabWindow` - окно для открепленных вкладок

#### 2.2. Создать `views/panels/` папку для панелей

**`views/panels/projects_panel.py`**
- `ProjectsPanel` - левая панель со списком проектов
- Методы: `create_projects_panel()`, `update_projects_list()`, `on_project_tree_double_clicked()`, `show_project_context_menu()`

**`views/panels/tabs_panel.py`**
- `TabsPanel` - центральная панель с вкладками
- Методы: `create_tabs_panel()`, работа с вкладками, контекстное меню вкладок

#### 2.3. Создать `views/tree/` папку для работы с деревьями данных

**`views/tree/tree_builder.py`**
- `TreeBuilder` - построение деревьев из данных
- Методы: `build_tree_from_data()`, `create_tree_item()`, `load_project_data_to_tree()`

**`views/tree/tree_config.py`**
- `TreeConfig` - конфигурация деревьев (заголовки, маппинг колонок)
- Методы: `configure_tree_headers()`, `_configure_tree_headers_for_widget()`, `hide_zero_columns_in_tree()`

**`views/tree/tree_handlers.py`**
- `TreeHandlers` - обработчики событий деревьев
- Методы: `on_tree_item_clicked()`, `on_tree_selection_changed()`, `show_tree_context_menu()`, `show_tree_header_context_menu()`

#### 2.4. Создать `views/errors/` папку для работы с ошибками

**`views/errors/errors_manager.py`**
- `ErrorsManager` - управление отображением ошибок
- Методы: `load_errors_to_tab()`, `_check_budget_errors()`, `_check_consolidated_errors()`, `_update_errors_table()`, `_export_errors()`

#### 2.5. Создать `views/metadata/` папку для метаданных

**`views/metadata/metadata_panel.py`**
- `MetadataPanel` - панель метаданных
- Методы: `load_metadata()`, `_get_metadata_widgets()`

#### 2.6. Создать `views/menu/` папку для меню

**`views/menu/menu_bar.py`**
- `MenuBar` - создание меню-бара
- Методы: `create_menu_bar()`, все действия меню

**`views/menu/toolbar.py`**
- `ToolBar` - создание тулбара
- Методы: `create_toolbar()`

#### 2.7. Обновить `views/main_window.py`
**Остается как главное окно:**
- Инициализация всех панелей и компонентов
- Координация между компонентами
- Основные обработчики событий верхнего уровня
- Управление состоянием окна

---

## Детальная структура файлов

```
controllers/
├── __init__.py
├── main_controller.py          # Координатор (уменьшен до ~300-400 строк)
├── project_controller.py        # (существующий)
├── document_controller.py      # (существующий)
├── solution_controller.py      # (существующий)
├── revision_controller.py      # НОВЫЙ: управление ревизиями
├── form_controller.py           # НОВЫЙ: работа с формами
├── reference_controller.py     # НОВЫЙ: справочники
├── calculation_controller.py  # НОВЫЙ: расчеты и экспорт
└── tree_controller.py          # НОВЫЙ: построение дерева проектов

views/
├── __init__.py
├── main_window.py              # Главное окно (уменьшено до ~500-700 строк)
├── widgets/                    # НОВАЯ ПАПКА
│   ├── __init__.py
│   ├── custom_headers.py       # WrapHeaderView
│   ├── custom_delegates.py     # WordWrapItemDelegate
│   └── detached_tab_window.py  # DetachedTabWindow
├── panels/                     # НОВАЯ ПАПКА
│   ├── __init__.py
│   ├── projects_panel.py      # ProjectsPanel
│   └── tabs_panel.py           # TabsPanel
├── tree/                       # НОВАЯ ПАПКА
│   ├── __init__.py
│   ├── tree_builder.py         # TreeBuilder
│   ├── tree_config.py          # TreeConfig
│   └── tree_handlers.py        # TreeHandlers
├── errors/                     # НОВАЯ ПАПКА
│   ├── __init__.py
│   └── errors_manager.py       # ErrorsManager
├── metadata/                   # НОВАЯ ПАПКА
│   ├── __init__.py
│   └── metadata_panel.py       # MetadataPanel
└── menu/                       # НОВАЯ ПАПКА
    ├── __init__.py
    ├── menu_bar.py             # MenuBar
    └── toolbar.py               # ToolBar
```

---

## Преимущества такой структуры

1. **Разделение ответственности (SRP)**: Каждый модуль отвечает за одну область функциональности
2. **Упрощение тестирования**: Легче писать unit-тесты для отдельных компонентов
3. **Улучшение читаемости**: Меньшие файлы проще понимать и поддерживать
4. **Переиспользование**: Компоненты можно использовать в других частях приложения
5. **Параллельная разработка**: Разные разработчики могут работать над разными модулями
6. **Легче находить код**: Логика находится в предсказуемых местах

---

## План миграции

### Этап 1: Подготовка
1. Создать новую структуру папок
2. Создать `__init__.py` файлы

### Этап 2: Рефакторинг контроллеров
1. Вынести `revision_controller.py`
2. Вынести `form_controller.py`
3. Вынести `reference_controller.py`
4. Вынести `calculation_controller.py`
5. Вынести `tree_controller.py`
6. Обновить `main_controller.py` для использования новых контроллеров

### Этап 3: Рефакторинг представлений
1. Вынести кастомные виджеты в `widgets/`
2. Вынести панели в `panels/`
3. Вынести работу с деревьями в `tree/`
4. Вынести работу с ошибками в `errors/`
5. Вынести метаданные в `metadata/`
6. Вынести меню в `menu/`
7. Обновить `main_window.py` для использования новых компонентов

### Этап 4: Тестирование и очистка
1. Протестировать все функции
2. Удалить неиспользуемый код
3. Обновить импорты во всех файлах

---

## Альтернативный подход (более агрессивный)

Если нужна еще более модульная структура, можно использовать паттерн MVC/MVP более строго:

```
controllers/
├── base_controller.py          # Базовый класс контроллера
├── main_controller.py          # Главный координатор
└── [специализированные контроллеры]

views/
├── base_view.py                # Базовый класс представления
├── main_window.py             # Главное окно
└── [специализированные представления]

presenters/                     # НОВАЯ ПАПКА (опционально)
└── [презентеры для связи view-controller]
```

Но для текущего проекта первый вариант (умеренный рефакторинг) будет более безопасным и достаточным.

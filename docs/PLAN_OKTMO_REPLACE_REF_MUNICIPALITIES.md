# План: замена ref_municipalities на ОКТМО в диалоге проекта

**Выбранные варианты:** 2.B (дата = дата создания проекта), 3.a (в проекте хранится `oktmo_code`). **Учёт:** из проекта убираются `municipality_id` и связь с `ref_municipalities`.

---

## Цель
В диалоговом окне создания/редактирования проекта использовать онлайн-справочник **oktmo** вместо **ref_municipalities**. В комбобоксе МО — только записи ОКТМО с **8 разрядами в коде**. Список ОКТМО фильтровать по **дате создания проекта**; при изменении даты создания — перезагружать справочник. В проекте хранить **только** `oktmo_code`; поле `municipality_id` и привязка к `ref_municipalities` из проекта **удаляются**.

---

## 1. Требования к данным

| Что | Детали |
|-----|--------|
| Источник в диалоге | Таблица `oktmo`. |
| Фильтр по разрядам | Только коды с 8 разрядами: `LENGTH(REPLACE(TRIM(code), ' ', '')) = 8` (или аналог по формату в БД). |
| Фильтр по дате | По `startdate`/`enddate` в oktmo. Дата = дата создания проекта (`created_at`); при создании проекта — текущая дата. При изменении даты создания — перезагрузка списка ОКТМО. |
| Хранение в проекте | Только `oktmo_code` (TEXT, 8 символов). `municipality_id` из проекта и таблицы `projects` убирается. |

---

## 2. План работ по шагам

### 2.1. БД: миграция проектов

- [ ] **Добавить колонку `oktmo_code`** в таблицу `projects` (TEXT, nullable).
- [ ] **Миграция данных (опционально):** для существующих проектов по `municipality_id` взять код/имя из `ref_municipalities`, найти соответствующий код в `oktmo` (8 разрядов, по дате проекта) и записать в `oktmo_code`. Если сопоставление невозможно — оставить NULL.
- [ ] **Удалить колонку `municipality_id`** из таблицы `projects` (ALTER или пересоздание таблицы с новой схемой).

### 2.2. Модель Project (base_models.py)

- [ ] **Добавить** атрибут `oktmo_code: Optional[str]` в класс `Project`.
- [ ] **Удалить** атрибут `municipality_id` и все обращения к нему в `Project` (from_dict, to_dict, сериализация).
- [ ] В загрузке/сохранении проекта в БД читать и писать только `oktmo_code`.

### 2.3. DatabaseManager (database.py)

- [ ] **Добавить** метод `load_oktmo_for_municipality(filter_date: str, code_length: int = 8)` — загрузка из `oktmo` с фильтром по дате (логика как в `get_filtered_view`) и по длине кода 8 разрядов. Возвращать список пар (code, name) или список объектов с полями code, name.
- [ ] **Добавить** метод `get_oktmo_name_by_code(code: str, filter_date: Optional[str] = None) -> Optional[str]` — по коду и дате (startdate/enddate) вернуть `name` из `oktmo`.
- [ ] **В загрузке/сохранении проектов:** убрать чтение/запись `municipality_id`; добавить чтение/запись `oktmo_code` (в `load_projects`, создание/обновление проекта, `_load_project_data` при необходимости).
- [ ] **Удалить** методы `load_municipalities`, `get_or_create_municipality`, `get_municipality_by_id`, `save_municipalities_bulk`, все обращения к таблице `ref_municipalities` (инициализация, подсчёт, загрузка/сохранение). Таблица и справочник «Муниципальные образования» из приложения убираются полностью; в этом документе остаётся только описание полей старой таблицы (раздел 5).

### 2.4. Диалог проекта (ProjectDialog)

- [ ] **Дата создания проекта:** добавить в форму поле «Дата создания» (QDateEdit), привязанное к `project.created_at`. При создании проекта — значение по умолчанию текущая дата; при редактировании — подставлять `project.created_at`.
- [ ] **Загрузка списка МО:** заменить `_load_municipalities()` на `_load_oktmo_for_dialog(filter_date: str)`: вызывать `db_manager.load_oktmo_for_municipality(filter_date, code_length=8)`, заполнять комбобокс форматом «код — наименование», в `currentData()` хранить `code`.
- [ ] **При открытии диалога:** определить дату для фильтра: при редактировании — `project.created_at` (в формате YYYY-MM-DD), при создании — текущая дата. Вызвать `_load_oktmo_for_dialog(дата)`.
- [ ] **Перезагрузка при смене даты:** подписаться на `dateChanged` поля «Дата создания»; в обработчике пересчитать дату и вызвать `_load_oktmo_for_dialog(новая_дата)`, обновить комбо, при возможности сохранить выбранный код (если он есть в новом списке).
- [ ] **set_project:** выбирать в комбо элемент по `project.oktmo_code` (findData(oktmo_code)); убрать логику по `municipality_id`.
- [ ] **get_project_data:** возвращать `oktmo_code` из текущего выбора комбо (currentData()); убрать `municipality_id` и обращение к `_municip_cache`/ref_municipalities. При необходимости возвращать также обновлённую дату создания для сохранения в проект.
- [ ] Удалить `_municip_cache` и вызовы `load_municipalities`, `get_or_create_municipality` из диалога проекта.

### 2.5. Отображение МО (главное окно, дерево)

- [ ] **main_controller (project_info):** убрать использование `municipality_id` и `load_municipalities()`. Для подписи МО вызывать `get_oktmo_name_by_code(project.oktmo_code, project.created_at)`; если `oktmo_code` пустой — показывать «—».
- [ ] **tree_controller:** убрать загрузку `load_municipalities()` и словарь по `municipality_id`. Для узла проекта брать название МО через `get_oktmo_name_by_code(project.oktmo_code, project.created_at)`; при отсутствии кода — «—».

### 2.6. project_controller и прочее

- [ ] **project_controller:** при загрузке/обновлении проекта не читать и не записывать `municipality_id`; использовать только `oktmo_code` (и дату создания при необходимости).

### 2.7. Управление справочниками

- [ ] **Убрать** из диалога «Управление справочниками» справочник «Муниципальные образования» (ref_municipalities): удалить запись из списка справочников, все обработчики и загрузку данных по этой таблице. В приложении ref_municipalities больше не используется; для справки остаётся только описание полей старой таблицы (раздел 5).

### 2.8. Тесты и проверки

- [ ] Создание нового проекта: дата создания = текущая, выбор МО из ОКТМО (8 разрядов), сохранение — в БД только `oktmo_code`.
- [ ] Редактирование: подставляются дата создания и выбранный ОКТМО; при смене даты создания список ОКТМО перезагружается.
- [ ] Дерево и главное окно: название МО выводится по `oktmo_code` из oktmo.
- [ ] В проекте нигде не остаётся использования `municipality_id` и ref_municipalities для привязки проекта к МО.

---

## 3. Зависимости

- Таблица `oktmo` и логика фильтрации по дате (`get_filtered_view`, SORT_RULES для oktmo) уже есть.
- Условие «8 разрядов» для кода задать один раз (формат кода в `oktmo`: с/без пробелов).

---

## 4. Краткое резюме

| Компонент | Действие |
|-----------|----------|
| Проект (модель/БД) | Добавить `oktmo_code`. Удалить `municipality_id` из проекта и из таблицы `projects`. |
| Диалог проекта | Список МО из oktmo (8 разрядов, дата = дата создания). При смене даты создания — перезагрузка списка. Сохранение только `oktmo_code`. |
| Дерево / главное окно | Название МО только из oktmo по `oktmo_code` и дате создания. |
| ref_municipalities | Убирается из приложения полностью (БД, DatabaseManager, диалог справочников). В документе остаётся только описание полей старой таблицы (раздел 5). |

---

## 5. Описание старой таблицы ref_municipalities (по полям)

Для справки — структура таблицы `ref_municipalities`, которая удаляется из приложения.

| Поле | Тип | Описание |
|------|-----|----------|
| id | INTEGER PRIMARY KEY AUTOINCREMENT | Идентификатор записи. |
| code | VARCHAR(3) UNIQUE | Код МО (3 символа). |
| name | TEXT NOT NULL | Наименование. |
| municipality_type_code | VARCHAR(1) | Ссылка на ref_municipality_types. |
| municipality_code | VARCHAR(3) | Код муниципалитета. |
| genitive_case | TEXT | Наименование в родительном падеже. |
| council_address | TEXT | Адрес совета. |
| administration_address | TEXT | Адрес администрации. |
| council_email | VARCHAR(50) | Email совета. |
| administration_email | VARCHAR(50) | Email администрации. |
| council_position | VARCHAR(30) | Должность председателя совета. |
| council_surname | VARCHAR(30) | Фамилия председателя совета. |
| council_first_name | VARCHAR(30) | Имя председателя совета. |
| council_patronymic | VARCHAR(30) | Отчество председателя совета. |
| administration_position | VARCHAR(30) | Должность главы администрации. |
| administration_surname | VARCHAR(30) | Фамилия главы администрации. |
| administration_first_name | VARCHAR(30) | Имя главы администрации. |
| administration_patronymic | VARCHAR(30) | Отчество главы администрации. |
| agreement_date | DATE | Дата соглашения. |
| decision_date | DATE | Дата решения. |
| decision_number | VARCHAR(50) | Номер решения. |
| initial_income | REAL | Начальные доходы. |
| initial_expense | REAL | Начальные расходы. |
| initial_deficit | REAL | Начальный дефицит. |
| is_active | INTEGER NOT NULL DEFAULT 1 | Признак активности. |

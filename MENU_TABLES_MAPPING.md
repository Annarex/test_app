# Соответствие пунктов меню и таблиц базы данных

## Меню "Справочники"

### 1. "Справочники..." (Ctrl+R)
**Диалог:** `ReferencesManagementDialog` (`views/references_management_dialog.py`)  
**Окно:** "Справочники" (с кнопкой максимизации)

#### Обычные справочники (с загрузкой из Excel):
- **Коды доходов** → `income_reference_records`
  - Загрузка: через `ReferenceController.load_reference_file()` (тип: 'доходы')
  - Использование: КРИТИЧЕСКИ ВАЖЕН для расчетов доходов
  
- **Коды источников** → `source_reference_records`
  - Загрузка: через `ReferenceController.load_reference_file()` (тип: 'источники')
  - Использование: КРИТИЧЕСКИ ВАЖЕН для расчетов источников

#### Справочники конфигурации (редактируемые):
- **Годы** → `ref_years`
- **Муниципальные образования** → `ref_municipalities`
- **Типы форм** → `ref_form_types`
- **Периоды** → `ref_periods`

#### Онлайн справочники (только просмотр):
- **ОКТМО** → `oktmo`
- **Классификаторы доходов бюджета ФУ** → `budgetclastypeinc` (+ `npa`)
- **Классификаторы доходов бюджета МО** → `budgetclassubtypincmo` (+ `npa`)
- **Администраторы бюджета ФУ** → `budgetclasgabs` (+ `npa`)
- **Администраторы бюджета МО** → `budgetclasgabsmo` (+ `npa`)
- **Распорядители бюджета ФУ** → `budgetclasgrbs` (+ `npa`)
- **Распорядители бюджета МО** → `budgetclasgrbsmo` (+ `npa`)
- **Классификаторы расходов бюджета ФУ** → `budgetclascosts` (+ `npa`)
- **Классификаторы расходов бюджета МО** → `budgetclascostsmo` (+ `npa`)
- **Источники финансирования дефицита ФУ** → `budgetclasgaiffb` (+ `npa`)
- **Источники финансирования дефицита МО** → `budgetclasgaifmo` (+ `npa`)
- **Классификаторы источников финансирования ФУ** → `budgetclassources` (+ `npa`)
- **Классификаторы источников финансирования МО** → `budgetclassourcesmo` (+ `npa`)

#### Объединенные представления (только просмотр):
- **Классификаторы доходов (объединенные)** → `v_budgetclastypeinc_merged` (+ `npa`)
- **Классификаторы расходов (объединенные)** → `v_budgetclascosts_merged` (+ `npa`)
- **Распорядители бюджета (объединенные)** → `v_budgetclasgrbs_merged` (+ `npa`)
- **Администраторы бюджета (объединенные)** → `v_budgetclasgabs_merged` (+ `npa`)
- **Источники финансирования дефицита (объединенные)** → `v_budgetclasgaiffb_merged` (+ `npa`)
- **Классификаторы источников финансирования (объединенные)** → `v_budgetclassources_merged` (+ `npa`)

### 2. "Обновить онлайн справочники..." (Ctrl+Shift+R)
**Диалог:** `BudgetReferencesUpdateDialog` (`views/budget_references_update_dialog.py`)  
**Сервис:** `services.budget_references_service.BudgetReferencesService`

**Таблицы (обновление через API):**
- `oktmo`
- `budgetclastypeinc`
- `budgetclassubtypincmo`
- `budgetclasgabs`
- `budgetclasgabsmo`
- `budgetclasgrbs`
- `budgetclasgrbsmo`
- `budgetclascosts`
- `budgetclascostsmo`
- `budgetclasgaiffb`
- `budgetclasgaifmo`
- `budgetclassources`
- `budgetclassourcesmo`

## Примечания

### Загрузка справочников:
- **Коды доходов** и **Коды источников** загружаются через `ReferenceController.load_reference_file()` (специальная обработка из Excel)
- **Онлайн справочники** обновляются через API бюджетной системы (`services.budget_references_service`)
- **Справочники конфигурации** можно редактировать прямо в таблице диалога

### Структура проекта:
- **Сервисы:** `services/budget_references_service.py`, `services/budget_level_processor.py`
- **Утилиты:** `utils/db_utils.py`, `utils/level_utils.py`
- **Модели:** `models/database.py` (определение схемы БД и VIEW)

### Особенности:
- Онлайн справочники с `has_npa: True` показывают данные из связанной таблицы `npa` (нормативно-правовые акты)
- Объединенные представления (`v_*_merged`) автоматически создаются при инициализации БД и объединяют данные ФУ и МО
- Все онлайн справочники можно обновить через диалог "Обновить онлайн справочники..."

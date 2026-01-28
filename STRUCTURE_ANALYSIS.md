# Анализ структуры проекта

## Выявленные проблемы

### 1. Сервисы находятся в `models/` вместо `services/`

**Проблема:**
- `models/references/budget_references_service.py` - сервис для работы с API
- `models/utils/budget_level_processor.py` - процессор обработки данных

**Должно быть:**
- `services/budget_references_service.py`
- `services/budget_level_processor.py`

**Причина:** Сервисы содержат бизнес-логику и состояние, не являются моделями данных.

---

### 2. Общие утилиты находятся в `models/utils/` вместо корневого `utils/`

**Проблема:**
- `models/utils/db_utils.py` - общие функции для работы с БД
- `models/utils/level_utils.py` - общие функции для уровней

**Должно быть:**
- `utils/db_utils.py`
- `utils/level_utils.py`

**Причина:** Это чистые функции-утилиты, не привязанные к конкретным моделям.

**Оставить в `models/utils/`:**
- `code_utils.py` - специфичные для моделей (коды форм)
- `form_utils.py` - специфичные для моделей (работа с формами)

---

### 3. Пустая структура

**Проблема:**
- `views/controllers/` - только `__init__.py`, не используется

**Решение:** Удалить папку или использовать по назначению.

---

### 4. Несоответствие назначения папки `models/references/`

**Проблема:**
- `models/references/` содержит только сервис, нет моделей данных

**Решение:** После перемещения сервиса в `services/`, папку можно удалить или использовать для моделей данных справочников.

---

## Предлагаемая структура

```
services/                          # Сервисы с бизнес-логикой
  ├── error_checker_service.py     # уже есть
  ├── budget_references_service.py # ← переместить из models/references/
  └── budget_level_processor.py   # ← переместить из models/utils/

utils/                             # Общие утилиты
  ├── numeric_utils.py            # уже есть
  ├── db_utils.py                  # ← переместить из models/utils/
  └── level_utils.py              # ← переместить из models/utils/

models/utils/                      # Утилиты специфичные для моделей
  ├── code_utils.py               # работа с кодами форм
  └── form_utils.py               # работа с формами

models/references/                 # Удалить или использовать для моделей данных
```

---

## План реорганизации

1. Переместить `budget_references_service.py` → `services/`
2. Переместить `budget_level_processor.py` → `services/`
3. Переместить `db_utils.py` → `utils/`
4. Переместить `level_utils.py` → `utils/`
5. Обновить все импорты
6. Удалить пустую папку `views/controllers/` (если не используется)
7. Удалить `models/references/` (если не будет моделей данных)

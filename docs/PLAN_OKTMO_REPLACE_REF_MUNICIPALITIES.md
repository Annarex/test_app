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

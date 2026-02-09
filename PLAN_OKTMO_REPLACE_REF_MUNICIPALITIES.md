## 5. Описание старой таблицы ref_municipalities (по полям)

Для справки — структура таблицы `ref_municipalities`, которая удаляется из приложения.

| Поле | Тип | Описание |
|------|-----|----------|
| id | INTEGER PRIMARY KEY AUTOINCREMENT | Идентификатор записи. |
| code_oktmo | TEXT | Код муниципального органа. |
| date_actuality | TEXT | Дата актуальности. |
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
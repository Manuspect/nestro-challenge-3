# Robot for nestro-challenge, 3 track

## "Приложение 1.xlsx" и "Приложение 2.xlsx"

Возможности изменений:

Строки: содержание, кол-во, порядок - да 

Столбцы: названия, порядок - нет

[Постановка задачи для дэшбордов и требования к автоматизации в приложениях](https://autumn-athlete-fea.notion.site/b0f7889d5f774bfea0f683e0ef71b654?pvs=4)

## Запуск робота и автоматизации

app_1_2_preprocessor собирает данные из браузера, приложения 2 и обновляет приложение 1

```console
rcc task run --robot ./robot.yaml --task app_1_2_preprocessor
```

app_1_postprocessor собирает данные из приложения 1 и заполняет файл `./output/shared/workitems.json`, отвечающий за отображение графиков с помощью flutter

```console
rcc task run --robot ./robot.yaml --task app_1_postprocessor
```
## Задачи решаемые с помощью `app_1_2_preprocessor``
- курс usd monthly_exchange_rate, quarterly_exchange_rate, [источник](https://www.cbr.ru/currency_base/dynamics/?UniDbQuery.Posted=True&UniDbQuery.so=1&UniDbQuery.mode=1&UniDbQuery.date_req1=%7Bstart_date%7D&UniDbQuery.date_req2=%7Bend_date%7D&UniDbQuery.VAL_NM_RQ=R01235)
- Анализ_БК+ББ, Y, курс на дату реализации, источник - 'monthly_exchange_rate'
- Анализ_БК+ББ, Z, курс на дату поступления ден средств (погашения ДЗ), источник - 'monthly_exchange_rate'
- Company ABC_факт_НДПИ (Platts), C7:N7, собирает ставку вывозной таможенной пошлины  на нефть сырую, [источник](https://www.economy.gov.ru/material/directions/vneshneekonomicheskaya_deyatelnost/tamozhenno_tarifnoe_regulirovanie/o_vyvoznyh_tamozhennyh_poshlinah_na_neft_i_otdelnye_kategorii_tovarov_vyrabotannyh_iz_nefti_na_period_s_1_po_28_fevralya_2022_goda.html)
- Анализ_БК+ББ, C, котировки Spread (Корректировка), $/барр., [источник](https://ru.investing.com/commodities/brent-wti-crude-spread-futures-historical-data)
- Анализ_БК+ББ, B, Цена Brent, $/барр. [источник](https://ru.investing.com/commodities/brent-oil-historical-data)
- Компания 1_факт_НДПИ (Platts), C14:N14;C171:N171, Цена Юралс, $/bbl [источник](https://economy.gov.ru/material/departments/d12/konyunktura_mirovyh_tovarnyh_rynkov/o_sredney_cene_na_neft_sorta_yurals_za_yanvar_2022_goda_.html)
- Анализ_БК+ББ, AB, Цена Brent, $/барр. источник - 'Приложение 2.xlsx'

## User flow:
1) Запустить приложение с дэшбордами
2) Указать путь к таблице
3) Выбрать - автоматически обновлять таблицу или нет?
Да: запускаем скрипт предобработки, ждем окончания и запускаем скрипт постобработки 
Нет: запускаем скрипт постобработки
4) Подождать пока постобработка с помощью робота закончится
5) Дэшборды отображены

После первого запуска, информация о наличии данных и пути к таблице сохраняется и при повторном запуске приложения робот заново не запускается, а сразу показываются дэшборды

## Архитектура

![Alt text](image.png)


## Запуск дэшбордов

Подробности для запуска дэшбордов в [этом репозитории](https://github.com/alexeynau/flutter-dashboard)

EXCEL_FILE_PATH содержит путь к  файлу Excel.
```
import os

excel_file_path = os.environ.get("EXCEL_FILE_PATH")
# Теперь excel_file_path содержит путь к файлу "Приложение 1.xlsx"

```
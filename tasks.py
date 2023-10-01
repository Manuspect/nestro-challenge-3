from bs4 import BeautifulSoup
from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import Files as Excel
import openpyxl
from robocorp.tasks import task
from RPA.Desktop.OperatingSystem import OperatingSystem
from collections import defaultdict
from selenium.webdriver.common.keys import Keys
from openpyxl import Workbook
import pandas as pd
import numpy as np
import logging
import os
import json
from dotenv import load_dotenv

load_dotenv()

try:
    os.remove('parse_log.log')
except OSError:
    pass

logging.basicConfig(filename='parse_log.log', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s',
                    encoding='utf-8')

output_directory = os.environ.get("ROBOT_ARTIFACTS")
shared_directory = os.path.join(output_directory, "shared")
app1_file_path = os.environ.get("APP1_FILE_PATH")
app2_file_path = os.environ.get("APP2_FILE_PATH")
print("APP1_FILE_PATH", app1_file_path, "APP2_FILE_PATH", app2_file_path)


@task
def web_preprocessor():
    from first_table import pars_usd, customs_duties_parser, parcing_first_table, parsing_brent_cost, urals_parser, second_table
    print("Start parser")
    logging.info("Start parser")
    # pars_usd.usd_kurs()
    # customs_duties_parser.start(app1_file_path)
    # parcing_first_table.start(app1_file_path)
    # parsing_brent_cost.start(app1_file_path)

    urals_parser.start(app1_file_path)

    second_table.to_application_1(app1_file_path, app2_file_path)


@task
def after_update_postprocessor():
    # Загрузите данные из monthly_exchange_rate.xlsx
    monthly_data = pd.read_excel('monthly_exchange_rate.xlsx').T

    months = {'январь': 1, 'февраль': 2, 'март': 3, 'апрель': 4, 'май': 5,
              'июнь': 6, 'июль': 7, 'август': 8, 'сентябрь': 9, 'октябрь': 10,
              'ноябрь': 11, 'декабрь': 12}

    logging.info(monthly_data)

    # Загрузите данные из Приложение_1.xlsx
    app1_data = pd.read_excel(app1_file_path, header=[0, 1])
    excel = Excel()
    excel.open_workbook(app1_file_path)
    excel.set_active_worksheet('Анализ_БК+ББ')
    counter = 4
    shipment_date = excel.get_cell_value(counter, 'L')
    while excel.get_cell_value(counter, 'L') is not None:
        shipment_date = excel.get_cell_value(counter, 'L')
        reciving_date = excel.get_cell_value(counter, 'O')
        excel.set_cell_value(counter, "Y", monthly_data[months[shipment_date.lower(
        ).replace(" ", "")]].loc['Средний курс'])
        excel.set_cell_value(counter, "Z", monthly_data[months[reciving_date.lower(
        ).replace(" ", "")]].loc['Средний курс'])
        counter += 1
    excel.save_workbook(app1_file_path)

    # Загрузите данные из цены_нефти.xlsx
    oil_price_data = pd.read_excel('цены_нефти.xlsx')

    # TODO: read price_data from 'Компания 1_факт_НДПИ (Platts)'
    # oil_price_data = pd.read_excel(
    #     app1_file_path, sheet_name='Компания 1_факт_НДПИ (Platts)')
    # print(oil_price_data.iloc[10:14])

    # Загрузите данные из Приложение_1.xlsx
    app1_data = pd.read_excel(app1_file_path)

    # Задайте столбцы, с которыми нужно сопоставить данные
    columns_to_match = ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль", "Август", "Сентябрь", "Октябрь",
                        "Ноябрь", "Декабрь"]

    # Обновите данные в Приложение_1.xlsx
    for column in columns_to_match:
        app1_data[column] = oil_price_data[column]

    # Сохраните обновленные данные в Приложение_1.xlsx
    app1_data.to_excel(app1_file_path, index=False)

    # Загрузите данные из quarterly_exchange_rate.xlsx
    quarterly_data = pd.read_excel('quarterly_exchange_rate.xlsx')

    # Загрузите данные из Приложение_1.xlsx
    app1_data = pd.read_excel(app1_file_path)

    # Задайте столбцы кварталов, с которыми нужно сопоставить данные
    quarters_to_match = ["1 кв", "2 кв", "3 кв", "4 кв"]

    # Обновите данные в Приложение_1.xlsx
    for quarter in quarters_to_match:
        app1_data[quarter] = quarterly_data["Средний курс"]

    # Сохраните обновленные данные в Приложение_1.xlsx
    app1_data.to_excel(app1_file_path, index=False)


def add_to_json(df, title, plot_type, chats_data):
    data_name = [dict_["name"] for dict_ in chats_data['data']]
    data = []
    for column in list(df.columns):
        if column in data_name:
            continue
        value = {
            "name": column,
            "series": list(df[column])
        }
        data.append(value)
    plot_name = {
        "plot_name": title,
        "x": df.columns[0],
        "y": [column for column in list(df.columns[1:])]
    }
    chats_data["data"] = chats_data["data"] + data
    chats_data["charts"][plot_type].append(plot_name)
    return chats_data


def is_numeric(val):
    try:
        float(val)
        return True
    except ValueError:
        return False


@task
def get_data_from_excel():
    chats_data = {"data": [], "charts": defaultdict(list)}
    exchange_rate_df = pd.read_excel('monthly_exchange_rate.xlsx')
    exchange_rate_df['Месяц'] = exchange_rate_df['Месяц'].replace({1: 'Январь',
                                                                   2: 'Февраль',
                                                                   3: 'Март',
                                                                   4: 'Апрель',
                                                                   5: 'Май',
                                                                   6: 'Июнь',
                                                                   7: 'Июль',
                                                                   8: 'Август',
                                                                   9: 'Сентябрь',
                                                                   10: 'Октябрь',
                                                                   11: 'Ноябрь',
                                                                   12: 'Декабрь'})
    exchange_rate_df['Средний курс'] = exchange_rate_df['Средний курс'].apply(
        lambda x: round(x, 2))
    chats_data = add_to_json(
        exchange_rate_df, "Курс доллара по месяцам", "plots", chats_data=chats_data)

    excel = Excel()
    excel.open_workbook(app1_file_path, data_only=True)
    excel.set_active_worksheet('Анализ_БК+ББ')

    freight_usd = []
    freight_rub = []
    date = []

    rows = excel.read_worksheet_as_table()
    freight_column_usd = rows.get_column('AD')
    freight_column_rub = rows.get_column('AV')
    date_column = rows.get_column('M')

    for item_usd, item_rub, item_date in zip(list(freight_column_usd.values())[3:],
                                             list(freight_column_rub.values())[3:], list(date_column.values())[3:]):
        if (item_usd is None) or (item_rub is None) or (item_date is None):
            break
        freight_usd.append(float(item_usd))
        freight_rub.append(float(item_rub))
        date.append(item_date)
    freight_df = pd.DataFrame({'Дата': date,
                               'Freight_usd': freight_usd,
                               'Freight_rub': freight_rub})
    freight_df['Дата'] = pd.to_datetime(
        freight_df['Дата']).dt.strftime('%Y-%m-%d')
    freight_df[['Freight_usd', 'Freight_rub']] = freight_df[[
        'Freight_usd', 'Freight_rub']].map(lambda x: x if is_numeric(x) else 0)
    freight_df = freight_df.groupby('Дата', as_index=False)[
        ['Freight_usd', 'Freight_rub']].mean()
    freight_df[['Freight_usd', 'Freight_rub']] = freight_df[[
        'Freight_usd', 'Freight_rub']].apply(lambda x: round(x, 2))
    chats_data = add_to_json(
        freight_df, "Курс Freight по дате отгрузки", "plots", chats_data=chats_data)

    # Spread
    spread_usd_col = rows.get_column('AC')
    spread_rub_col = rows.get_column('AL')

    spread_usd = []
    spread_rub = []
    date = []
    for item_usd, item_rub, item_date in zip(list(spread_usd_col.values())[3:], list(spread_rub_col.values())[3:],
                                             list(date_column.values())[3:]):
        if (item_usd is None) or (item_rub is None) or (item_date is None):
            break
        spread_usd.append(item_usd)
        spread_rub.append(item_rub)
        date.append(item_date)

    spread_df = pd.DataFrame({'Дата spread': date,
                              'Spread $/барр.': spread_usd,
                              'Spread руб./т.': spread_rub})
    spread_df['Дата spread'] = pd.to_datetime(
        spread_df['Дата spread']).dt.strftime('%Y-%m-%d')
    spread_df[['Spread $/барр.', 'Spread руб./т.']] = spread_df[['Spread $/барр.',
                                                                 'Spread руб./т.']].map(
        lambda x: x if is_numeric(x) else 0)
    spread_df = spread_df.groupby('Дата spread', as_index=False)[
        ['Spread $/барр.', 'Spread руб./т.']].mean()
    spread_df[['Spread $/барр.', 'Spread руб./т.']
              ] = spread_df[['Spread $/барр.', 'Spread руб./т.']].apply(lambda x: round(x, 2))
    chats_data = add_to_json(
        spread_df, "Spread на дату отгрузки", "plots", chats_data=chats_data)

    # Brent
    brent_usd_col = rows.get_column('AB')
    brent_rub_col = rows.get_column('AK')
    brent_usd = []
    brent_rub = []
    date = []

    for item_usd, item_rub, item_date in zip(list(brent_usd_col.values())[3:], list(brent_rub_col.values())[3:],
                                             list(date_column.values())[3:]):
        if (item_usd is None) or (item_rub is None) or (item_date is None):
            break
        brent_usd.append(item_usd)
        brent_rub.append(item_rub)
        date.append(item_date)

    brent_df = pd.DataFrame({'Дата brent': date,
                             'Brent $/барр.': brent_usd,
                             'Brent руб./т.': brent_rub})
    brent_df['Дата brent'] = pd.to_datetime(
        brent_df['Дата brent']).dt.strftime('%Y-%m-%d')
    brent_df[['Brent $/барр.', 'Brent руб./т.']] = brent_df[['Brent $/барр.',
                                                             'Brent руб./т.']].map(lambda x: x if is_numeric(x) else 0)
    brent_df = brent_df.groupby('Дата brent', as_index=False)[
        ['Brent $/барр.', 'Brent руб./т.']].mean()
    brent_df[['Brent $/барр.', 'Brent руб./т.']] = brent_df[['Brent $/барр.',
                                                             'Brent руб./т.']].apply(lambda x: round(x, 2))
    chats_data = add_to_json(
        brent_df, "Brent на дату отгрузки", "plots", chats_data=chats_data)

    # курсовая разница
    profit_before_col = rows.get_column('AJ')
    profit_after_col = rows.get_column('AR')
    profit_usd_col = rows.get_column('AI')
    diff_col = rows.get_column('AS')
    date_after_col = rows.get_column('P')
    usd_before_col = rows.get_column('Y')
    usd_after_col = rows.get_column('Z')
    profit_before = []
    profit_after = []
    profit_usd = []
    diff = []
    date_after = []
    date = []
    usd_before = []
    usd_after = []

    for item_before, item_after, item_diff, item_date_after, item_date, item_profit_usd, item_usd_before, item_usd_after in zip(
            list(profit_before_col.values())[3:], list(
                profit_after_col.values())[3:], list(diff_col.values())[3:],
            list(date_after_col.values())[3:], list(
                date_column.values())[3:],
            list(profit_usd_col.values())[3:], list(usd_before_col.values())[3:], list(usd_after_col.values())[3:]):
        if (item_before is None) or (item_after is None) or (item_diff is None) or (item_date_after is None) or (
                item_date is None):
            break
        profit_before.append(item_before)
        profit_after.append(item_after)
        diff.append(item_diff)
        date_after.append(item_date_after)
        date.append(item_date)
        profit_usd.append(item_profit_usd)
        usd_before.append(item_usd_before)
        usd_after.append(item_usd_after)

    profit_df = pd.DataFrame({'Сумма реализации usd': profit_usd,
                              'Сумма реализация руб.': profit_before,
                              'Сумма поступления руб.': profit_after,
                              'Дата отгрузки': date,
                              'Дата поступления выручки': date_after,
                              'Курс на дату реализации': usd_before,
                              'Курс на дату поступления': usd_after,
                              'Курсовая разница руб.': diff})

    profit_df['Дата отгрузки'] = pd.to_datetime(
        profit_df['Дата отгрузки']).dt.strftime('%Y-%m-%d')
    profit_df['Дата поступления выручки'] = pd.to_datetime(
        profit_df['Дата поступления выручки']).dt.strftime('%Y-%m-%d')
    profit_df[['Сумма реализация руб.', 'Сумма поступления руб.', 'Курсовая разница руб.', 'Сумма реализации usd',
               'Курс на дату реализации', 'Курс на дату поступления']] = profit_df[
        ['Сумма реализация руб.', 'Сумма поступления руб.', 'Курсовая разница руб.', 'Сумма реализации usd',
         'Курс на дату реализации', 'Курс на дату поступления']].map(lambda x: x if is_numeric(x) else 0)

    data = [
        {"name": "Сумма реализации usd", "series": list(
            profit_df["Сумма реализации usd"])},
        {"name": "Сумма реализация руб.", "series": list(
            profit_df["Сумма реализация руб."])},
        {"name": "Сумма поступления руб.", "series": list(
            profit_df["Сумма поступления руб."])},
        {"name": "Дата отгрузки", "series": list(profit_df["Дата отгрузки"])},
        {"name": "Дата поступления выручки", "series": list(
            profit_df["Дата поступления выручки"])},
        {"name": "Курс на дату реализации", "series": list(
            profit_df["Курс на дату реализации"])},
        {"name": "Курс на дату поступления", "series": list(
            profit_df["Курс на дату поступления"])},
        {"name": "Курсовая разница руб.", "series": list(
            profit_df["Курсовая разница руб."])},
    ]
    plot = {
        "plot_name": "Разница в итоговой прибыли при вариации сроков реализации этапов",
        "x": "Сумма реализации usd",
        "y": ["Сумма реализация руб.", "Сумма поступления руб."],
        "hidden": [
            [
                "Курс на дату реализации",
                "Дата отгрузки"
            ],
            [
                "Курс на дату поступления",
                "Дата поступления выручки"
            ]
        ]
    }

    chats_data["data"] = chats_data["data"] + data
    chats_data["charts"]['plots'].append(plot)

    excel = Excel()
    excel.open_workbook(app1_file_path, data_only=True)
    excel.set_active_worksheet('Анализ_БК+ББ')
    rows = excel.read_worksheet_as_table()

    app1_data = pd.read_excel(
        app1_file_path, sheet_name='Анализ_БК+ББ')
    # company_head = app1_data[]

    # Компании по которым мы анализизируем клиентов
    companys_names = ['Компания 1', 'Company ABC',
                      'A-Нефтегаз', 'Компания ААА', 'Компания АВА']

    companys_clients = app1_data[app1_data.columns[9]]  # 'Покупатель']
    sub_columns = app1_data.iloc[0]

    companys_index = companys_clients.index[companys_clients.isin(
        companys_names)]

    conditions_revenues = []
    # for i in range(len(companys_names)):
    for i in range(1):
        company_client = companys_clients[companys_index[i]                                          : companys_index[i + 1]]

        date_clients = app1_data[app1_data.columns[12]
                                 ].iloc[companys_index[i]: companys_index[i + 1]]

        company_rows = rows.get_slice(companys_index[i], companys_index[i + 1])
        fob = company_rows.get_column('BR').values()
        revenue = company_rows.get_column('CS').values()
        conditions = company_rows.get_column('K').values()

        # Creating Dictionary
        frame = {
            companys_names[i]: company_client,
            app1_data.columns[12]: date_clients,
            'Цена реализации FOB, $/bb': fob,
            'revenue': revenue,
            'conditions': conditions
        }
        # Creating Dataframe
        client_df = pd.DataFrame(frame)
        client_df = client_df.dropna()
        # logging.info(client_df.head())
        conditions_revenue = client_df.groupby(['conditions'], as_index=False)[
            'revenue'].mean()
        # logging.info(company_client)
    chats_data = add_to_json(
        conditions_revenue, "Средняя выручка по условиям договора", "bar chart", chats_data=chats_data)

    # Селектор по компаниям
    counter = 1
    row = rows.get_row(counter)
    black_list = ['Переоценка остатков ДС', 'Продажа валюты',
                  'Переоценка ДЗ и КЗ на конец каждого периода']
    companies_list = ['Компания 1', 'Company ABC',
                      'A-Нефтегаз', 'Компания ААА']
    companies = defaultdict(dict)

    try:
        while True:
            if (row['A'] is None) and (row['Y'] is None) and (row['Z'] is None):
                company_name = row['J']
                if (company_name not in companies.keys()) and (company_name not in black_list) and (
                        rows.get_row(counter + 1)['J'] is not None) and (company_name in companies_list):
                    companies[company_name]['row_index'] = counter + 1
            counter += 1
            row = rows.get_row(counter)
    except IndexError:
        print(f'Компании закончились')

    # Создание списка колонок для всех компаний
    def add_colum_to_list(companies, rows, column_name, name, month=False):
        for company_elem in companies.keys():
            company_row_index = companies[company_elem]['row_index']
            added_column = rows.get_column(column_name)
            columns_list = []
            for item in list(added_column.values())[company_row_index:]:
                if month and (item is not None) and (item != '#DIV/0!'):
                    columns_list.append(item.strftime('%B'))
                    continue
                if (item is not None) and (item != '#DIV/0!'):
                    columns_list.append(item)
            companies[company_elem][name] = columns_list
        return companies

    # Достаем все нужные колонки
    companies = add_colum_to_list(
        companies, rows, 'L', 'Месяц отгрузки', month=True)
    companies = add_colum_to_list(
        companies, rows, 'O', 'Месяц поступления выручки', month=True)
    companies = add_colum_to_list(
        companies, rows, 'CZ', 'Коэффициент перевода барр./т.', month=False)
    companies = add_colum_to_list(
        companies, rows, 'DB', "Допдоход к рынку (к цене НДПИ Platts) млн руб", month=False)
    companies = add_colum_to_list(
        companies, rows, 'DA', "Допдоход к рынку (к цене НДПИ Argus) млн руб", month=False)
    companies = add_colum_to_list(
        companies, rows, 'DC', "Зачисление процентов,млн руб", month=False)
    companies = add_colum_to_list(
        companies, rows, 'DD', "Макропараметры (brent/ discount/ escalation) млн руб.", month=False)
    companies = add_colum_to_list(
        companies, rows, 'DE', "Макропараметры (spread) млн руб.", month=False)
    companies = add_colum_to_list(
        companies, rows, 'DF', "Freight, млн руб.", month=False)
    companies = add_colum_to_list(
        companies, rows, 'DG', "Коммерческие за искл. Freight, млн руб.", month=False)
    companies = add_colum_to_list(
        companies, rows, 'DH', "Экспортная пошлина, млн руб.", month=False)
    companies = add_colum_to_list(
        companies, rows, 'DI', "Курсовые разницы (отгрузка / оплата)млн руб", month=False)
    companies = add_colum_to_list(
        companies, rows, 'DJ', "откл НДПИ (макропараметры)", month=False)
    companies = add_colum_to_list(
        companies, rows, 'DK', "откл НДПИ уплач от расч НДПИ (курс оплаты)млн руб", month=False)
    companies = add_colum_to_list(
        companies, rows, 'DL', "откл НДПИ (Argus) и НДПИ (Plats Urals Med), млн руб.", month=False)
    companies = add_colum_to_list(
        companies, rows, 'DM', "Эффект", month=False)

    # Отклонение по макропараметрам
    columns = ["Допдоход к рынку (к цене НДПИ Platts) млн руб",
               "Допдоход к рынку (к цене НДПИ Argus) млн руб",
               "Зачисление процентов,млн руб",
               "Макропараметры (brent/ discount/ escalation) млн руб.",
               "Макропараметры (spread) млн руб.",
               "Freight, млн руб.",
               "Коммерческие за искл. Freight, млн руб.",
               "Экспортная пошлина, млн руб.",
               "Курсовые разницы (отгрузка / оплата)млн руб",
               "откл НДПИ (макропараметры)",
               "откл НДПИ уплач от расч НДПИ (курс оплаты)млн руб",
               "откл НДПИ (Argus) и НДПИ (Plats Urals Med), млн руб.",
               "Эффект"]
    for company in list(companies.keys())[:2]:
        min_length = min([len(companies[company][item])
                          for item in columns])
        df = pd.DataFrame(data={item: companies[company][item][:min_length]
                                for item in columns})
        df["Эффект"] = df["Эффект"].apply(lambda x: None)
        df = df.transpose().reset_index().rename(
            columns={'index': 'Макропараметры'})
        num_columns = len(df.columns) - 1
        new_column_names = [f'сделка_{i}' for i in range(1, num_columns + 1)]
        df = df.rename(columns={col: new_name for col, new_name in zip(
            df.columns[1:], new_column_names)})
        chats_data = add_to_json(
            df, f"Отклонения в макропараметрах для компании для {company}", "waterfall", chats_data=chats_data)

    # Коэффициент перевода
    columns = ['Месяц отгрузки', 'Коэффициент перевода барр./т.']
    for company in list(companies.keys())[:2]:
        min_length = min([len(companies[company][item])
                          for item in columns])
        df = pd.DataFrame(data={item: companies[company][item][:min_length]
                                for item in columns})
        df = df.groupby('Месяц отгрузки', as_index=False).mean()
        chats_data = add_to_json(
            df, f"Коэффициент перевода барр./т. в месяц отгрузки {company}", "bar_chart", chats_data=chats_data)

    columns = ['Месяц поступления выручки', 'Коэффициент перевода барр./т.']
    for company in list(companies.keys())[:2]:
        min_length = min([len(companies[company][item])
                          for item in columns])
        df = pd.DataFrame(data={item: companies[company][item][:min_length]
                                for item in columns})
        df = df.groupby('Месяц поступления выручки', as_index=False).mean()
        chats_data = add_to_json(
            df, f"Коэффициент перевода барр./т. в месяц поступления выручки {company}", "bar_chart",
            chats_data=chats_data)

    column = 'Месяц отгрузки'
    for company in list(companies.keys())[:2]:
        df = pd.DataFrame(data={column: companies[company][column]})
        df = df.groupby('Месяц отгрузки', as_index=False).size()
        df = df.rename(columns={"size": "Количество сделок"})
        chats_data = add_to_json(
            df, f"Количество сделок по месяцам {company}", "bar_chart", chats_data=chats_data)

        # data = json.dumps(chats_data)
    with open(os.path.join(shared_directory, 'workitems.json'), "w", encoding='utf-8') as outfile:
        json.dump(chats_data, outfile, ensure_ascii=False)

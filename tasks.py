from collections import defaultdict
from bs4 import BeautifulSoup
from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import Files as Excel
from robocorp.tasks import task
from robocorp import workitems
from selenium.webdriver.common.keys import Keys
from openpyxl import Workbook
import pandas as pd
import logging
import os
import time
import re
import json
import openpyxl


output_directory = os.environ.get("ROBOT_ARTIFACTS")
shared_directory = os.path.join(output_directory, "shared")


try:
    os.remove('parse_log.log')
except OSError:
    pass

logging.basicConfig(filename='parse_log.log', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s',
                    encoding='utf-8')


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
def web_preprocessor():
    from first_table import pars_usd, customs_duties_parser, parcing_first_table, parsing_brent_cost, urals_parser, second_table
    print("Start parser")
    logging.info("Start parser")
    pars_usd.usd_kurs()
    customs_duties_parser.test()
    parcing_first_table.test()
    parsing_brent_cost.test()
    urals_parser.test()
    second_table.to_application_1()


@task
def after_update_postprocessor():
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
    excel.open_workbook('./first_table/Приложение 1.xlsx', data_only=True)
    excel.set_active_worksheet('Анализ_БК+ББ')

    freight_usd = []
    freight_rub = []
    date = []

    rows = excel.read_worksheet_as_table()
    freight_column_usd = rows.get_column('AD')
    freight_column_rub = rows.get_column('AV')
    date_column = rows.get_column('M')
    for item_usd, item_rub, item_date in zip(list(freight_column_usd.values())[3:], list(freight_column_rub.values())[3:], list(date_column.values())[3:]):
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
    for item_usd, item_rub, item_date in zip(list(spread_usd_col.values())[3:], list(spread_rub_col.values())[3:], list(date_column.values())[3:]):
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
                                                                 'Spread руб./т.']].map(lambda x: x if is_numeric(x) else 0)
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

    for item_usd, item_rub, item_date in zip(list(brent_usd_col.values())[3:], list(brent_rub_col.values())[3:], list(date_column.values())[3:]):
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
    for item_before, item_after, item_diff, item_date_after, item_date, item_profit_usd, item_usd_before, item_usd_after in zip(list(profit_before_col.values())[3:], list(profit_after_col.values())[3:], list(diff_col.values())[3:],
                                                                                                                                list(date_after_col.values())[3:], list(
                                                                                                                                    date_column.values())[3:],
                                                                                                                                list(profit_usd_col.values())[3:], list(usd_before_col.values())[3:], list(usd_after_col.values())[3:]):
        if (item_before is None) or (item_after is None) or (item_diff is None) or (item_date_after is None) or (item_date is None):
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
               'Курс на дату реализации', 'Курс на дату поступления']] = profit_df[['Сумма реализация руб.', 'Сумма поступления руб.', 'Курсовая разница руб.', 'Сумма реализации usd',
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
    excel.open_workbook('./first_table/Приложение 1.xlsx', data_only=True)
    excel.set_active_worksheet('Анализ_БК+ББ')
    rows = excel.read_worksheet_as_table()

    app1_data = pd.read_excel(
        './first_table/Приложение 1.xlsx', sheet_name='Анализ_БК+ББ')
    # company_head = app1_data[]
    logging.info(app1_data.info())
    logging.info(app1_data.head())

    # Компании по которым мы анализизируем клиентов
    companys_names = ['Компания 1', 'Company ABC',
                      'A-Нефтегаз', 'Компания ААА', 'Компания АВА']
    logging.info((app1_data.columns[9], app1_data.columns[12]))
    companys_clients = app1_data[app1_data.columns[9]]  # 'Покупатель']
    sub_columns = app1_data.iloc[0]

    companys_index = companys_clients.index[companys_clients.isin(
        companys_names)]

    conditions_revenues = []
    # for i in range(len(companys_names)):
    for i in range(1):
        company_client = companys_clients[companys_index[i]
            : companys_index[i+1]]
        logging.info([companys_index[i], companys_index[i+1]])

        date_clients = app1_data[app1_data.columns[12]
                                 ].iloc[companys_index[i]: companys_index[i+1]]

        company_rows = rows.get_slice(companys_index[i], companys_index[i+1])
        fob = company_rows.get_column('BR').values()
        revenue = company_rows.get_column('CS').values()
        conditions = company_rows.get_column('K').values()
        logging.info(fob)
        logging.info(revenue)
        logging.info(conditions)

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

    # data = json.dumps(chats_data)
    with open(os.path.join(shared_directory, 'workitems.json'), "w", encoding='utf-8') as outfile:
        json.dump(chats_data, outfile, ensure_ascii=False)
    logging.info(chats_data)

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


output_directory = os.environ.get("ROBOT_ARTIFACTS")
shared_directory = os.path.join(output_directory, "shared")


logging.basicConfig(filename='parse_log.log', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s',
                    encoding='utf-8')

@task
def add_urals():
    parser = Selenium()
    urls = {
        'декабрь 2021': 'https://www.economy.gov.ru/material/directions/vneshneekonomicheskaya_deyatelnost/tamozhenno_tarifnoe_regulirovanie/o_vyvoznyh_tamozhennyh_poshlinah_na_neft_i_otdelnye_kategorii_tovarov_vyrabotannyh_iz_nefti_na_period_s_1_po_31_yanvarya_2022_goda.html',
        'январь': 'https://www.economy.gov.ru/material/directions/vneshneekonomicheskaya_deyatelnost/tamozhenno_tarifnoe_regulirovanie/o_vyvoznyh_tamozhennyh_poshlinah_na_neft_i_otdelnye_kategorii_tovarov_vyrabotannyh_iz_nefti_na_period_s_1_po_28_fevralya_2022_goda.html',
        'февраль': 'https://www.economy.gov.ru/material/directions/vneshneekonomicheskaya_deyatelnost/tamozhenno_tarifnoe_regulirovanie/o_vyvoznyh_tamozhennyh_poshlinah_na_neft_i_otdelnye_kategorii_tovarov_vyrabotannyh_iz_nefti_na_period_s_1_po_31_marta_2022_goda.html',
        'март': 'https://www.economy.gov.ru/material/directions/vneshneekonomicheskaya_deyatelnost/tamozhenno_tarifnoe_regulirovanie/o_vyvoznyh_tamozhennyh_poshlinah_na_neft_i_otdelnye_kategorii_tovarov_vyrabotannyh_iz_nefti_na_period_s_1_po_30_aprelya_2022_goda.html',
        'апрель': 'https://www.economy.gov.ru/material/directions/vneshneekonomicheskaya_deyatelnost/tamozhenno_tarifnoe_regulirovanie/o_vyvoznyh_tamozhennyh_poshlinah_na_neft_i_otdelnye_kategorii_tovarov_vyrabotannyh_iz_nefti_na_period_s_1_po_31_maya_2022_goda.html',
        'май': 'https://www.economy.gov.ru/material/directions/vneshneekonomicheskaya_deyatelnost/tamozhenno_tarifnoe_regulirovanie/o_vyvoznyh_tamozhennyh_poshlinah_na_neft_i_otdelnye_kategorii_tovarov_vyrabotannyh_iz_nefti_na_period_s_1_po_30_iyunya_2022_goda.html',
        'июнь': 'https://www.economy.gov.ru/material/directions/vneshneekonomicheskaya_deyatelnost/tamozhenno_tarifnoe_regulirovanie/o_vyvoznyh_tamozhennyh_poshlinah_na_neft_i_otdelnye_kategorii_tovarov_vyrabotannyh_iz_nefti_na_period_s_1_po_31_iyulya_2022_goda.html',
        'июль': 'https://www.economy.gov.ru/material/directions/vneshneekonomicheskaya_deyatelnost/tamozhenno_tarifnoe_regulirovanie/o_vyvoznyh_tamozhennyh_poshlinah_na_neft_i_otdelnye_kategorii_tovarov_vyrabotannyh_iz_nefti_na_period_s_1_po_31_avgusta_2022_goda.html',
        'август': 'https://www.economy.gov.ru/material/directions/vneshneekonomicheskaya_deyatelnost/tamozhenno_tarifnoe_regulirovanie/o_vyvoznyh_tamozhennyh_poshlinah_na_neft_i_otdelnye_kategorii_tovarov_vyrabotannyh_iz_nefti_na_period_s_1_po_30_sentyabrya_2022_goda.html',
        'сентябрь': 'https://www.economy.gov.ru/material/directions/vneshneekonomicheskaya_deyatelnost/tamozhenno_tarifnoe_regulirovanie/o_vyvoznyh_tamozhennyh_poshlinah_na_neft_i_otdelnye_kategorii_tovarov_vyrabotannyh_iz_nefti_na_period_s_1_po_31_oktyabrya_2022_goda.html',
        'октябрь': 'https://www.economy.gov.ru/material/directions/vneshneekonomicheskaya_deyatelnost/tamozhenno_tarifnoe_regulirovanie/o_vyvoznyh_tamozhennyh_poshlinah_na_neft_i_otdelnye_kategorii_tovarov_vyrabotannyh_iz_nefti_na_period_s_1_po_30_noyabrya_2022_goda.html',
        'ноябрь': 'https://www.economy.gov.ru/material/directions/vneshneekonomicheskaya_deyatelnost/tamozhenno_tarifnoe_regulirovanie/o_vyvoznyh_tamozhennyh_poshlinah_na_neft_i_otdelnye_kategorii_tovarov_vyrabotannyh_iz_nefti_na_period_s_1_po_31_dekabrya_2022_goda.html',
        'декабрь': 'https://www.economy.gov.ru/material/directions/vneshneekonomicheskaya_deyatelnost/tamozhenno_tarifnoe_regulirovanie/o_vyvoznyh_tamozhennyh_poshlinah_na_neft_i_otdelnye_kategorii_tovarov_vyrabotannyh_iz_nefti_na_period_s_1_po_31_yanvarya_2023_goda.html'
    }
    payload = {}
    urals_payload = {}
    tax_rates = {}
    keys = list(urls.keys())
    for index, month in enumerate(keys):
        parser.open_available_browser(urls[month])
        if index != 0:
            urals = parser.get_text("xpath=//table/tbody/tr[2]/td/p")
            urals = urals.replace(',', '.')
            urals_payload[month] = float(urals.split(' ')[0])

        if index != (len(urls)-1):
            tax_rate = parser.get_text("xpath=//table/tbody/tr/td[3]/p")
            tax_rate = tax_rate.replace(',', '.')
            tax_rates[keys[index+1]] = float(tax_rate)
        parser.close_browser()
    payload = {
        "urals_payload": urals_payload,
        "tax_rates": tax_rates
    }
    # workitems.outputs.create(urals_payload)
    with open(os.path.join(shared_directory, 'workitems.json'), "w") as outfile:
        json.dump(payload, outfile)


@task
def load_new_external_data_to_excel():
    app1_data = pd.read_excel('Приложение_1.xlsx', sheet_name='Анализ_БК+ББ')
    # company_head = app1_data[]
    logging.info(app1_data.info())
    logging.info(app1_data.head())

    # Компании по которым мы анализизируем клиентов
    companys_names = ['Компания 1', 'Company ABC',
                      'A-Нефтегаз', 'Компания ААА', 'Компания АВА']

    logging.info(app1_data.columns)
    companys_clients = app1_data['Покупатель']
    sub_columns = app1_data.iloc[0]

    companys_index = companys_clients.index[companys_clients.isin(
        companys_names)]

    clients = []
    for i in range(len(companys_names)):
        client = companys_clients[companys_index[i]: companys_index[i+1]]
        logging.info([companys_index[i], companys_index[i+1]])
        logging.info(client)
        clients.append(client)

    # Добавление новых клиентов в шаблон
    app2_data = pd.read_excel('компания 1 заказчик.xlsx')
    logging.info(app2_data.head())
    new_clients = app2_data['Клиент']
    logging.info(app2_data.columns)

    # TODO: append new_clients to template


# @task
# def load_input_to_excel():

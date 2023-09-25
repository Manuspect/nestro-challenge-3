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
def usd_kurs():
    # Определите даты начала и конца
    start_date = "01.01.2022"
    end_date = "31.12.2022"
    try:
        # Инициализируем Selenium
        browser = Selenium()

        # Создаем URL с параметрами дат
        url = f"https://www.cbr.ru/currency_base/dynamics/?UniDbQuery.Posted=True&UniDbQuery.so=1&UniDbQuery.mode=1&UniDbQuery.date_req1={start_date}&UniDbQuery.date_req2={end_date}&UniDbQuery.VAL_NM_RQ=R01235"
        browser.open_available_browser(url)
        logging.info(f"Открыли сайт")
        # Дождитесь полной загрузки страницы
        browser.wait_until_element_is_visible(
            "//*[@id='content']/div/div/div/div[2]/div[1]/table/tbody/tr[1]", 15)

        # Создаем пустой список для хранения данных
        data = []

        # Создаем словарь для хранения кварталов
        quarters = {1: [], 2: [], 3: [], 4: []}

        # Итерируемся по дням от 239 до 1
        for i in range(239, 0, -1):
            # Получаем текст из ячейки таблицы
            row = browser.get_text(
                f'//*[@id="content"]/div/div/div/div[2]/div[1]/table/tbody/tr[{i}]')
            date = row[:10]
            rate = row[12:]
            month = date[3:5]
            logging.info(f"Текущий месяц: {month}")
            rate = rate.replace(',', '.')
            # Добавляем день и его курс в общий список
            try:
                data.append([date, float(rate)])

                # Добавляем день в соответствующий квартал
                day_number = int(date[:2])
                if 1 <= day_number <= 31:
                    quarters[1].append(float(rate))
                elif 32 <= day_number <= 61:
                    quarters[2].append(float(rate))
                elif 62 <= day_number <= 92:
                    quarters[3].append(float(rate))
                elif 93 <= day_number <= 123:
                    quarters[4].append(float(rate))
            except ValueError:
                logging.info(f"Not a float: {month}")

        browser.close_browser()

        # Создаем DataFrame из данных
        df = pd.DataFrame(data, columns=["Дата", "Курс"])

        # Добавляем столбец с месяцем
        df['Месяц'] = df['Дата'].str[3:5]

        # Считаем средний курс по месяцам
        monthly_avg = df.groupby('Месяц')['Курс'].mean().reset_index()
        monthly_avg.columns = ["Месяц", "Средний курс"]

        # Считаем средний курс по кварталам
        quarterly_avg = {}
        for quarter, days in quarters.items():
            if days:
                avg_rate = sum(days) / len(days)
                quarterly_avg[f"Квартал {quarter}"] = [avg_rate]

        quarterly_avg_df = pd.DataFrame.from_dict(
            quarterly_avg, orient='index', columns=["Средний курс"])

        # Сохраняем результаты в Excel файлы
        monthly_avg.to_excel("monthly_exchange_rate.xlsx", index=False)
        quarterly_avg_df.to_excel("quarterly_exchange_rate.xlsx")

    except Exception as e:
        print(f"Произошла ошибка: {str(e)}")
        logging.error(f"Произошла ошибка: {str(e)}")
    finally:
        # Закрываем браузер после использования
        browser.close_browser()


@task
def parse_chrome():
    # Создаем экземпляр браузера Selenium
    browser = Selenium()
    params_chrome = "2022 сентябрь"
    try:
        # Открываем Google и вводим запрос
        browser.open_available_browser("https://www.google.com")
        browser.input_text("name=q", f"цена нефти urals {params_chrome}")
        browser.press_keys("name=q", Keys.ENTER)

        browser.wait_until_element_is_visible("//*[@alt='Google']", 15)

        # Ждем, пока страница загрузится
        page_source = browser.get_source()

        # Используем BeautifulSoup для парсинга страницы
        soup = BeautifulSoup(page_source, "html.parser")

        # Находим элемент с классами HwtpBd gsrt PZPZlf kTOYnf
        result_element = soup.find("div", class_="HwtpBd gsrt PZPZlf kTOYnf")

        # Внутри этого элемента находим элемент с классом IZ6rdc (цена)
        price_element = result_element.find("div", class_="IZ6rdc")

        # Получаем текст из элемента с ценой и выводим его
        price = price_element.get_text()
        logging.info(f"Цена нефти Urals за сентябрь 2022 года: {price}")

        # Создаем новый файл Excel
        wb = Workbook()

        # Активируем лист
        sheet = wb.active

        # Добавляем заголовки для колонок
        sheet.append(["Год Месяц", "Цена"])

        # Сохраняем новый файл Excel
        wb.save("цены_нефти.xlsx")

        current_date = params_chrome

        # Разделяем текущую дату на месяц и год
        date = params_chrome

        # Добавляем запись в таблицу Excel
        row = [date, price]
        sheet.append(row)

        # Сохраняем изменения
        wb.save("цены_нефти.xlsx")

    finally:
        # Закрываем браузер после использования
        browser.close_browser()


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
def pars_exel():
    # Загрузите данные из monthly_exchange_rate.xlsx
    monthly_data = pd.read_excel('monthly_exchange_rate.xlsx').T

    months = {'январь': 1, 'февраль': 2, 'март': 3, 'апрель': 4, 'май': 5,
              'июнь': 6, 'июль': 7, 'август': 8, 'сентябрь': 9, 'октябрь': 10,
              'ноябрь': 11, 'декабрь': 12}

    logging.info(monthly_data)

    # Загрузите данные из Приложение_1.xlsx
    app1_data = pd.read_excel('Приложение 1.xlsx', header=[0, 1])
    excel = Excel()
    excel.open_workbook('Приложение 1.xlsx')
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
    excel.save_workbook('./result.xlsx')

    # Загрузите данные из цены_нефти.xlsx
    oil_price_data = pd.read_excel('цены_нефти.xlsx')

    # Загрузите данные из Приложение_1.xlsx
    app1_data = pd.read_excel('Приложение_1.xlsx')

    # Задайте столбцы, с которыми нужно сопоставить данные
    columns_to_match = ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль", "Август", "Сентябрь", "Октябрь",
                        "Ноябрь", "Декабрь"]
# Обновите данные в Приложение_1.xlsx
    for column in columns_to_match:
        app1_data[column] = oil_price_data[column]

    # Сохраните обновленные данные в Приложение_1.xlsx
    app1_data.to_excel('Приложение_1.xlsx', index=False)

    # Загрузите данные из quarterly_exchange_rate.xlsx
    quarterly_data = pd.read_excel('quarterly_exchange_rate.xlsx')

    # Загрузите данные из Приложение_1.xlsx
    app1_data = pd.read_excel('Приложение_1.xlsx')

    # Задайте столбцы кварталов, с которыми нужно сопоставить данные
    quarters_to_match = ["1 кв", "2 кв", "3 кв", "4 кв"]

    # Обновите данные в Приложение_1.xlsx
    for quarter in quarters_to_match:
        app1_data[quarter] = quarterly_data["Средний курс"]

    # Сохраните обновленные данные в Приложение_1.xlsx
    app1_data.to_excel('Приложение_1.xlsx', index=False)


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

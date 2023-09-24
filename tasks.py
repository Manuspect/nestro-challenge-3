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


@task
def producer():
    """Split Excel rows into multiple output Work Items for the next step."""
    for item in workitems.inputs:
        output_directory = os.environ.get("ROBOT_ARTIFACTS")
        name = "orders.xlsx"
        path = item.get_file(name, os.path.join(output_directory, name))

        excel = Excel()
        excel.open_workbook(path)
        rows = excel.read_worksheet_as_table(header=True)

        for row in rows:
            payload = {
                "Name": row["Name"],
                "Zip": row["Zip"],
                "Product": row["Item"],
            }
            workitems.outputs.create(payload)


@task
def consumer():
    """Process all the produced input Work Items from the previous step."""
    for item in workitems.inputs:
        try:
            name = item.payload["Name"]
            address = item.payload["Zip"]
            product = item.payload["Product"]
            print(f"Processing order: {name}, {address}, {product}")
            item.done()
        except KeyError as err:
            item.fail(code="MISSING_VALUE", message=str(err))

@task
def usd_kurs():
    # Определите даты начала и конца
    start_date = "01.01.2022"
    end_date = "31.12.2022"

    logging.basicConfig(filename='parse_log.log', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s',
                        encoding='utf-8')

    try:
        # Инициализируем Selenium
        browser = Selenium()

        # Создаем URL с параметрами дат
        url = f"https://www.cbr.ru/currency_base/dynamics/?UniDbQuery.Posted=True&UniDbQuery.so=1&UniDbQuery.mode=1&UniDbQuery.date_req1={start_date}&UniDbQuery.date_req2={end_date}&UniDbQuery.VAL_NM_RQ=R01235"
        browser.open_available_browser(url)
        logging.info(f"Открыли сайт")

        # Дождитесь полной загрузки страницы
        browser.wait_until_element_is_visible("//*[@id='content']/div/div/div/div[2]/div[1]/table/tbody/tr[1]")

        # Создаем пустой список для хранения данных
        data = []

        # Создаем словарь для хранения кварталов
        quarters = {1: [], 2: [], 3: [], 4: []}

        # Итерируемся по дням от 239 до 1
        for i in range(239, 0, -1):
            # Получаем текст из ячейки таблицы
            row = browser.get_text(f'//*[@id="content"]/div/div/div/div[2]/div[1]/table/tbody/tr[{i}]')
            date = row[:10]
            rate = row[12:]
            month = date[3:5]
            logging.info(f"Текущий месяц: {month}")

            # Добавляем день и его курс в общий список
            data.append([date, float(rate.replace(',', '.'))])

            # Добавляем день в соответствующий квартал
            day_number = int(date[:2])
            if 1 <= day_number <= 31:
                quarters[1].append(float(rate.replace(',', '.')))
            elif 32 <= day_number <= 61:
                quarters[2].append(float(rate.replace(',', '.')))
            elif 62 <= day_number <= 92:
                quarters[3].append(float(rate.replace(',', '.')))
            elif 93 <= day_number <= 123:
                quarters[4].append(float(rate.replace(',', '.')))

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

        quarterly_avg_df = pd.DataFrame.from_dict(quarterly_avg, orient='index', columns=["Средний курс"])

        # Сохраняем результаты в Excel файлы
        monthly_avg.to_excel("monthly_exchange_rate.xlsx", index=False)
        quarterly_avg_df.to_excel("quarterly_exchange_rate.xlsx")

    except Exception as e:
        print(f"Произошла ошибка: {str(e)}")
        logging.error(f"Произошла ошибка: {str(e)}")


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

        # Ждем, пока страница загрузится
        time.sleep(60)
        page_source = browser.get_source()

        # Используем BeautifulSoup для парсинга страницы
        soup = BeautifulSoup(page_source, "html.parser")

        # Находим элемент с классами HwtpBd gsrt PZPZlf kTOYnf
        result_element = soup.find("div", class_="HwtpBd gsrt PZPZlf kTOYnf")

        # Внутри этого элемента находим элемент с классом IZ6rdc (цена)
        price_element = result_element.find("div", class_="IZ6rdc")

        # Получаем текст из элемента с ценой и выводим его
        price = price_element.get_text()
        print("Цена нефти Urals за сентябрь 2022 года:", price)

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
def pars_exel():
    # Загрузите данные из monthly_exchange_rate.xlsx
    monthly_data = pd.read_excel('monthly_exchange_rate.xlsx')

    # Загрузите данные из Приложение_1.xlsx
    app1_data = pd.read_excel('Приложение_1.xlsx')

    # Задайте столбцы, с которыми нужно сопоставить данные
    columns_to_match = ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль", "Август", "Сентябрь", "Октябрь",
                        "Ноябрь", "Декабрь"]

    # Обновите данные в Приложение_1.xlsx
    for column in columns_to_match:
        app1_data[column] = monthly_data[column]

    # Сохраните обновленные данные в Приложение_1.xlsx
    app1_data.to_excel('Приложение_1.xlsx', index=False)

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


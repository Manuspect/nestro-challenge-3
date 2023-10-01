from RPA.Browser.Selenium import Selenium
import logging
import pandas as pd


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
            "//*[@id='content']/div/div/div/div[2]/div[1]/table/tbody/tr[1]", 30)

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

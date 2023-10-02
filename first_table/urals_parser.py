from bs4 import BeautifulSoup as bs
from urllib.request import Request, urlopen
import openpyxl
import logging


def get_prices_for_month(month):
    url = f"https://www.economy.gov.ru/material/departments/d12/konyunktura_mirovyh_tovarnyh_rynkov/o_sredney_cene_na_neft_sorta_yurals_za_{month}_2022_goda.html"
    req = Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    web_byte = urlopen(req).read()
    webpage = web_byte.decode('utf-8')
    soup = bs(webpage, "html.parser")
    price_text = [x.text for x in soup.findAll('p') if "США за баррель" in x.text]
    if price_text:
        return float(price_text[0].split()[0].replace(",", "."))
    return None


def start(file_path: str):
    logging.info('urals_parser')
    months = ['yanvar', 'fevral', 'mart', 'aprel', 'may', 'iyun', 'iyul',
              'avgust', 'sentyabr', 'oktyabr', 'noyabr', 'dekabr']

    prices = []

    for month in months:
        price = get_prices_for_month(month)
        if price is not None:
            prices.append(price)

    wb = openpyxl.load_workbook(file_path)
    sheet = wb['Компания 1_факт_НДПИ (Platts)']

    for i, price in enumerate(prices):
        col_letter = chr(ord('B') + i)  # Преобразование индекса в букву столбца
        start_cell = f"{col_letter}14"
        end_cell = f"{col_letter}171"

        print(f"Setting value {price} in cells {start_cell} to {end_cell}")

        sheet[start_cell].value = price
        sheet[end_cell].value = price

    wb.save(file_path)
    wb.close()


from bs4 import BeautifulSoup as bs
from urllib.request import Request, urlopen
import numpy as np
import pandas as pd
import openpyxl
import logging


def start(file_path: str):
    logging.info('urals_parser')
    url = ["https://www.economy.gov.ru/material/departments/d12/konyunktura_mirovyh_tovarnyh_rynkov/o_sredney_cene_na_neft_sorta_yurals_za_", "_2022_goda", ".html"]
    months = ['yanvar', 'fevral', 'mart', 'aprel', 'may', 'iyun', 'iyul',
              'avgust', 'sentyabr', 'oktyabr', 'noyabr', 'dekabr']

    prices = []

    for i, month in enumerate(months):
        print(month)
        if i == 0:
            url_link = url[0]+month+url[1]+'_'+url[2]
        else:
            url_link = url[0]+month+url[1]+url[2]
        print(url_link)
        req = Request(url_link, headers={'User-Agent': 'Mozilla/5.0'})
        web_byte = urlopen(req).read()

        webpage = web_byte.decode('utf-8')
        soup = bs(webpage, "html.parser")
        prices.append(float([x for x in soup.findAll(
            'p') if "США за баррель" in x.text][0].text.split()[0].replace(",", ".")))

    wb = openpyxl.load_workbook(file_path)
    sheet = wb['Компания 1_факт_НДПИ (Platts)']

    print(prices)

    for i in range(len(prices)):
        # sheet[] = quotes[i]
        print(f"{chr(67+i)}{14}", f"{chr(67+i)}{171}")
        sheet[f"{chr(67+i)}{14}"].value = prices[i]
        sheet[f"{chr(67+i)}{171}"].value = prices[i]

    wb.save(file_path)
    wb.close()

from bs4 import BeautifulSoup as bs
from urllib.request import Request, urlopen
import numpy as np
import pandas as pd
import openpyxl
import logging


def spread_quotes(url):
    req = Request(url, headers={'User-Agent': 'Mozilla/5.0'})
    web_byte = urlopen(req).read()

    webpage = web_byte.decode('utf-8')
    soup = bs(webpage, "html.parser")
    quotes = soup.find_all('td', class_=['datatable_cell__LJp3C datatable_cell--align-end__qgxDQ datatable_cell--up__hIuZF text-right text-sm font-normal leading-5 align-middle min-w-[77px] text-[#007C32]',
                                         'datatable_cell__LJp3C datatable_cell--align-end__qgxDQ datatable_cell--down___c4Fq text-right text-sm font-normal leading-5 align-middle min-w-[77px] text-[#D91400]'])
    quotes = np.array([float(x.text.replace(',', '.')) for x in quotes])
    return quotes


def start(file_path: str):
    logging.info('parcing_first_table')
    url = 'https://ru.investing.com/commodities/brent-wti-crude-spread-futures-historical-data'
    quotes = spread_quotes(url)
    print(quotes)

    wb = openpyxl.load_workbook(file_path)
    sheet = wb['Анализ_БК+ББ']

    for i in range(len(quotes)):
        # sheet[] = quotes[i]
        cell = sheet[f"C{4+i}"]
        print(f"C{4+i}", cell.value, quotes[i])
        logging.info(f"C{4+i}, {cell.value}, {quotes[i]}")
        set_cell_value(f"C{4+i}", quotes[i], sheet)

    wb.save(file_path)
    wb.close()


def set_cell_value(cell_key, new_value, sheet):
    from openpyxl.cell.cell import MergedCell
    cell = sheet[cell_key]
    if not isinstance(cell, MergedCell):
        cell.value = new_value
        return
    for merged_cells_range in sheet.merged_cells.ranges:
        if cell_key in merged_cells_range:
            print('merged cells range', merged_cells_range)
            logging.info(f'merged cells range: {merged_cells_range}')
            merged_cells_range.start_cell.value = new_value
            break

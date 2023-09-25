from bs4 import BeautifulSoup
import requests
import pandas as pd
import openpyxl
import numpy as np
import datetime


def generate_strings(letter, num_1, num_2):
    s = list([])
    for i in range(num_1, num_2 + 1):
        s.append(str(letter) + str(i))
    return s


def fill_xlsx(cell, letter):
    wb = openpyxl.load_workbook('Приложение 1.xlsx')
    sheet = wb['Анализ_БК+ББ']
    d = sheet[cell].value.date()
    d = str(d).split('-')[::-1]
    d = '.'.join(d)
    message = 0
    cost_table = pd.DataFrame(columns=["Date", "Brent_oil_cost"])
    url = 'https://ru.investing.com/commodities/brent-oil-historical-data'
    headers = {'user-Agent': 'Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:72.0) Gecko/20100101 Firefox/72.0',
               'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8'
               }
    response = requests.get(url, headers=headers)
    soup = BeautifulSoup(response.text, 'html.parser')
    rows = soup.findAll("tr",
                        class_="h-[41px] hover:bg-[#F5F5F5] relative after:absolute after:bottom-0 after:bg-[#ECEDEF] after:h-px after:left-0 after:right-0 historical-data-v2_price__atUfP")
    for el in rows:
        date = el.find("td",
                       class_="datatable_cell__LJp3C text-left align-middle overflow-hidden text-v2-black text-ellipsis whitespace-nowrap text-sm font-semibold leading-4 min-w-[106px] left-0 sticky bg-white sm:bg-inherit").text
        cost = el.find("td", class_=lambda value: value and value.startswith(
            "datatable_cell__LJp3C datatable_cell--align-end__qgxDQ")).text
        cost = str(cost).replace(',', '.')
        cost_table = pd.concat([cost_table, pd.Series({"Date": date, "Brent_oil_cost": cost}).to_frame().T],
                               ignore_index=True)
    cost_table['Brent_oil_cost'] = cost_table['Brent_oil_cost'].astype(float)
    value = cost_table[cost_table.Date == d].Brent_oil_cost
    if len(value) == 0:
        value = pd.Series([message])
    value = np.array(value)[0]
    cell_vigruz = str(letter) + str(cell[1:])
    print(cell_vigruz)
    sheet[cell_vigruz] = value
    wb.save("Приложение 1.xlsx")
    wb.close()


cells = generate_strings('N', 4, 41)
for cell in cells:
    fill_xlsx(cell, 'B')
wb = openpyxl.load_workbook('Приложение 1.xlsx')
sheet = wb['Анализ_БК+ББ']
first_not_null = 4
for i in range(first_not_null, 29):
    if sheet[f'B{i}'].value != 0:
        first_not_null = i
        break
if sheet['B4'].value == 0:
    sheet['B4'].value = sheet[f'B{first_not_null}'].value
for i in range(4, 29):
    if sheet[f'B{i + 1}'].value == 0:
        sheet[f'B{i + 1}'].value = sheet[f'B{i}'].value
wb.save("Приложение 1.xlsx")
wb.close()


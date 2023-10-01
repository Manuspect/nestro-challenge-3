from robocorp.tasks import task
from RPA.Excel.Files import Files as Excel


def to_application_1(app1_file_path: str, app2_file_path: str):
    excel = Excel()
    excel.open_workbook(app1_file_path, data_only=True)
    excel.set_active_worksheet('Анализ_БК+ББ')
    oil_brand = excel.get_cell_value(1, "AB")
    rows = excel.read_worksheet_as_table()
    months_col = rows.get_column('L')
    months = []
    month_dict = {1: 'январь', 2: 'февраль', 3: 'март', 4: 'апрель', 5: 'май', 6: 'июнь',
                  7: 'июль', 8: 'август', 9: 'сентябрь', 10: 'октябрь', 11: 'ноябрь', 12: 'декабрь'}
    for month in list(months_col.values())[3:]:
        if month is None:
            break
        months.append(month)
    months = [month_dict[month.month] for month in months]
    excel.open_workbook(app2_file_path, data_only=True)
    excel.set_active_worksheet('2023')
    rows = excel.read_worksheet_as_table()
    key_column = None

    row = rows.get_row(2)
    for key, item in row.items():
        if item == oil_brand:
            key_column = key
            break
    if key_column is None:
        raise "нет котировки нужного сорта"
    values = rows.get_column(key_column)
    oil_price = []
    for price in list(values.values())[4:]:
        if price is None:
            break
        oil_price.append(price)
    excel.open_workbook(app1_file_path, data_only=True)
    excel.set_active_worksheet('Анализ_БК+ББ')
    oil_price = {month_dict[index+1]: value for index,
                 value in enumerate(oil_price)}
    result_oil_price = [oil_price[month] for month in months]
    # print(result_oil_price)
    for index, elem in enumerate(result_oil_price):
        excel.set_cell_value(4+index, 'AB', elem)
    excel.save_workbook(app1_file_path)
    excel.close_workbook()

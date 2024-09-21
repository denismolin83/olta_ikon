import openpyxl as op
import os


def xlsx_filename(folder_path):
    for file_name in os.listdir(folder_path):
        if file_name.endswith('.xlsx'):  # если файл - xlsx
            file_path = os.path.join(folder_path, file_name)  # получаем полный путь к файлу
            return file_path  # возвращаем путь к файлу
    else:
        return None  # если xlsx файлы не найдены


file_olta_in = xlsx_filename("olta_in")
file_sale_by_period = xlsx_filename("sale_by_period")

olta_in_wb = op.load_workbook(file_olta_in, data_only=True)
olta_in_sheet = olta_in_wb.active

sale_by_period = op.load_workbook(file_sale_by_period, data_only=True)
sale_by_period_sheet = sale_by_period.active

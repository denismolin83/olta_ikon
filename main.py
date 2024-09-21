import openpyxl as op

from utils.file_open import xlsx_filename

file_olta_in = xlsx_filename("olta_in")
file_sale_by_period = xlsx_filename("sale_by_period")

olta_in_wb = op.load_workbook(file_olta_in, data_only=True)
olta_in_sheet = olta_in_wb.active

sale_by_period = op.load_workbook(file_sale_by_period, data_only=True)
sale_by_period_sheet = sale_by_period.active

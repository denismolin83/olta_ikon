import openpyxl as op

from utils.file_open import xlsx_filename

file_olta_in = xlsx_filename("olta_in")
file_result_xls = xlsx_filename("result")
# file_sale_by_period = xlsx_filename("sale_by_period")

olta_in_wb = op.load_workbook(file_olta_in, data_only=True)
olta_in_sheet = olta_in_wb.active

result_xls_wb = op.load_workbook(file_result_xls, data_only=True)
result_xls_wb_sheet = result_xls_wb.active

# sale_by_period = op.load_workbook(file_sale_by_period, data_only=True)
# sale_by_period_sheet = sale_by_period.active

list_tyres_dict = {}

for i in range(3, olta_in_sheet.max_row):
    tyre_count = olta_in_sheet.cell(row=i, column=3).value
    if tyre_count is not None:
        tyre_name = str(olta_in_sheet.cell(row=i, column=2).value).lower().replace('автошина', '').strip()
        # print(f"{tyre_name} - {tyre_count}")
        if tyre_name in list_tyres_dict:
            list_tyres_dict[tyre_name] += int(tyre_count)
        else:
            list_tyres_dict[tyre_name] = int(tyre_count)

print(list_tyres_dict)

result_xls_wb_sheet.cell(row=1, column=1).value = 'Наименование'
result_xls_wb_sheet.cell(row=1, column=2).value = 'Кол-во получено'
i = 2
for item, count in list_tyres_dict.items():
    result_xls_wb_sheet.cell(row=i, column=1).value = item
    result_xls_wb_sheet.cell(row=i, column=2).value = count
    i += 1

result_xls_wb_sheet.cell(row=i, column=2).value = sum(list_tyres_dict.values())
result_xls_wb.save(file_result_xls)



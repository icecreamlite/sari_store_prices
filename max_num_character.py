#!/home/b/projects/sari_store_prices/venv/bin/python3

#get length of the item with the most number of characters
#for calibration purposes of ssp.py

from openpyxl import load_workbook

excel_file = load_workbook('/home/b/projects/sari_store_prices/ssp.xlsx')
sheet1 = excel_file.get_sheet_by_name('Sheet1')
num_items = len(list(sheet1.rows)) #get number of rows

max = 0
for i in range(1, num_items):
    ch_num = 0
    for ch in sheet1['A' + str(i+1)].value:
        ch_num += 1
    if ch_num > max:
        max = ch_num
print(max, row)
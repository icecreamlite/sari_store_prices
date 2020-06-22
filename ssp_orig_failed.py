#!/home/b/projects/sari_store_prices/venv/bin/python3

#link for openpyxl tutorial
#https://code.tutsplus.com/tutorials/how-to-work-with-excel-documents-using-python--cms-25698

from openpyxl import load_workbook
from tkinter import *
import sys, re

excel_file = load_workbook('/home/b/projects/sari_store_prices/ssp.xlsx')
sheet1 = excel_file.get_sheet_by_name('Sheet1')
num_items = len(list(sheet1.rows))

firstAttempt = True

def printPrice(event):
    if ent.get() != '':
        regex = re.compile(ent.get(), re.I)
        
        name_text = ''
        price_text = ''
        height = 1

        for i in range(1, num_items):

            if regex.search(sheet1['A' + str(i+1)].value) != None:
                name_text += sheet1['A' + str(i+1)].value + '\n'
                price_text += str(sheet1['B' + str(i+1)].value) + '\n'
                height += 1

                #Label(name_frame, text=sheet1['A' + str(i+1)].value).grid(row=row_counter)
                #Label(price_frame, text=str(sheet1['B' + str(i+1)].value)).grid(row=row_counter)
                #row_counter += 1
                #makeItemLabel(sheet1['A' + str(i+1)].value, str(sheet1['B' + str(i+1)].value, i)
                #print(sheet1['A' + str(i+1)].value + '\t' + str(sheet1['B' + str(i+1)].value))

        global firstAttempt
        if not firstAttempt:
            nT.grid_forget()
            pT.grid_forget()
        
        else:
            firstAttempt = False
            nT = Text(name_frame, height=height, width=45)
            pT = Text(price_frame, height=height, width=5)

            nT.grid(row=0)
            pT.grid(row=0, sticky=E)

            nT.insert(END, name_text)
            pT.insert(END, price_text)


root = Tk()
root.title('Item finder')

ent = Entry(root, bd=3, bg='white', fg='black')
ent.grid(row=0, columnspan=2)
ent.focus_set()

name_frame = Frame(root)
price_frame = Frame(root)


name_frame.grid(row=1)
price_frame.grid(row=1, column=1)

root.bind('<KeyRelease>', printPrice)

root.mainloop()

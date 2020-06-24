#!/home/b/projects/sari_store_prices/venv/bin/python3

#link for openpyxl tutorial
#https://code.tutsplus.com/tutorials/how-to-work-with-excel-documents-using-python--cms-25698

from openpyxl import load_workbook
from tkinter import *
import re


sspDir = '/home/b/projects/sari_store_prices/'

excel_file = load_workbook(sspDir + 'ssp.xlsx')
sheet1 = excel_file.get_sheet_by_name('Sheet1')
num_items = len(list(sheet1.rows)) #get number of rows

col_title = '-' * 59 + '\n' + '||' + 'ITEM'.center(46) + '||' + 'PRICE'.center(7) + '||'\
                + '\n' + '-' * 59 + '\n'

def printPrice(event):
    list_text = ''

    if ent.get() != '':
        regex = re.compile(ent.get(), re.I) #set regex to what's typewritten in ent
        for i in range(1, num_items):
            if regex.search(sheet1['A' + str(i+1)].value) != None: #append regex matches to list_text
                list_text += '||' + sheet1['A' + str(i+1)].value.center(46) + '||' + str(sheet1['B' + str(i+1)].value).center(7) + '||\n'

    #update tkinter price_list text box
    global col_title
    price_list.config(state='normal')
    price_list.delete(1.0, 'end')
    price_list.insert(1.0, col_title + list_text)
    price_list.config(state='disabled')

def clearEnt(event):
    ent.delete(0, 'end')


root = Tk()
root.title('Item finder')
root.resizable(0,0) #remove maximize button

# ws = root.winfo_screenwidth() #get device screeen width
hs = root.winfo_screenheight() #get device screen height

root.geometry(f'487x{hs-63}+0+0') #initialize window position

ent = Entry(root, bd=3, bg='white', fg='black', width=53)
ent.grid(row=0)
ent.focus_set()

price_list = Text(root, width=59, height=57)
price_list.grid(row=2, pady=10)

price_list.insert(1.0, col_title)
price_list.config(state='disabled')

root.bind('<KeyRelease>', printPrice) #call printPrice everytime a key is pressed and released
root.bind('<Escape>', clearEnt)

root.mainloop()
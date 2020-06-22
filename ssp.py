#!/home/b/projects/sari_store_prices/venv/bin/python3

#link for openpyxl tutorial
#https://code.tutsplus.com/tutorials/how-to-work-with-excel-documents-using-python--cms-25698

from openpyxl import load_workbook
from tkinter import *
from os import system
import sys, re

excel_file = load_workbook('/home/b/projects/sari_store_prices/ssp.xlsx')
sheet1 = excel_file.get_sheet_by_name('Sheet1')
num_items = len(list(sheet1.rows)) #get number of rows
clear = lambda: system('clear') #function that clears console everytime Entry ent updates

def printPrice(event):
    clear() #clear console
    if ent.get() != '':
        regex = re.compile(ent.get(), re.I) #set regex to what's typewritten in ent
        
        name_text = ''
        price_text = ''
        height = 1
        print('-' * 59)
        print('||' + 'ITEM'.center(46) + '||' + 'PRICE'.center(7) + '||')
        print('-' * 59)
        for i in range(1, num_items):

            if regex.search(sheet1['A' + str(i+1)].value) != None: #print regex match to console
                print('||' + sheet1['A' + str(i+1)].value.center(46) + '||' + str(sheet1['B' + str(i+1)].value).center(7) + '||')

root = Tk()
root.title('Item finder')
root.resizable(0,0) #remove maximize button

#ws = root.winfo_screenwidth() #get device screeen width
#hs = root.winfo_screenheight() #get device screen height

root.geometry(f'+{0}+{0}') #initialize window position

ent = Entry(root, bd=3, bg='white', fg='black', width=47)
ent.pack()
ent.focus_set() #focus the entry to type

root.bind('<KeyRelease>', printPrice) #call printPrice everytime a key is pressed and released

root.mainloop()
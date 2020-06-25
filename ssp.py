#!/home/b/projects/sari_store_prices/venv/bin/python3

#link for openpyxl tutorial
#https://code.tutsplus.com/tutorials/how-to-work-with-excel-documents-using-python--cms-25698

from openpyxl import load_workbook
from tkinter import *
import re


sspDir = '/home/b/projects/sari_store_prices/'
list_excel_rows = [] #list for excel row number of displayed items
selected_line = 0 #initialize line number of selected item

excel_file = load_workbook(sspDir + 'ssp.xlsx')
sheet1 = excel_file.get_sheet_by_name('Sheet1')
num_items = len(list(sheet1.rows)) #get number of rows

col_title = '-' * 59 + '\n' + '||' + 'ITEM'.center(46) + '||' + 'PRICE'.center(7) + '||'\
                + '\n' + '-' * 59 + '\n'


def updateText():
    list_text = ''
    global list_excel_rows
    list_excel_rows = []
    if ent.get() != '':
        regex = re.compile(ent.get(), re.I) #set regex to what's typewritten in ent
        for i in range(1, num_items):
            if regex.search(sheet1['A' + str(i+1)].value) != None: #append regex matches to list_text
                list_excel_rows.append(str(i+1))
                list_text += '||' + sheet1['A' + str(i+1)].value.center(46) + '||' + str(sheet1['B' + str(i+1)].value).center(7) + '||\n'

    #update tkinter price_list text box
    price_list.config(state='normal')
    price_list.delete(1.0, 'end')
    price_list.insert(1.0, col_title + list_text)
    price_list.config(state='disabled')

def mapKey(event):
    current_focus = str(root.focus_get())

    if current_focus == '.!entry':

        if event.keysym == 'Escape': ent.delete(0, 'end') #clear entry
        updateText()

    elif current_focus == '.!frame.!text':

        if event.keysym == 'Escape':
            ent.focus_set()
            updateText()


def selectItem(event):
    line_num = int(float(event.widget.index(CURRENT)))
    if line_num > 3 and line_num < len(list_excel_rows) + 4:
        global selected_line
        selected_line = line_num
        event.widget.focus_set()
        updateText()
        event.widget.tag_add('selected', float(line_num), line_num + 0.59)
        event.widget.tag_config('selected', background='DodgerBlue2')


root = Tk()
root.title('Item finder')
root.resizable(0,0) #remove maximize button

# ws = root.winfo_screenwidth() #get device screeen width
hs = root.winfo_screenheight() #get device screen height

root.geometry(f'487x{hs-63}+0+0') #initialize window position

ent = Entry(root, bd=3, bg='white', fg='black', width=53)
ent.grid(row=0)
ent.focus_set()

text_frame = Frame(root)
text_frame.grid(row=1)

price_list = Text(text_frame, width=59, height=57, cursor='arrow')
price_list.grid(row=0, pady=10)

price_list.insert(1.0, col_title)
price_list.config(state='disabled')

root.bind('<KeyRelease>', mapKey)
price_list.bind('<1>', selectItem)

root.mainloop()
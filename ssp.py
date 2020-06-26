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


def updateExcel(itemEntry, priceEntry, editWindow):
    updateBool = True
    iE_bg = 'white'
    pE_bg = 'white'

    item = itemEntry.get()
    price = priceEntry.get()

    #if same data in entry, exit and do nothing
    if sheet1['A' + list_excel_rows[selected_line - 4]].value == item and\
        str(sheet1['B' + list_excel_rows[selected_line - 4]].value) == str(price):
        editWindowWithdraw()

    #else check validity and update excel file
    else:
        #checks if new price is valid
        try:
            price = int(price)
            assert price > 0
        except:
            try:
                price = float(price)
                assert price > 0
            except:
                updateBool = False
                pE_bg = 'red'

        #checks if item entry is empty (invalid input)
        if item == '':
            updateBool = False
            iE_bg = 'red'

        itemEntry.config(bg=iE_bg)
        priceEntry.config(bg=pE_bg)

        #if both entry fields are valid, update the file
        #save changes to tracker text file
        #close editWindow
        if updateBool:
            with open(sspDir + 'change_tracker.txt', 'a') as ct:
                ct.write(sheet1['A' + list_excel_rows[selected_line - 4]].value + '\t' +  str(sheet1['B' + list_excel_rows[selected_line - 4]].value)\
                    + '  =>  ' + item + '\t' + str(price) + '\n')
            sheet1['A' + list_excel_rows[selected_line - 4]].value = item
            sheet1['B' + list_excel_rows[selected_line - 4]].value = price
            excel_file.save(sspDir + 'ssp.xlsx')
            editWindowWithdraw()
            updateText()


def editWindowWithdraw():
    editWindow.withdraw()
    ent.config(state='normal')


def mapKey(event):
    if editWindow.state() == 'withdrawn':

        current_focus = str(root.focus_get())

        if current_focus == '.!entry':

            if event.keysym == 'Escape': ent.delete(0, 'end') #clear entry
            updateText()

        elif current_focus == '.!frame.!text':

            if event.keysym == 'Escape':
                ent.focus_set()
                updateText()

            if event.keysym == 'Return':

                ent.config(state='disabled')

                editWindow.deiconify()

                itemEntry.delete(0, 'end')
                priceEntry.delete(0, 'end')
                itemEntry.insert(END, sheet1['A' + list_excel_rows[selected_line - 4]].value)
                priceEntry.insert(END, str(sheet1['B' + list_excel_rows[selected_line - 4]].value))

                editWindow.protocol('WM_DELETE_WINDOW', editWindowWithdraw)


def selectItem(event):
    if editWindow.state() == 'withdrawn':
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

##create withdrawn editWindow
editWindow = Toplevel(root)
editWindow.title('Edit Item')
editWindow.resizable(0,0) #remove maximize button

ws = root.winfo_screenwidth() #get device screeen width
editWindow.geometry(f'400x160+{ws//2-200}+{hs//2-200}') #initialize window position

#create labelframes in editWindow
itemLabelFrame = LabelFrame(editWindow, text='Item')
itemLabelFrame.grid(row=0)
priceLabelFrame = LabelFrame(editWindow, text='Price')
priceLabelFrame.grid(row=1, pady=5)

#create submitButton in editWindow
submitButton = Button(editWindow, text='Submit', bg='dim gray', activebackground='dim gray',\
    command=lambda: updateExcel(itemEntry, priceEntry, editWindow))
submitButton.grid(row=2)

#center widgets in editWindow
editWindow.columnconfigure(0, weight=1)
editWindow.rowconfigure(0, weight=1)
editWindow.rowconfigure(1, weight=1)
editWindow.rowconfigure(2, weight=1)

#create entries in editWindow
itemEntry = Entry(itemLabelFrame, bd=3, bg='white', fg='black', width=37, justify=CENTER)
itemEntry.grid(row=0, padx=5, pady=5)
priceEntry = Entry(priceLabelFrame, bd=3, bg='white', fg='black', width=37, justify=CENTER)
priceEntry.grid(row=1, padx=5, pady=5)

editWindow.withdraw()
##end of editWindow creation


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
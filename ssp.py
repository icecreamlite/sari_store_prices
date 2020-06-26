#!/home/b/projects/sari_store_prices/venv/bin/python3

#link for openpyxl tutorial
#https://code.tutsplus.com/tutorials/how-to-work-with-excel-documents-using-python--cms-25698

from openpyxl import load_workbook
from tkinter import *
from tkinter import ttk
import re
from sys import path

sspDir = '/home/b/projects/sari_store_prices/'

path.append(sspDir + 'modules')

from tagloc_generate import generate

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
    price_list['state'] = 'normal'
    price_list.delete(1.0, 'end')
    price_list.insert(1.0, col_title + list_text)
    price_list['state'] = 'disabled'


def updateExcel(itemEntry, priceEntry, editWindow):
    #setting default
    updateBool = True
    iE_bg = 'white'
    pE_bg = 'white'
    Cbtheme = 'white'

    item = itemEntry.get()
    price = priceEntry.get()
    tag = tagCb.get()
    location = locationCb.get()

    #if same data in entry, exit and do nothing
    selected_item = str(sheet1['A' + list_excel_rows[selected_line - 4]].value)
    selected_price = str(sheet1['B' + list_excel_rows[selected_line - 4]].value)
    selected_tag = str(sheet1['C' + list_excel_rows[selected_line - 4]].value)
    selected_location = str(sheet1['D' + list_excel_rows[selected_line - 4]].value)
    if selected_item == item and selected_price == str(price)\
        and selected_tag == tag and selected_location == location: editWindowWithdraw()

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
        if tag == '' or location == '':
            updateBool = False
            Cbtheme = 'red'

        #update options based on error occurence in input fields
        itemEntry['bg'] = iE_bg
        priceEntry['bg'] = pE_bg
        style.theme_use(Cbtheme)

        #if both entry fields are valid, update the file
        #save changes to tracker text file
        #close editWindow
        if updateBool:

            sheet1['A' + list_excel_rows[selected_line - 4]].value = item
            sheet1['B' + list_excel_rows[selected_line - 4]].value = price
            sheet1['C' + list_excel_rows[selected_line - 4]].value = tag
            sheet1['D' + list_excel_rows[selected_line - 4]].value = location
            excel_file.save(sspDir + 'ssp.xlsx')
            editWindowWithdraw()
            updateText()

            #generate tag and location text files only if they're updated
            if selected_tag != tag or selected_location != location:
                generate(sspDir)

            #track changes
            with open(sspDir + 'text_files/change_tracker.txt', 'a') as ct:
                ct.write(selected_item + '\t' +  selected_price + '\t' + selected_tag + '\t' + selected_location\
                    + '  =>  ' + item + '\t' + str(price) + '\t' + tag + '\t' + location + '\n')


def editWindowWithdraw():
    editWindow.withdraw()
    ent['state'] = 'normal'
    itemEntry['bg'] = 'white'
    priceEntry['bg'] = 'white'
    style.theme_use('white')


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

                ent['state'] = 'disabled'

                #fetch list of tags
                tags = []
                for tag in open(sspDir + 'text_files/tags.txt', 'r'):
                    tag = tag[:-1]
                    tags.append(tag)
                    if str(sheet1['C' + list_excel_rows[selected_line - 4]].value) == tag:
                        tagCurInd = len(tags) - 1

                #fetch list of locations
                locations = []
                for location in open(sspDir + 'text_files/locations.txt', 'r'):
                    location = location[:-1]
                    locations.append(location)
                    if str(sheet1['D' + list_excel_rows[selected_line - 4]].value) == location:
                        locCurInd = len(locations) - 1

                ##Set editWindow Fields
                tagCb['value'] = tags
                tagCb.current(tagCurInd)
                locationCb['value'] = locations
                locationCb.current(locCurInd)

                itemEntry.delete(0, 'end')
                priceEntry.delete(0, 'end')
                itemEntry.insert(END, sheet1['A' + list_excel_rows[selected_line - 4]].value)
                priceEntry.insert(END, str(sheet1['B' + list_excel_rows[selected_line - 4]].value))
                ##

                editWindow.deiconify()

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
editWindow.geometry(f'400x300+{ws//2-200}+{hs//2-200}') #initialize window position

#create labelframes in editWindow
itemLF = LabelFrame(editWindow, text='Item')
itemLF.grid(row=0)
priceLF = LabelFrame(editWindow, text='Price')
priceLF.grid(row=1)
tagLF = LabelFrame(editWindow, text='Tag')
tagLF.grid(row=2)
locationLF = LabelFrame(editWindow, text='Location')
locationLF.grid(row=3, pady=1)

#create submitButton in editWindow
submitButton = Button(editWindow, text='Submit', bg='dim gray', activebackground='dim gray',\
    command=lambda: updateExcel(itemEntry, priceEntry, editWindow))
submitButton.grid(row=4)

#center widgets in editWindow
editWindow.columnconfigure(0, weight=1)
editWindow.rowconfigure(0, weight=1)
editWindow.rowconfigure(1, weight=1)
editWindow.rowconfigure(2, weight=1)
editWindow.rowconfigure(3, weight=1)
editWindow.rowconfigure(4, weight=1)

#create entries in editWindow
itemEntry = Entry(itemLF, bd=3, bg='white', fg='black', width=37, justify=CENTER)
itemEntry.grid(row=0, padx=5, pady=5)
priceEntry = Entry(priceLF, bd=3, bg='white', fg='black', width=37, justify=CENTER)
priceEntry.grid(row=0, padx=5, pady=5)

#create comboboxes in editWindow
style = ttk.Style()
style.theme_create('red', parent='alt', 
    settings = {'TCombobox': 
                {'configure': 
                    {'fieldbackground': 'red'
                    }}}
)
style.theme_create('white', parent='alt', 
    settings = {'TCombobox': 
                {'configure': 
                    {'fieldbackground': 'white'
                    }}}
)

tagCb = ttk.Combobox(tagLF, foreground='black', width=37, justify=CENTER)
tagCb.grid(row=0, padx=5, pady=5)
locationCb = ttk.Combobox(locationLF, width=37, justify=CENTER)
locationCb.grid(row=0, padx=5, pady=5)

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
price_list['state'] = 'disabled'

root.bind('<KeyRelease>', mapKey)
price_list.bind('<1>', selectItem)

root.mainloop()
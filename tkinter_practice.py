#!/home/b/projects/sari_store_prices/venv/bin/python3

from tkinter import *

root = Tk()
root.title('Item finder')
root.resizable(0,0)

root_label = Label(root, text='Search: ')
root_label.pack(side=LEFT)

ent = Entry(root, bd=3)
ent.pack(side=RIGHT)
ent.focus_set()

root.bind('<Return>', lambda x: print(ent.get()))

root.mainloop()
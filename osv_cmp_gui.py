
import tkinter as tk
import tkinter.ttk as ttk
import tkinter.messagebox as messagebox

from tkinter import filedialog
from osv_cmp import load_osv_smeta, load_osv_1c


class App(tk.Tk):
    def bt_pick_file1(self, event):
        messagebox.showinfo('Title', 'Message')

    def bt_pick_file2(self, event):
        pass

    def __init__(self):
        super().__init__()
        
        button = ttk.Button(self, text='Выбрать файл 1')
        button.grid(column=1, row=1)
        button.bind('<1>', self.bt_pick_file1)

        self.entry1 = ttk.Entry(self, width=100)
        self.entry1.grid(column=2, row=1, sticky=tk.EW)


        ttk.Button(self, text='Загрузить/проверить').grid(column=3, row=1)
        
        ttk.Button(self, text='Выбрать файл 2').grid(column=1, row=2)
        
        self.entry2 = ttk.Entry(self, width=100)
        self.entry2.grid(column=2, row=2, sticky=tk.EW)


        ttk.Button(self, text='Загрузить/проверить').grid(column=3, row=2)
        ttk.Button(self, text='Сравнить').grid(column=4, row=1, rowspan=2, sticky=tk.NS)

        tk.Text().grid(column=1, row=4, columnspan=5, sticky=tk.EW)
        ttk.Button(self, text='Очистить отчет').grid(column=1, row=5)
        ttk.Button(self, text='Сохранить отчет').grid(column=2, row=5)

app = App()
app.mainloop()


import tkinter as tk
import tkinter.ttk as ttk
import tkinter.messagebox as messagebox

from tkinter import filedialog
from osv_cmp import load_osv_smeta, load_osv_1c


class App(tk.Tk):
    def bt_pick_file1(self, event):
        self.filename1 = filedialog.askopenfilename(filetypes=[('Документ Excel', '*.xls')])
        self.entry1.delete(0, tk.END)
        self.entry1.insert(0, self.filename1)

    def bt_pick_file2(self, event):
        self.filename2 = filedialog.askopenfilename(filetypes=[('Документ Excel', '*.xls')])
        self.entry2.delete(0, tk.END)
        self.entry2.insert(0, self.filename2)

    def bt_clear_entry1(self, event):
        self.entry1.delete(0, tk.END)

    def bt_clear_entry2(self, event):
        self.entry2.delete(0, tk.END)

    def __init__(self):
        super().__init__()
        
        button = ttk.Button(self, text='Выбрать файл 1')
        button.grid(column=1, row=1)
        button.bind('<1>', self.bt_pick_file1)

        self.entry1 = ttk.Entry(self, width=100)
        self.entry1.grid(column=2, row=1, sticky=tk.EW)

        button = ttk.Button(self, text='X')
        button.grid(column=3, row=1)
        button.bind('<1>', self.bt_clear_entry1)

        button = ttk.Button(self, text='Выбрать файл 2')
        button.grid(column=1, row=2)
        button.bind('<1>', self.bt_pick_file2)
        
        self.entry2 = ttk.Entry(self, width=100)
        self.entry2.grid(column=2, row=2, sticky=tk.EW)

        button = ttk.Button(self, text='X')
        button.grid(column=3, row=2)
        button.bind('<1>', self.bt_clear_entry2)
        
        ttk.Button(self, text='Загрузить/\nперечитать').grid(column=4, row=1, rowspan=2, sticky=tk.NS)

        ttk.Button(self, text='Сравнить').grid(column=5, row=1, rowspan=2, sticky=tk.NS)

        tk.Text().grid(column=1, row=4, columnspan=5, sticky=tk.EW)
        
        ttk.Button(self, text='Сохранить отчет').grid(column=1, row=5)

app = App()
app.mainloop()

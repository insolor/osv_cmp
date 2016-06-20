
import tkinter as tk
import tkinter.ttk as ttk
import tkinter.messagebox as messagebox
import xlrd

from tkinter import filedialog
from osv_cmp import load_osv_smeta, load_osv_1c, check_format


class Report(tk.Text):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

    def clear(self):
        self.delete(0.0, tk.END)

    def print(self, text='', end='\n'):
        self.insert(tk.END, text + end)


class App(tk.Tk):
    def load_file(self, filename):
        self.report.print("Загрузка файла %r" % filename)

        wb = xlrd.open_workbook(filename, formatting_info=True)
        sheet = wb.sheet_by_index(0)

        fmt = check_format(sheet)
        self.report.print("Формат: %s" % fmt)

        if fmt == '1c':
            return load_osv_1c(sheet)
        elif fmt == 'Smeta':
            return load_osv_smeta(sheet)
        else:
            self.report.print("Формат документа не опознан. Загрузка прекращена.")
            return None

    def bt_pick_file1(self, event):
        self.filename1 = filedialog.askopenfilename(filetypes=[('Документ Excel', '*.xls')])
        self.entry1.delete(0, tk.END)
        self.entry1.insert(0, self.filename1)

        self.osv1 = self.load_file(self.filename1)

    def bt_pick_file2(self, event):
        self.filename2 = filedialog.askopenfilename(filetypes=[('Документ Excel', '*.xls')])
        self.entry2.delete(0, tk.END)
        self.entry2.insert(0, self.filename2)

        self.osv2 = self.load_file(self.filename2)

    def bt_clear_entry1(self, event):
        self.entry1.delete(0, tk.END)
        self.osv1 = None

    def bt_clear_entry2(self, event):
        self.entry2.delete(0, tk.END)
        self.osv2 = None

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

        self.report = Report()
        self.report.grid(column=1, row=4, columnspan=5, sticky=tk.EW)
        
        ttk.Button(self, text='Сохранить отчет').grid(column=1, row=5)

        self.osv1 = None
        self.osv2 = None

app = App()
app.mainloop()


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
        if not filename:
            return

        self.report.print("Загрузка файла %r" % filename)

        wb = xlrd.open_workbook(filename, formatting_info=True)
        sheet = wb.sheet_by_index(0)

        fmt = check_format(sheet)
        self.report.print("Формат: %s" % fmt)

        if fmt == '1c':
            osv = load_osv_1c(sheet)
            self.report.print("Файл загружен.")
            return osv
        elif fmt == 'Smeta':
            osv = load_osv_smeta(sheet)
            self.report.print("Файл загружен.")
            return osv
        else:
            self.report.print("Формат документа не опознан. Загрузка прекращена.")
            return None

    def bt_pick_file(self, i, event):
        self.filename[i] = filedialog.askopenfilename(filetypes=[('Документ Excel', '*.xls')])
        if not self.filename[i]:
            return

        self.entry[i].delete(0, tk.END)
        self.entry[i].insert(0, self.filename[i])

        self.osv[i] = self.load_file(self.filename[i])

    def bt_clear_entry(self, i, event):
        self.entry[i].delete(0, tk.END)
        self.osv[i] = None

    def bt_reread(self, event):
        self.report.clear()
        for i in range(2):
            self.filename[i] = self.entry[i].get()
            self.osv[i] = self.load_file(self.filename[i])

    def __init__(self):
        super().__init__()
        
        button = ttk.Button(self, text='Выбрать файл 1')
        button.grid(column=1, row=1)
        button.bind('<1>', lambda event: self.bt_pick_file(0, event))

        self.entry = [None, None]

        self.entry[0] = ttk.Entry(self, width=100)
        self.entry[0].grid(column=2, row=1, sticky=tk.EW)

        button = ttk.Button(self, text='X')
        button.grid(column=3, row=1)
        button.bind('<1>', lambda event: self.bt_clear_entry(0, event))

        button = ttk.Button(self, text='Выбрать файл 2')
        button.grid(column=1, row=2)
        button.bind('<1>', lambda event: self.bt_pick_file(1, event))

        self.entry[1] = ttk.Entry(self, width=100)
        self.entry[1].grid(column=2, row=2, sticky=tk.EW)

        button = ttk.Button(self, text='X')
        button.grid(column=3, row=2)
        button.bind('<1>', lambda event: self.bt_clear_entry(1, event))
        
        button = ttk.Button(self, text='Загрузить/\nперечитать')
        button.grid(column=4, row=1, rowspan=2, sticky=tk.NS)
        button.bind('<1>', self.bt_reread)

        ttk.Button(self, text='Сравнить').grid(column=5, row=1, rowspan=2, sticky=tk.NS)

        self.report = Report()
        self.report.grid(column=1, row=4, columnspan=5, sticky=tk.EW)
        
        ttk.Button(self, text='Сохранить отчет').grid(column=1, row=5)

        self.osv = [None, None]
        self.filename = [None, None]

app = App()
app.mainloop()


import tkinter as tk
import tkinter.ttk as ttk
import tkinter.messagebox as messagebox
import xlrd
import json
import pprint

from tkinter import filedialog
from osv_cmp import load_osv_smeta, load_osv_1c, check_format
from collections import OrderedDict


class Report(tk.Text):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

    def clear(self):
        self.delete(0.0, tk.END)

    def print(self, *objects, sep=' ', end='\n'):
        self.insert(tk.END, sep.join(str(item) for item in objects) + end)


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
            osv, log = load_osv_1c(sheet)
        elif fmt == 'Smeta':
            osv, log = load_osv_smeta(sheet)
        else:
            self.report.print("Формат документа не опознан. Загрузка прекращена.")
            return None
        
        self.report.print(*log, sep='\n')
        
        self.report.print("Файл загружен.")
        self.report.print("Загружено счетов: %d" % len(osv))
        self.report.print("Загружено подчиненных записей: %d" % sum(len(records) for records in osv.values()))
        
        self.report.print()
        return osv

    def bt_pick_file(self, i, event):
        self.filename[i] = filedialog.askopenfilename(filetypes=[('Документ Excel', '*.xls')])
        if not self.filename[i]:
            return

        self.entry[i].delete(0, tk.END)
        self.entry[i].insert(0, self.filename[i])

        self.osv[i] = self.load_file(self.filename[i])

    def bt_clear_entry(self, i, event):
        self.entry[i].delete(0, tk.END)
        self.filename[i] = ''
        self.osv[i] = None

    def bt_reread(self, event):
        self.report.clear()
        for i in range(2):
            self.filename[i] = self.entry[i].get()
            self.osv[i] = self.load_file(self.filename[i])
    
    def bt_compare(self, event):
        if not (self.osv[0] and self.osv[1]):
            return
        
        accs = [set(item.keys()) for item in self.osv]
        if accs[0] == accs[1]:
            self.report.print('Различий в наборе загруженных счетов нет.')
        else:
            self.report.print('Различия в наборе счетов:')
            # Что пропало (из того что было вычесть то что осталось)
            for item in sorted(accs[0] - accs[1], key=lambda x: x.split('.')):
                self.report.print('-', item)
            
            # Что появилось (из того что стало вычесть то что было)
            for item in sorted(accs[1] - accs[0], key=lambda x: x.split('.')):
                self.report.print('+', item)
        
        self.report.print('\nСравнение набора подчиненных записей для каждого счета из исходного документа:')
        diffs = OrderedDict()
        osv = self.osv
        for acc in osv[0]:
            if acc in osv[1]:
                records = [set(osv[i][acc].keys()) for i in range(2)]
                if records[0] == records[1]:
                    continue
                diffs[acc] = (sorted(records[0] - records[1]), sorted(records[1] - records[0]))
        
        if not diffs:
            self.report.print('Различий нет.')
        else:
            for acc, (absent, new) in diffs.items():
                self.report.print('%s:' % acc)
                
                for item in absent:
                    self.report.print(' - %r' % item)
                
                for item in new:
                    self.report.print(' + %r' % item)
        
        self.report.print('\nСравнение сумм:')
        diffs = OrderedDict()
        for acc in osv[0]:
            if acc in osv[1]:
                for record, row in osv[0][acc].items():
                    if record in osv[1][acc]:
                        row2 = osv[1][acc][record]
                        if row[:4] == row2[:4]:
                            continue
                        elif (row[0]-row[1], row[2]-row[3]) == (row2[0]-row2[1], row2[2]-row2[3]):
                            continue
                        else:
                            if acc not in diffs:
                                diffs[acc] = OrderedDict()
                            
                            diffs[acc][record] = (row[:4], row2[:4])
        
        if not diffs:
            self.report.print('Недопустимых различий нет.')
        else:
            for acc, records in diffs.items():
                self.report.print('%s:' % acc)
                
                for record, diff in records.items():
                    self.report.print(' %r' % record)
                    self.report.print('  --- | %15.2f | %15.2f | %15.2f | %15.2f | ...' % tuple(diff[0]))
                    self.report.print('  +++ | %15.2f | %15.2f | %15.2f | %15.2f | ...' % tuple(diff[1]))
                    self.report.print()
    
    def __init__(self):
        super().__init__()

        self.entry = [None, None]

        button = ttk.Button(self, text='Выбрать файл 1')
        button.grid(column=1, row=1)
        button.bind('<1>', lambda event: self.bt_pick_file(0, event))

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

        button = ttk.Button(self, text='Сравнить')
        button.grid(column=5, row=1, rowspan=2, sticky=tk.NS)
        button.bind('<1>', self.bt_compare)

        self.report = Report()
        self.report.grid(column=1, row=4, columnspan=5, sticky=tk.EW)
        
        ttk.Button(self, text='Сохранить отчет').grid(column=1, row=5)

        self.osv = [None, None]
        self.filename = [None, None]

app = App()
app.mainloop()

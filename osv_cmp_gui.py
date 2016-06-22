
import tkinter as tk
import tkinter.ttk as ttk
import tkinter.messagebox as messagebox
import xlrd
import json
import pprint

from tkinter import filedialog
from osv_cmp import load_osv_smeta, load_osv_1c, check_format, osv_compare
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

    def bt_pick_file(self, event, i):
        self.filename[i] = filedialog.askopenfilename(filetypes=[('Документ Excel', '*.xls')])
        if not self.filename[i]:
            return

        self.entry[i].delete(0, tk.END)
        self.entry[i].insert(0, self.filename[i])

        self.osv[i] = self.load_file(self.filename[i])

    def bt_clear_entry(self, event, i):
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
        
        diffs = osv_compare(*self.osv)
        
        if not diffs['accs']:
            self.report.print('Различий в наборе загруженных счетов нет.')
        else:
            self.report.print('Различия в наборе счетов:')
            
            for item in diffs['accs'][0]:
                self.report.print('- %r' % item)
            
            for item in diffs['accs'][1]:
                self.report.print('+ %r' % item)
        
        self.report.print('\nСравнение набора подчиненных записей для каждого счета из исходного документа:')
        diff_records = diffs['records']
        if not diff_records :
            self.report.print('Различий нет.')
        else:
            for acc, (absent, new) in diff_records.items():
                self.report.print('%s:' % acc)
                
                for item in absent:
                    self.report.print(' - %r' % item)
                
                for item in new:
                    self.report.print(' + %r' % item)
        
        self.report.print('\nСравнение сумм:')
        diff_sums = diffs['sums']
        if not diff_sums:
            self.report.print('Недопустимых различий нет.')
        else:
            for acc, records in diff_sums.items():
                self.report.print('%s:' % acc)
                
                for record, diff in records.items():
                    self.report.print(' %r' % record)
                    self.report.print('  --- | %15.2f | %15.2f | %15.2f | %15.2f | ...' % tuple(diff[0]))
                    self.report.print('  +++ | %15.2f | %15.2f | %15.2f | %15.2f | ...' % tuple(diff[1]))
                    self.report.print()
    
    def __init__(self):
        def init_header(parent):
            self.entry = [None, None]
            
            for i in range(2):
                button = ttk.Button(parent, text='Выбрать файл %d' % (i+1))
                button.grid(column=0, row=i)
                button.bind('<1>', lambda event, i=i: self.bt_pick_file(event, i))

                self.entry[i] = ttk.Entry(parent, width=100)
                self.entry[i].grid(column=1, row=i, sticky=tk.EW)

                button = ttk.Button(parent, text='X')
                button.grid(column=2, row=i)
                button.bind('<1>', lambda event, i=i: self.bt_clear_entry(event, i))
            
            button = ttk.Button(parent, text='Загрузить/\nперечитать')
            button.grid(column=3, row=0, rowspan=2, sticky=tk.NS)
            button.bind('<1>', self.bt_reread)

            button = ttk.Button(parent, text='Сравнить')
            button.grid(column=4, row=0, rowspan=2, sticky=tk.NS)
            button.bind('<1>', self.bt_compare)
        
        def init_report_area(parent):
            scrollbar = ttk.Scrollbar(parent)
            scrollbar.pack(side='right', fill='y')
            
            self.report = Report(parent)
            self.report.pack(side='left', fill='both', expand=1)
            
            scrollbar['command'] = self.report.yview
            self.report['yscrollcommand'] = scrollbar.set
        
        def init_footer(parent):
            button = ttk.Button(parent, text='Сохранить отчет')
            button.pack()
        
        super().__init__()
        
        header = tk.Frame()
        header.pack(side='top', fill='x')
        init_header(header)
        
        footer = tk.Frame()
        footer.pack(side='bottom', fill='x')
        init_footer(footer)
        
        report_area = tk.Frame()
        report_area.pack()
        init_report_area(self)
        
        
        self.osv = [None, None]
        self.filename = [None, None]

app = App()
app.mainloop()

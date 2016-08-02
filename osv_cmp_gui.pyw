
import tkinter as tk
import tkinter.ttk as ttk
import tkinter.messagebox as messagebox
import xlrd
import json

from tkinter import filedialog
from osv_cmp import load_osv_smeta, load_osv_1c, check_format, osv_compare, osv_sum


class Report(tk.Text):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

    def clear(self):
        self.delete(0.0, tk.END)

    def print(self, *objects, sep=' ', end='\n'):
        self.insert(tk.END, sep.join(str(item) for item in objects) + end)


class App(tk.Tk):
    def load_file(self, filename, i):
        self.reports[i].clear()

        if not filename:
            return

        self.reports[i].print("Загрузка файла %r" % filename)

        wb = xlrd.open_workbook(filename, formatting_info=True)
        sheet = wb.sheet_by_index(0)
        sheet = [sheet.row_values(i) for i in range(sheet.nrows)]

        fmt = check_format(sheet)
        self.reports[i].print("Формат: %s" % fmt)

        if fmt == '1c':
            osv, log = load_osv_1c(sheet)
        elif fmt == 'Smeta':
            osv, log = load_osv_smeta(sheet)
        else:
            self.report.print("Формат документа не опознан. Загрузка прекращена.")
            return None
        
        self.reports[i].print(*log, sep='\n')
        
        self.reports[i].print("Файл загружен.")
        self.reports[i].print("Загружено счетов: %d" % len(osv))
        self.reports[i].print("Загружено подчиненных записей: %d" % sum(len(records) for records in osv.values()))

        # todo: не учитывать забалансовые счета
        s = osv_sum(osv)
        self.reports[i].print("Сумма по документу (включая забалансовые счета):")
        self.reports[i].print("[%s]" % ', '.join('%.2f' % item for item in s))

        return osv

    def bt_pick_file(self, _, i):
        self.filename[i] = filedialog.askopenfilename(filetypes=[('Документ Excel', '*.xls')])
        if not self.filename[i]:
            return

        self.notebook.select(i)
        self.entry[i].delete(0, tk.END)
        self.entry[i].insert(0, self.filename[i])

        self.osv[i] = self.load_file(self.filename[i], i)

    def bt_clear_entry(self, _, i):
        self.entry[i].delete(0, tk.END)
        self.filename[i] = ''
        self.osv[i] = None
        self.reports[i].clear()

    def bt_reread(self, _):
        for i in range(2):
            self.filename[i] = self.entry[i].get()
            self.osv[i] = self.load_file(self.filename[i], i)

        self.reports[2].clear()
        self.notebook.select(0)
    
    def bt_compare(self, _):
        if not (self.osv[0] and self.osv[1]):
            messagebox.showwarning('Нужно два документа',
                                   'Для сравнения нужно загрузить два документа')
            return

        if self.filename[0] == self.filename[1]:
            messagebox.showwarning('Один и тот же документ',
                                   'Вы пытаетесь сравнить один и тот же документ с самим собой')
            return

        self.notebook.select(2)
        self.report.clear()

        diffs = osv_compare(*self.osv)
        
        if not diffs['accs']:
            self.report.print('Различий в наборе загруженных счетов нет.')
        else:
            self.report.print('Различия в наборе счетов:')
            
            for i, sign in enumerate(('-', '+')):
                for item in diffs['accs'][i]:
                    self.report.print('%s %r' % (sign, item))
                    for subrecord, values in self.osv[i][item].items():
                        self.report.print('   %-22r [%s]' % (subrecord, ', '.join('%.2f' % n for n in values)))
        
        self.report.print('\nСравнение набора подчиненных записей для каждого счета из исходного документа:')
        diff_records = diffs['records']
        if not diff_records:
            self.report.print('Различий нет.')
        else:
            for acc, (absent, new) in diff_records.items():
                self.report.print('%s:' % acc)
                
                s = set(('-', item, tuple(values[:4])) for item, values in absent.items()) | \
                    set(('+', item, tuple(values[:4])) for item, values in new.items())
                
                for sign, item, values in sorted(s, key=lambda x: (x[2], -ord(x[0]), x[1])):
                    self.report.print(' %s %-22r [%s, ...]' % (sign, item, ', '.join('%.2f' % n for n in values)))
        
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

    def bt_save_report(self, _):
        if not any(part.get(1.0, tk.END).strip() for part in self.reports):
            messagebox.showwarning('Пустой отчет', 'Отчет пуст: не загружен ни один файл и не произведено сравнение')
            return

        filename = filedialog.asksaveasfilename(filetypes=[('Текстовый документ', '*.txt')], defaultextension='.txt')
        if filename:
            with open(filename, mode='wt') as fn:
                print(('-' * 80 + '\n').join(part.get(1.0, tk.END) for part in self.reports), file=fn)
    
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
            
            report = Report(parent)
            report.pack(fill='both', expand=1)
            
            scrollbar['command'] = report.yview
            report['yscrollcommand'] = scrollbar.set
            return report
        
        def init_footer(parent):
            button = ttk.Button(parent, text='Сохранить отчет')
            button.pack()
            button.bind('<1>', self.bt_save_report)

        def init_notebook(*tabs):
            notebook = ttk.Notebook()
            self.reports = [None, None, None]
            for i in range(3):
                f1 = tk.Frame(notebook)
                notebook.add(f1, text=tabs[i])
                notebook.pack(fill='both', expand=1)
                self.reports[i] = init_report_area(f1)
            return notebook

        super().__init__()
        
        header = tk.Frame()
        header.pack(side='top', fill='x')
        init_header(header)
        
        footer = tk.Frame()
        footer.pack(side='bottom', fill='x')
        init_footer(footer)

        self.notebook = init_notebook('Отчет по файлу 1', 'Отчет по файлу 2', 'Сравнение')

        self.report = self.reports[2]
        
        self.osv = [None, None]
        self.filename = [None, None]

app = App()
app.mainloop()

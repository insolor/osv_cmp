import json
from collections import OrderedDict

import xlrd


def load_osv_smeta(sheet: xlrd.sheet.Sheet):
    sheet_dict = OrderedDict()
    current_acc = None
    for i in range(8, sheet.nrows):
        row = sheet.row_values(i)
        key = row[0].strip()

        parts = key.count('.') + 1
        if key.startswith('Итого'):
            break
        elif parts == 1 and len(key) > 0:
            pass
        elif parts == 3:
            current_acc = key
            sheet_dict[current_acc] = OrderedDict()
        else:
            assert current_acc is not None
            key = ''.join(key.split('.'))

            if len(key) == 20:
                assert key.startswith('000'), "20-digit KBK started not from 000: %s" % key
                key = key[3:]

            assert key not in sheet_dict[current_acc], "Double KBK %s in account %s" % (key, current_acc)
            sheet_dict[current_acc][key] = row[1:]

    return sheet_dict


def load_osv_1c(sheet: xlrd.sheet.Sheet):
    sheet_dict = OrderedDict()
    current_kfo = None
    current_acc = None
    for i in range(12, sheet.nrows):
        row = sheet.row_values(i)
        key = row[0]
        row = [0.0 if not item else item for item in (row[3], row[6], row[9], row[14], row[16], row[19])]
        if isinstance(key, float):
            key = key
        else:
            key = key

        print(repr(key), row)
        if key == 'Итого':
            break
        elif isinstance(key, float):
            if current_kfo is None or key == current_kfo + 1:
                current_kfo = int(key)
        else:
            pass


wb = xlrd.open_workbook(r'c:\Users\ret\YandexDisk\ОСВ Тихвинский сс\OSV_VED_1.xls', formatting_info=True)
sheet = wb.sheet_by_index(0)

osv_smeta = json.dumps(load_osv_smeta(sheet), indent=2)

wb = xlrd.open_workbook(r'c:\Users\ret\YandexDisk\ОСВ Тихвинский сс\Тихвинский ОСВ - после свертки.xls',
                        formatting_info=True)
sheet = wb.sheet_by_index(0)

load_osv_1c(sheet)

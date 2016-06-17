from collections import OrderedDict

import xlrd


def load_osv_smeta(sheet: xlrd.sheet.Sheet):
    sheet_dict = OrderedDict()
    current_acc = None
    for i in range(8, sheet.nrows):
        row = sheet.row_values(i)
        key = row[0].strip()
        # print(repr(key), row[1:])
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
                assert key.startswith('000')
                key = key[3:]

            assert key not in sheet_dict[current_acc], "Double KBK %s in account %s" % (key, current_acc)
            sheet_dict[current_acc][key] = row[1:]

    return sheet_dict


rb = xlrd.open_workbook(r'c:\Users\ret\YandexDisk\ОСВ Тихвинский сс\OSV_VED_1.xls', formatting_info=True)
sheet = rb.sheet_by_index(0)

print(dict(load_osv_smeta(sheet)))

import json
import pprint
import xlrd

from collections import OrderedDict


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
        elif 1 < parts <= 3:
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
        
        if key == 'Итого':
            break
        elif isinstance(key, float):
            next_key = str(sheet.row_values(i+1)[0]) if i < sheet.nrows - 1 else ''
            if current_kfo is None or int(key) == current_kfo + 1 and not next_key.startswith('%02d' % key):
                current_kfo = int(key)
            else:
                key = '%02d' % key
                current_acc = key
        elif '.' in key:
            current_acc = key
        else:
            acc = '%s.%s' % (current_kfo, current_acc)
            if acc not in sheet_dict:
                sheet_dict[acc] = OrderedDict()
            
            if key in sheet_dict[acc]:
                print("Double KBK %s in account %s" % (key, acc))
                j = 1
                candidate = '%s_%d' % (key, j)
                while candidate in sheet_dict[acc]:
                    j += 1
                    candidate = '%s_%d' % (key, j)
                key = candidate
            sheet_dict[acc][key] = row
    return sheet_dict


def main():
    pp = pprint.PrettyPrinter()

    wb = xlrd.open_workbook(r'ОСВ Тихвинский сс\OSV_VED_1.xls', formatting_info=True)
    sheet = wb.sheet_by_index(0)

    osv_smeta = load_osv_smeta(sheet)
    pp.pprint([(key, len(val)) for key, val in osv_smeta.items()])

    wb = xlrd.open_workbook(r'ОСВ Тихвинский сс\Тихвинский ОСВ - после свертки.xls', formatting_info=True)
    sheet = wb.sheet_by_index(0)

    osv_1c = load_osv_1c(sheet)
    pp.pprint([(key, len(val)) for key, val in osv_1c.items()])

main()

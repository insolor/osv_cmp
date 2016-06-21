import json
import pprint
import sys
import xlrd

from collections import OrderedDict


def load_osv_1c(sheet: xlrd.sheet.Sheet):
    sheet_dict = OrderedDict()
    current_kfo = 0
    current_acc = None
    for i in range(9, sheet.nrows):
        row = sheet.row_values(i)
        key = row[0]
        row = [0.0 if not item else item for item in (row[3], row[6], row[9], row[14], row[16], row[19])]

        if key == 'Итого':
            break
        elif isinstance(key, float):
            next_key = str(sheet.row_values(i + 1)[0]) if i < sheet.nrows - 1 else ''
            # if current_kfo is None or int(key) == current_kfo + 1 and not next_key.startswith('%02d' % key):
            if current_kfo < int(key) <= 5 and not next_key.startswith('%02d' % key):
                current_kfo = int(key)
            else:
                current_acc = '%02d' % key
                acc = '%s.%s' % (current_kfo, current_acc)
                if acc in sheet_dict:
                    value = sheet_dict[acc]
                    new_acc = '%s.%03d' % (current_kfo, key)
                    sheet_dict[new_acc] = value
                    del(sheet_dict[acc])
        elif '.' in key or 'Н' in key or key == 'ОЦИ':
            current_acc = key
        else:
            acc = '%s.%s' % (current_kfo, current_acc)
            if acc not in sheet_dict:
                sheet_dict[acc] = OrderedDict()

            if key in sheet_dict[acc]:
                print("Double KBK %r in account %s, line #%d" % (key, acc, i + 1))
                j = 1
                candidate = '%s_%d' % (key, j)
                while candidate in sheet_dict[acc]:
                    j += 1
                    candidate = '%s_%d' % (key, j)
                key = candidate
            sheet_dict[acc][key] = row
    return sheet_dict


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
        elif 1 < parts <= 3 and len(key) < 17 or 'Н' in key:
            current_acc = key
            sheet_dict[current_acc] = OrderedDict()
        else:
            assert current_acc is not None, "line #%d" % (i+1)
            key = ''.join(key.split('.'))

            if len(key) == 20 and key.startswith('000'):
                key = key[3:]

            assert key not in sheet_dict[current_acc], "Double KBK %s in account %s, line #%d" % (key, current_acc, i+1)
            sheet_dict[current_acc][key] = row[1:]

    return sheet_dict


def load_osv_general(sheet: xlrd.sheet.Sheet):
    if sheet.row_values(1)[0].startswith('Оборотно-сальдовая ведомость'):
        return load_osv_1c(sheet)
    else:
        return load_osv_smeta(sheet)


def check_format(sheet: xlrd.sheet.Sheet):
    cell00 = sheet.row_values(0)[0]
    cell10 = sheet.row_values(1)[0]
    if cell00.startswith('Оборотно-сальдовая ведомость') or cell10.startswith('Оборотно-сальдовая ведомость'):
        return '1c'
    elif cell10 == 'ОБОРОТНО-САЛЬДОВАЯ ВЕДОМОСТЬ':
        return 'Smeta'
    else:
        return 'unknown'


def main(file1, file2):
    pp = pprint.PrettyPrinter()

    wb = xlrd.open_workbook(file1, formatting_info=True)
    sheet = wb.sheet_by_index(0)
    print('Loading File1')
    osv_smeta = load_osv_1c(sheet)
    print('File1 loaded')

    #pp.pprint([(key, len(val)) for key, val in osv_smeta.items()])

    wb = xlrd.open_workbook(file2, formatting_info=True)
    sheet = wb.sheet_by_index(0)

    osv_1c = load_osv_smeta(sheet)
    print('File2 loaded')

    #pp.pprint([(key, len(val)) for key, val in osv_1c.items()])

    # Сравнение набора счетов
    accs1 = set(osv_smeta.keys())
    accs2 = set(osv_1c.keys())

    for item in sorted(accs1 - accs2, key=lambda x: x.split('.')):
        print('-', item)

    for item in sorted(accs2 - accs1, key=lambda x: x.split('.')):
        print('+', item)

    print()

if __name__ == '__main__':
    if len(sys.argv) >= 3:
        file1, file2 = sys.argv[1], sys.argv[2]
    else:
        file1 = r'МАУК Нижнематренский\МАУК Нижнематренский ОСВ за 2016 г. - после.xls'
        file2 = r'МАУК Нижнематренский\OSV_VED_1.xls'

    main(file1, file2)

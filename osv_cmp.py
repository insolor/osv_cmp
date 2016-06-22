import json
import pprint
import sys
import xlrd

from collections import OrderedDict


def load_osv_1c(sheet: xlrd.sheet.Sheet):
    log = []
    sheet_dict = OrderedDict()
    current_kfo = 0
    current_acc = None
    for i in range(12, sheet.nrows):
        row = sheet.row_values(i)
        key = row[0]
        row = [0.0 if not item else item for item in (row[3], row[6], row[9], row[14], row[16], row[19])]
        assert current_acc is None or 'None' not in current_acc, 'Line #%d' % i
        if key == 'Итого':
            break
        elif isinstance(key, float):
            next_key = str(sheet.row_values(i + 1)[0]) if i < sheet.nrows - 1 else ''
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
            assert current_acc is not None, "Line #%d" % i
            acc = '%s.%s' % (current_kfo, current_acc)
            if acc not in sheet_dict:
                sheet_dict[acc] = OrderedDict()

            if key in sheet_dict[acc]:
                log.append("Double KBK %r in account %s, line #%d" % (key, acc, i + 1))
                j = 1
                candidate = '%s_%d' % (key, j)
                while candidate in sheet_dict[acc]:
                    j += 1
                    candidate = '%s_%d' % (key, j)
                key = candidate
            sheet_dict[acc][key] = row
    return sheet_dict, log


def load_osv_smeta(sheet: xlrd.sheet.Sheet):
    log = []
    sheet_dict = OrderedDict()
    current_acc = None
    for i in range(8, sheet.nrows):
        row = sheet.row_values(i)
        key = row[0].strip()

        parts = key.count('.') + 1
        if key.startswith('Итого'):
            break
        elif parts == 1 and 0 < len(key) < 17:
            pass
        elif 1 < parts <= 3 and len(key) < 17 or 'Н' in key:
            current_acc = key
            sheet_dict[current_acc] = OrderedDict()
        else:
            assert current_acc is not None, "line #%d" % (i+1)
            key = ''.join(key.split('.'))

            if len(key) == 20 and key.startswith('000'):
                key = key[3:]
            
            if key in sheet_dict[current_acc]:
                log.append("Double KBK %r in account %s, line #%d" % (key, current_acc, i + 1))
                j = 1
                candidate = '%s_%d' % (key, j)
                while candidate in sheet_dict[current_acc]:
                    j += 1
                    candidate = '%s_%d' % (key, j)
                key = candidate
            
            sheet_dict[current_acc][key] = row[1:]

    return sheet_dict, log


def check_format(sheet: xlrd.sheet.Sheet):
    cell00 = sheet.row_values(0)[0]
    cell10 = sheet.row_values(1)[0]
    if cell00.startswith('Оборотно-сальдовая ведомость') or cell10.startswith('Оборотно-сальдовая ведомость'):
        return '1c'
    elif cell10 == 'ОБОРОТНО-САЛЬДОВАЯ ВЕДОМОСТЬ':
        return 'Smeta'
    else:
        return 'unknown'


if __name__ == '__main__':
    pass

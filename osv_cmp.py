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
    
    start = None
    for i in range(0, sheet.nrows):
        if sheet.row_values(i)[0] == 'КПС':
            start = i+1
            break
    
    for i in range(start, sheet.nrows):
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
        elif current_acc is None:
            current_kfo = 0
        else:
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
    if any(sheet.row_values(i)[0].startswith('Оборотно-сальдовая ведомость') for i in range(2)):
        return '1c'
    elif cell10 == 'ОБОРОТНО-САЛЬДОВАЯ ВЕДОМОСТЬ':
        return 'Smeta'
    else:
        return 'unknown'


def osv_compare(*osv):
    assert len(osv) == 2
    
    # Compare accounts
    accs = [set(item.keys()) for item in osv]
    
    if accs[0] == accs[1]:
        diff_accs = None
    else:
        diff_accs = (
            # Что пропало (из того что было вычесть то что осталось)
            sorted(accs[0] - accs[1], key=lambda x: x.split('.')),
            # Что появилось (из того что стало вычесть то что было)
            sorted(accs[1] - accs[0], key=lambda x: x.split('.'))
        )
    
    # Compare subrecords
    diffs = OrderedDict()
    osv = osv
    for acc in osv[0]:
        if acc in osv[1]:
            records = [set(osv[i][acc].keys()) for i in range(2)]
            if records[0] == records[1]:
                continue
            diffs[acc] = (sorted(records[0] - records[1]), sorted(records[1] - records[0]))
    diff_records = diffs
    
    # Compare sums
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
    
    diff_sums = diffs
    
    return dict(accs=diff_accs, records=diff_records, sums=diff_sums)


if __name__ == '__main__':
    pass

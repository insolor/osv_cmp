from collections import OrderedDict


class KBK:
    def __init__(self, s):
        self.original = str(s)
    
    @staticmethod
    def normalize(kbk):
        key = ''.join(str(kbk).split('.'))
        if len(key) == 20:
            key = key[3:]
        return key
    
    @property
    def normalized(self):
        return KBK.normalize(self.original)
    
    def __str__(self):
        return self.original
    
    def __eq__(self, other):
        return self.normalized == KBK.normalize(str(other))
    
    def __hash__(self):
        return self.normalized.__hash__()


def load_osv_1c(sheet):
    log = []
    sheet_dict = OrderedDict()
    current_kfo = 0
    current_acc = None
    
    start = None
    for i, row in enumerate(sheet):
        if row[0] == 'КПС':
            start = i+1
            break
    
    for i, row in enumerate(sheet[start:]):
        key = row[0]
        row = [0.0 if not item else item for item in (row[3], row[6], row[9], row[14], row[16], row[19])]
        assert current_acc is None or 'None' not in current_acc, 'Line #%d' % i
        if key == 'Итого':
            break
        elif isinstance(key, float):
            next_key = str(sheet[i + 1][0]) if i < len(sheet) - 1 else ''
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
                log.append("Дублирующуяся запись %r в счете %s, строка #%d" % (key, acc, i + 1))
                j = 1
                candidate = '%s_%d' % (key, j)
                while candidate in sheet_dict[acc]:
                    j += 1
                    candidate = '%s_%d' % (key, j)
                key = candidate
            sheet_dict[acc][key] = row
    return sheet_dict, log


def load_osv_smeta(sheet):
    log = []
    sheet_dict = OrderedDict()
    current_acc = None
    heads = set()
    for i, row in enumerate(sheet[8:]):
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
            if current_acc is None:
                log.append("Не удалось определить текущий счет, строка #%d.\n"
                           "Возможно, при формировании оборотно-сальдовой ведомости не были выбраны "
                           "необходимые пункты группировки. Загрузка прервана." % (i + 1))
                return None, log

            if key and '.' not in key:
                log.append("КБК %r без точек в счете %s, строка #%d. "
                           "Необходимо заполнить поля данного КБК в Смете-СМАРТ." % (key, current_acc, i + 1))

            head = key.partition('.')[0]
            if len(head) == 3:
                heads.add(head)
            key = ''.join(key.split('.'))

            if len(key) == 20:
                key = key[3:]
            
            if key in sheet_dict[current_acc]:
                log.append("Дублирующуяся запись %r в счете %s, строка #%d" % (key, current_acc, i + 1))
                j = 1
                candidate = '%s_%d' % (key, j)
                while candidate in sheet_dict[current_acc]:
                    j += 1
                    candidate = '%s_%d' % (key, j)
                key = candidate
            
            sheet_dict[current_acc][key] = row[1:]

    log.append("Коды главы в оборотно-сальдовой ведомости: {}\n".format(', '.join(heads)))
    return sheet_dict, log


def check_format(sheet):
    if any(sheet[i][0].startswith('Оборотно-сальдовая ведомость') for i in range(2)):
        return '1c'
    elif sheet[1][0] == 'ОБОРОТНО-САЛЬДОВАЯ ВЕДОМОСТЬ':
        return 'Smeta'
    else:
        return 'unknown'


def symm_diff_dicts(d1, d2):
    sd1 = set(d1.keys())
    sd2 = set(d2.keys())
    absent_keys = sd1 - sd2
    absent_records = {key: value for key, value in d1.items() if key in absent_keys}
    new_keys = sd2 - sd1
    new_records = {key: value for key, value in d2.items() if key in new_keys}
    return absent_records, new_records


def osv_compare(*osv):
    assert len(osv) == 2
    
    # Compare accounts
    accs = [set(item.keys()) for item in osv]
    
    if accs[0] == accs[1]:
        diff_accs = None
    else:
        diff_accs = [sorted(x, key=lambda x: x.split('.')) for x in (
                accs[0] - accs[1],  # Что пропало (из того что было вычесть то что осталось)
                accs[1] - accs[0],  # Что появилось (из того что стало вычесть то что было)
            )
        ]
    
    # Compare subrecords
    diffs = OrderedDict()
    for acc in osv[0]:
        if acc in osv[1]:
            records = [osv[i][acc] for i in range(2)]
            if set(records[0].keys()) == set(records[1].keys()):
                continue
            diffs[acc] = symm_diff_dicts(records[0], records[1])
    diff_records = diffs
    
    def money_is_equal(x, y, eps=0.001):
        return abs(x - y) < eps
    
    # Compare sums
    diffs = OrderedDict()
    for acc in osv[0]:
        if acc in osv[1]:
            for record, row in osv[0][acc].items():
                if record in osv[1][acc]:
                    row2 = osv[1][acc][record]
                    if row[:4] == row2[:4]:
                        continue
                    elif money_is_equal(row[0]-row[1], row2[0]-row2[1]) and money_is_equal(row[2]-row[3], row2[2]-row2[3]):
                        continue
                    else:
                        if acc not in diffs:
                            diffs[acc] = OrderedDict()
                        
                        diffs[acc][record] = (row[:4], row2[:4])

    diff_sums = diffs
    
    return dict(accs=diff_accs, records=diff_records, sums=diff_sums)


def sum_lists(s):
    s = iter(s)
    x = list(next(s))
    for row in s:
        assert len(x) == len(row), "Row lengths must be the same"
        for i, item in enumerate(row):
            x[i] += item
    return x


def osv_sum(osv: dict):
    return sum_lists(operations[:6] for subrecord in osv.values() for operations in subrecord.values())


if __name__ == '__main__':
    pass

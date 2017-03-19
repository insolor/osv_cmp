from collections import OrderedDict
from collections.abc import *


class KBK:
    def __init__(self, s, suffix=''):
        self.original = str(s)
        self.suffix = suffix
    
    @staticmethod
    def normalize(kbk):
        key = ''.join(str(kbk).split('.'))
        if len(key) == 20:
            key = key[3:]
        return key
    
    @property
    def normalized(self):
        return KBK.normalize(self.original) + self.suffix
    
    def __len__(self):
        return len(''.join(self.original.split('.')))
    
    def __str__(self):
        return self.original + self.suffix

    def __repr__(self):
        return repr(self.original) + self.suffix
    
    def __eq__(self, other):
        return self.normalized == KBK.normalize(str(other))

    def __lt__(self, other):
        return self.normalized < KBK.normalize(str(other))
    
    def __hash__(self):
        return self.normalized.__hash__()


def load_osv_1c(rows: Sequence):
    log = []
    data_dict = OrderedDict()
    current_dep = ''
    current_kfo = 0
    current_acc = None
    
    start = None
    for i, row in enumerate(rows):
        if row[0] == 'КПС':
            start = i+1
            break
    
    for i, row in enumerate(rows[start:]):
        key = row[0]  # type: str
        row = [float(item) if item else 0.0 for item in (row[3], row[6], row[9], row[14], row[16], row[19])]
        assert current_acc is None or 'None' not in current_acc, 'Line #%d' % i
        if key == 'Итого':
            break
        elif ' ' in key:  # Наименование учреждения
            current_dep = key
        elif len(key) == 1 and key.isdigit():  # КФО (N)
            current_kfo = key
        elif key and len(key) <= 6:  # Счет (NNN.MM)
            current_acc = key
        elif len(key) <= 17:  # КПС или пусто (17N)
            if current_dep not in data_dict:
                data_dict[current_dep] = OrderedDict()

            current_section = data_dict[current_dep]

            acc = '%s.%s' % (current_kfo, current_acc)
            if acc not in current_section:
                current_section[acc] = OrderedDict()

            if key in current_section[acc]:
                log.append("Дублирующуяся запись %r в счете %s, строка #%d" % (key, acc, i + 1))
                j = 1
                candidate = '%s_%d' % (key, j)
                while candidate in current_section[acc]:
                    j += 1
                    candidate = '%s_%d' % (key, j)
                key = candidate
            current_section[acc][key] = row

    return data_dict, log


def load_osv_smeta(rows: Sequence):
    log = []
    data_dict = OrderedDict()
    current_dep = ''
    current_acc = None
    heads = set()

    start = None
    for i, row in enumerate(rows):
        if row[0] == 'Субсчет':
            start = i+2
            break

    for i, row in enumerate(rows[start:]):
        key = row[0].strip()
        parts = key.count('.') + 1
        key_plain = ''.join(key.split('.'))
        if key.startswith('Итого'):
            break
        elif len(key) == 1 and key.isdigit():  # КФО (N)
            pass
        elif parts <= 3 and 1 < len(key_plain) < 17:  # Счет 6 знаков (N.MMM.KK)
            current_acc = key
        elif ' ' in key:  # Наименование учреждения (считаем, что это несколько слов, разделенных пробелами)
            current_dep = key
        else:  # КБК или пусто (могут быть цифры, буквы, разделенные или не разделенные точками)
            if current_acc is None:
                log.append("Не удалось определить текущий счет, строка #%d.\n"
                           "Возможно, при формировании оборотно-сальдовой ведомости не были выбраны "
                           "необходимые пункты группировки. Загрузка прервана." % (i + 1))
                return None, log

            if current_dep not in data_dict:
                data_dict[current_dep] = OrderedDict()

            current_section = data_dict[current_dep]

            acc = current_acc
            if acc not in current_section:
                current_section[acc] = OrderedDict()

            if key:
                if '.' not in key:
                    log.append("КБК %r без точек в счете %s, строка #%d. "
                               "Необходимо заполнить поля данного КБК в Смете-СМАРТ." % (key, acc, i + 1))

                if len(key_plain) == 17:
                    log.append("Слишком короткий (старый) КБК: '%s' (%d цифр) в счете %s, строка #%d" %
                               (key, len(key_plain), acc, i + 1))
                elif len(key_plain) < 20:
                    log.append("Слишком короткий КБК: '%s' (%d цифр) в счете %s, строка #%d" %
                               (key, len(key_plain), acc, i + 1))
            
            head = key.partition('.')[0]
            if len(head) == 3:
                heads.add(head)

            key = KBK(key)
            
            if key in current_section[acc]:
                log.append("Дублирующаяся запись %r в счете %s, строка #%d" % (key, acc, i + 1))
                j = 1
                candidate = KBK(key, '(%s)' % j)
                while candidate in current_section[acc]:
                    j += 1
                    candidate = KBK(key, '(%s)' % j)
                key = candidate
            
            current_section[acc][key] = [float(item) for item in row[1:]]

    log.append("Коды главы в оборотно-сальдовой ведомости: {}\n".format(', '.join(heads)))
    return data_dict, log


def check_format(rows: Sequence):
    if any(rows[i][0].startswith('Оборотно-сальдовая ведомость') for i in range(2)):
        return '1c'
    elif rows[1][0] == 'ОБОРОТНО-САЛЬДОВАЯ ВЕДОМОСТЬ':
        return 'Smeta'
    else:
        return 'unknown'


def symm_diff_dicts(d1: dict, d2: dict):
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
                    elif (money_is_equal(row[0]-row[1], row2[0]-row2[1]) and
                          money_is_equal(row[2]-row[3], row2[2]-row2[3])):
                        continue
                    else:
                        if acc not in diffs:
                            diffs[acc] = OrderedDict()
                        
                        diffs[acc][record] = (row[:4], row2[:4])

    diff_sums = diffs
    
    return dict(accs=diff_accs, records=diff_records, sums=diff_sums)


def sum_lists(s: iter):
    x = list(next(s))
    for row in s:
        assert len(x) == len(row), "Row lengths must be the same"
        for i, item in enumerate(row):
            x[i] += item
    return x


def osv_sum(osv: dict):
    return sum_lists(operations[:6]
                     for dep_accounts in osv.values()
                     for account in dep_accounts.values()
                     for operations in account.values())


if __name__ == '__main__':
    pass

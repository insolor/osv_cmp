import pytest

from osv_cmp import KBK

def test_kbk():
    kbk = KBK('123.345345.34635745.345')
    assert str(kbk) == '123.345345.34635745.345'

    assert kbk.normalized == '34534534635745345'

    d = dict()
    d[kbk] = None

    assert kbk in d
    assert str(kbk) not in d
    assert kbk.normalized in d

import pytest
from deallens.lbo import Assumptions,calculate_lbo,_irr

def test_lbo():
    r=calculate_lbo(Assumptions())
    assert r['entry_ev']==pytest.approx(27.3)
    assert r['entry_debt']==pytest.approx(11.7)
    assert len(r['projection'])==5
    assert r['moic']>0

def test_irr():
    assert _irr([-100,0,121])==pytest.approx(.1)

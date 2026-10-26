from deallens.memo import normalize_result,MEMO_KEYS

def test_reject_out_of_context():
    r=normalize_result({'company_name':'T','memo':{},'historical':[{'year':2025,'revenue':10,'source_page':99}]},[1])
    assert not r['historical']
    assert all(r['memo'].get(k) for k in MEMO_KEYS)

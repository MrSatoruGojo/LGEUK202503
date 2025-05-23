# tests/test_string_utils.py

from src.utils.string_utils import sanitize_filename, extract_short_name, truncate_middle

def test_sanitize_filename():
    bad = r'inva|id:na*me?.txt'
    clean = sanitize_filename(bad)
    # Should replace illegal chars with underscore
    assert ':' not in clean and '*' not in clean and '?' not in clean

def test_extract_short_name():
    assert extract_short_name("VeryLongCustomerNameHere", 12) == "VERYLONGCUST"
    assert extract_short_name("", 12) == ""

def test_truncate_middle():
    s = "abcdefghijklmnopqrstuvwxyz"
    t = truncate_middle(s, max_len=10)
    # Should be length 10
    assert len(t) == 10
    # Should contain an ellipsis
    assert '…' in t

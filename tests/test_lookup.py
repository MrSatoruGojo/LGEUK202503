# tests/test_lookup.py

import pandas as pd
from src.utils.lookup import create_lookup_dict, lookup_value, lookup_customer

def make_df():
    return pd.DataFrame([
        {'Promotion No': 1, 'FieldA': 'X', 'Bill To Name': 'Customer1', 'Product Code': 'P1'},
        {'Promotion No': 1, 'FieldA': 'Y', 'Bill To Name': 'Customer2', 'Product Code': 'P2'},
        {'Promotion No': 2, 'FieldA': 'Z', 'Bill To Name': 'Customer1', 'Product Code': 'P3'},
    ])

def test_create_and_lookup():
    df = make_df()
    lookup = create_lookup_dict(df, 'Promotion No', ['FieldA', 'Bill To Name', 'Product Code'])
    # For promo 1, first record has FieldA X
    assert lookup_value(1, 'FieldA', lookup) == 'X'
    # Non-existent promo returns default
    assert lookup_value(999, 'FieldA', lookup, default='NOT') == 'NOT'
    # lookup_customer with existing name
    val = lookup_customer(1, 'FieldA', lookup, customer_name='Customer1')
    assert val == 'X'

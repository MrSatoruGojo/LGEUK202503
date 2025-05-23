# tests/test_tracker.py

import pandas as pd
from src.processing.tracker import enrich_tracker_data

def test_enrich_tracker_data_basic():
    # Claim with PSI and closed orders already present
    claim = pd.DataFrame([{
        'Bill To Name Short': 'CUSTX',
        'Product Code SPMS': 'P1',
        'SELL-OUT': 2,
        'SELL-IN': 3,
        'INVENTORY': 4,
        'Total Closed Orders': 10,
        'Q': 1
    }])

    # Tracker data: one matching row
    tracker = pd.DataFrame([{
        'Customer': 'CustXtraName',
        'Model': 'P1',
        'Claim Volume': 4
    }])

    out = enrich_tracker_data(claim, tracker)

    assert out.at[0, 'Tracker']      == 4
    assert out.at[0, 'PSI TOTAL']    == 2+3+4
    assert out.at[0, 'CO CHECK']     == 10 - 4 - 1
    assert out.at[0, 'PSI CHECK']    == 2 - 1
    assert out.at[0, 'ID PSI CHECK'] == (2+3+4) - 1

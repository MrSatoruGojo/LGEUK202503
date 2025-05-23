# tests/test_psi.py

import pandas as pd
from src.processing.psi import enrich_psi_data

def test_enrich_psi_data_basic():
    # One claim row with a single week label
    claim_df = pd.DataFrame([{
        'Bill To Name Short': 'CUSTSHORT',
        'Product Code SPMS': 'P1',
        "Week's range": 'WEEK1'
    }])

    # Build PSI DataFrame: 6 metadata cols + 1 weekly col
    data = [
        {'Channel': 'CustShortXYZ', 'Model.Suffix': 'P1', 'Measure': 'Sell-Out FCST_KAM [R+F]',
         'M1':0,'M2':0,'M3':0, 'WEEK1': 5},
        {'Channel': 'CustShortXYZ', 'Model.Suffix': 'P1', 'Measure': 'Sell-In FCST_KAM [R+F]',
         'M1':0,'M2':0,'M3':0, 'WEEK1': 10},
        {'Channel': 'CustShortXYZ', 'Model.Suffix': 'P1', 'Measure': 'Ch. Inventory_Sellable',
         'M1':0,'M2':0,'M3':0, 'WEEK1': 15},
    ]
    psi_df = pd.DataFrame(data, columns=['Channel','Model.Suffix','Measure','M1','M2','M3','WEEK1'])

    out = enrich_psi_data(claim_df, psi_df)

    assert out.at[0, 'SELL-OUT']   == 5
    assert out.at[0, 'SELL-IN']    == 10
    assert out.at[0, 'INVENTORY']  == 15

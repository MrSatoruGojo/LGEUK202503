# tests/test_claim.py

import pandas as pd
from src.processing.claim import enrich_claim_data

def test_enrich_claim_data_minimal():
    # One promotion with known SPMS mapping
    df = pd.DataFrame([{
        'Promotion No': 1,
        'Bill To Name': 'CustomerNameHere',
        'Product Code': 'Prod1'
    }])
    spms_map = {
        1: [{
            'Cancel Flag': 'Y',
            'Recreate Flag': 'N',
            'Promotion Start YYYYMMDD': '20210101',
            'Promotion End YYYYMMDD':   '20210114',
            'Bill To Name': 'CustomerNameHere',
            'Product Code': 'Prod1',
            'Sales PGM NO': 'SPGM1'
        }]
    }
    spms2_map = {}

    enriched = enrich_claim_data(df, spms_map, spms2_map)

    # Flags
    assert enriched.at[0, 'Cancel Flag'] == 1
    assert enriched.at[0, 'Recreate Flag'] == 0

    # Dates preserved
    assert enriched.at[0, 'Promotion Start Date'] == '20210101'
    assert enriched.at[0, 'Promotion End Date']   == '20210114'

    # Year conversion
    assert enriched.at[0, 'Promotion Start Year'] == 2021
    assert enriched.at[0, 'Promotion End Year']   == 2021

    # SPMS fields
    assert enriched.at[0, 'Bill To Name SPMS']    == 'CustomerNameHere'
    assert enriched.at[0, 'Product Code SPMS']    == 'Prod1'
    assert enriched.at[0, 'Sales PGM NO']         == 'SPGM1'

    # Short name is first 12 uppercase chars
    assert enriched.at[0, 'Bill To Name Short']   == 'CUSTOMERNAME'

    # Week’s range is a non-empty string
    wr = enriched.at[0, "Week's range"]
    assert isinstance(wr, str) and len(wr) > 0

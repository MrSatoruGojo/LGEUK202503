# tests/test_orders.py

import pandas as pd
import os
from src.processing.orders import merge_closed_orders

def test_merge_closed_orders_with_temp_excel(tmp_path):
    # Create a dummy closed-orders sheet
    orders = pd.DataFrame({
        'Bill To Name': ['Cust1','Other'],
        'Model': ['M1','X'],
        'Order Qty': [10, 5]
    })
    file_path = tmp_path / "2021 CLOSED ORDERS.xlsx"
    orders.to_excel(file_path, index=False)

    # Claim DataFrame with one matching row
    claim_df = pd.DataFrame([{
        'Bill To Name Short': 'CUST1',
        'Product Code SPMS': 'M1'
    }])

    result = merge_closed_orders(claim_df, [str(file_path)])

    # Check that the yearly column exists and sums correctly
    col = 'Order Qty 2021'
    assert col in result.columns
    assert result.at[0, col] == 10

    # Total Closed Orders should equal the same
    assert result.at[0, 'Total Closed Orders'] == 10

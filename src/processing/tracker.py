# src/processing/tracker.py

import logging
import pandas as pd

from src.utils.string_utils import extract_short_name

def enrich_tracker_data(
    claim_df: pd.DataFrame,
    tracker_df: pd.DataFrame,
    cust_short_key: str = 'Bill To Name Short',
    prod_key: str = 'Product Code SPMS',
    tracker_customer_col: str = 'Customer',
    tracker_model_col: str = 'Model',
    tracker_volume_col: str = 'Claim Volume',
    claim_qty_col: str = 'Q'
) -> pd.DataFrame:
    """
    Enrich the DataFrame with old‐tracker sums and compute check columns.

    - Summarizes `tracker_df` by 12-char customer & model into a 'Tracker' column.
    - Computes 'PSI TOTAL', 'CO CHECK', 'PSI CHECK', and 'ID PSI CHECK'.

    Parameters
    ----------
    claim_df : pd.DataFrame
        The DataFrame after PSI enrichment, must contain:
        - cust_short_key (12-char customer)
        - prod_key (product code)
        - 'SELL-OUT', 'SELL-IN', 'INVENTORY', 'Total Closed Orders', and claim_qty_col
    tracker_df : pd.DataFrame
        Raw tracker data, expected to have:
        - tracker_customer_col (full customer name)
        - tracker_model_col    (model code)
        - tracker_volume_col   (numeric volumes)
    cust_short_key : str
        Column in claim_df for the 12-char customer key.
    prod_key : str
        Column in claim_df for the model code.
    tracker_customer_col : str
        Column in tracker_df with full customer names.
    tracker_model_col : str
        Column in tracker_df with model codes.
    tracker_volume_col : str
        Column in tracker_df with the numeric claim volumes.
    claim_qty_col : str
        Column in claim_df with the original claim quantity ('Q').

    Returns
    -------
    pd.DataFrame
        Copy of claim_df with new columns:
        - 'Tracker'
        - 'PSI TOTAL'
        - 'CO CHECK'
        - 'PSI CHECK'
        - 'ID PSI CHECK'
    """
    df = claim_df.copy()
    tdf = tracker_df.copy()
    logging.info("Starting tracker enrichment…")

    # Prepare tracker_df: 12-char key and numeric volumes
    tdf['Customer Short'] = (
        tdf[tracker_customer_col]
        .fillna('')
        .apply(lambda x: extract_short_name(x, 12))
    )
    tdf[tracker_volume_col] = pd.to_numeric(
        tdf[tracker_volume_col], errors='coerce'
    ).fillna(0)

    # Summarize by keys
    # Initialize the Tracker column
    df['Tracker'] = 0

    # For each row, sum any tracker_df rows whose Customer startswith our short name
    for idx, row in df.iterrows():
        cust = str(row[cust_short_key]).upper()
        prod = row[prod_key]
        mask = (
            tdf[tracker_customer_col]
                .fillna('')
                .str.upper()
                .str.startswith(cust)
            &
            (tdf[tracker_model_col] == prod)
        )
        total_volume = tdf.loc[mask, tracker_volume_col].sum()
        df.at[idx, 'Tracker'] = int(total_volume)

    # Compute check columns
    # Ensure claim_qty_col exists
    if claim_qty_col not in df.columns:
        logging.warning(f"Column '{claim_qty_col}' not found in claim_df; using 0 for checks.")
        df[claim_qty_col] = 0

    df['PSI TOTAL'] = df.get('SELL-OUT', 0) + df.get('SELL-IN', 0) + df.get('INVENTORY', 0)
    df['CO CHECK']  = df.get('Total Closed Orders', 0) - df['Tracker'] - df[claim_qty_col]
    df['PSI CHECK'] = df.get('SELL-OUT', 0) - df[claim_qty_col]
    df['ID PSI CHECK'] = df['PSI TOTAL'] - df[claim_qty_col]

    logging.info("Finished tracker enrichment.")
    return df

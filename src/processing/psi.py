# src/processing/psi.py

import logging
import pandas as pd
from src.utils.string_utils import extract_short_name

def enrich_psi_data(
    claim_df: pd.DataFrame,
    psi_df: pd.DataFrame,
    promo_key: str = 'Promotion No',
    cust_short_key: str = 'Bill To Name Short',
    prod_key: str = 'Product Code SPMS'
) -> pd.DataFrame:
    """
    Enrich the claim DataFrame with SELL-OUT, SELL-IN, and INVENTORY values
    pulled from PSI weekly forecasts/inventory.

    Parameters
    ----------
    claim_df : pd.DataFrame
        The DataFrame coming out of merge_closed_orders(),
        must include cust_short_key, prod_key, and "Week's range".
    psi_df : pd.DataFrame
        Raw PSI DataFrame, expected to have columns:
        ['Channel', 'Model.Suffix', 'Measure', <weekly cols...>]
    promo_key : str
        Name of the promotion-number column (unused here but kept for symmetry).
    cust_short_key : str
        Column in claim_df with the 12-char short customer name.
    prod_key : str
        Column in claim_df with the product code to match PSI 'Model.Suffix'.

    Returns
    -------
    pd.DataFrame
        Copy of claim_df with three new columns: 'SELL-OUT', 'SELL-IN', 'INVENTORY'.
    """
    df = claim_df.copy()
    logging.info("Starting PSI enrichment…")

    # Validate presence of key PSI columns
    if 'Channel' not in psi_df.columns or 'Model.Suffix' not in psi_df.columns:
        logging.error("PSI enrichment skipped: 'Channel' or 'Model.Suffix' missing in PSI data")
        return df

    # Determine weekly columns (assume first 6 cols are metadata)
    weekly_cols = list(psi_df.columns[6:])

    # Define the measures and target column names
    measures = [
        ('Sell-Out FCST_KAM [R+F]', 'SELL-OUT'),
        ('Sell-In FCST_KAM [R+F]', 'SELL-IN'),
        ('Ch. Inventory_Sellable', 'INVENTORY'),
    ]

    for measure_name, target_col in measures:
        logging.info(f"Enriching '{target_col}' using measure '{measure_name}'")
        # Filter PSI for this measure
        psi_subset = psi_df[psi_df['Measure'] == measure_name].copy()

        # Initialize target column
        df[target_col] = 0

        # Iterate rows to sum relevant weeks
        for idx, row in df.iterrows():
            cust = row.get(cust_short_key, '')
            prod = row.get(prod_key, '')
            weeks_range = row.get("Week's range", '')

            if not weeks_range or not cust or not prod:
                continue

            # Parse labels, only keep those actually in PSI
            labels = [w for w in weeks_range.split(', ') if w in weekly_cols]
            if not labels:
                continue

            # Filter by channel & model
            matched = psi_subset[
                psi_subset['Channel']
                    .fillna('')
                    .str.upper()
                    .str.startswith(cust)
                &
                (psi_subset['Model.Suffix'] == prod)
            ]
            if matched.empty:
                continue

            # Sum across the selected weekly columns
            try:
                total = matched[labels].sum(axis=1).sum()
                df.at[idx, target_col] = total
            except KeyError as e:
                logging.warning(f"Some week columns not found for {target_col}: {e}")

        logging.info(f"Finished enriching '{target_col}'")

    logging.info("All PSI measures enriched.")
    return df

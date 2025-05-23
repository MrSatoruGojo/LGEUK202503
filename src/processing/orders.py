# src/processing/orders.py

import os
import re
import logging
import pandas as pd

from src.io.file_ops import read_excel_file
from src.utils.string_utils import extract_short_name

def merge_closed_orders(
    claim_df: pd.DataFrame,
    closed_files: list[str]
) -> pd.DataFrame:
    """
    Merge closed-orders into the claim dataframe.

    Parameters
    ----------
    claim_df : pd.DataFrame
        The enriched claim DataFrame (must include 'Bill To Name Short' and 'Product Code SPMS').
    closed_files : list[str]
        List of file paths to the closed-orders Excel workbooks.

    Returns
    -------
    pd.DataFrame
        A copy of claim_df with:
        - One 'Order Qty <year>' column per file (summing Order Qty)
        - A 'Total Closed Orders' column summing across years.
    """
    df = claim_df.copy()
    logging.info("Starting merge_closed_orders…")

    # Prepare lookup sets
    customers = set(df['Bill To Name Short'].dropna().unique())
    models    = set(df['Product Code SPMS'].dropna().unique())

    for filepath in closed_files:
        filename = os.path.basename(filepath)
        # Extract leading 4-digit year, fallback to 'unknown'
        match = re.match(r'(\d{4})', filename)
        year = match.group(1) if match else 'unknown'

        try:
            logging.info(f"Reading closed-orders file {filename}")
            sheets = read_excel_file(filepath, sheet_name=None)

            # Get the first sheet as a DataFrame
            if isinstance(sheets, dict):
                sheet_df = list(sheets.values())[0].copy()
            else:
                sheet_df = sheets.copy()

            # Verify required columns
            required = {'Bill To Name', 'Model', 'Order Qty'}
            if not required.issubset(sheet_df.columns):
                logging.warning(f"Missing columns in {filename}, expected {required}. Skipping.")
                continue

            # Build short name for filtering
            sheet_df['Bill To Name Short'] = sheet_df['Bill To Name'] \
                .apply(lambda x: extract_short_name(x, 12))

            # Filter on matching customers & models
            filtered = sheet_df[
                sheet_df['Bill To Name Short'].isin(customers) &
                sheet_df['Model'].isin(models)
            ]

            if filtered.empty:
                logging.warning(f"No matching records in {filename}.")
                continue

            # Aggregate quantities
            agg = (
                filtered
                .groupby(['Bill To Name Short', 'Model'])['Order Qty']
                .sum()
                .reset_index()
            )
            agg['Year'] = year

            # Merge back into df
            df = df.merge(
                agg,
                how='left',
                left_on=['Bill To Name Short', 'Product Code SPMS'],
                right_on=['Bill To Name Short', 'Model']
            )

            col_name = f"Order Qty {year}"
            df[col_name] = df['Order Qty'].fillna(0).astype(int)

            # Clean up
            df.drop(columns=['Order Qty', 'Model', 'Year'], inplace=True)

            logging.info(f"Merged closed-orders for year {year}.")

        except Exception as e:
            logging.error(f"Error processing {filename}: {e}")

    # Sum across all Order Qty columns
    qty_cols = [c for c in df.columns if c.startswith('Order Qty ')]
    df['Total Closed Orders'] = df[qty_cols].sum(axis=1)

    logging.info("Finished merge_closed_orders.")
    return df

# src/processing/output_splits.py

import os
from datetime import datetime
import pandas as pd
from src.utils.string_utils import sanitize_filename
from src.output.formatter import format_workbook
from src.output.formatter import format_workbook
from src.output.enhancements import add_closed_orders_by_year
def split_by_bebs(
    cleaned_df: pd.DataFrame,
    tracker_part6_df: pd.DataFrame,  # formerly old_tracker_df
    tracker_part5_df: pd.DataFrame,  # your new Part 5
    spms_df: pd.DataFrame,
    output_folder: str,
    prefix: str = 'CLAIM'
):
    """
    For each unique BEBS code in cleaned_df:
      1. Write a "{prefix}_{BEBS}_{user}_{timestamp}.xlsx" file
      2. Sheet "VERIFICATION": the filtered cleaned_df rows
      3. Sheet "Tracker": old_tracker rows matching customer–model
      4. Sheet "SPMS": SPMS rows for those promotions, with BEBS in column A
      5. Apply workbook styling via format_workbook()
    """
    user = os.getlogin()
    ts   = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        # copy the incoming Part 6 tracker and build the Customer-Short key
    tracker = tracker_part6_df.copy()
    tracker['Customer Short'] = tracker['Customer']\
                                   .fillna('')\
                                   .astype(str)\
                                   .str[:12]\
                                   .str.upper()

    for bebs in cleaned_df['BEBS'].dropna().unique():
        subset = cleaned_df[cleaned_df['BEBS'] == bebs]
        if subset.empty:
            continue

        safe = sanitize_filename(bebs)
        fname = f"{prefix}_{safe}_{user}_{ts}.xlsx"
        path  = os.path.join(output_folder, fname)

        # 1) VERIFICATION sheet
        with pd.ExcelWriter(path, engine='openpyxl') as writer:
            subset.to_excel(writer, index=False, sheet_name='VERIFICATION')

                # 2) Part 6 sheet (formerly “Tracker”)
        pairs = subset[['Bill To Name Short','Product Code SPMS']].drop_duplicates()
        # join against the enriched tracker copy (with its Customer-Short column)
        part6_filtered = (
            tracker
            .merge(
                pairs.rename(columns={
                    'Bill To Name Short':'Customer Short',
                    'Product Code SPMS':'Model'
                }),
                on=['Customer Short','Model'], how='inner'
            )
        )
        with pd.ExcelWriter(path, engine='openpyxl', mode='a') as writer:
            part6_filtered.to_excel(
                writer, index=False, sheet_name='Part 6'
            )

        # 3) Part 5 sheet
        part5_filtered = (
            tracker_part5_df
            .merge(
                pairs.rename(columns={
                    'Bill To Name Short':'Customer Short',
                    'Product Code SPMS':'Model'
                }),
                on=['Customer Short','Model'], how='inner'
            )
        )
        with pd.ExcelWriter(path, engine='openpyxl', mode='a') as writer:
            part5_filtered.to_excel(
                writer, index=False, sheet_name='Part 5'
            )

        # 3) SPMS sheet
        promos = subset['Promotion No'].unique()
        spms_subset = spms_df[spms_df['Promotion No'].isin(promos)].copy()
        bebs_map    = subset[['Promotion No','BEBS']].drop_duplicates()
        spms_with_bebs = spms_subset.merge(bebs_map, on='Promotion No')
        # move BEBS to front
        spms_with_bebs.insert(0, 'BEBS', spms_with_bebs.pop('BEBS'))
        with pd.ExcelWriter(path, engine='openpyxl', mode='a') as writer:
            spms_with_bebs.to_excel(writer, index=False, sheet_name='SPMS')

         # 4) Styling existing sheets
        format_workbook(path)

         # 5) Closed Order Filtering 2: add raw closed-order detail by year
        #    subset is the DataFrame for this BEBS (from VERIFICATION)
        add_closed_orders_by_year(
            workbook_path=path,
            subset=subset,
            date_col='Order Date'   # adjust this if your date column has a different header
        )

        # 6) Re-style the newly added year‐sheets
        format_workbook(path)

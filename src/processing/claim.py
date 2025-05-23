# src/processing/claim.py

import logging
import pandas as pd

from src.utils.date_utils import date_to_week, date_to_year, generate_weeks_range_monday
from src.utils.string_utils import extract_short_name

from src.utils.lookup import (
    lookup_value,
    lookup_customer,
    lookup_customer2,
    lookup_SPGM
)

def enrich_claim_data(
    claim_df: pd.DataFrame,
    spms_map: dict,
    spms2_map: dict,
    promo_key: str = 'Promotion No',
    customer_field: str = 'Bill To Name',
    product_field: str = 'Product Code'
) -> pd.DataFrame:
    """
    Enrich the raw claim DataFrame with SPMS lookups, flags, date/week/year fields,
    and a 'Week's range' column for downstream processing.
    """
    df = claim_df.copy()
    logging.info("Starting claim enrichment…")

    # Mark which promotions exist in SPMS
    df['_has_spms'] = df[promo_key].isin(spms_map)

    # Define row‐level enrichment function
    def _enrich(row):
        promo = row[promo_key]
        cust = row.get(customer_field, '')
        prod = row.get(product_field, '')

        if not row['_has_spms']:
            # Default zeros when no SPMS record found
            row['Cancel Flag'] = 0
            row['Recreate Flag'] = 0
            row['Promotion Start Date'] = 0
            row['Promotion End Date'] = 0
            row['Promotion Start Week'] = 0
            row['Promotion End Week'] = 0
            row['Promotion Start Year'] = 0
            row['Promotion End Year'] = 0
            row['Bill To Name SPMS'] = ''
            row['Product Code SPMS'] = ''
            row['Sales PGM NO'] = ''
        else:
            # Cancel / Recreate flags (Y → 1, else 0)
            row['Cancel Flag'] = 1 if lookup_value(promo, 'Cancel Flag', spms_map) == 'Y' else 0
            row['Recreate Flag'] = 1 if lookup_value(promo, 'Recreate Flag', spms_map) == 'Y' else 0

            # Promotion start / end dates
            start_date = lookup_value(promo, 'Promotion Start YYYYMMDD', spms_map)
            end_date   = lookup_value(promo, 'Promotion End YYYYMMDD', spms_map)
            row['Promotion Start Date'] = start_date
            row['Promotion End Date']   = end_date

            # Weeks and years
            try:
                row['Promotion Start Week'] = date_to_week(start_date)
                row['Promotion End Week']   = date_to_week(end_date)
                row['Promotion Start Year'] = date_to_year(start_date)
                row['Promotion End Year']   = date_to_year(end_date)
            except Exception as e:
                logging.warning(f"Week/year conversion failed for promo {promo}: {e}")
                row['Promotion Start Week'] = row['Promotion End Week'] = 0
                row['Promotion Start Year'] = row['Promotion End Year'] = 0

            # Look up SPMS customer and product fields
            row['Bill To Name SPMS'] = lookup_customer(
                promo,
                'Bill To Name',
                spms_map,
                customer_name=cust[:12],
                secondary_lookup_dict=spms2_map
            )
            row['Product Code SPMS'] = lookup_customer2(
                promo,
                'Product Code',
                spms_map,
                product_code=prod,
                secondary_lookup_dict=spms2_map
            )
            # Sales PGM NO
            row['Sales PGM NO'] = lookup_SPGM(
                promo,
                cust,
                spms_map,
                secondary_lookup_dict=spms2_map
            )

        return row

    # Apply enrichment
    df = df.apply(_enrich, axis=1)

    # Clean up helper column
    df.drop(columns=['_has_spms'], inplace=True)

    # Derive a 12‐char uppercase short name for closed‐orders matching
    df['Bill To Name Short'] = df['Bill To Name SPMS'].fillna('').apply(lambda x: extract_short_name(x, 12))

    # Build the week's range string for each promotion
    df["Week's range"] = df.apply(
        lambda r: ', '.join(
            generate_weeks_range_monday(
                r.get('Promotion Start Date', 0),
                r.get('Promotion End Date', 0)
            )
        ),
        axis=1
    )

    logging.info("Finished claim enrichment.")
    return df

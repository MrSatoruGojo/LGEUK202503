# src/utils/lookup.py
"""
Lookup dictionary creation and lookup functions for SPMS data.
"""
from typing import Any, Dict, List
import logging

def create_lookup_dict(
    df,
    key_col: str,
    value_cols: List[str]
) -> Dict[Any, List[Dict[str, Any]]]:
    """
    Build a lookup dict mapping key_col to list of records (dicts of value_cols).

    Args:
        df: pandas DataFrame containing data
        key_col: column to use as dictionary key
        value_cols: list of columns to include in each record dict

    Returns:
        Dict mapping each key to a list of dicts
    """
        # Only use the columns actually present in df
    available = [col for col in value_cols if col in df.columns]
    missing = [col for col in value_cols if col not in df.columns]
    if missing:
        logging.warning(f"create_lookup_dict: missing columns {missing}")
    # Safely build the lookup dict
    try:
        grouped = (
            df.groupby(key_col)[available]
              .apply(lambda x: x.to_dict(orient='records'))
              .to_dict()
        )
    except Exception as e:
        logging.error(f"create_lookup_dict error: {e}")
        grouped = {}
    return grouped


def lookup_value(
    promo_no: Any,
    field_name: str,
    lookup_dict: Dict[Any, List[Dict[str, Any]]],
    default: Any = None
) -> Any:
    """
    Retrieve the first matching field value for a promotion number.
    """
    records = lookup_dict.get(promo_no, [])
    if not records:
        return default
    return records[0].get(field_name, default)


def lookup_customer(
    promo_no: Any,
    field_name: str,
    lookup_dict: Dict[Any, List[Dict[str, Any]]],
    customer_name: str = None,
    secondary_lookup_dict: Dict[Any, List[Dict[str, Any]]] = None,
    default: Any = None
) -> Any:
    """
    Lookup a field by promo_no and optional customer_name filter, falling back to secondary dict.
    """
    def filter_records(ld: Dict[Any, List[Dict[str, Any]]]):
        recs = ld.get(promo_no, [])
        if customer_name:
            recs = [r for r in recs if r.get('Bill To Name', '').startswith(customer_name)]
        return recs

    records = filter_records(lookup_dict)
    if not records and secondary_lookup_dict:
        records = filter_records(secondary_lookup_dict)
    if not records:
        return default
    return records[0].get(field_name, default)


def lookup_customer2(
    promo_no: Any,
    field_name: str,
    lookup_dict: Dict[Any, List[Dict[str, Any]]],
    product_code: str = None,
    secondary_lookup_dict: Dict[Any, List[Dict[str, Any]]] = None,
    default: Any = None
) -> Any:
    """
    Similar to lookup_customer but filters on 'Product Code'.
    """
    def filter_records(ld: Dict[Any, List[Dict[str, Any]]]):
        recs = ld.get(promo_no, [])
        if product_code:
            recs = [r for r in recs if r.get('Product Code') == product_code]
        return recs

    records = filter_records(lookup_dict)
    if not records and secondary_lookup_dict:
        records = filter_records(secondary_lookup_dict)
    if not records:
        return default
    return records[0].get(field_name, default)


def lookup_SPGM(
    promo_no: Any,
    bill_to_name: str,
    lookup_dict: Dict[Any, List[Dict[str, Any]]],
    secondary_lookup_dict: Dict[Any, List[Dict[str, Any]]] = None,
    default: Any = None
) -> Any:
    """
    Retrieve 'Sales PGM NO' by promo_no and bill_to_name prefix.
    """
    def filter_records(ld: Dict[Any, List[Dict[str, Any]]]):
        recs = ld.get(promo_no, [])
        return [r for r in recs if r.get('Bill To Name', '').startswith(bill_to_name)]

    records = filter_records(lookup_dict)
    if not records and secondary_lookup_dict:
        records = filter_records(secondary_lookup_dict)
    if not records:
        return default
    return records[0].get('Sales PGM NO', default)

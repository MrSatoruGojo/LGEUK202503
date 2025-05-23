# src/main.py
"""
Orchestrator script: reads inputs, runs each processing module, and saves outputs.
Now dynamically reads all INPUT_PARTS from config, supports multiple closed-orders
and tracker inputs, writes intermediates, a Log sheet, and per-BEBS splits.
"""

import logging
import os
import time
from datetime import datetime

import pandas as pd
from config.config import LOGGING_CONFIG, INPUT_PARTS, SHEET_CLAIM
import config.config as cfg

from pathlib import Path
import json

from src.io.file_ops import (
    read_data,
    write_excel_file,
    create_unique_folder,
    read_excel_file,
)
from src.utils.lookup import create_lookup_dict
from src.processing.claim import enrich_claim_data
from src.processing.orders import merge_closed_orders
from src.processing.psi import enrich_psi_data
from src.processing.tracker import enrich_tracker_data
from src.processing.output_splits import split_by_bebs
from src.output.formatter import format_workbook

def auto_read(path, sheet=None, header=None):
    """
    Load a single-sheet Excel/CSV/JSON file:
     - if sheet is None, read the *first* sheet (sheet_name=0)
     - if header is None, use header row = 0
    """
    ext = os.path.splitext(path)[1].lower()
    # determine pandas args
    sheet_arg  = sheet if sheet is not None else 0
    header_arg = header if header is not None else 0

    if ext in ('.xls', '.xlsx', '.xlsb'):
        return pd.read_excel(path, sheet_name=sheet_arg, header=header_arg)
    elif ext == '.csv':
        return pd.read_csv(path, header=header_arg)
    elif ext == '.json':
        return pd.read_json(path)
    else:
        raise ValueError(f"Unsupported extension {ext}")

def setup_logging():
    """Ensure log directory exists and configure logging."""
    log_dir = os.path.dirname(LOGGING_CONFIG['filename'])
    os.makedirs(log_dir, exist_ok=True)
    logging.basicConfig(**LOGGING_CONFIG)


def main(progress_callback=None):
    """
    Main pipeline with seven steps:
      1. Load Inputs
      2. Enrich Claim Data
      3. Merge Closed Orders
      4. Enrich PSI Data
      5. Enrich Tracker Data
      6. Write & Format Master + Log sheet
      7. Split per-BEBS
    Reports progress via progress_callback(name, step, elapsed_seconds).
    """

    # ─── Setup ────────────────────────────────────────────
    setup_logging()
    start_time = time.time()
    log_records = []



    # ─── Load GUI sheet/header config ─────────────────────────────
    gui_cfg = {}
    cfg_path = Path.home() / ".claim_verificator_config.json"
    if cfg_path.exists():
        try:
            gui_cfg = json.loads(cfg_path.read_text())
        except Exception:
            gui_cfg = {}


    # Build output folder
    folder_path = create_unique_folder('Log', path=cfg.OUTPUT_DIR)
    logging.info(f"Output folder: {folder_path}")

    # ─── Step 1: Load Inputs (sub-steps) ───────────────────
    required = {'CLAIM_FILE', 'SPMS_FILE', 'PSI_FILE', 'CLOSED_ORDERS_DIR'}
    inputs = {}
    substep = 1

    for key, label in INPUT_PARTS:
        # a) path + optional sheets override from GUI
        setting = gui_cfg.get(key, {})
        if isinstance(setting, dict):
            path   = setting.get('path')
            sheets = setting.get('sheets', {})
        else:
            path   = getattr(cfg, key)
            sheets = getattr(cfg, f"{key}_SHEETS", {}) or {}

        # b) required?
        if key in required and not path:
            raise ValueError(f"Missing required input: {label} ({key})")
        if not path:
            continue

        # c) SPMS special case: skip here, load below
        if key == 'SPMS_FILE':
            substep += 1
            continue

        # d) load: multi‐file → read_data, else Excel/CSV/JSON with sheet headers
        if isinstance(path, (list, tuple)):
            df = read_data(path)
        else:
            if sheets:
                frames = []
                for sheet_name, hdr in sheets.items():
                    frames.append(pd.read_excel(
                        path,
                        sheet_name=sheet_name,
                        header=hdr-1
                    ))
                df = pd.concat(frames, ignore_index=True)
            else:
                df = auto_read(path, sheet=None, header=None)

        # e) store + progress
        inputs[key] = df
        if progress_callback:
            elapsed = time.time() - start_time
            progress_callback(f"1.{substep} Load {label}",
                              substep, elapsed, df=df)
        substep += 1

    # Extract core DataFrames
    claim_df = inputs['CLAIM_FILE']
    psi_df   = inputs['PSI_FILE']
    
    # ─── SPECIAL-CASE SPMS: load both sheets via read_excel_file ───────────
    spms_df = read_excel_file(
        cfg.SPMS_FILE,
        sheet_name=cfg.SHEET_SPMS_MAIN,
        skiprows=3
    )
    spms2_df = read_excel_file(
        cfg.SPMS_FILE,
        sheet_name=cfg.SHEET_SPMS_SECONDARY,
        skiprows=2
    )
    


    # ─── Audit loaded column names ───────────────────────
    logging.info(f"CLAIM_FILE columns: {claim_df.columns.tolist()}")
    logging.info(f"SPMS_FILE columns:  {spms_df.columns.tolist()}")
    logging.info(f"PSI_FILE columns:   {psi_df.columns.tolist()}")


    # strip whitespace from headers so “Promotion No” matches
        # normalize header names: replace NBSP, BOM, unicode-normalize, then strip
    claim_df.columns = (
        claim_df.columns
           .str.replace('\u00A0', ' ')
           .str.replace('\ufeff', '')
           .str.normalize('NFKC')
           .str.strip()
    )
    spms_df.columns = (
        spms_df.columns
           .str.replace('\u00A0', ' ')
           .str.replace('\ufeff', '')
           .str.normalize('NFKC')
           .str.strip()
    )

    # (you can optionally preview the stripped-headers claim_df)
    if progress_callback:
        elapsed = time.time() - start_time
        progress_callback(f"1.{substep} Clean column names",
                          substep,
                          elapsed,
                          df=claim_df)
    substep += 1

        # normalize column names so 'Promotion No' really exists
    claim_df.columns = claim_df.columns.str.strip()
    spms_df.columns  = spms_df.columns.str.strip()


    # Collect closed-orders DataFrames (years 2020–2025)
    closed_list = inputs['CLOSED_ORDERS_DIR']  # list of paths

    # ─── Collect tracker DataFrames ─────────────────────
    tracker_parts = []
    for tk_key in ('TRACKER_HA','TRACKER_ID','TRACKER_CURRYS','TRACKER_JLP',
                   'TRACKER_P3_TV','TRACKER_P5_EU','TRACKER_P5_APAC',
                   'TRACKER_P6_TV','NEW_TRACKER_1'):
        if tk_key in inputs:
            tracker_parts.append(inputs[tk_key])

    # Concatenate into one DataFrame (so enrich_tracker_data() can index columns)
    if tracker_parts:
        tracker_df = pd.concat(tracker_parts, ignore_index=True)
    else:
        tracker_df = pd.DataFrame()
    # Log and save intermediate
    elapsed = time.time() - start_time
    log_records.append({
        "step": 1,
        "name": "Load Inputs",
        "added_columns": list(claim_df.columns),
        "rows": len(claim_df),
        "elapsed_seconds": round(elapsed, 2)
    })
    claim_df.to_excel(os.path.join(folder_path, "1_raw_claim.xlsx"), index=False)
    if progress_callback:
        progress_callback("Load Inputs", 1, elapsed)

       # ─── Step 2: Enrich Claim Data (sub-steps) ────────────
    # 2.1 Build SPMS lookup
    if progress_callback:
        elapsed = time.time() - start_time
        progress_callback(f"2.{substep} Build SPMS lookup",
                          substep,
                          elapsed,
                          df=spms_df)
    spms_map = create_lookup_dict(spms_df, 'Promotion No', cfg.SPMS_FIELDS)
    substep += 1

    # 2.2 Build secondary SPMS lookup
    if progress_callback:
        elapsed = time.time() - start_time
        progress_callback(f"2.{substep} Build SPMS2 lookup",
                          substep,
                          elapsed,
                          df=spms_df)
    spms2_map = create_lookup_dict(spms2_df, 'Promotion No', cfg.SPMS2_FIELDS)
    substep += 1

    # 2.3 Apply enrich_claim_data
    if progress_callback:
        elapsed = time.time() - start_time
        progress_callback(f"2.{substep} Enrich claim data",
                          substep,
                          elapsed,
                          df=claim_df)
    enriched_claim = enrich_claim_data(claim_df, spms_map, spms2_map)
    substep += 1

    # record & save
    new_cols = sorted(set(enriched_claim.columns) - set(claim_df.columns))
    elapsed = time.time() - start_time
    log_records.append({
        "step": 2,
        "name": "Enrich Claim Data",
        "added_columns": new_cols,
        "rows": len(enriched_claim),
        "elapsed_seconds": round(elapsed, 2)
    })
    enriched_claim.to_excel(os.path.join(folder_path, "2_enriched_claim.xlsx"),
                            index=False)

    # 2.4 Preview enriched_claim
    if progress_callback:
        progress_callback(f"2.{substep} Preview enriched claim",
                          substep,
                          elapsed,
                          df=enriched_claim)
    substep += 1

    # ─── Step 3: Merge Closed Orders ──────────────────────
    enriched_orders = merge_closed_orders(enriched_claim, closed_list)
    new_cols = sorted(set(enriched_orders.columns) - set(enriched_claim.columns))
    elapsed = time.time() - start_time
    log_records.append({
        "step": 3,
        "name": "Merge Closed Orders",
        "added_columns": new_cols,
        "rows": len(enriched_orders),
        "elapsed_seconds": round(elapsed, 2)
    })
    enriched_orders.to_excel(os.path.join(folder_path, "3_after_closed_orders.xlsx"), index=False)
    if progress_callback:
        progress_callback("Merge Closed Orders", 3, elapsed)

    # ─── Step 4: Enrich PSI Data ──────────────────────────
    enriched_psi = enrich_psi_data(enriched_orders, psi_df)
    new_cols = sorted(set(enriched_psi.columns) - set(enriched_orders.columns))
    elapsed = time.time() - start_time
    log_records.append({
        "step": 4,
        "name": "Enrich PSI Data",
        "added_columns": new_cols,
        "rows": len(enriched_psi),
        "elapsed_seconds": round(elapsed, 2)
    })
    enriched_psi.to_excel(os.path.join(folder_path, "4_after_psi.xlsx"), index=False)
    if progress_callback:
        progress_callback("Enrich PSI Data", 4, elapsed)

    # ─── Step 5: Enrich Tracker Data ─────────────────────
    # with
    final_df = enrich_tracker_data(enriched_psi, tracker_df)
    new_cols = sorted(set(final_df.columns) - set(enriched_psi.columns))
    elapsed = time.time() - start_time
    log_records.append({
        "step": 5,
        "name": "Enrich Tracker Data",
        "added_columns": new_cols,
        "rows": len(final_df),
        "elapsed_seconds": round(elapsed, 2)
    })
    final_df.to_excel(os.path.join(folder_path, "5_after_tracker.xlsx"), index=False)
    if progress_callback:
        progress_callback("Enrich Tracker Data", 5, elapsed)


# ─── Prep Part 5 DF ───────────────────────────────────
    # Grab the single “Part 5” tracker (may be empty)
    part5_df = inputs.get('TRACKER_P5_EU', pd.DataFrame())



    # ─── Step 6: Write & Format Master + Log Sheet ───────
    master_path = os.path.join(folder_path, "6_master_verified.xlsx")
    log_df = pd.DataFrame(log_records)
    with pd.ExcelWriter(master_path, engine="openpyxl") as writer:
        final_df.to_excel(writer, sheet_name="Verified", index=False)
        log_df.to_excel(writer, sheet_name="Log",      index=False)
    format_workbook(master_path)

    elapsed = time.time() - start_time
    log_records.append({
        "step": 6,
        "name": "Write & Format Master",
        "added_columns": [],
        "rows": len(final_df),
        "elapsed_seconds": round(elapsed, 2)
    })
    if progress_callback:
        progress_callback("Write & Format Master", 6, elapsed)
       # ─── Add Part 5 tracker to the master workbook ─────────
    if not part5_df.empty:
        with pd.ExcelWriter(master_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            part5_df.to_excel(writer, sheet_name='Part 5', index=False)
        # re-run formatting to pick up the new sheet
        format_workbook(master_path)
        # ─── Step 7: Split per-BEBS (Parts 6+5) ───────────────
    # Build Part 5 DataFrame (single file)
    part5_df = inputs.get('TRACKER_P5_EU', pd.DataFrame())# or whichever key you used for Part 5

    split_by_bebs(
        cleaned_df=final_df,
        tracker_part6_df=tracker_df,
        tracker_part5_df=part5_df,
        spms_df=spms_df,
        output_folder=folder_path
    )

    elapsed = time.time() - start_time
    logging.info(f"Processing completed in {elapsed:.1f}s")
    if progress_callback:
        progress_callback("Split per-BEBS", 7, elapsed)

    print(f"All done! Files saved in: {folder_path}")


if __name__ == "__main__":
    main()

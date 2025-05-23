# config/config.py

"""
Centralized configuration for file paths, constants, and logging settings.
"""
import os
import json
from datetime import datetime

# Determine project root
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# Data and output directories (can be overridden by GUI)
DATA_DIR   = os.getenv('DATA_DIR', os.path.join(BASE_DIR, 'data'))
OUTPUT_DIR = os.getenv('OUTPUT_DIR', os.path.join(BASE_DIR, 'output'))
LOG_DIR    = os.getenv('LOG_DIR', os.path.join(BASE_DIR, 'logs'))

# ─── INPUT PARTS ──────────────────────────────────────────────────────────────
# Each tuple is (CONFIG_KEY, Label for GUI)
# ─── INPUT PARTS ──────────────────────────────────────────────────────────────
INPUT_PARTS = [
    ('CLAIM_FILE',        'Claim File'),
    ('SPMS_FILE',         'SPMS Files'),
    ('PSI_FILE',          'PSI File'),
    ('CLOSED_ORDERS_DIR', 'Closed Orders'),
    ('TRACKER_HA',        'Tracker HA'),
    ('TRACKER_ID',        'Tracker ID'),
    ('TRACKER_CURRYS',    'Part 2 Tracker TV Currys'),
    ('TRACKER_JLP',       'Part 2 Tracker TV JLP'),
    ('TRACKER_P3_TV',     'Part 3 Tracker TV'),
    ('TRACKER_P5',        'Part 5 Tracker '),
    ('TRACKER_P6_TV',     'Part 6 Tracker TV'),
    ('NEW_TRACKER_1',     'New Tracker 1'),
]
# ────────────────────────────────────────────────────────────────────────────────

# Default input file names (overridden by GUI)
CLAIM_FILE         = os.getenv('CLAIM_FILE',         os.path.join(DATA_DIR, 'claim.xlsx'))
SPMS_FILE          = os.getenv('SPMS_FILE',          os.path.join(DATA_DIR, 'SPMS.xlsx'))
PSI_FILE           = os.getenv('PSI_FILE',           os.path.join(DATA_DIR, 'PSI.xlsx'))
OLD_TRACKER_FILE   = os.getenv('OLD_TRACKER_FILE',   os.path.join(DATA_DIR, 'old_tracker.xlsx'))

# For closed‐orders and trackers we default to None; GUI will override these
CLOSED_ORDERS_2020 = None
CLOSED_ORDERS_2021 = None
CLOSED_ORDERS_2022 = None
CLOSED_ORDERS_2023 = None
CLOSED_ORDERS_2024 = None
CLOSED_ORDERS_2025 = None

TRACKER_HA       = None
TRACKER_ID       = None
TRACKER_CURRYS   = None
TRACKER_JLP      = None
TRACKER_P3_TV    = None
TRACKER_P5       = None
TRACKER_P6_TV    = None
NEW_TRACKER_1    = None

# SPMS fields used for lookups
SPMS_FIELDS = [
    'Promotion Start YYYYMMDD',
    'Promotion End YYYYMMDD',
    'Promotion Name',
    'Promotion Status Code',
    'Cancel Flag',
    'Recreate Flag',
    'Original Promotion No',
    'Sales PGM NO',
    'Sales PGM Status',
    'Promotion Property',
    'Alloc Div Code',
    'Apply Month_YYYYMM',
    'Bill To Name',
    'Claim Line Flag',
    'Customer Code',
    'Division Code',
    'Product Code',
    'Expected Qty',
    'Dc Operand',
    'Expected Cost'
]
SPMS2_FIELDS = [
    'Sales PGM NO',
    'Bill To Name',
    'Product Code',
    'Expected Qty',
    'Dc Operand',
    'Expected Cost'
]

# Excel sheet names
SHEET_CLAIM          = 'CLAIM'
SHEET_SPMS_MAIN      = 'Report 1'
SHEET_SPMS_SECONDARY = 'Report 7'
SHEET_PSI            = None  # default to first sheet
SHEET_OLDTRACKER     = None  # default to first sheet

# Filename patterns for closed-orders (if you ever scan a dir)
CLOSED_ORDER_PATTERNS = ['*.xls', '*.xlsx', '*.xlsb']

# Logging configuration
LOGGING_CONFIG = {
    'level':    os.getenv('LOG_LEVEL', 'INFO'),
    'filename': os.path.join(
                    LOG_DIR,
                    f"process_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
                ),
    'filemode': 'w',
    'format':   '%(asctime)s - %(levelname)s - %(message)s'
}

# Frontend defaults (if you extend to a web UI)
FRONTEND_CONFIG = {
    'host': '127.0.0.1',
    'port': 8050,
    'debug': True
}

# Filename template for export
EXPORT_FILENAME_TEMPLATE = 'claim_processed_{timestamp}.xlsx'

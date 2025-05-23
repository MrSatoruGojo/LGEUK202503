import argparse
import sys
import config.config as cfg
from src.main import main

def parse_args():
    parser = argparse.ArgumentParser(
        description="Run the Claim Verificator pipeline end-to-end."
    )
    parser.add_argument(
        "--claim",
               help="Path to Claim Excel file",
        default=cfg.CLAIM_FILE,
    )
    parser.add_argument(
        "--spms",
        help="Path to SPMS Excel file",
        default=cfg.SPMS_FILE,
    )
    parser.add_argument(
        "--psi",
        help="Path to PSI Excel file",
        default=cfg.PSI_FILE,
    )
    parser.add_argument(
        "--tracker",
        help="Path to old Tracker Excel file",
        default=cfg.OLD_TRACKER_FILE,
    )
    parser.add_argument(
        "--closed",
        nargs="+",
        help="One or more closed-orders Excel files",
        default=None,
    )
    parser.add_argument(
        "--outdir",
        help="Directory where timestamped output folders will be created",
        default=cfg.OUTPUT_DIR,
    )
    return parser.parse_args()

def run_cli():
    args = parse_args()

    # Override the config module values
    cfg.CLAIM_FILE          = args.claim
    cfg.SPMS_FILE           = args.spms
    cfg.PSI_FILE            = args.psi
    cfg.OLD_TRACKER_FILE    = args.tracker
    cfg.OUTPUT_DIR          = args.outdir

    # If the user passed closed-order files on CLI, use them
    if args.closed:
        cfg.CLOSED_ORDERS_DIR = args.closed

    # Now call the main pipeline
    try:
        main()
    except Exception as e:
        print(f"Pipeline failed: {e}", file=sys.stderr)
        sys.exit(1)

if __name__ == "__main__":
    run_cli()

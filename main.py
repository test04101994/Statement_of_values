"""
CLI entry point for the SOV Excel parser.

Usage:
    python main.py <sov_file.xlsx>
    python main.py <sov_file.xlsx> --profile-arn arn:aws:bedrock:...
    python main.py <sov_file.xlsx> --sheets 0,2
    python main.py <sov_file.xlsx> --config my_config.yaml
"""

from __future__ import annotations

import argparse
import json
import logging
import sys
import warnings
from pathlib import Path

from sov.engine import configure_logging, parse_sov_file

warnings.filterwarnings("ignore")
logger = logging.getLogger(__name__)

# All relative paths resolve from the script's own directory
SCRIPT_DIR = Path(__file__).resolve().parent


def _resolve(path_str: str) -> str:
    """Resolve a path relative to the script directory."""
    p = Path(path_str)
    if not p.is_absolute():
        p = SCRIPT_DIR / p
    return str(p)


def _parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Parse a property SOV Excel file.")
    p.add_argument("filepath", nargs="?", default="sample_sov.xlsx", help=".xlsx or .xls file")
    p.add_argument("--config", default=None, help="Path to config.yaml")
    p.add_argument("--profile-arn", dest="inference_profile_arn", help="Bedrock inference profile ARN")
    p.add_argument("--region", dest="aws_region", help="AWS region")
    p.add_argument("--aws-profile", dest="aws_profile", help="Named AWS CLI profile")
    p.add_argument("--sheets", default=None, help="Comma-separated sheet indices or names")
    return p.parse_args(argv)


def main(argv: list[str] | None = None) -> None:
    configure_logging(level=logging.INFO)
    args = _parse_args(argv)

    filepath = _resolve(args.filepath)
    config_path = _resolve(args.config) if args.config else None

    overrides: dict = {}
    if args.inference_profile_arn:
        overrides["inference_profile_arn"] = args.inference_profile_arn
    if args.aws_region:
        overrides["aws_region"] = args.aws_region
    if args.aws_profile:
        overrides["aws_profile"] = args.aws_profile
    if args.sheets:
        parts = [s.strip() for s in args.sheets.split(",")]
        overrides["sheets"] = [int(s) if s.isdigit() else s for s in parts]

    df, metadata = parse_sov_file(
        filepath,
        config_path=config_path,
        config_overrides=overrides or None,
    )

    logger.info("Preview (first 10 rows):\n%s", df.head(10).to_string(index=False))

    # --- Save outputs next to the input file ---
    base = filepath.rsplit(".", 1)[0]

    out_xlsx = base + "_cleaned.xlsx"
    df.to_excel(out_xlsx, index=False)
    logger.info("Saved cleaned data : %s", out_xlsx)

    out_json = base + "_llm_responses.json"
    with open(out_json, "w", encoding="utf-8") as f:
        json.dump(metadata, f, indent=2, default=str)
    logger.info("Saved LLM responses: %s", out_json)


if __name__ == "__main__":
    main()

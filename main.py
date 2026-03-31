"""
CLI entry point for the SOV Excel parser.

Usage:
    python main.py <file_or_folder>
    python main.py /path/to/folder              ← processes all .xlsx/.xls files in folder
    python main.py sov_file.xlsx                 ← processes single file
    python main.py sov_file.xlsx --profile-arn arn:aws:bedrock:...
    python main.py /path/to/folder --sheets 0,2
"""

from __future__ import annotations

import argparse
import json
import logging
import sys
import warnings
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

from sov.engine import configure_logging, parse_sov_file

warnings.filterwarnings("ignore")
logger = logging.getLogger(__name__)

SCRIPT_DIR = Path(__file__).resolve().parent


def _resolve(path_str: str) -> str:
    """Resolve a path relative to this script's directory when not absolute."""
    p = Path(path_str)
    if not p.is_absolute():
        p = SCRIPT_DIR / p
    return str(p)


def _parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    """Parse CLI arguments; ``argv`` defaults to ``sys.argv[1:]`` when omitted."""
    p = argparse.ArgumentParser(description="Parse property SOV Excel file(s).")
    p.add_argument(
        "path", nargs="?", default=".", help=".xlsx/.xls file OR folder of files"
    )
    p.add_argument("--config", default=None, help="Path to config.yaml")
    p.add_argument(
        "--profile-arn",
        dest="inference_profile_arn",
        help="Bedrock inference profile ARN",
    )
    p.add_argument("--region", dest="aws_region", help="AWS region")
    p.add_argument("--aws-profile", dest="aws_profile", help="Named AWS CLI profile")
    p.add_argument("--rebuild-embeddings", action="store_true", help="Force rebuild embedding cache from config synonyms")
    p.add_argument(
        "--sheets", default=None, help="Comma-separated sheet indices or names"
    )
    return p.parse_args(argv)


def _collect_files(path: str) -> list[Path]:
    """Return list of .xlsx/.xls files from a path (file or directory)."""
    p = Path(path)
    if p.is_file():
        return [p]
    if p.is_dir():
        files = sorted(
            f
            for f in p.iterdir()
            if f.suffix.lower() in (".xlsx", ".xls")
            and not f.name.startswith("~$")  # skip Excel temp files
        )
        return files
    raise FileNotFoundError(f"Path not found: {path}")


def _append_sheets_to_source(
    source_path: str,
    df: pd.DataFrame,
    df_sources: pd.DataFrame,
) -> None:
    """Append ``Cleaned Data`` and ``Source References`` sheets to the original Excel file."""
    wb = load_workbook(source_path)

    # Remove existing output sheets if re-running
    for name in ("Cleaned Data", "Source References"):
        if name in wb.sheetnames:
            del wb[name]

    wb.save(source_path)

    # Now append using ExcelWriter in append mode
    with pd.ExcelWriter(
        source_path, engine="openpyxl", mode="a", if_sheet_exists="replace"
    ) as writer:
        df.to_excel(writer, sheet_name="Cleaned Data", index=False)
        df_sources.to_excel(writer, sheet_name="Source References", index=False)


def process_file(
    filepath: str,
    config_path: str | None,
    overrides: dict | None,
    output_dir: str | None = None,
) -> bool:
    """Process a single SOV file. Returns True on success."""
    logger.info("=" * 70)
    logger.info("Processing: %s", filepath)
    logger.info("=" * 70)

    try:
        df, df_sources, metadata = parse_sov_file(
            filepath,
            config_path=config_path,
            config_overrides=overrides,
        )

        if df.empty:
            logger.warning("No data extracted from %s", filepath)
            return False

        logger.info("Extracted %d rows, %d columns", len(df), len(df.columns))

        # Append output sheets to the original file
        _append_sheets_to_source(filepath, df, df_sources)
        logger.info("Added sheets to source: %s", filepath)

        # Save LLM responses JSON to output directory
        fname = Path(filepath).stem + "_llm_responses.json"
        out_json = (
            str(Path(output_dir) / fname)
            if output_dir
            else filepath.rsplit(".", 1)[0] + "_llm_responses.json"
        )
        Path(out_json).parent.mkdir(parents=True, exist_ok=True)
        with open(out_json, "w", encoding="utf-8") as f:
            json.dump(metadata, f, indent=2, default=str)
        logger.info("Saved LLM responses: %s", out_json)

        return True

    except Exception as exc:  # pylint: disable=broad-exception-caught
        # SOV pipeline may fail for I/O, AWS, or parsing reasons; log and continue batch.
        logger.error("Failed to process %s: %s", filepath, exc, exc_info=True)
        return False


def main(argv: list[str] | None = None) -> None:
    """Run the CLI: parse files, append sheets, and write LLM JSON metadata to ``output/``."""
    configure_logging(level=logging.INFO)
    args = _parse_args(argv)

    target = _resolve(args.path)
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

    # Rebuild embeddings if explicitly asked, then exit
    if args.rebuild_embeddings:
        from sov.engine import load_config, rebuild_embedding_cache
        cfg = load_config(config_path)
        rebuild_embedding_cache(cfg)
        logger.info("Embedding cache rebuilt. Run again without --rebuild-embeddings to process files.")
        return

    files = _collect_files(target)
    if not files:
        logger.error("No .xlsx/.xls files found in %s", target)
        sys.exit(1)

    # JSON output goes to output/ subfolder next to the input files
    input_dir = Path(target) if Path(target).is_dir() else Path(target).parent
    output_dir = str(input_dir / "output")
    Path(output_dir).mkdir(exist_ok=True)
    logger.info(
        "Found %d file(s) to process. JSON output → %s/", len(files), output_dir
    )

    success = 0
    failed = 0
    for f in files:
        ok = process_file(str(f), config_path, overrides or None, output_dir=output_dir)
        if ok:
            success += 1
        else:
            failed += 1

    logger.info("")
    logger.info("Done: %d succeeded, %d failed, %d total", success, failed, len(files))


if __name__ == "__main__":
    main()

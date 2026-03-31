"""
SOV (Statement of Values) package.

Extracts normalized property data from messy SOV spreadsheets using Claude on
AWS Bedrock. Configuration lives in ``config.yaml`` at the project root.
"""

from sov.engine import load_config, parse_sov_file

__all__ = ["parse_sov_file", "load_config"]

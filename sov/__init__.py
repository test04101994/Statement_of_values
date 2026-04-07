"""
SOV (Statement of Values) package.

Extracts normalized property data from messy SOV spreadsheets using Claude on
AWS Bedrock. Field definitions and model settings live in ``config/config.yaml``.

Public API
----------
``parse_sov_file``
    End-to-end parse of one or more workbook sheets into DataFrames.
``load_config``
    Load and normalize the YAML configuration (schema, synonyms, Bedrock).
"""

from sov.engine import load_config, parse_sov_file

__all__ = ["parse_sov_file", "load_config"]

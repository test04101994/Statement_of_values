"""Unit tests for SOV engine helpers (smoke tests).

Run from the project root::

    python -m unittest discover -s tests -v
"""

from __future__ import annotations

import unittest
from pathlib import Path

from sov.engine import _clean_cell, _normalize_header, load_config


class TestNormalizeHeader(unittest.TestCase):
    """Tests for :func:`sov.engine._normalize_header`."""

    def test_collapses_whitespace(self) -> None:
        """Lowercases and collapses internal spaces."""
        self.assertEqual(_normalize_header("  Foo   BAR  "), "foo bar")


class TestCleanCell(unittest.TestCase):
    """Tests for :func:`sov.engine._clean_cell`."""

    def test_amount_strips_currency(self) -> None:
        """Amount type parses dollar-formatted strings."""
        self.assertEqual(_clean_cell("$1,234.50", "amount"), 1234.5)

    def test_boolean_yes(self) -> None:
        """Boolean type maps common affirmatives to Yes."""
        self.assertEqual(_clean_cell("Y", "boolean"), "Yes")


class TestLoadConfig(unittest.TestCase):
    """Tests for :func:`sov.engine.load_config`."""

    def test_missing_file_raises(self) -> None:
        """A non-existent path raises ``FileNotFoundError``."""
        bad = Path(__file__).resolve().parent / "nonexistent_config_xyz.yaml"
        with self.assertRaises(FileNotFoundError):
            load_config(bad)

    def test_default_config_loads(self) -> None:
        """Project ``config/config.yaml`` loads with a non-empty ``fields`` map."""
        cfg = load_config()
        self.assertIsInstance(cfg.get("fields"), dict)
        self.assertGreaterEqual(len(cfg["fields"]), 1)


if __name__ == "__main__":
    unittest.main()

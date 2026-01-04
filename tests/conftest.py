"""Shared pytest fixtures."""

import tempfile
from pathlib import Path
from unittest.mock import patch

import pytest

from lotus123.core import Spreadsheet
from lotus123.formula import FormulaParser


@pytest.fixture(autouse=True)
def isolate_config():
    """Redirect config to temp directory so tests don't affect user's real config."""
    with tempfile.TemporaryDirectory() as tmpdir:
        config_dir = Path(tmpdir) / "lotus123"
        config_file = config_dir / "config.json"
        with (
            patch("lotus123.ui.config.CONFIG_DIR", config_dir),
            patch("lotus123.ui.config.CONFIG_FILE", config_file),
        ):
            yield


@pytest.fixture
def spreadsheet() -> Spreadsheet:
    return Spreadsheet()


@pytest.fixture
def parser(spreadsheet: Spreadsheet) -> FormulaParser:
    return FormulaParser(spreadsheet)

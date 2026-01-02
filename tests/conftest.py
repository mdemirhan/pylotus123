"""Shared pytest fixtures."""

import pytest

from lotus123.core import Spreadsheet
from lotus123.formula import FormulaParser


@pytest.fixture
def spreadsheet() -> Spreadsheet:
    return Spreadsheet()


@pytest.fixture
def parser(spreadsheet: Spreadsheet) -> FormulaParser:
    return FormulaParser(spreadsheet)

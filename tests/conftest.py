"""Shared pytest fixtures for agent-xlsx security tests."""

import pytest
from openpyxl import Workbook


@pytest.fixture
def sample_xlsx(tmp_path):
    """Minimal .xlsx with known cell content for command output tests."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"] = "header"
    ws["A2"] = "value1"
    ws["B1"] = "amount"
    ws["B2"] = 100
    p = tmp_path / "sample.xlsx"
    wb.save(p)
    return p


@pytest.fixture
def sample_xlsm(tmp_path):
    """Minimal macro-enabled .xlsm for VBA gate tests (no actual VBA content needed)."""
    wb = Workbook()
    ws = wb.active
    ws["A1"] = "test"
    p = tmp_path / "sample.xlsm"
    wb.save(p)
    return p

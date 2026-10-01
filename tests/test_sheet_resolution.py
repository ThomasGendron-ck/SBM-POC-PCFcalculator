"""Tests for tolerant Excel sheet-name resolution in io_sbm.load_sheet."""
import pandas as pd
import pytest

from sbm_pcf_calculation.io_sbm import load_sheet, resolve_sheet_name


@pytest.fixture
def workbook(tmp_path):
    path = tmp_path / "wb.xlsx"
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        pd.DataFrame({"A": [1]}).to_excel(writer, sheet_name="MasterBase_Products", index=False)
        pd.DataFrame({"B": [2]}).to_excel(writer, sheet_name="other sheet", index=False)
    return path


def test_exact_match(workbook):
    assert resolve_sheet_name(str(workbook), "MasterBase_Products") == "MasterBase_Products"


def test_case_and_space_tolerant(workbook):
    assert resolve_sheet_name(str(workbook), "masterbase products") == "MasterBase_Products"


def test_underscore_tolerant(workbook):
    assert resolve_sheet_name(str(workbook), "MASTERBASEProducts") == "MasterBase_Products"


def test_missing_sheet_lists_available(workbook):
    with pytest.raises(ValueError) as exc:
        resolve_sheet_name(str(workbook), "nope")
    assert "nope" in str(exc.value) and "other sheet" in str(exc.value)


def test_load_sheet_uses_resolved_name(workbook):
    df = load_sheet(str(workbook), "MASTERBASE Products", 1)
    assert list(df.columns) == ["A"]

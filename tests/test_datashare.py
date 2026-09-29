"""Tests de la mise à jour v0.92 du fichier de collecte LM (datashare)."""

from pathlib import Path

from openpyxl import load_workbook

from pcf_extraction.datashare import (
    BLOCK_COLORS_V092,
    DQR_GUIDE_SHEET,
    apply_v092,
    block_of,
)

INPUT = Path(__file__).resolve().parents[1] / "input" / "Data_collection_v091.xlsx"


def _make_workbook(tmp_path):
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb["Sheet"]
    ws.title = "Product"
    headers = [
        "Product SKU",
        "Product Details",
        "Supplier PCF value",
        "Supplier PDS",
        "Supplier PCF",
        "Transformation Process Name",
        "Transformation DQR value",
        "Data validation flag",
        "Data validation flag impact",
    ]
    for col, h in enumerate(headers, start=1):
        ws.cell(row=1, column=col, value=h)
    ug = wb.create_sheet("UserGuide")
    ug.cell(row=1, column=2, value="SBM - PCF - LM references - Data collection")
    path = tmp_path / "in.xlsx"
    wb.save(path)
    return path


def test_block_of():
    assert block_of("Product SKU", "Product") == "Produit"
    assert block_of("Transformation GHG", "Product") == "Transformation"
    assert block_of("Supplier PCF value", "Component") == "Supplier PCF"
    assert block_of("Supplier DQR", "Component") == "Supplier PCF"
    assert block_of("Data validation flag", "Component") == "Flag"
    assert block_of("Component SKU", "Component") == "Composants"
    assert block_of("Net Weight", "Component") == "Description du composant"
    assert block_of("Raw Material (MB Product)", "Component") == "Description du composant"


def test_colors_and_sheets(tmp_path):
    in_path = _make_workbook(tmp_path)
    out_path = tmp_path / "out.xlsx"
    apply_v092(in_path, out_path)

    wb = load_workbook(out_path)
    assert "Product" in wb.sheetnames
    assert DQR_GUIDE_SHEET in wb.sheetnames
    assert "UserGuide" in wb.sheetnames

    ws = wb["Product"]
    by_header = {c.value: c for c in ws[1]}
    assert by_header["Product SKU"].fill.fgColor.rgb == BLOCK_COLORS_V092["Produit"]
    assert by_header["Supplier PCF value"].fill.fgColor.rgb == BLOCK_COLORS_V092["Supplier PCF"]
    assert by_header["Transformation DQR value"].fill.fgColor.rgb == BLOCK_COLORS_V092["Transformation"]
    assert by_header["Data validation flag"].fill.fgColor.rgb == BLOCK_COLORS_V092["Flag"]
    # marqueurs GROUPING conservés en gris
    assert by_header["Product Details"].fill.fgColor.rgb == "FFD9D9D9"
    assert by_header["Supplier PCF"].fill.fgColor.rgb == "FFD9D9D9"

    dqr = wb[DQR_GUIDE_SHEET]
    values = {c.value for row in dqr.iter_rows() for c in row if c.value}
    assert "PACT Methodology 3.0 - Page 66" in values
    assert "Technological representativeness (TECH DQR)" in values
    assert "Overall DQR evaluation" in values
    assert "Excellent quality" in values
    assert "Poor quality" in values

    ug = wb["UserGuide"]
    texts = [c.value for row in ug.iter_rows() for c in row if c.value]
    assert any("How to collect the data from suppliers" in str(t) for t in texts)
    assert any("Do you hold the PCF of [SKU]?" in str(t) for t in texts)
    assert any("Transformation Process Scrap Rate" in str(t) for t in texts)
    assert any("Risks associated with data estimation" in str(t) for t in texts)


def test_real_file_if_available(tmp_path):
    if not INPUT.is_file():
        return
    out_path = tmp_path / "out.xlsx"
    apply_v092(INPUT, out_path)
    wb = load_workbook(out_path)
    assert wb.sheetnames == ["UserGuide", "Product", "Component", DQR_GUIDE_SHEET]
    comp = wb["Component"]
    by_header = {c.value: c for c in comp[1]}
    assert by_header["Component SKU"].fill.fgColor.rgb == BLOCK_COLORS_V092["Composants"]
    assert by_header["Net Weight"].fill.fgColor.rgb == BLOCK_COLORS_V092["Description du composant"]
    assert by_header["Supplier PDS"].fill.fgColor.rgb == BLOCK_COLORS_V092["Supplier PCF"]
    assert by_header["Transformation Process Name"].fill.fgColor.rgb == BLOCK_COLORS_V092["Transformation"]
    assert by_header["Component Details"].fill.fgColor.rgb == "FFD9D9D9"
    prod = wb["Product"]
    by_header_p = {c.value: c for c in prod[1]}
    assert by_header_p["Product SKU"].fill.fgColor.rgb == BLOCK_COLORS_V092["Produit"]
    assert by_header_p["Product SKU"].font.color.rgb == "FFFFFFFF"

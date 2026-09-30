"""Construction du fichier de collecte LM selon les spécifications v0.94.

Source : « POC Calculateur - Spécifications - v0.93.xlsx » (Spec_CollectionFile).
- onglets Product (50 colonnes) et Component (53 colonnes) reconstruits selon la
  spec v0.93 : blocs Supplier PCF (framework/scope/declared unit ajoutés) et
  Transformation (declared unit, location, Energy Type/Consumption, Comment
  ajoutés, champs Energy avant Process) ;
- formules Excel explicites (Transformation DQR value, Transformation GHG,
  Data validation flag) conformes aux règles de validation de la spec ;
- listes déroulantes (framework, scope, external review, DQR 1-5, Energy Type) ;
- couleurs des en-têtes par bloc (identiques à la maquette « Fichier de
  collecte » de la spec, inchangée entre v0.92 et v0.93) ;
- onglets UserGuide et DQR_Guide repris tels quels du template validé par SBM
  (contenu, couleurs de cellules et polices, hauteurs de lignes, largeurs de
  colonnes, fusions de cellules, bordures) ;
- ordre des onglets : UserGuide, Product, Component, DQR_Guide.
"""

from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

HEADER_ROW = 2
MANDATORY_ROW = 1
DATA_START_ROW = 3
DATA_END_ROW = 202
DEFAULT_TEMPLATE = Path(__file__).resolve().parents[2] / "templates" / "Data_collection_template.xlsx"

GROUPING = "GROUPING"
CALCULATED_FIELDS = {"Transformation DQR value", "Transformation GHG"}
CALCULATED_COLOR = "FFA6A6A6"

BLOCK_COLORS: dict[str, str] = {
    "Produit": "FF13501B",
    "Composants": "FF75A67C",
    "Description du composant": "FFA3C4A7",
    "Supplier PCF": "FFFBE3D6",
    "Transformation": "FFF6C6AD",
    "Calculé": CALCULATED_COLOR,
    "Flag": "FFFFC000",
    GROUPING: "FFD9D9D9",
}
BLOCK_FONT_COLORS: dict[str, str] = {
    "Produit": "FFFFFFFF",
    "Calculé": "FF000000",
    "Composants": "FF000000",
    "Description du composant": "FF000000",
    "Supplier PCF": "FF000000",
    "Transformation": "FF000000",
    "Flag": "FF000000",
    GROUPING: "FF404040",
}
BLOCK_COLORS_V092 = BLOCK_COLORS
BLOCK_FONT_COLORS_V092 = BLOCK_FONT_COLORS

# Onglet Component : colonnes descriptives -> bloc « Description du composant ».
COMPONENT_DESCRIPTION_FIELDS = {
    "Pack unit box",
    "Net Weight",
    "Net Weight Unit",
    "Gross Weight",
    "Gross Weight Unit",
    "Stock unit",
    "Raw Material (MB Product)",
    "Raw Material - Carbon Footprint",
    "UVP description",
}

SUPPLIER_PCF_FIELDS = [
    "Supplier PCF framework",
    "Supplier PCF scope",
    "Supplier PCF declared unit",
    "Supplier PCF value",
    "Supplier PCF Unit",
    "Supplier PDS",
    "Supplier DQR",
    "Supplier PCF date",
    "Supplier PCF source",
    "Supplier PCF external review",
]
TRANSFORMATION_FIELDS = [
    "Transformation Process declared unit",
    "Transformation Process location",
    "Transformation Energy Type",
    "Transformation Energy Consumption",
    "Transformation Energy Unit",
    "Transformation Energy EF Name",
    "Transformation Energy EF Value",
    "Transformation Energy EF Unit",
    "Transformation Energy EF Source",
    "Transformation Process Name",
    "Transformation Process EF Name",
    "Transformation Process EF Value",
    "Transformation Process EF Unit",
    "Transformation Process EF Source",
    "Transformation Process Scrap Rate",
    "Transformation Process Comment",
    "Transformation PDS value",
    "Transformation GEO DQR",
    "Transformation TECH DQR",
    "Transformation TEMP DQR",
    "Transformation DQR value",
    "Transformation GHG",
    "Transformation GHG Unit",
]
FLAG_FIELDS = ["Data validation flag", "Data validation flag impact"]

# Mandatory field (spec v0.94, Spec_CollectionFile colonne I) par onglet.
MANDATORY: dict[str, dict[str, str]] = {
    "Product": {
        "Product SKU": "Yes",
        "Product Designation": "Yes",
        "Category Code": "Yes",
        "Category description": "Yes",
        "Supplier Code": "Yes",
        "Supplier Name": "Yes",
        "Net Weight": "Yes",
        "Net Weight Unit": "Yes",
        "Gross Weight": "Yes",
        "Gross Weight Unit": "Yes",
        "Stock unit": "Yes",
        "Supplier PCF framework": "Yes, if PCF section filled",
        "Supplier PCF scope": "Yes, if PCF section filled",
        "Supplier PCF declared unit": "Yes, if PCF section filled",
        "Supplier PCF value": "Yes, if PCF section filled",
        "Supplier PCF Unit": "Yes, if PCF section filled",
        "Supplier PDS": "Yes, if PCF section filled",
        "Supplier PCF date": "Yes, if PCF section filled",
        "Supplier PCF source": "Yes, if PCF section filled",
        "Supplier PCF external review": "Yes, if PCF section filled",
        "Transformation Process declared unit": "Yes, if transformation section filled",
        "Transformation Process location": "Yes, if transformation section filled",
        "Transformation Energy Type": "Yes, if energy section filled",
        "Transformation Energy Consumption": "Yes, if energy section filled",
        "Transformation Energy Unit": "Yes, if energy section filled",
        "Transformation Process Name": "Yes, if process section filled",
        "Transformation Process Scrap Rate": "Yes, if process section filled",
        "Transformation DQR value": "Automatic",
        "Transformation GHG": "Automatic",
        "Transformation GHG Unit": "Yes, if transformation section filled",
        "Data validation flag": "Automatic",
        "Data validation flag impact": "Automatic",
    },
    "Component": {
        "Component SKU": "Yes",
        "Component Designation": "Yes",
        "Category Code": "Yes",
        "Category description": "Yes",
        "Carbon category": "Yes",
        "Carbon sub category": "Yes",
        "Supplier Code": "Yes",
        "Supplier Name": "Yes",
        "UVP description": "Yes",
        "Net Weight": "Yes",
        "Net Weight Unit": "Yes",
        "Gross Weight": "Yes",
        "Gross Weight Unit": "Yes",
        "Stock unit": "Yes",
        "Supplier PCF framework": "Yes, if PCF section filled",
        "Supplier PCF value": "Yes, if PCF section filled",
        "Supplier PCF Unit": "Yes, if PCF section filled",
        "Supplier PDS": "Yes, if PCF section filled",
        "Supplier PCF date": "Yes, if PCF section filled",
        "Supplier PCF source": "Yes, if PCF section filled",
        "Supplier PCF external review": "Yes, if PCF section filled",
        "Transformation Process declared unit": "Yes, if transformation section filled",
        "Transformation Process location": "Yes, if transformation section filled",
        "Transformation Energy Type": "Yes, if energy section filled",
        "Transformation Energy Consumption": "Yes, if energy section filled",
        "Transformation Energy Unit": "Yes, if energy section filled",
        "Transformation Process Name": "Yes, if process section filled",
        "Transformation Process Scrap Rate": "Yes, if process section filled",
        "Transformation DQR value": "Automatic",
        "Transformation GHG": "Automatic",
        "Transformation GHG Unit": "Yes, if transformation section filled",
        "Data validation flag": "Automatic",
        "Data validation flag impact": "Automatic",
    },
}


# Groupings Plan (spec v0.94, Spec_CollectionFile) : (première, dernière) colonne
# du groupe, bornes incluses, par onglet.
GROUPING_RANGES: dict[str, list[tuple[str, str]]] = {
    "Product": [
        ("Category Code", "Stock unit"),
        ("Supplier PDS", "Supplier PCF external review"),
        ("Transformation Process declared unit", "Transformation DQR value"),
    ],
    "Component": [
        ("Category Code", "Stock unit"),
        ("Supplier PDS", "Supplier PCF external review"),
        ("Transformation Process declared unit", "Transformation DQR value"),
    ],
}

# Colonnes par onglet, dans l'ordre de la spec v0.93 (Spec_CollectionFile).
PRODUCT_COLUMNS: list[tuple[str, str]] = (
    [
        ("Product SKU", "Produit"),
        ("Product Designation", "Produit"),
        ("Category Code", "Produit"),
        ("Category description", "Produit"),
        ("Supplier Code", "Produit"),
        ("Supplier Name", "Produit"),
        ("Pack Unit Box", "Produit"),
        ("Net Weight", "Produit"),
        ("Net Weight Unit", "Produit"),
        ("Gross Weight", "Produit"),
        ("Gross Weight Unit", "Produit"),
        ("Stock unit", "Produit"),
        ("Product Details", GROUPING),
    ]
    + [(h, "Supplier PCF") for h in SUPPLIER_PCF_FIELDS]
    + [("Supplier PCF", GROUPING)]
    + [(h, "Transformation") for h in TRANSFORMATION_FIELDS]
    + [("Transformation Details", GROUPING)]
    + [(h, "Flag") for h in FLAG_FIELDS]
)
COMPONENT_COLUMNS: list[tuple[str, str]] = (
    [
        ("Component SKU", "Composants"),
        ("Component Designation", "Composants"),
        ("Category Code", "Composants"),
        ("Category description", "Composants"),
        ("Carbon category", "Composants"),
        ("Carbon sub category", "Composants"),
        ("Supplier Code", "Composants"),
        ("Supplier Name", "Composants"),
        ("UVP description", "Description du composant"),
        ("Raw Material (MB Product)", "Description du composant"),
        ("Raw Material - Carbon Footprint", "Description du composant"),
        ("Pack unit box", "Description du composant"),
        ("Net Weight", "Description du composant"),
        ("Net Weight Unit", "Description du composant"),
        ("Gross Weight", "Description du composant"),
        ("Gross Weight Unit", "Description du composant"),
        ("Stock unit", "Description du composant"),
        ("Component Details", GROUPING),
    ]
    + [(h, "Supplier PCF") for h in SUPPLIER_PCF_FIELDS]
    + [("Supplier PCF", GROUPING)]
    + [(h, "Transformation") for h in TRANSFORMATION_FIELDS]
    + [("Transformation Details", GROUPING)]
    + [(h, "Flag") for h in FLAG_FIELDS]
)

# Renommage des données d'entrée (v0.9x) vers les en-têtes v0.93 :
# en-tête v0.93 -> en-tête équivalent du fichier d'entrée.
HEADER_RENAMES = {"Transformation Energy Type": "Transformation Energy Name"}

DQR_GUIDE_SHEET = "DQR_Guide"
SHEET_ORDER = ["UserGuide", DQR_GUIDE_SHEET, "Product", "Component"]
# Couleur des onglets (nom des onglets), validée par SBM (fichier v0.92_mod2).
TAB_COLORS: dict[str, str] = {
    "UserGuide": "FF1F497D",
    DQR_GUIDE_SHEET: "FF1F497D",
    "Product": "FF13501B",
    "Component": "FF75A67C",
}


def block_of(header: str, sheet: str = "Product") -> str:
    """Bloc spec (couleur d'en-tête) associé à un champ."""
    if header in FLAG_FIELDS:
        return "Flag"
    if header in SUPPLIER_PCF_FIELDS or header == "Supplier PCF":
        return "Supplier PCF"
    if header.startswith("Transformation"):
        return "Transformation"
    if sheet == "Component":
        if header in COMPONENT_DESCRIPTION_FIELDS:
            return "Description du composant"
        return "Composants"
    return "Produit"


def _column_of(columns: list[tuple[str, str]], header: str) -> int:
    for idx, (name, _) in enumerate(columns, start=1):
        if name == header:
            return idx
    raise KeyError(header)


def _cell(columns: list[tuple[str, str]], header: str, row: int) -> str:
    return f"{get_column_letter(_column_of(columns, header))}{row}"


def _header_style(cell, block: str) -> None:
    cell.fill = PatternFill("solid", fgColor=BLOCK_COLORS[block])
    cell.font = Font(color=BLOCK_FONT_COLORS[block], bold=True)
    if block == GROUPING:
        cell.alignment = Alignment(vertical="center", textRotation=90, wrap_text=True)
    else:
        cell.alignment = Alignment(vertical="center", wrap_text=True)


def _defaults(sheet: str, columns: list[tuple[str, str]], row: int) -> dict[str, object]:
    pcf_unit = "kgCO2e/product" if sheet == "Product" else "kgCO2e/component"
    values: dict[str, object] = {
        "Net Weight Unit": "kg",
        "Gross Weight Unit": "kg",
        "Supplier PCF Unit": pcf_unit,
        "Supplier PCF source": "Supplier",
        "Supplier PCF external review": "No",
        "Transformation Energy Unit": "kW",
        "Transformation Energy EF Unit": "kgCO2e/kW",
        "Transformation Process EF Unit": "kgCO2e",
        "Transformation PDS value": 0,
        "Transformation GHG Unit": pcf_unit,
    }
    geo = _cell(columns, "Transformation GEO DQR", row)
    tech = _cell(columns, "Transformation TECH DQR", row)
    temp = _cell(columns, "Transformation TEMP DQR", row)
    values["Transformation DQR value"] = (
        f'=IF(COUNT({geo},{tech},{temp})=0,"",ROUND(AVERAGE({geo},{tech},{temp}),3))'
    )
    ef = _cell(columns, "Transformation Process EF Value", row)
    net = _cell(columns, "Net Weight", row)
    scrap = _cell(columns, "Transformation Process Scrap Rate", row)
    values["Transformation GHG"] = (
        f'=IF({ef}="","",ROUND({ef}*IF({net}="",0,{net})*(1+IF({scrap}="",0,{scrap})),5))'
    )
    values["Data validation flag"] = _flag_formula(sheet, columns, row)
    return values


def _flag_formula(sheet: str, columns: list[tuple[str, str]], row: int) -> str:
    def c(header: str) -> str:
        return _cell(columns, header, row)

    sku = c("Product SKU" if sheet == "Product" else "Component SKU")
    parts = [
        f'IF(AND({sku}<>"",{c("Supplier Code")}=""),"Supplier code empty (High);","")',
        _pds_flag(c("Supplier PDS")),
        _dqr_flag(c("Supplier DQR")),
        (
            f'IF(AND({c("Supplier PCF value")}<>"",OR({c("Supplier PCF value")}<=0,'
            f'{c("Supplier PCF value")}>100)),"Unusual PCF value might be wrong (High);","")'
        ),
        _pds_flag(c("Transformation PDS value")),
        _dqr_flag(c("Transformation DQR value")),
        (
            f'IF(AND({c("Transformation Process Scrap Rate")}<>"",'
            f'OR({c("Transformation Process Scrap Rate")}>1,'
            f'{c("Transformation Process Scrap Rate")}<0)),"Incorrect scrap rate (High);",'
            f'IF(AND({sku}<>"",{c("Transformation Process Scrap Rate")}=0),'
            f'"No scrap rate (No impact);",""))'
        ),
    ]
    if sheet == "Component":
        parts.append(
            f'IF(AND({sku}<>"",{c("UVP description")}="",'
            f'{c("Raw Material (MB Product)")}="",'
            f'{c("Raw Material - Carbon Footprint")}=""),'
            f'"No raw material available (High);","")'
        )
    return "=" + "&".join(parts) + '&""'


def _pds_flag(ref: str) -> str:
    return (
        f'IF(AND({ref}<>"",OR({ref}>1,{ref}<0)),"Incorrect PDS value (High);",'
        f'IF(AND({ref}<>"",{ref}>0.5),"High PDS value (Medium);",""))'
    )


def _dqr_flag(ref: str) -> str:
    return (
        f'IF(AND({ref}<>"",OR({ref}>5,{ref}<1)),"Incorrect DQR value (High);",'
        f'IF(AND({ref}<>"",{ref}>4),"Incorrect DQR value (Medium);",'
        f'IF(AND({ref}<>"",{ref}>3),"Incorrect DQR value (Low);","")))'
    )


def _number_formats(columns: list[tuple[str, str]]) -> dict[int, str]:
    formats = {
        "Supplier PDS": "0.00%",
        "Supplier DQR": "0.000",
        "Supplier PCF date": "mm-dd-yy",
        "Transformation Process Scrap Rate": "0.00%",
        "Transformation PDS value": "0.00%",
        "Transformation GEO DQR": "0.000",
        "Transformation TECH DQR": "0.000",
        "Transformation TEMP DQR": "0.000",
        "Transformation DQR value": "0.000",
    }
    return {
        _column_of(columns, header): fmt for header, fmt in formats.items()
    }


def _build_sheet(wb, sheet: str, columns: list[tuple[str, str]], data_end_row: int = DATA_END_ROW):
    if sheet in wb.sheetnames:
        del wb[sheet]
    ws = wb.create_sheet(sheet)
    mandatory = MANDATORY.get(sheet, {})
    for idx, (header, block) in enumerate(columns, start=1):
        if header in CALCULATED_FIELDS:
            block = "Calculé"
        cell = ws.cell(row=HEADER_ROW, column=idx, value=header)
        _header_style(cell, block)
        mcell = ws.cell(row=MANDATORY_ROW, column=idx, value=mandatory.get(header))
        mcell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        mcell.font = Font(italic=True, color="FF404040")
        if mandatory.get(header) == "Yes":
            mcell.font = Font(italic=True, bold=True, color="FFC00000")
        letter = get_column_letter(idx)
        if block == GROUPING:
            ws.column_dimensions[letter].width = 3.43
        elif header == "Data validation flag":
            ws.column_dimensions[letter].width = 45
        elif header == "Data validation flag impact":
            ws.column_dimensions[letter].width = 15
        else:
            ws.column_dimensions[letter].width = 18
    ws.row_dimensions[MANDATORY_ROW].height = 30
    ws.row_dimensions[HEADER_ROW].height = 60
    formats = _number_formats(columns)
    for row in range(DATA_START_ROW, data_end_row + 1):
        values = _defaults(sheet, columns, row)
        for idx, (header, _) in enumerate(columns, start=1):
            if header in values:
                ws.cell(row=row, column=idx, value=values[header])
            if idx in formats:
                ws.cell(row=row, column=idx).number_format = formats[idx]
    list_rules = {
        "Supplier PCF framework": '"PACT,TfS,ISO14067,other"',
        "Supplier PCF scope": '"Cradle-to-Gate,Cradle-to-Grave"',
        "Supplier PCF external review": '"Yes,No"',
        "Transformation Energy Type": '"electricity,renewable certified electricity,natural gas"',
    }
    for header, formula in list_rules.items():
        dv = DataValidation(type="list", formula1=formula, allow_blank=True)
        col = _column_of(columns, header)
        dv.add(f"{get_column_letter(col)}{DATA_START_ROW}:{get_column_letter(col)}{data_end_row}")
        ws.add_data_validation(dv)
    dv = DataValidation(type="list", formula1='"1,2,3,4,5"', allow_blank=True)
    for header in ("Transformation GEO DQR", "Transformation TECH DQR", "Transformation TEMP DQR"):
        col = _column_of(columns, header)
        dv.add(f"{get_column_letter(col)}{DATA_START_ROW}:{get_column_letter(col)}{data_end_row}")
    ws.add_data_validation(dv)
    for first, last in GROUPING_RANGES.get(sheet, []):
        first_col = _column_of(columns, first)
        last_col = _column_of(columns, last)
        for col in range(first_col, last_col + 1):
            letter = get_column_letter(col)
            dim = ws.column_dimensions[letter]
            dim.outlineLevel = 1
            dim.hidden = False
    ws.sheet_properties.outlinePr.summaryRight = False
    ws.freeze_panes = "C2"
    return ws


def _copy_input_data(wb, input_path: Path) -> None:
    """Recopie les données saisies du fichier d'entrée vers le nouveau layout."""
    src = load_workbook(input_path, data_only=False)
    for sheet, columns in (("Product", PRODUCT_COLUMNS), ("Component", COMPONENT_COLUMNS)):
        if sheet not in src.sheetnames or sheet not in wb.sheetnames:
            continue
        src_ws, dst_ws = src[sheet], wb[sheet]
        src_headers = {
            (c.value or "").strip(): c.column
            for c in src_ws[1]
            if c.value is not None
        }
        src_data_start = 2
        for row in src_ws.iter_rows(min_row=src_data_start, max_row=src_ws.max_row):
            if row[0].row - src_data_start + DATA_START_ROW > DATA_END_ROW:
                break
            if not any(c.value is not None for c in row):
                continue
            src_r = row[0].row
            dst_r = src_r - src_data_start + DATA_START_ROW
            for idx, (header, _) in enumerate(columns, start=1):
                old = HEADER_RENAMES.get(header, header)
                col = src_headers.get(old)
                if col is None:
                    continue
                value = src_ws.cell(row=src_r, column=col).value
                if value is not None and not str(value).startswith("="):
                    dst_ws.cell(row=dst_r, column=idx, value=value)


def _write_rows(ws, columns: list[tuple[str, str]], rows: list[dict]) -> None:
    """Écrit des lignes de données pré-remplies à partir de la ligne 2."""
    for offset, row in enumerate(rows, start=DATA_START_ROW):
        for header, value in row.items():
            if value is None:
                continue
            try:
                col = _column_of(columns, header)
            except KeyError:
                continue
            ws.cell(row=offset, column=col, value=value)


def apply_v093(
    input_path: str | Path,
    output_path: str | Path,
    template_path: str | Path | None = None,
    spec_path: str | Path | None = None,
    material_path: str | Path | None = None,
) -> tuple[int, int] | None:
    """Génère le fichier de collecte v0.93 depuis un fichier v0.9x.

    Le template (guides UserGuide/DQR_Guide formatés par SBM, ordre des onglets)
    sert de base ; les onglets Product et Component sont reconstruits selon la
    spec v0.93 puis les données du fichier d'entrée y sont recopiées.

    Si spec_path et material_path sont fournis (spécifications v0.93 + fichier
    « Material and Packaging - ExtractPourPCF »), les onglets Product et
    Component sont pré-remplis avec les références LM et leurs composants.
    Retourne alors (nombre de produits, nombre de composants) pré-remplis.
    """
    template = Path(template_path) if template_path else DEFAULT_TEMPLATE
    wb = load_workbook(template)
    data_end_row = DATA_END_ROW
    product_rows = component_rows = None
    if spec_path and material_path:
        from .prefill import build_prefill_rows

        product_rows, component_rows = build_prefill_rows(Path(spec_path), Path(material_path))
        data_end_row = max(
            DATA_START_ROW + len(product_rows) - 1,
            DATA_START_ROW + len(component_rows) - 1,
        )
    _build_sheet(wb, "Product", PRODUCT_COLUMNS, data_end_row)
    _build_sheet(wb, "Component", COMPONENT_COLUMNS, data_end_row)
    order = {name: i for i, name in enumerate(SHEET_ORDER)}
    wb._sheets.sort(key=lambda ws: order.get(ws.title, len(order)))
    for ws in wb.worksheets:
        if ws.title in TAB_COLORS:
            ws.sheet_properties.tabColor = TAB_COLORS[ws.title]
    _copy_input_data(wb, Path(input_path))
    counts = None
    if product_rows is not None:
        _write_rows(wb["Product"], PRODUCT_COLUMNS, product_rows)
        _write_rows(wb["Component"], COMPONENT_COLUMNS, component_rows)
        counts = (len(product_rows), len(component_rows))
    wb.save(output_path)
    return counts

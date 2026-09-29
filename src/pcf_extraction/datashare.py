"""Application des spécifications v0.92 au fichier de collecte LM (Product/Component).

Mise à jour depuis « POC Calculateur - Spécifications - v0.92.xlsx » :
- couleurs de remplissage des en-têtes par bloc (résolution des couleurs de thème
  accent1/accent2/accent3/dk2 avec tints exacts de la spec) ;
- onglet « DQR_Guide » dédié (grille PACT TECH/GEO/TEMP + niveaux de qualité) ;
- UserGuide enrichi du guide de collecte ligne par ligne (travail de l'onglet
  « Introduction » de la spec, traduit en anglais).
"""

from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

HEADER_ROW = 1
GROUPING_FILL = "FFD9D9D9"

# Couleurs des blocs, résolues depuis les couleurs de thème de la spec v0.92
# (thème Office : accent1=156082, accent2=E97132, accent3=196B24, dk2=0E2841).
# Palette de l'onglet « Fichier de collecte » de la spec v0.92 :
#   Produit = accent3 -0.25 -> 13501B (police blanche)
#   Composants = accent3 +0.4 -> 75A67C
#   Description du composant = accent3 +0.6 -> A3C4A7
#   PCF fournisseur = accent2 +0.8 -> FBE3D6
#   Impact fabrication fournisseur = accent2 +0.6 -> F6C6AD
BLOCK_COLORS_V092: dict[str, str] = {
    "Produit": "FF13501B",
    "Composants": "FF75A67C",
    "Description du composant": "FFA3C4A7",
    "Supplier PCF": "FFFBE3D6",
    "Transformation": "FFF6C6AD",
    "Flag": "FFFFC000",
}
BLOCK_FONT_COLORS_V092: dict[str, str] = {
    "Produit": "FFFFFFFF",
    "Composants": "FF000000",
    "Description du composant": "FF000000",
    "Supplier PCF": "FF000000",
    "Transformation": "FF000000",
    "Flag": "FF000000",
}

# Bloc (déterminé par le nom de champ) -> couleur de remplissage de l'en-tête.
def block_of(header: str, sheet: str = "Product") -> str:
    if header in ("Data validation flag", "Data validation flag impact"):
        return "Flag"
    if header.startswith("Supplier PCF") or header in ("Supplier PDS", "Supplier DQR"):
        return "Supplier PCF"
    if header.startswith("Transformation"):
        return "Transformation"
    if sheet == "Component" and header in COMPONENT_DESCRIPTION_FIELDS:
        return "Description du composant"
    if sheet == "Component":
        return "Composants"
    return "Produit"

# Onglet Component : colonnes descriptives (poids, unités, matières) -> bloc
# « Description du composant » de la spec, les autres -> « Composants ».
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

# Onglet DQR_Guide : grille PACT (source : onglet PACT_DQRSpec de la spec v0.92,
# PACT Methodology 3.0 - page 66).
DQR_GUIDE_SHEET = "DQR_Guide"
DQR_GUIDE_ROWS: list[tuple[str, str, str, str, str, str, str]] = [
    ("Source", "PACT Methodology 3.0 - Page 66", "", "", "", "", ""),
    ("Link", "https://docs.carbon-transparency.org/", "", "", "", "", ""),
    ("", "", "", "", "", "", ""),
    ("Technological representativeness (TECH DQR)", "", "", "Geographical representativeness (GEO DQR)", "", "", ""),
    ("Score", "Simplified wording", "", "Score", "Simplified wording", "", ""),
    ("1", "Dataset reflects the exact technology employed (plant-specific process).", "", "1", "Country subdivision where the product is manufactured (e.g., US state, region).", "", ""),
    ("2", "Same technology, company/site-specific but not necessarily plant-specific.", "", "2", "Country average for the manufacturing country.", "", ""),
    ("3", "Average for an equivalent technology to the one employed (same technology group).", "", "3", "Regional average including the site (e.g., Europe, Asia, North America).", "", ""),
    ("4", "Technological proxy (similar but not the same technology, regardless of the supplier).", "", "4", "Global average.", "", ""),
    ("5", "Different or unknown technology versus the technology actually employed.", "", "5", "Unknown geographical scope, or a country/region not including the manufacturing site.", "", ""),
    ("", "", "", "", "", "", ""),
    ("Temporal / Time representativeness (TEMP DQR)", "", "", "Overall DQR evaluation", "", "", ""),
    ("Score", "Simplified wording", "", "Overall DQR", "Quality level", "", ""),
    ("1", "Difference of 1 year or less (366 days) between the dataset's reference year and the PCF reference year.", "", "DQR < 1.5", "Excellent quality", "", ""),
    ("2", "Difference of more than 1 year and up to 2 years (731 days).", "", "1.5 < DQR < 2.0", "Very good quality", "", ""),
    ("3", "Difference of more than 2 years and up to 3 years (1,096 days).", "", "2.0 < DQR < 3.0", "Good quality", "", ""),
    ("4", "Difference of more than 3 years and up to 4 years (1,461 days).", "", "3.0 < DQR < 4.0", "Fair quality", "", ""),
    ("5", "Difference of more than 4 years, or unknown.", "", "DQR > 4.0", "Poor quality", "", ""),
]
DQR_GUIDE_HEADER_COLS = (1, 4)
DQR_GUIDE_SUBHEADER_COLS = (1, 2, 4, 5)

# UserGuide : guide de collecte ligne par ligne (traduction de l'onglet
# « Introduction » de la spec v0.92).
USERGUIDE_INTRODUCTION_ROWS: list[tuple[str, str, str, str]] = [
    ("section", "How to collect the data from suppliers", "", ""),
    ("text", "To compute robust PCFs for the SBM range, energy data must be collected on the components purchased by SBM.", "", ""),
    ("text", "Data collection is structured in 2 quality levels:", "", ""),
    ("text", "Level 1: The supplier already holds the PCF of its component.", "", ""),
    ("text", "If the supplier has already carried out PCFs or LCAs, it holds the precise figure for the component. The information to request is:", "", ""),
    ("table_header", "Question to ask", "Field to fill", "Person in charge of the collection", ""),
    ("table_row", "Do you hold the PCF of [SKU]?", "Component SKU", "SBM", ""),
    ("table_row", "What is the calculation methodology?", "Supplier PCF framework", "Supplier", ""),
    ("table_row", "What is the scope of the PCF?", "Supplier PCF scope", "Supplier", ""),
    ("table_row", "What is the declared (functional) unit of the calculation?", "Supplier PCF declared unit", "Supplier", ""),
    ("table_row", "What is the PCF value?", "Supplier PCF value", "Supplier", ""),
    ("table_row", "Do you have the PDS calculation of the PCF? If yes, provide it. (Mandatory data)", "Supplier PDS", "Supplier", ""),
    ("table_row", "Do you have the DQR calculation of the PCF? If yes, provide it.", "Supplier DQR", "Supplier", ""),
    ("table_row", "In which year was the PCF calculated?", "Supplier PCF date", "Supplier", ""),
    ("table_row", "Has this PCF been certified?", "Supplier PCF external review", "Supplier", ""),
    ("table_row", "Where does the collected data come from?", "Supplier PCF source", "SBM", ""),
    ("text", "", "", ""),
    ("text", "Level 2: The supplier knows the energy associated with the manufacture of the component, or its production process.", "", ""),
    ("text", "The supplier holds the specific energy figure for the component.", "", ""),
    ("text", "The supplier knows the production process of the component very well.", "", ""),
    ("text", "The supplier knows its emissions per production line or per site.", "", ""),
    ("table_header", "Question to ask", "Field to fill", "Person in charge of the collection", ""),
    ("table_row", "Which component are we referring to?", "Component SKU", "SBM", ""),
    ("table_row", "What is the unit of the monitored or computed transformation?", "Transformation Process declared unit", "Supplier or SBM", ""),
    ("table_row", "In which country does the transformation take place?", "Transformation Process location", "Supplier", ""),
    ("table_row", "What type of energy is used to manufacture the component?", "Transformation Energy Type", "Supplier", ""),
    ("table_row", "How much energy is used?", "Transformation Energy Consumption", "Supplier", ""),
    ("table_row", "What is the name of the transformation process used?", "Transformation Process Name", "Supplier", ""),
    ("table_row", "Is there a scrap rate associated with the process?", "Transformation Process Scrap Rate", "Supplier", ""),
    ("table_row", "Does the supplier have a comment on the robustness of the information?", "Transformation Process Comment", "Supplier", ""),
    ("table_row", "Source of the data (monitored, computed or estimated)", "Transformation Energy EF Source", "SBM", ""),
    ("text", "", "", ""),
    ("section", "Risks associated with data estimation", "", ""),
    ("text", "1. Using a proxy far from reality, therefore over- or under-estimating the value.", "", ""),
    ("text", "2. Not finding a proxy for the requested process.", "", ""),
    ("text", "3. Making a site-level consumption approximation, which under-estimates high-consuming products and over-estimates low-consuming ones.", "", ""),
    ("text", "The main risk for SBM is to over/under-evaluate products. PCF quality is lost and there is a risk of missing the right orders of magnitude for some product typologies.", "", ""),
]


GROUPING_MARKERS = {"Product Details", "Component Details", "Supplier PCF", "Transformation Details"}


def _set_block_fill(ws, header: str, cell, sheet: str) -> None:
    block = block_of(header, sheet)
    fill = PatternFill("solid", fgColor=BLOCK_COLORS_V092[block])
    color = BLOCK_FONT_COLORS_V092[block]
    cell.fill = fill
    cell.font = Font(color=color, bold=True)
    cell.alignment = Alignment(vertical="center", wrap_text=True)


def update_colors(wb) -> None:
    """Applique les couleurs v0.92 aux en-têtes Product/Component."""
    for sheet in ("Product", "Component"):
        if sheet not in wb.sheetnames:
            continue
        ws = wb[sheet]
        for cell in ws[HEADER_ROW]:
            header = cell.value
            if header is None:
                continue
            if header in GROUPING_MARKERS:
                cell.fill = PatternFill("solid", fgColor=GROUPING_FILL)
                cell.font = Font(color="FF404040", bold=True)
                cell.alignment = Alignment(vertical="center", textRotation=90, wrap_text=True)
                continue
            _set_block_fill(ws, header, cell, sheet)


def add_dqr_guide(wb) -> None:
    """Crée l'onglet DQR_Guide dédié (grille PACT + niveaux de qualité)."""
    if DQR_GUIDE_SHEET in wb.sheetnames:
        del wb[DQR_GUIDE_SHEET]
    ws = wb.create_sheet(DQR_GUIDE_SHEET)
    for r, row_vals in enumerate(DQR_GUIDE_ROWS, start=1):
        for c, val in enumerate(row_vals, start=1):
            if val != "":
                ws.cell(row=r, column=c, value=val)
    for col_idx in DQR_GUIDE_SUBHEADER_COLS:
        cell = ws.cell(row=5, column=col_idx)
        cell.font = Font(bold=True)
    for r in (4, 12):
        for col_idx in (1, 4):
            cell = ws.cell(row=r, column=col_idx)
            cell.font = Font(bold=True)
            cell.fill = PatternFill("solid", fgColor=BLOCK_COLORS_V092["Produit"])
            cell.font = Font(bold=True, color="FFFFFFFF")
    for r in (5, 13):
        for col_idx in DQR_GUIDE_SUBHEADER_COLS:
            cell = ws.cell(row=r, column=col_idx)
            cell.font = Font(bold=True)
    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 60
    ws.column_dimensions["C"].width = 3
    ws.column_dimensions["D"].width = 10
    ws.column_dimensions["E"].width = 60
    ws.column_dimensions["F"].width = 3
    ws.column_dimensions["G"].width = 18
    for row in ws.iter_rows(min_row=4, max_row=18, min_col=1, max_col=7):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    thin = Side(style="thin", color="FFB0B0B0")
    for row in ws.iter_rows(min_row=5, max_row=18, min_col=1, max_col=5):
        for cell in row:
            cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)


def update_userguide(wb) -> None:
    """Ajoute le guide de collecte ligne par ligne (onglet Introduction traduit) au UserGuide."""
    if "UserGuide" not in wb.sheetnames:
        return
    ws = wb["UserGuide"]
    start = ws.max_row + 2
    bold_font = Font(bold=True, size=14)
    header_fill = PatternFill("solid", fgColor=BLOCK_COLORS_V092["Produit"])
    thin = Side(style="thin", color="FFB0B0B0")
    for i, (kind, *vals) in enumerate(USERGUIDE_INTRODUCTION_ROWS):
        r = start + i
        if kind == "section":
            cell = ws.cell(row=r, column=2, value=vals[0])
            cell.font = bold_font
        elif kind == "text":
            cell = ws.cell(row=r, column=2, value=vals[0])
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            ws.row_dimensions[r].height = 30 if len(vals[0]) > 90 else None
        elif kind == "table_header":
            for j, v in enumerate(vals[:3]):
                cell = ws.cell(row=r, column=2 + j, value=v)
                cell.font = Font(bold=True, color="FFFFFFFF")
                cell.fill = header_fill
                cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)
        elif kind == "table_row":
            for j, v in enumerate(vals[:3]):
                cell = ws.cell(row=r, column=2 + j, value=v)
                cell.alignment = Alignment(wrap_text=True, vertical="top")
                cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)
    ws.column_dimensions["B"].width = 80
    ws.column_dimensions["C"].width = 40
    ws.column_dimensions["D"].width = 30


def apply_v092(input_path: str | Path, output_path: str | Path) -> None:
    """Met à jour un fichier de collecte LM (v0.9x) vers la mise en forme v0.92."""
    wb = load_workbook(input_path)
    update_colors(wb)
    add_dqr_guide(wb)
    update_userguide(wb)
    wb.save(output_path)

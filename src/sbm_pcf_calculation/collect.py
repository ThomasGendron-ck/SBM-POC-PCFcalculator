"""Pipeline « Fichier de collecte » PCF (v0.3).

Construit, pour chaque produit LM de l'échantillon, une ligne par composant de sa
nomenclature (BOM alternative 1, statut actif), enrichie des facteurs d'émission
(CK_MaterialPurchase d'abord, CK_EF_Packaging en secours), des scores DQR/PDS et
des indicateurs PCF, conformément aux onglets « Clé de détermination » et
« Fichier de collecte » du fichier « POC Calculateur - règles de calcul.xlsx ».
"""

from pathlib import Path

import numpy as np
import pandas as pd

from . import config
from .io_sbm import load_sheet

HEADER_EF_PACKAGING = 4
HEADER_MATERIAL_PURCHASE = 12
HEADER_MB_BOM = 4
HEADER_MB_PRODUCTS = 3

BOM_EXCLUDED_ALTERNATIVES = (2, 9)
BOM_ACTIVE_STATUS = 2

INPUT_PATTERNS = {
    "produits_lm": "Référencement*.xlsx",
    "mb_products": "Masterbase_Product*.xlsx",
    "mb_bom": "Masterbase_BOM*.xlsx",
    "material": "*Material and Packaging*.xlsx",
    "freight": "*Freight*.xlsx",
    "lcia": "Cut-off Cumulative LCIA*.xlsx",
}

SHEET_MB_PRODUCTS = "MASTERBASE Products"
SHEET_MB_BOM = "MASTERBASE BOM"

FLAG_PRODUIT_INTROUVABLE = "PRODUIT INTROUVABLE"
FLAG_FOURNISSEUR_INTROUVABLE = "FOURNISSEUR INTROUVABLE"
FLAG_QTE_NULLE = "QTE NULLE (oubli de suppression ?)"
FLAG_POIDS_COHERENCE = "POIDS > 120% ITEM WEIGHT"
FLAG_PAS_DE_MATIERE = "PAS DE MATIERE ASSOCIEE"
FLAG_PAS_DE_FE = "PAS DE FE ASSOCIE A LA MATIERE"
FLAG_PAS_DE_BOM = "PAS DE BOM"
FLAG_PAS_DE_TRAJET_FRET = "PAS DE TRAJET FRET"

UNIT_GHG = "kgCO2e/composant"
UNIT_PCF = "kg CO2eq"

COLLECTE_COLUMNS: list[str] = [
    col
    for _, block in config.BLOCK_HEADERS
    for col in block
]



def _text(value) -> str | None:
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return None
    if pd.isna(value):
        return None
    text = str(value).strip()
    return text or None


def _num(value) -> float:
    return pd.to_numeric(pd.Series([value]), errors="coerce").iloc[0]


def _geo_dqr(geography: str | None) -> float:
    if not geography or geography == "0":
        return 5.0
    if geography == "GLO":
        return 4.0
    if geography == "RoW":
        return 3.0
    if geography == "RER":
        return 2.0
    return 1.0


def _tech_dqr(source: str | None) -> float:
    if not source:
        return 5.0
    if "PROXY" in source.upper():
        return 3.0
    if "CK" in source.upper():
        return 1.0
    return 2.0


def compute_dqr(geography: str | None, source: str | None, has_ef: bool) -> dict[str, float]:
    """Calcule RM GEO/TECH/TEMP DQR et RM DQR value selon la Clé de détermination v0.4.

    TEMP DQR = 1 si un FE est associé, 5 sinon (calibré sur l'exemple SORHOY15
    du Fichier de collecte).
    """
    geo = _geo_dqr(geography)
    tech = _tech_dqr(source)
    temp = 1.0 if has_ef else 5.0
    return {
        "RM GEO DQR": geo,
        "RM TECH DQR": tech,
        "RM TEMP DQR": temp,
        "RM DQR value": (geo + tech + temp) / 3,
    }


def _pick_ck_row(achats: pd.DataFrame, component_ref: str) -> pd.Series | None:
    """Sélectionne la ligne CK_MaterialPurchase de référence d'un composant.

    Priorité aux lignes avec un FE valide (> 0), puis au poids acheté le plus
    élevé (fournisseur principal).
    """
    rows = achats[achats["PRODUCT"] == component_ref]
    if rows.empty:
        return None
    with_ef = rows[rows["Prod_EF_Value"].fillna(0) > 0]
    candidates = with_ef if not with_ef.empty else rows
    return candidates.sort_values("Weight_KGTotal", ascending=False, na_position="last").iloc[0]


def _lookup_ef_packaging(materiaux: pd.DataFrame, matiere: str, description: str | None) -> pd.Series | None:
    """Cherche le FE packaging dans CK_EF_Packaging par matière + description."""
    if not matiere:
        return None
    exact = materiaux[materiaux["RawMat_SubFamily"] == matiere]
    if exact.empty:
        return None
    if description:
        by_desc = exact[exact["CATEGORIE DESCRIPTION"] == description]
        if not by_desc.empty:
            return by_desc.iloc[0]
    return exact.iloc[0]


def load_freight_tables(freight_path: str | Path) -> tuple[dict[str, pd.DataFrame], pd.DataFrame]:
    """Charge les onglets Freight_Raw_In_* et CK_FreightConsolidation du fichier Freight.

    Retourne :
    - raw : dict flux -> DataFrame (IRIS / LS Eu / Fert) avec colonnes
n      composant, fournisseur, UniqueKey et GHG_perunit normalisées ;
    - ck_inbound : lignes inbound de CK_FreightConsolidation indexées par UniqueKey.
    """
    raw: dict[str, pd.DataFrame] = {}
    for flux, spec in config.FREIGHT_SHEETS.items():
        df = load_sheet(freight_path, spec["sheet"], spec["header_row"])
        df = df[[
            spec["component_col"],
            spec["supplier_col"],
            config.FREIGHT_UNIQUEKEY_COL,
            config.FREIGHT_GHG_PERUNIT_COL,
        ]].copy()
        df.columns = ["component", "supplier", "uniquekey", "ghg_perunit"]
        for col in ("component", "supplier", "uniquekey"):
            df[col] = df[col].map(_text)
        df["ghg_perunit"] = pd.to_numeric(df["ghg_perunit"], errors="coerce")
        df = df.dropna(subset=["component", "uniquekey"])
        df["ghg_perunit"] = df.groupby("uniquekey")["ghg_perunit"].transform("mean")
        raw[flux] = df

    ck = load_sheet(freight_path, config.FREIGHT_CK_SHEET, config.FREIGHT_CK_HEADER_ROW)
    ck = ck[ck[config.FREIGHT_FRET_TYPE_COL].fillna("").str.startswith(config.FREIGHT_FRET_TYPE_PREFIX)].copy()
    ck[config.FREIGHT_FRET_TYPE_COL] = (
        ck[config.FREIGHT_FRET_TYPE_COL]
        .str.replace(config.FREIGHT_FRET_TYPE_PREFIX, "", regex=False)
        .str.strip()
    )
    for col in (config.FREIGHT_UNIQUEKEY_COL, config.FREIGHT_ADDRESSKEY_COL, config.FREIGHT_MODE_COL):
        ck[col] = ck[col].map(_text)
    ck_dedup = (
        ck.groupby(config.FREIGHT_UNIQUEKEY_COL, as_index=False)[
            [config.FREIGHT_FRET_TYPE_COL, config.FREIGHT_ADDRESSKEY_COL, config.FREIGHT_MODE_COL]
        ]
        .first()
    )
    ck_dedup = ck_dedup.set_index(config.FREIGHT_UNIQUEKEY_COL, drop=False)
    return raw, ck_dedup


def _lookup_freight(
    raw: dict[str, pd.DataFrame],
    ck_inbound: pd.DataFrame,
    component_ref: str | None,
    supplier_code: str | None,
) -> dict:
    """Rattache le trajet fret d'un composant : raw (composant+fournisseur) -> UniqueKey -> CK.

    Retourne un dict avec route (ADDRESS_KEY), mode (Transportation_Mode),
    flux (colonne FRET TYPE de CK_FreightConsolidation) et ghg_perunit
    (GHG_perunit (kgCO2e/kg) de l'onglet Freight_Raw_In_* correspondant).
    """
    if not component_ref:
        return {}
    for flux, df in raw.items():
        rows = df[df["component"] == component_ref]
        if supplier_code:
            by_supplier = rows[rows["supplier"] == supplier_code]
            if not by_supplier.empty:
                rows = by_supplier
        if rows.empty:
            continue
        row = rows.iloc[0]
        ck_row = ck_inbound.loc[ck_inbound.index == row["uniquekey"]]
        if ck_row.empty:
            continue
        ck_row = ck_row.iloc[0]
        return {
            "Freight Route": ck_row[config.FREIGHT_ADDRESSKEY_COL],
            "Freight Transportation Mode": ck_row[config.FREIGHT_MODE_COL],
            "Flux fret": ck_row[config.FREIGHT_FRET_TYPE_COL],
            "ghg_perunit": row["ghg_perunit"],
        }
    return {}


def load_transformation_saisie(path: str | Path) -> pd.DataFrame:
    """Charge le fichier de saisie de transformation rempli par SBM.

    Une ligne saisie par couple (Product SKU, Component SKU) : ses colonnes
    transformation remplacent les valeurs vides du Fichier de collecte.
    """
    saisie = pd.read_excel(path, sheet_name=config.SAISIE_TRANSFORMATION_SHEET)
    saisie.columns = [str(c).strip() for c in saisie.columns]
    return saisie


def _apply_transformation(collecte: pd.DataFrame, saisie: pd.DataFrame) -> pd.DataFrame:
    """Fusionne les données de transformation saisies par SBM dans la collecte."""
    if saisie is None or saisie.empty:
        return collecte
    saisie = saisie.copy()
    saisie["Product SKU"] = saisie["Product SKU"].map(_text)
    saisie["Component SKU"] = saisie["Component SKU"].map(_text)
    collecte = collecte.copy()
    collecte["Product SKU"] = collecte["Product SKU"].map(_text)
    collecte["Component SKU"] = collecte["Component SKU"].map(_text)
    merged = collecte.merge(
        saisie,
        on=["Product SKU", "Component SKU"],
        how="left",
        suffixes=("", "_saisie"),
    )
    for col in (
        "Transformation Process Name",
        "Transformation Process EF Name",
        "Transformation Process EF Value",
        "Transformation Process EF Unit",
        "Transformation EF Source",
        "Transformation Energy Name",
        "Transformation Energy EF Name",
        "Transformation Energy EF Value",
        "Transformation Energy EF Unit",
        "Transformation Energy EF Source",
        "Scrap Rate",
    ):
        saisie_col = f"{col}_saisie"
        if saisie_col in merged.columns:
            merged[col] = merged[saisie_col].combine_first(merged[col])
    merged["Transformation Process EF Value"] = pd.to_numeric(
        merged["Transformation Process EF Value"], errors="coerce"
    )
    if "Scrap Rate" in merged.columns:
        merged["Scrap Rate"] = pd.to_numeric(merged["Scrap Rate"], errors="coerce")
    scrap = merged["Scrap Rate"] if "Scrap Rate" in merged.columns else pd.Series(dtype=float)
    scrap_factor = 1.0 + scrap.fillna(0.0) if len(scrap) == len(merged) else 1.0
    has_transfo = merged["Transformation Process EF Value"].notna()
    ghg_transfo = merged["Transformation Process EF Value"] * merged["Quantity"] * merged["Net Weight"] * scrap_factor
    merged.loc[has_transfo, "Transformation GHG"] = ghg_transfo[has_transfo]
    merged.loc[has_transfo, "Transformation EF Source"] = (
        merged.loc[has_transfo, "Transformation EF Source"].fillna("Saisie SBM")
    )
    return merged


def _compute_transformation_ghg(collecte: pd.DataFrame) -> pd.DataFrame:
    """Transformation GHG = FE procédé x Quantity x Net Weight x (1 + Scrap Rate)."""
    fe_transfo = pd.to_numeric(collecte["Transformation Process EF Value"], errors="coerce")
    has_transfo = fe_transfo.notna()
    scrap = pd.to_numeric(collecte["Scrap Rate"], errors="coerce") if "Scrap Rate" in collecte.columns else pd.Series(np.nan, index=collecte.index)
    scrap_factor = 1.0 + scrap.fillna(0.0)
    ghg = fe_transfo * collecte["Quantity"] * collecte["Net Weight"] * scrap_factor
    collecte["Transformation GHG"] = ghg.where(has_transfo)
    collecte.loc[has_transfo, "Transformation GHG Unit"] = config.UNIT_GHG_TRANSFORMATION
    return collecte


def _join_flags(flags: list[str]) -> str | None:
    unique = sorted(set(flags))
    return " | ".join(unique) if unique else None


def resolve_inputs(input_dir: str | Path) -> dict[str, Path]:
    """Résout les fichiers source attendus dans le dossier input.

    Les noms contiennent des dates de mise à jour mensuelles : on cherche par
    motif et on garde le plus récent (tri lexicographique, format AAAA MM JJ).
    """
    directory = Path(input_dir)
    if not directory.is_dir():
        raise FileNotFoundError(f"Dossier input introuvable : {directory}")
    resolved: dict[str, Path] = {}
    for key, pattern in INPUT_PATTERNS.items():
        matches = sorted(directory.glob(pattern))
        if not matches:
            raise FileNotFoundError(f"Aucun fichier correspondant à '{pattern}' dans {directory}")
        resolved[key] = matches[-1]
    return resolved


def build_fe_overrides(matching: pd.DataFrame) -> dict[str, dict]:
    """Construit les overrides de FE depuis le matching ecoinvent validé.
    Chaque composant MATCHÉ reçoit le FE proposé (GWP100 EF v3.1, kg CO2e/kg).
    """
    overrides: dict[str, dict] = {}
    for _, row in matching.iterrows():
        if row["Statut"] != "MATCHÉ" or pd.isna(row["FE proposé (kg CO2e/kg)"]):
            continue
        overrides[str(row["Component SKU"])] = {
            "name": row["Dataset ecoinvent"],
            "value": float(row["FE proposé (kg CO2e/kg)"]),
            "unit": "kgCO2e/kg",
            "source": "EcoInvent 3.12 cut-off (validé SBM)",
            "geo": row["Géographie"],
        }
    return overrides


def agg_freight(df: pd.DataFrame, id_col: str, weight_col: str, ghg_col: str) -> pd.DataFrame:
    """Agrège les lignes d'achat fret par article.

    GHG_perunit est un facteur kgCO2e/kg : l'émission totale par ligne = facteur * poids.
    On somme ensuite par article pour obtenir l'émission fret totale du composant (kgCO2e),
    et on garde aussi le poids total transporté.
    """
    df = df.copy()
    df[id_col] = df[id_col].astype(str).str.strip()
    for col in (weight_col, ghg_col):
        df[col] = pd.to_numeric(df[col], errors="coerce")
    df["emission_line"] = df[ghg_col] * df[weight_col]
    grouped = df.groupby(id_col, dropna=False).agg(
        poids_total=(weight_col, "sum"),
        emission_totale=("emission_line", "sum"),
        nb_lignes_achat=(id_col, "size"),
    ).reset_index()
    return grouped


def build_collecte(
    input_dir: str | Path,
    transformation: str | Path | None = None,
    fe_overrides: dict[str, dict] | None = None,
) -> pd.DataFrame:
    """Construit le Fichier de collecte à partir des sources du dossier input.

    Si `transformation` est fourni (fichier de saisie rempli par SBM), ses
    valeurs alimentent les colonnes du bloc Impact fabrication fournisseur.
    Si `fe_overrides` est fourni (matching ecoinvent validé), les composants
    sans FE reçoivent le FE proposé avant calcul du DQR, du GHG et du PCF.
    """
    paths = resolve_inputs(input_dir)

    ef_packaging = load_sheet(paths["material"], "CK_EF_Packaging", HEADER_EF_PACKAGING)
    achats = load_sheet(paths["material"], "CK_MaterialPurchase", HEADER_MATERIAL_PURCHASE)
    bom = load_sheet(paths["mb_bom"], SHEET_MB_BOM, HEADER_MB_BOM)
    produits = load_sheet(paths["mb_products"], SHEET_MB_PRODUCTS, HEADER_MB_PRODUCTS)
    freight_raw, ck_inbound = load_freight_tables(paths["freight"])

    saisie = None
    if transformation is not None:
        saisie = load_transformation_saisie(transformation)

    ef_packaging = ef_packaging.copy()
    ef_packaging["RawMat_SubFamily"] = ef_packaging["RawMat_SubFamily"].map(_text)
    ef_packaging["CATEGORIE DESCRIPTION"] = ef_packaging["CATEGORIE DESCRIPTION"].map(_text)

    achats = achats.copy()
    achats["PRODUCT"] = achats["PRODUCT"].map(_text)
    for col in ("Weight_KGTotal", "Prod_EF_Value", "Taux Recyclé"):
        achats[col] = pd.to_numeric(achats[col], errors="coerce")
    achats["RawMat_SubFamily"] = achats["RawMat_SubFamily"].map(_text)

    produits = produits.copy()
    produits["SKU"] = produits["SKU"].map(_text)
    produits = produits.dropna(subset=["SKU"]).drop_duplicates(subset="SKU").set_index("SKU", drop=False)

    bom = bom.copy()
    for col in ("ITMREF", "CPNITMREF"):
        bom[col] = bom[col].map(_text)
    bom_active = bom[(~bom["BOMALT"].isin(BOM_EXCLUDED_ALTERNATIVES)) & (bom["USESTA_0"] == BOM_ACTIVE_STATUS)]
    bom_active = bom_active.sort_values(["ITMREF", "CPNITMREF", "BOMALT"])
    bom_active = bom_active.drop_duplicates(subset=["ITMREF", "CPNITMREF"], keep="first")

    sample = pd.read_excel(paths["produits_lm"], sheet_name="Export")
    sample.columns = [str(c).strip() for c in sample.columns]
    sample = sample.rename(
        columns={
            "Num Reference fournisseur": "SKU",
            "Designation article": "Designation",
        }
    )
    sample["SKU"] = sample["SKU"].map(_text)
    sample = sample.dropna(subset=["SKU"]).drop_duplicates(subset="SKU")

    rows_out: list[dict] = []
    for _, lm in sample.iterrows():
        sku = _text(lm["SKU"])
        flags_produit: list[str] = []

        prod = produits.loc[sku] if sku in produits.index else None
        if prod is None:
            flags_produit.append(FLAG_PRODUIT_INTROUVABLE)

        sage_prod = _text(prod["Category"]) if prod is not None else None
        desc_prod, _ = config.SAGE_CATEGORY_MAP.get(sage_prod or "", ("", ""))

        fournisseur_prod = _text(prod["BPSNAM"]) if prod is not None else None
        code_prod = _text(prod["BPSNUM"]) if prod is not None else None

        base_produit = {
            "Product SKU": sku,
            "Product Designation": _text(lm["Designation"]) or (_text(prod["SKU Designation"]) if prod is not None else None),
            "Product Category Code": sage_prod,
            "Product Category description": desc_prod or None,
            "Pack Unit Box": _num(prod["Pack unit box"]) if prod is not None else np.nan,
            "Product Net Weight": _num(prod["NET_WEIGHT0"]) if prod is not None else np.nan,
            "Product Gross Weight": _num(prod["GROSS_WEIGHT0"]) if prod is not None else np.nan,
            "Product Supplier Code": code_prod,
            "Product Supplier Name": fournisseur_prod,
        }
        if not fournisseur_prod:
            flags_produit.append(FLAG_FOURNISSEUR_INTROUVABLE)

        bom_prod = bom_active[bom_active["ITMREF"] == sku]
        if bom_prod.empty:
            flags_produit.append(FLAG_PAS_DE_BOM)
            rows_out.append({**base_produit, "Flag": _join_flags(flags_produit)})
            continue

        comp_rows: list[dict] = []
        for _, bom_row in bom_prod.iterrows():
            row = dict(base_produit)
            flags = list(flags_produit)

            qty = _num(bom_row["UVQTY"])
            if pd.isna(qty) or qty == 0:
                flags.append(FLAG_QTE_NULLE)

            comp_ref = _text(bom_row["CPNITMREF"])
            comp = produits.loc[comp_ref] if comp_ref in produits.index else None
            if comp is None:
                flags.append(FLAG_PRODUIT_INTROUVABLE)

            ck_row = _pick_ck_row(achats, comp_ref)

            item_weight = _num(comp["Item weight"]) if comp is not None else np.nan
            gross_weight = _num(comp["GROSS_WEIGHT0"]) if comp is not None else np.nan
            if pd.notna(gross_weight) and gross_weight == 0:
                gross_weight = item_weight
            if pd.notna(gross_weight) and pd.notna(item_weight) and gross_weight > 1.2 * item_weight:
                flags.append(FLAG_POIDS_COHERENCE)

            sage_comp = _text(comp["Category"]) if comp is not None else None
            desc_comp, cf_cat = config.SAGE_CATEGORY_MAP.get(sage_comp or "", ("", ""))

            fournisseur = _text(comp["BPSNAM"]) if comp is not None else None
            code_fournisseur = _text(comp["BPSNUM"]) if comp is not None else None
            if not fournisseur and ck_row is not None:
                fournisseur = _text(ck_row["SUPPLIER_NAME"])
                code_fournisseur = _text(ck_row["SUPPLIER_CODE"])
            if not fournisseur:
                flags.append(FLAG_FOURNISSEUR_INTROUVABLE)

            matiere = _text(comp["ZCODMAT2"]) if comp is not None else None
            if not matiere and ck_row is not None:
                matiere = _text(ck_row["RawMat_SubFamily"])
            if not matiere and cf_cat == "Raw Material" and desc_comp:
                matiere = desc_comp
            if not matiere:
                flags.append(FLAG_PAS_DE_MATIERE)

            recycle = _num(comp["ZRECYCLE"]) if comp is not None else np.nan
            if pd.isna(recycle) and ck_row is not None:
                recycle = _num(ck_row["Taux Recyclé"])
            if pd.isna(recycle):
                recycle = 0.0

            fe_nom = fe_val = fe_unit = fe_src = fe_geo = None
            if ck_row is not None and pd.notna(ck_row["Prod_EF_Value"]) and ck_row["Prod_EF_Value"] > 0:
                fe_nom = _text(ck_row["Prod_EF_Name"])
                fe_val = float(ck_row["Prod_EF_Value"])
                fe_unit = _text(ck_row["Prod_EF_Unit"])
                fe_src = _text(ck_row["Prod_EF_Source"])
                fe_geo = _text(ck_row["Prod_EF_Geography"])
            else:
                ef_pack = _lookup_ef_packaging(ef_packaging, matiere, desc_comp) if cf_cat == "PACKAGING" else None
                if ef_pack is not None:
                    if recycle > 0 and pd.notna(ef_pack["Recycled_EF_Value"]):
                        fe_nom = _text(ef_pack["Recycled1_EF_Name"]) or f"FE calculé par CK pour {matiere} recyclé"
                        fe_val = float(ef_pack["Recycled_EF_Value"])
                        fe_unit = _text(ef_pack["Recycled1_EF_Unit"])
                        fe_src = "CK_EF_Packaging"
                    else:
                        fe_nom = _text(ef_pack["Virgin1_EF_Name2"]) or f"FE calculé par CK pour {matiere}"
                        fe_val = float(ef_pack["Virgin_EF_Value"])
                        fe_unit = _text(ef_pack["Virgin1_EF_Unit"])
                        fe_src = "CK_EF_Packaging"
                    fe_geo = "FR"
            if fe_val is None and fe_overrides and comp_ref in fe_overrides:
                override = fe_overrides[comp_ref]
                fe_nom = override["name"]
                fe_val = override["value"]
                fe_unit = override["unit"]
                fe_src = override["source"]
                fe_geo = override["geo"]
            has_ef = fe_val is not None and pd.notna(fe_val)
            if not has_ef:
                flags.append(FLAG_PAS_DE_FE)

            freight_info = _lookup_freight(freight_raw, ck_inbound, comp_ref, code_fournisseur)
            if not freight_info:
                flags.append(FLAG_PAS_DE_TRAJET_FRET)

            activity_pds = 1.0 if (_text(comp["SKU Designation"]) if comp is not None else None) else 0.0
            dqr = compute_dqr(fe_geo, fe_src, has_ef)

            ghg = np.nan
            if has_ef and pd.notna(qty) and pd.notna(item_weight):
                ghg = item_weight * qty * fe_val

            row.update(
                {
                    "Component SKU": comp_ref,
                    "Component Designation": _text(comp["SKU Designation"]) if comp is not None else None,
                    "Category Code": sage_comp,
                    "Category description": desc_comp or None,
                    "Carbon category": cf_cat or None,
                    "Pack unit box": _num(comp["Pack unit box"]) if comp is not None else np.nan,
                    "Quantity": qty,
                    "RM PDS Activity Data": activity_pds,
                    "Net Weight": item_weight,
                    "Net Weight Unit": _text(comp["Weight unit"]) if comp is not None else None,
                    "Stock unit": _text(comp["PCU0"]) if comp is not None else None,
                    "Gross Weight": gross_weight,
                    "Gross Weight Unit": _text(comp["Weight unit"]) if comp is not None else None,
                    "Supplier code": code_fournisseur,
                    "Supplier Name": fournisseur,
                    "Raw Material": matiere,
                    "Recycled %": recycle,
                    "Scrap Rate": None,
                    "Supplier PCF value": None,
                    "Supplier PCF Unit": None,
                    "Supplier PDS": None,
                    "Supplier DQR": None,
                    "Supplier PCF source": None,
                    "Supplier PCF external review": None,
                    "RM EF Name": fe_nom,
                    "RM EF Value": fe_val,
                    "RM EF Unit": fe_unit,
                    "RM EF Source": fe_src,
                    "RM EF PDS": 0.0,
                    "RM PDS value": activity_pds * 0.0,
                    "Prod_EF_Geography": fe_geo,
                    **dqr,
                    "RM GHG": ghg,
                    "RM GHG Unit": UNIT_GHG,
                    "Transformation Process Name": None,
                    "Transformation Process EF Name": None,
                    "Transformation Process EF Value": None,
                    "Transformation Process EF Unit": None,
                    "Transformation EF Source": None,
                    "Transformation Energy Name": None,
                    "Transformation Energy EF Name": None,
                    "Transformation Energy EF Value": None,
                    "Transformation Energy EF Unit": None,
                    "Transformation Energy EF Source": None,
                    "Transformation GHG": None,
                    "Transformation GHG Unit": None,
                    "Freight Supplier Code": code_fournisseur,
                    "Freight Supplier Name": fournisseur,
                    "Freight Route": freight_info.get("Freight Route"),
                    "Freight Transportation Mode": freight_info.get("Freight Transportation Mode"),
                    "Freight GHG": (
                        freight_info["ghg_perunit"] * qty * item_weight
                        if "ghg_perunit" in freight_info
                        and pd.notna(freight_info.get("ghg_perunit"))
                        and pd.notna(qty)
                        and pd.notna(item_weight)
                        else np.nan
                    ),
                    "Freight GHG Unit": config.UNIT_GHG_TRANSPORT if "ghg_perunit" in freight_info else None,
                }
            )
            row["Flag"] = _join_flags(flags)
            comp_rows.append(row)

        poids_totaux = sum(
            r["Net Weight"] * r["Quantity"]
            for r in comp_rows
            if pd.notna(r["Net Weight"]) and pd.notna(r["Quantity"])
        )
        pcf_value = sum(r["RM GHG"] for r in comp_rows if pd.notna(r["RM GHG"]))
        dqr_terms = [
            (r["RM DQR value"], r["Net Weight"] * r["Quantity"] / poids_totaux)
            for r in comp_rows
            if pd.notna(r["RM DQR value"]) and poids_totaux and pd.notna(r["Net Weight"]) and pd.notna(r["Quantity"])
        ]
        dqr_product = sum(d * p for d, p in dqr_terms) / sum(p for _, p in dqr_terms) if dqr_terms else np.nan
        for r in comp_rows:
            poids_comp = r["Net Weight"] * r["Quantity"] if pd.notna(r["Net Weight"]) and pd.notna(r["Quantity"]) else np.nan
            r["Part du composant dans le produit"] = poids_comp / poids_totaux if pd.notna(poids_comp) and poids_totaux else np.nan
            r["PCF Value"] = pcf_value if pcf_value else np.nan
            r["PCF Unit"] = UNIT_PCF if pcf_value else None
            r["DQR Product"] = dqr_product
            r["PDS Product"] = 0.0 if pcf_value else np.nan
            rows_out.append(r)

    collecte = pd.DataFrame(rows_out)
    for col in COLLECTE_COLUMNS:
        if col not in collecte.columns:
            collecte[col] = None
        if collecte[col].dtype == float and col in config.OBJECT_COLUMNS:
            collecte[col] = collecte[col].astype(object)
    collecte = collecte[COLLECTE_COLUMNS]
    collecte = _compute_transformation_ghg(collecte)
    if saisie is not None:
        collecte = _apply_transformation(collecte, saisie)
        collecte = collecte[COLLECTE_COLUMNS]
    return collecte

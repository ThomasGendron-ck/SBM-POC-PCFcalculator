"""Pré-remplissage des onglets Product et Component du fichier de collecte.

Sources (bibliothèque SBM) :
- « POC Calculateur - Spécifications - v0.93.xlsx », onglet « Produits LM » :
  la liste des références Leroy Merlin (une ligne par produit fini) ;
- « SBM LS Europe - BC FY24-25 - Material and Packaging - vF - ExtractPourPCF.xlsx » :
  MasterBase_Products (attributs produit/composant), MasterBase_BOM (nomenclatures),
  CK_MaterialPurchase (fournisseurs, matières premières), Category_Param
  (description des catégories SAGE).

Règles de périmètre (identiques au pipeline pcf-collecte v0.74) :
- composants = alternatives BOMALT toutes sauf 2 et 9, statut actif (USESTA_0 = 2) ;
- dédoublement : un couple produit/composant présent dans plusieurs alternatives
  est retenu à la plus petite alternative ;
- onglet Component : une ligne par composant unique du périmètre.
"""

from pathlib import Path

import pandas as pd

BOM_EXCLUDED_ALTERNATIVES = {2, 9}
BOM_ACTIVE_STATUS = 2
LM_SPEC_SHEET = "Produits LM"
LM_SKU_COL = "Num Reference fournisseur "
LM_DESIGNATION_COL = "Designation article "


def _text(value) -> str | None:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    text = str(value).strip()
    return text or None


def _num(value) -> float | None:
    number = pd.to_numeric(value, errors="coerce")
    if pd.isna(number):
        return None
    return float(number)


MB_PRODUCT_SHEET = "MASTERBASE Products"
MB_EXTRACT_SHEET = "MasterBase_Products"


def _sku_key(sku: str) -> str:
    """Clé de lookup SKU : les zéros de tête ne sont pas significatifs
    (la BOM référence « 000001 » alors que certains extraits stockent « 1 »)."""
    return sku.lstrip("0") or sku


def _load_products_frame(path: Path) -> pd.DataFrame:
    """Charge les produits Masterbase depuis un extrait complet
    (« MASTERBASE Products », en-tête sur la ligne 3) ou depuis le fichier
    « ExtractPourPCF » (« MasterBase_Products », en-tête ligne 12)."""
    sheets = pd.ExcelFile(path).sheet_names
    sheet = next(
        (name for name in sheets if name.strip().lower() == MB_PRODUCT_SHEET.lower()),
        None,
    )
    if sheet is None:
        sheet = next(
            (name for name in sheets if name.strip().lower() == MB_EXTRACT_SHEET.lower()),
            None,
        )
    if sheet is None:
        raise ValueError(
            f"{path} : aucun onglet produits Masterbase trouvé "
            f"(« {MB_PRODUCT_SHEET} » ou « {MB_EXTRACT_SHEET} »)."
        )
    preview = pd.read_excel(path, sheet_name=sheet, header=None, nrows=15)
    header_row = next(
        i for i in range(len(preview)) if _text(preview.iloc[i, 0]) == "SKU"
    )
    products = pd.read_excel(path, sheet_name=sheet, header=header_row, dtype=str)
    products["SKU"] = products["SKU"].map(_text)
    products = products.dropna(subset=["SKU"]).drop_duplicates(subset="SKU")
    products.index = products["SKU"].map(_sku_key)
    products = products[~products.index.duplicated(keep="first")]
    return products


LM_EXPORT_SHEET = "Export"
LM_EXPORT_SKU_COL = "Num Reference fournisseur"
LM_EXPORT_DESIGNATION_COL = "Designation article"


def load_lm_products(source_path: str | Path) -> pd.DataFrame:
    """Liste des références LM depuis la spec (onglet « Produits LM ») ou
    directement depuis le fichier « Référencement LM » (onglet « Export »)."""
    source_path = Path(source_path)
    sheets = pd.ExcelFile(source_path).sheet_names
    if LM_SPEC_SHEET in sheets:
        df = pd.read_excel(source_path, sheet_name=LM_SPEC_SHEET)
        df = df.rename(columns={LM_SKU_COL: "SKU", LM_DESIGNATION_COL: "Designation"})
    elif LM_EXPORT_SHEET in sheets:
        df = pd.read_excel(source_path, sheet_name=LM_EXPORT_SHEET)
        df.columns = [str(c).strip() for c in df.columns]
        df = df.rename(
            columns={LM_EXPORT_SKU_COL: "SKU", LM_EXPORT_DESIGNATION_COL: "Designation"}
        )
    else:
        raise ValueError(
            f"{source_path} : ni l'onglet « {LM_SPEC_SHEET} » (spec) ni "
            f"« {LM_EXPORT_SHEET} » (Référencement LM) n'a été trouvé."
        )
    df["SKU"] = df["SKU"].map(_text)
    df = df.dropna(subset=["SKU"]).drop_duplicates(subset="SKU")
    return df[["SKU", "Designation"]]


def _load_masterbase(
    material_path: Path, mb_product_path: Path | None = None
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    products = _load_products_frame(Path(mb_product_path) if mb_product_path else material_path)
    bom = pd.read_excel(material_path, sheet_name="MasterBase_BOM", header=11)
    purchases = pd.read_excel(material_path, sheet_name="CK_MaterialPurchase", header=11)
    categories = pd.read_excel(material_path, sheet_name="Category_Param")
    bom = bom.dropna(subset=["ITMREF", "CPNITMREF"])
    for col in ("ITMREF", "CPNITMREF"):
        bom[col] = bom[col].map(_text)
    bom["BOMALT"] = pd.to_numeric(bom["BOMALT"], errors="coerce")
    bom["USESTA_0"] = pd.to_numeric(bom["USESTA_0"], errors="coerce")
    bom = bom.sort_values(["ITMREF", "CPNITMREF", "BOMALT"])
    purchases["PRODUCT"] = purchases["PRODUCT"].map(_text)
    purchases = purchases.dropna(subset=["PRODUCT"])
    return products, bom, purchases, categories


def build_prefill_rows(
    spec_path: str | Path,
    material_path: str | Path,
    mb_product_path: str | Path | None = None,
) -> tuple[list[dict], list[dict]]:
    """Construit les lignes Product et Component à pré-remplir.

    Retourne (product_rows, component_rows) : une ligne par référence LM et une
    ligne par composant unique du périmètre BOM élargi. Les attributs produits
    proviennent de la Masterbase complète (mb_product_path) si fournie, sinon de
    l'extrait ExtractPourPCF.
    """
    lm = load_lm_products(spec_path)
    products, bom, purchases, categories = _load_masterbase(
        Path(material_path), Path(mb_product_path) if mb_product_path else None
    )
    category_names = dict(zip(categories["Category_Code"], categories["Category_Name"]))

    bom_active = bom[
        (~bom["BOMALT"].isin(BOM_EXCLUDED_ALTERNATIVES)) & (bom["USESTA_0"] == BOM_ACTIVE_STATUS)
    ]
    bom_active = bom_active.drop_duplicates(subset=["ITMREF", "CPNITMREF"], keep="first")

    raw_material_cf = (
        purchases.dropna(subset=["RawMat_Hypothesis"])
        .sort_values("Weight_KGTotal", ascending=False)
        .drop_duplicates(subset=["PRODUCT"], keep="first")
        .set_index("PRODUCT", drop=False)
    )

    def _product_row(sku: str):
        key = _sku_key(sku)
        return products.loc[key] if key in products.index else None

    def supplier_of(sku: str) -> tuple[str | None, str | None]:
        """Supplier Code/Name : MB_Product (BPSNUM/BPSNAM) uniquement (spec v0.96)."""
        row = _product_row(sku)
        if row is None:
            return None, None
        code = _text(row["BPSNUM"]) if "BPSNUM" in products.columns else None
        name = _text(row["BPSNAM"]) if "BPSNAM" in products.columns else None
        return code, name

    def raw_material_of(sku: str) -> str | None:
        """Matière première : MB_Product (ZCODMAT2) uniquement (spec v0.96)."""
        row = _product_row(sku)
        if row is None or "ZCODMAT2" not in products.columns:
            return None
        return _text(row["ZCODMAT2"])

    def raw_material_cf_of(sku: str) -> str | None:
        """Raw Material - Carbon Footprint : CK_MaterialPurchase
        (RawMat_Hypothesis), ligne de plus gros poids (spec v0.96)."""
        if sku in raw_material_cf.index:
            return _text(raw_material_cf.loc[sku]["RawMat_Hypothesis"])
        key = _sku_key(sku)
        if key != sku and key in raw_material_cf.index:
            return _text(raw_material_cf.loc[key]["RawMat_Hypothesis"])
        return None


    product_rows: list[dict] = []
    for _, lmr in lm.iterrows():
        sku = lmr["SKU"]
        prod = _product_row(sku)

        def attr(column: str):
            if prod is None:
                return None
            return _text(prod[column]) if column in products.columns else None

        category = attr("Category")
        supplier_code, supplier_name = supplier_of(sku)
        product_rows.append(
            {
                "Product SKU": sku,
                "Product Designation": _text(lmr["Designation"]) or attr("SKU Designation"),
                "Category Code": category,
                "Category description": category_names.get(category),
                "Supplier Code": supplier_code,
                "Supplier Name": supplier_name,
                "Pack Unit Box": None if prod is None else _num(prod["Pack unit box"]),
                "Net Weight": None if prod is None else _num(prod["Item weight"]),
                "Net Weight Unit": attr("Weight unit"),
                "Gross Weight": None if prod is None else _num(prod["GROSS_WEIGHT0"]),
                "Gross Weight Unit": attr("Weight unit"),
                "Stock unit": attr("PCU0"),
            }
        )

    lm_skus = set(lm["SKU"])
    component_refs = sorted(set(bom_active[bom_active["ITMREF"].isin(lm_skus)]["CPNITMREF"]))
    component_rows: list[dict] = []
    for sku in component_refs:
        prod = _product_row(sku)

        def attr(column: str):
            if prod is None:
                return None
            return _text(prod[column]) if column in products.columns else None

        category = attr("Category")
        supplier_code, supplier_name = supplier_of(sku)
        raw_material = raw_material_of(sku)
        component_rows.append(
            {
                "Component SKU": sku,
                "Component Designation": attr("SKU Designation"),
                "Category Code": category,
                "Category description": category_names.get(category),
                "Supplier Code": supplier_code,
                "Supplier Name": supplier_name,
                "Pack unit box": None if prod is None else _num(prod["Pack unit box"]),
                "Net Weight": None if prod is None else _num(prod["Item weight"]),
                "Net Weight Unit": attr("Weight unit"),
                "Gross Weight": None if prod is None else _num(prod["GROSS_WEIGHT0"]),
                "Gross Weight Unit": attr("Weight unit"),
                "Stock unit": attr("PCU0"),
                "Raw Material (MB Product)": raw_material,
                "Raw Material - Carbon Footprint": raw_material_cf_of(sku),
                "UVP description": attr("ZUVP_DES"),
            }
        )
    return product_rows, component_rows

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


def _load_masterbase(material_path: Path) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    products = pd.read_excel(material_path, sheet_name="MasterBase_Products", header=11)
    bom = pd.read_excel(material_path, sheet_name="MasterBase_BOM", header=11)
    purchases = pd.read_excel(material_path, sheet_name="CK_MaterialPurchase", header=11)
    categories = pd.read_excel(material_path, sheet_name="Category_Param")
    products["SKU"] = products["SKU"].map(_text)
    products = products.dropna(subset=["SKU"]).drop_duplicates(subset="SKU").set_index("SKU", drop=False)
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
) -> tuple[list[dict], list[dict]]:
    """Construit les lignes Product et Component à pré-remplir.

    Retourne (product_rows, component_rows) : une ligne par référence LM et une
    ligne par composant unique du périmètre BOM élargi.
    """
    lm = load_lm_products(spec_path)
    products, bom, purchases, categories = _load_masterbase(Path(material_path))
    category_names = dict(zip(categories["Category_Code"], categories["Category_Name"]))

    bom_active = bom[
        (~bom["BOMALT"].isin(BOM_EXCLUDED_ALTERNATIVES)) & (bom["USESTA_0"] == BOM_ACTIVE_STATUS)
    ]
    bom_active = bom_active.drop_duplicates(subset=["ITMREF", "CPNITMREF"], keep="first")

    purchase_suppliers = (
        purchases.dropna(subset=["SUPPLIER_CODE"])
        .sort_values("SUPPLIER_CODE")
        .drop_duplicates(subset=["PRODUCT"], keep="first")
        .set_index("PRODUCT", drop=False)
    )

    raw_materials = (
        purchases.dropna(subset=["RawMat_SubFamily"])
        .sort_values("Weight_KGTotal", ascending=False)
        .drop_duplicates(subset=["PRODUCT"], keep="first")
        .set_index("PRODUCT", drop=False)
    )

    def supplier_of(sku: str) -> tuple[str | None, str | None]:
        """Supplier Code/Name : MB_Product (BPSNUM/BPSNAM) selon la spec,
        fallback CK_MaterialPurchase (SUPPLIER_CODE/SUPPLIER_NAME)."""
        if sku in products.index:
            row = products.loc[sku]
            code = _text(row["BPSNUM"]) if "BPSNUM" in products.columns else None
            name = _text(row["BPSNAM"]) if "BPSNAM" in products.columns else None
            if code or name:
                return code, name
        if sku in purchase_suppliers.index:
            row = purchase_suppliers.loc[sku]
            return _text(row["SUPPLIER_CODE"]), _text(row["SUPPLIER_NAME"])
        return None, None

    def raw_material_of(sku: str) -> str | None:
        """Matière première : MB_Product (ZCODMAT2) selon la spec,
        fallback CK_MaterialPurchase (RawMat_SubFamily)."""
        if sku in products.index:
            row = products.loc[sku]
            if "ZCODMAT2" in products.columns:
                material = _text(row["ZCODMAT2"])
                if material:
                    return material
        if sku in raw_materials.index:
            return _text(raw_materials.loc[sku]["RawMat_SubFamily"])
        return None


    product_rows: list[dict] = []
    for _, lmr in lm.iterrows():
        sku = lmr["SKU"]
        prod = products.loc[sku] if sku in products.index else None

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
        prod = products.loc[sku] if sku in products.index else None

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
                "UVP description": attr("ZUVP_DES"),
            }
        )
    return product_rows, component_rows

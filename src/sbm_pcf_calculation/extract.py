"""Extraction des composants pour le calcul des facteurs d'émission (PCF).

Pipeline :
1. Charger l'échantillon de produits à analyser (SKU).
2. Charger le référentiel produits (MasterBase_Products) et la nomenclature (MasterBase_BOM).
3. Pour chaque produit, extraire ses composants via la BOM (relation 1-n).
4. Enrichir chaque composant avec les émissions de CK_MaterialPurchase (Prod / Use / EoL).
5. Enrichir avec les émissions de fret amont (IRIS / Fert / LS Europe).
6. Produire un rapport détaillé : produits, composants, statistiques sur les FE manquants.
"""

import pandas as pd

from . import config
from .io_sbm import load_bom, load_components, load_freight, load_products, load_sample_products

PHASE_COLUMNS = {
    "Production": "Qty_Prod_GHG (kgCO2e)",
    "Usage": "Qty_Use_GHG (kgCO2e)",
    "Fin de Vie": "Qty_EoL_GHG (kgCO2e)",
}

FREIGHT_SOURCES = [
    ("Fret Amont - IRIS", "Freight_Raw_In_IRIS", "Référence Article", 13),
    ("Fret Amont - Fert", "Freight_Raw_In_Fert", "Référence Article", 14),
    ("Fret Amont - LS Europe", "Freight_Raw_In_LSEur", "Product", 13),
]

FREIGHT_ID_COL = "ID Unique Composant"


def _agg_freight(df: pd.DataFrame, id_col: str, weight_col: str, ghg_col: str) -> pd.DataFrame:
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


def extract(material_path: str, freight_path: str, sample_path: str | None = None) -> dict[str, pd.DataFrame]:
    """Exécute le pipeline complet et retourne les tables du rapport.

    Si sample_path est fourni, seuls les SKU de l'échantillon sont traités.
    """
    products = load_products(material_path)
    bom = load_bom(material_path)
    components = load_components(material_path)
    freight_sheets = {
        source[1]: load_freight(freight_path, source[1], source[3])
        for source in FREIGHT_SOURCES
    }

    products = products.copy()
    products["SKU"] = products["SKU"].astype(str).str.strip()
    bom = bom.copy()
    for col in ("ITMREF", "CPNITMREF"):
        bom[col] = bom[col].astype(str).str.strip()
    components = components.copy()
    components["PRODUCT"] = components["PRODUCT"].astype(str).str.strip()

    if sample_path:
        sample = load_sample_products(sample_path)
        sample_skus = set(sample["SKU"])
        products = products[products["SKU"].isin(sample_skus)]
        bom = bom[bom["ITMREF"].isin(sample_skus)]

    prod_cols = {
        "SKU": "ID Unique Produit",
        "SKU Designation": "Nom",
        "Des super segment": "Super Segment",
        "Segment_DES_TSI2": "Segment",
        "Sub Subsegment_DES_TSI3": "Sub Segment",
    }
    products_out = products[list(prod_cols)].rename(columns=prod_cols).copy()
    products_out = products_out.dropna(subset=["ID Unique Produit"]).drop_duplicates(subset="ID Unique Produit")

    bom_out = bom[["ITMREF", "CPNITMREF", "BOMQTY"]].rename(
        columns={
            "ITMREF": "ID Unique Produit",
            "CPNITMREF": "ID Unique Composant",
            "BOMQTY": "Quantité de composant dans produit",
        }
    )
    bom_out = bom_out.dropna(subset=["ID Unique Produit", "ID Unique Composant"])

    comp_cols = {
        "PRODUCT": "ID Unique Composant",
        "PRODUCT_NAME": "Nom",
        "Catégorie ACV": "Type composant",
        PHASE_COLUMNS["Production"]: "Emission - Production (kgCO2e)",
        PHASE_COLUMNS["Usage"]: "Emission - Usage (kgCO2e)",
        PHASE_COLUMNS["Fin de Vie"]: "Emission - Fin de Vie (kgCO2e)",
    }
    components_out = components[list(comp_cols)].rename(columns=comp_cols)
    for col in ("Emission - Production (kgCO2e)", "Emission - Usage (kgCO2e)", "Emission - Fin de Vie (kgCO2e)"):
        components_out[col] = pd.to_numeric(components_out[col], errors="coerce")
    components_out = components_out.dropna(subset=["ID Unique Composant"]).drop_duplicates(subset="ID Unique Composant")

    freight_by_component = {}
    for label, sheet, id_col, _ in FREIGHT_SOURCES:
        df = freight_sheets[sheet]
        weight_col = config.FREIGHT_WEIGHT_COLUMNS[sheet]
        ghg_col = "GHG_perunit (kgCO2e/kg)"
        agg = _agg_freight(df, id_col, weight_col, ghg_col)
        freight_by_component[label] = agg.rename(
            columns={
                id_col: FREIGHT_ID_COL,
                "poids_total": f"{label} - Poids total (t)",
                "emission_totale": f"{label} - Emission (kgCO2e)",
                "nb_lignes_achat": f"{label} - Nb lignes achat",
            }
        )

    relations = bom_out.merge(products_out, on="ID Unique Produit", how="left")
    relations = relations.merge(components_out, on="ID Unique Composant", how="left")
    relations = relations.rename(columns={"Nom_x": "Nom Produit", "Nom_y": "Nom Composant"})
    for label, agg in freight_by_component.items():
        relations = relations.merge(agg, on=FREIGHT_ID_COL, how="left")

    return {
        "produits": products_out,
        "relations": relations,
        "composants": components_out,
        "freight": freight_by_component,
    }

"""Lecture robuste des fichiers Excel SBM : détection d'en-tête, chargement en DataFrame."""

import pandas as pd


def load_sheet(path: str, sheet: str, header_row: int) -> pd.DataFrame:
    """Charge un onglet en DataFrame avec l'en-tête à la ligne donnée (1-based)."""
    df = pd.read_excel(path, sheet_name=sheet, header=header_row - 1)
    df = df.dropna(axis=1, how="all")
    df = df.loc[:, ~pd.Index(df.columns).duplicated()]
    df.columns = [str(c).strip() for c in df.columns]
    return df


def load_sample_products(path: str, sheet: str = None) -> pd.DataFrame:
    """Charge le fichier Echantillon - Produits à analyser.

    Détecte automatiquement l'onglet et la colonne SKU si non fournis.
    """
    xls = pd.ExcelFile(path)
    sheet = sheet or xls.sheet_names[0]
    df = pd.read_excel(path, sheet_name=sheet)
    df.columns = [str(c).strip() for c in df.columns]
    sku_col = next(
        (c for c in df.columns if "sku" in str(c).lower() or "product" in str(c).lower() or "id" in str(c).lower()),
        df.columns[0],
    )
    df = df.rename(columns={sku_col: "SKU"})
    df["SKU"] = df["SKU"].astype(str).str.strip()
    return df.drop_duplicates(subset="SKU")

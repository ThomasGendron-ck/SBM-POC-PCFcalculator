"""Lecture robuste des fichiers Excel SBM : détection d'en-tête, chargement en DataFrame."""

import pandas as pd


def _normalize_name(name: str) -> str:
    return "".join(str(name).lower().split()).replace("_", "")


def resolve_sheet_name(path: str, sheet: str) -> str:
    """Résout un nom d'onglet de manière tolérante (casse, espaces, underscores).

    Lève une erreur explicite listant les onglets disponibles si rien ne correspond.
    """
    xls = pd.ExcelFile(path)
    target = _normalize_name(sheet)
    for name in xls.sheet_names:
        if _normalize_name(name) == target:
            return name
    partial = [name for name in xls.sheet_names if target in _normalize_name(name)]
    if partial:
        return partial[0]
    available = ", ".join(xls.sheet_names)
    raise ValueError(
        f"Onglet '{sheet}' introuvable dans {path}. Onglets disponibles : {available}"
    )


def load_sheet(path: str, sheet: str, header_row: int) -> pd.DataFrame:
    """Charge un onglet en DataFrame avec l'en-tête à la ligne donnée (1-based).

    Le nom d'onglet est résolu de manière tolérante (casse, espaces, underscores).
    """
    sheet = resolve_sheet_name(path, sheet)
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

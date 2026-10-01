"""Lecture robuste des fichiers Excel SBM : détection d'en-tête, chargement en DataFrame."""

import pandas as pd


def _normalize_name(name: str) -> str:
    return "".join(str(name).lower().split()).replace("_", "")


def resolve_sheet_name(path, sheet: str) -> str:
    """Résout un nom d'onglet de manière tolérante (casse, espaces, underscores).

    `path` peut être un chemin ou un pd.ExcelFile déjà ouvert (évite de relire
    le classeur). Lève une erreur explicite listant les onglets disponibles.
    """
    if hasattr(path, "sheet_names"):
        names = list(path.sheet_names)
    else:
        names = list(pd.ExcelFile(path).sheet_names)
    target = _normalize_name(sheet)
    for name in names:
        if _normalize_name(name) == target:
            return name
    partial = [name for name in names if target in _normalize_name(name)]
    if partial:
        return partial[0]
    available = ", ".join(names)
    raise ValueError(
        f"Onglet '{sheet}' introuvable dans {path}. Onglets disponibles : {available}"
    )


def _excel_engine(path: str) -> str | None:
    """Moteur de lecture le plus rapide disponible (calamine), sinon défaut pandas."""
    try:
        import python_calamine  # noqa: F401

        return "calamine"
    except ImportError:
        return None


def load_sheet_columns(path: str, sheet: str, header_row: int, columns: list[str]) -> pd.DataFrame:
    """Charge un onglet puis n'en garde que les colonnes utiles.

    Utilise le moteur calamine quand il est installé (lecture plusieurs fois
    plus rapide qu'openpyxl). La sélection a lieu après normalisation des
    en-têtes (strip), les colonnes absentes ne sont pas une erreur ici :\n    l'appelant vérifie et liste les colonnes manquantes.
    """
    sheet = resolve_sheet_name(path, sheet)
    engine = _excel_engine(path)
    kwargs = {"engine": engine} if engine else {}
    df = pd.read_excel(path, sheet_name=sheet, header=header_row - 1, **kwargs)
    df = df.dropna(axis=1, how="all")
    df = df.loc[:, ~pd.Index(df.columns).duplicated()]
    df.columns = [str(c).strip() for c in df.columns]
    return df[[c for c in columns if c in df.columns]]


def load_sheet(path, sheet: str, header_row: int) -> pd.DataFrame:
    """Charge un onglet en DataFrame avec l'en-tête à la ligne donnée (1-based).

    `path` peut être un chemin ou un pd.ExcelFile déjà ouvert (lecture partagée
    d'un même classeur sans le re-parser). Nom d'onglet résolu tolérant.
    """
    sheet = resolve_sheet_name(path, sheet)
    if isinstance(path, pd.ExcelFile):
        df = path.parse(sheet_name=sheet, header=header_row - 1)
    else:
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

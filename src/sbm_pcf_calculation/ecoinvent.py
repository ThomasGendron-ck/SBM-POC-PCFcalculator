"""Matching matières sans FE -> ICV ecoinvent (GWP100 EF v3.1, cut-off v3.12).

Pour chaque composant sans facteur d'émission du Fichier de collecte, propose
un dataset ecoinvent :
1. règles de correspondance matière/désignation -> mots-clés de datasets
   (table MATCH_RULES, ordre = priorité, première règle qui matche gagne) ;
2. sélection géographique en cascade : pays du fournisseur (déduit du code
   BPSNUM/SUPPLIER_CODE, ex. EFR... -> FR) si ecoinvent le couvre, sinon
   RER -> RoW -> GLO ;
3. dans un dataset multi-produits, sélection du produit principal (celui dont
   le nom est contenu dans le nom de l'activité) plutôt qu'un co-produit.

L'indicateur retenu est la colonne « EF v3.1 / climate change / global warming
potential (GWP100) / kg CO2-Eq » du fichier Cut-off Cumulative LCIA, divisé
par le Reference Product Amount pour un FE par kg.
"""

import re

import pandas as pd

GEO_FALLBACK = ["RER", "RoW", "GLO"]

COUNTRY_GEO = {
    "FR": ["FR", "RER", "RoW", "GLO"],
    "DE": ["DE", "RER", "RoW", "GLO"],
    "IT": ["IT", "RER", "RoW", "GLO"],
    "ES": ["ES", "RER", "RoW", "GLO"],
    "GB": ["GB", "RER", "RoW", "GLO"],
    "NL": ["NL", "RER", "RoW", "GLO"],
    "BE": ["BE", "RER", "RoW", "GLO"],
    "CH": ["CH", "RER", "RoW", "GLO"],
    "PL": ["PL", "RER", "RoW", "GLO"],
    "US": ["US", "RNA", "RoW", "GLO"],
    "AD": ["RER", "RoW", "GLO"],
}

MATCH_RULES = [
    ("EAU (DEMIN|.*VILLE)", [r"market for tap water", r"water production, deionised"], "Eau potable / déminéralisée"),
    ("ACIDE ACETIQUE|HERBICLEAN", [r"market for acetic acid", r"acetic acid production"], "Acide acétique / désinfectant"),
    ("SILCOLAPSE|ANTIMOUSSE.*SILICO|POLYGLYCOL|EMERION|CYPREVERT|ADJUVANT|MOUILLANT|GENAGEN|SOPROPHOR|LUCRAMUL|DISPERSANT", [r"non-ionic surfactant production", r"market for non-ionic surfactant"], "Tensioactif non ionique"),
    ("PROXEL|ACTICIDE|BARDAC|BITREX|DENATONIUM|PREVENTOL|SPINOSAD|IMIDACLOPRID|PYRETHRUM|DELFIN|THURINGIENSIS|FT REZIST", [r"market for pesticide, unspecified"], "Biocide / pesticide (proxy)"),
    ("SULFATE DE FER|CITRATE FERR", [r"iron sulfate production", r"market for iron sulfate"], "Sulfate de fer"),
    ("SULFATE DE MAGNESIUM|EPSOMITE", [r"magnesium sulfate production", r"market for magnesium sulfate"], "Sulfate de magnésium"),
    ("CARBONATE DE MAGNESIE|MAGNESIE", [r"magnesite|magnesium carbonate production", r"dolomite production"], "Carbonate de magnésium"),
    ("SULFATE DE MANGANESE", [r"manganese sulfate", r"manganese carbonate production"], "Sulfate de manganèse (proxy carbonate)"),
    ("SULFATE DE ZINC", [r"primary zinc production", r"zinc sulphate production"], "Sulfate de zinc (proxy zinc)"),
    ("SULFATE DE POTASSIUM|POTASSE ORGANIQUE", [r"potassium sulfate production", r"market for potassium sulfate"], "Sulfate de potassium / potasse"),
    ("ARMICARB|BICARBONATE DE POTASSIUM", [r"market for potassium carbonate", r"potassium carbonate production"], "Bicarbonate de potassium (proxy carbonate)"),
    ("HUILE DE COLZA|RAPESEED OIL(?! METHYL)", [r"market for rape oil, crude", r"market for vegetable oil, refined"], "Huile de colza"),
    ("RAPESEED OIL METHYL|FAME", [r"market for fatty acid methyl ester"], "Ester méthylique d'huile (EMHC)"),
    ("HUILE VASELINE|VASSELINE|PARAFFIN", [r"lubricating oil production", r"petroleum refinery operation"], "Huile minérale (base pétrole)"),
    ("IONOL|BHT", [r"market for toluene, liquid", r"toluene production"], "BHT (proxy toluène)"),
    ("MONO PROPYLENE GLYCOL|PROPYLENE GLYCOL", [r"market for propylene glycol", r"propylene glycol production"], "Propylène glycol"),
    ("ARGIREC|NITRATE DE SODIUM", [r"sodium nitrate production", r"market for sodium nitrate"], "Nitrate de sodium"),
    ("LECITHIN", [r"market for soybean oil", r"soybean oil production"], "Lécithine (proxy huile de soja)"),
    ("D-LIMONENE|LIMONENE", [r"orange production, processing grade", r"orange production", r"terpene"], "D-limonène (proxy agrume)"),
    ("COLORANT|BLEU POUDRE|PIGMENT", [r"market for chemical, organic, unspecified"], "Colorant (proxy chimie organique)"),
    ("BICARBONATE", [r"market for sodium bicarbonate"], "Bicarbonate de sodium"),
    ("PET12|PET\\b|BOTEV", [r"polyethylene terephthalate production, granulate, bottle grade", r"polyethylene terephthalate production"], "PET"),
    ("HDPE|SOFTHDPE", [r"polyethylene production, high density", r"market for polyethylene, high density"], "PE-HD"),
    ("PP\\b|DOSEUR|CAPCR|BOUCHON|TRIGGER|VALVE|ATTACHE|ACCTA|ACCSP|ACC.*COLOUR", [r"polypropylene production", r"market for polypropylene"], "PP"),
    ("BOTPE|PEPA|BIDON|LAIZE", [r"polyethylene production, high density", r"market for polyethylene, high density"], "PE-HD"),
    ("LDPE|SAC POUBELLE|STRETCH HOOD|PALST", [r"polyethylene production, low density", r"packaging film production, low density"], "PE-LD / film"),
    ("FILM PET12.*PE50|LAMINATE|FILM M8534", [r"packaging film production", r"polyethylene terephthalate production"], "Film complexe"),
    ("HYDROSOLUBLE|FLXRL|PVOH", [r"packaging film production, low density"], "Film hydrosoluble (proxy film PE-LD)"),
    ("CASSH|CAISSE|CARTON|OUTERCASE", [r"corrugated board box production", r"market for corrugated board box"], "Caisse carton ondulé"),
    ("ETIQUETTE|LAB\\b|STICKER", [r"graphic paper production", r"offset printing"], "Étiquette (papier imprimé)"),
    ("NPK|N-P-K|BRODI|EVER 7|OSYR|BULK SUPPLIED", [r"NPK \(15-15-15\) fertiliser production", r"market for NPK"], "Engrais NPK"),
    ("PHOSPHATE NATUREL", [r"market for phosphate rock", r"single superphosphate"], "Phosphate naturel"),
    ("THIOVIT|SOUFRE", [r"sulfur production", r"market for sulfur"], "Soufre (fongicide)"),
    ("BORDO|BOUILLIE|CUIVRE|MICROSPRAY|ROC-MICRO", [r"copper sulfate production", r"market for copper sulfate"], "Cuivre (sulfate/bouillie)"),
    ("ARGIVERT|ARGILE|BENTONITE|TERRE DE FEU", [r"bentonite quarry operation", r"activated bentonite production"], "Argile / bentonite"),
    ("BLACK SOAP|SAVON", [r"soap production", r"market for soap"], "Savon noir"),
    ("CHICKEN MANURE|FUMIER|COMPOST", [r"market for compost", r"compost"], "Fumier / compost"),
    ("SEQUESTRENE|CHELATE", [r"market for EDTA", r"EDTA"], "Chélateur (proxy EDTA)"),
    ("SOSIRYL|DOSE\\b", [r"injection moulding"], "Dose plastique (transformation)"),
]

GWP_COLUMN = ("EF v3.1", "climate change", "global warming potential (GWP100)", "kg CO2-Eq")


def supplier_country(code: str | None) -> str | None:
    """Déduit le pays du code fournisseur (EFR06853 -> FR, XDE00017 -> DE)."""
    if code is None or pd.isna(code):
        return None
    c = str(code).strip().upper()
    if len(c) >= 5 and c[0] in ("E", "X"):
        return c[1:3]
    return None


def load_lcia_gwp(lcia_path: str) -> pd.DataFrame:
    """Charge le fichier Cut-off Cumulative LCIA et retourne un FE par kg."""
    lcia = pd.read_excel(lcia_path, sheet_name="LCIA", header=[0, 1, 2, 3])
    gwp = lcia[GWP_COLUMN]
    base = pd.DataFrame(
        {
            "name": lcia.iloc[:, 1],
            "geo": lcia.iloc[:, 2],
            "prod": lcia.iloc[:, 3],
            "unit": lcia.iloc[:, 4],
            "amount": lcia.iloc[:, 5],
        }
    )
    base["gwp_per_unit"] = pd.to_numeric(gwp, errors="coerce") / pd.to_numeric(base["amount"], errors="coerce")
    return base[base["unit"] == "kg"]


def _main_product(matches: pd.DataFrame) -> pd.Series:
    """Dans un dataset multi-produits, garde le produit principal."""
    main = matches[matches.apply(lambda r: str(r["prod"]).lower() in str(r["name"]).lower(), axis=1)]
    pool = main if not main.empty else matches
    market = pool[pool["name"].str.startswith("market for ")]
    if not market.empty:
        pool = market
    return pool.sort_values("name").iloc[0]


def _pick_geo(matches: pd.DataFrame, country: str | None) -> pd.Series | None:
    order = COUNTRY_GEO.get(country, GEO_FALLBACK)
    for geo in order:
        sub = matches[matches["geo"] == geo]
        if not sub.empty:
            return _main_product(sub)
    return None


def _keyword_matches(base: pd.DataFrame, kw: str,
                    cache: dict[str, pd.DataFrame] | None = None) -> pd.DataFrame:
    """Rows whose activity name matches a keyword (cached per run when provided)."""
    if cache is not None and kw in cache:
        return cache[kw]
    matches = base[base["name"].str.contains(kw, case=False, na=False, regex=True)]
    if cache is not None:
        cache[kw] = matches
    return matches


def pick_dataset(base: pd.DataFrame, keywords: list[str], country: str | None,
                 cache: dict[str, pd.DataFrame] | None = None) -> pd.Series | None:
    for kw in keywords:
        matches = _keyword_matches(base, kw, cache)
        if matches.empty:
            continue
        best = _pick_geo(matches, country)
        if best is not None:
            return best
    return None


def missing_fe_components(collecte: pd.DataFrame) -> pd.DataFrame:
    """Composants uniques sans FE, avec les champs de la spec MissingEF_Matching."""
    no_fe = collecte[collecte["RM EF Value"].isna() & collecte["Component SKU"].notna()].copy()
    no_fe["pays"] = no_fe["Supplier code"].map(supplier_country)
    aggs = {
        "design": ("Component Designation", "first"),
        "matiere": ("Raw Material", "first"),
        "fournis": ("Supplier Name", "first"),
        "pays": ("pays", "first"),
        "sage": ("Category Code", "first"),
        "cfdesc": ("Category description", "first"),
    }
    optional = {
        "code_fournis": ("Supplier code", "first"),
        "pack": ("Pack unit box", "first"),
        "recycle": ("Recycled %", "first"),
        "carbon_cat": ("Carbon category", "first"),
        "uvp": ("UVP description", "first"),
        "matiere_cf": ("Raw Material - Carbon Footprint", "first"),
    }
    for key, spec in optional.items():
        if spec[0] in no_fe.columns:
            aggs[key] = spec
    return no_fe.groupby("Component SKU").agg(**aggs).reset_index()


def match_missing_fe(collecte: pd.DataFrame, lcia_base: pd.DataFrame) -> pd.DataFrame:
    """Matche chaque composant sans FE avec un dataset ecoinvent.

    Accepte soit la collecte brute (lignes par produit x composant), soit
    la liste de composants uniques déjà produite par missing_fe_components
    (session.ef_matching / identify_missing_factors).
    """
    if "design" in collecte.columns and "RM EF Value" not in collecte.columns:
        comps = collecte
    else:
        comps = missing_fe_components(collecte)
    keyword_cache: dict[str, pd.DataFrame] = {}
    rows = []
    for _, comp in comps.iterrows():
        text = re.sub(
            r"\s+",
            " ",
            " ".join(str(x) for x in [comp["design"], comp["sage"], comp["cfdesc"], comp["matiere"]]).upper(),
        ).strip()
        matched, rule = None, None
        for pattern, keywords, comment in MATCH_RULES:
            if re.search(pattern, text):
                matched = pick_dataset(lcia_base, keywords, comp["pays"], keyword_cache)
                if matched is not None:
                    rule = comment
                    break
        rows.append(
            {
                "Component SKU": comp["Component SKU"],
                "Designation": comp["design"],
                "Matière": comp["matiere"],
                "Fournisseur": comp["fournis"],
                "Pays fournisseur": comp["pays"],
                "Dataset ecoinvent": None if matched is None else matched["name"],
                "Produit ecoinvent": None if matched is None else matched["prod"],
                "Géographie": None if matched is None else matched["geo"],
                "FE proposé (kg CO2e/kg)": None if matched is None else round(float(matched["gwp_per_unit"]), 4),
                "Règle de matching": rule,
                "Statut": "NON MATCHÉ" if matched is None else "MATCHÉ",
            }
        )
    return pd.DataFrame(rows)

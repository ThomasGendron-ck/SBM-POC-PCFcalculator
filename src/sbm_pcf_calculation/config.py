"""Centralisation des règles d'extraction issues de l'onglet CalculPCF du fichier SBM Material and Packaging.

Chaque entrée mappe un champ du PCF vers sa source : fichier, onglet, colonne.
Ligne en-tête = ligne Excel (1-based) où se trouve l'en-tête réel de la table.
"""

from dataclasses import dataclass


@dataclass(frozen=True)
class FieldMapping:
    section: str
    field: str
    source_file: str
    source_sheet: str
    source_column: str
    header_row: int


FILE_MATERIAL = "SBM_Material_Packaging.xlsx"
FILE_FREIGHT = "SBM_Freight.xlsx"

EXTRACTION_RULES: list[FieldMapping] = [
    FieldMapping("Produit Final", "ID Unique Produit", FILE_MATERIAL, "MasterBase_Products", "SKU", 12),
    FieldMapping("Produit Final", "Nom", FILE_MATERIAL, "MasterBase_Products", "SKU Designation", 12),
    FieldMapping("Produit Final", "Super Segment", FILE_MATERIAL, "MasterBase_Products", "Des super segment", 12),
    FieldMapping("Produit Final", "Segment", FILE_MATERIAL, "MasterBase_Products", "Segment_DES_TSI2", 12),
    FieldMapping("Produit Final", "Sub Segment", FILE_MATERIAL, "MasterBase_Products", "Sub Subsegment_DES_TSI3", 12),
    FieldMapping("Relation Produit - Composant", "ID Unique Produit", FILE_MATERIAL, "MasterBase_BOM", "ITMREF", 12),
    FieldMapping("Relation Produit - Composant", "ID Unique Composant", FILE_MATERIAL, "MasterBase_BOM", "CPNITMREF", 12),
    FieldMapping("Relation Produit - Composant", "Quantité de composant dans produit", FILE_MATERIAL, "MasterBase_BOM", "BOMQTY", 12),
    FieldMapping("Composant", "ID Unique Composant", FILE_MATERIAL, "CK_MaterialPurchase", "PRODUCT", 12),
    FieldMapping("Composant", "Nom", FILE_MATERIAL, "CK_MaterialPurchase", "PRODUCT_NAME", 12),
    FieldMapping("Composant", "Type composant", FILE_MATERIAL, "CK_MaterialPurchase", "Catégorie ACV", 12),
    FieldMapping("Composant", "Emission - Phase - Production", FILE_MATERIAL, "CK_MaterialPurchase", "Qty_Prod_GHG (kgCO2e)", 12),
    FieldMapping("Composant", "Emission - Phase - Usage", FILE_MATERIAL, "CK_MaterialPurchase", "Qty_Use_GHG (kgCO2e)", 12),
    FieldMapping("Composant", "Emission - Phase - Fin de Vie", FILE_MATERIAL, "CK_MaterialPurchase", "Qty_EoL_GHG (kgCO2e)", 12),
    FieldMapping("Composant - Fret - IRIS", "ID Unique Composant", FILE_FREIGHT, "Freight_Raw_In_IRIS", "Référence Article", 13),
    FieldMapping("Composant - Fret - IRIS", "Emission - Phase - Fret Amont", FILE_FREIGHT, "Freight_Raw_In_IRIS", "GHG_perunit (kgCO2e/kg)", 13),
    FieldMapping("Composant - Fret - Fert", "ID Unique Composant", FILE_FREIGHT, "Freight_Raw_In_Fert", "Référence Article", 13),
    FieldMapping("Composant - Fret - Fert", "Emission - Phase - Fret Amont", FILE_FREIGHT, "Freight_Raw_In_Fert", "GHG_perunit (kgCO2e/kg)", 13),
    FieldMapping("Composant - Fret - LS Europe", "ID Unique Composant", FILE_FREIGHT, "Freight_Raw_In_LSEur", "Product", 13),
    FieldMapping("Composant - Fret - LS Europe", "Emission - Phase - Fret Amont", FILE_FREIGHT, "Freight_Raw_In_LSEur", "GHG_perunit (kgCO2e/kg)", 13),
]

# Colonnes de poids utilisées pour convertir les FE fret en émissions absolues (par ligne d'achat)
FREIGHT_WEIGHT_COLUMNS = {
    "Freight_Raw_In_IRIS": "Weigh Final (in to)",
    "Freight_Raw_In_Fert": "Weigh Final (in to)",
    "Freight_Raw_In_LSEur": "Weigh Final (in to)",
}

# Correspondance CATEGORIE SAGE -> (CATEGORIE DESCRIPTION, CF_Category), issue de l'onglet "Catégorisation Cmpt"
# du fichier "POC Calculateur - règles de calcul.xlsx".
SAGE_CATEGORY_MAP: dict[str, tuple[str, str]] = {
    "EXPEN": ("Bought overheads", "INDIRECT PURCHASES"),
    "PDTAC": ("Accessories", "PACKAGING"),
    "PDTAE": ("Aerosol", "PACKAGING"),
    "PDTBO": ("Bottles", "PACKAGING"),
    "PDTCL": ("Closures", "PACKAGING"),
    "PDTCO": ("Containers", "PACKAGING"),
    "PDTDE": ("Composants Devices", "PACKAGING"),
    "PDTDI": ("Composants displays", "PACKAGING"),
    "PDTDP": ("SKU display", "PACKAGING"),
    "PDTEC": ("Packaging Components CARTON", "PACKAGING"),
    "PDTEL": ("LOGISTICS components", "PACKAGING"),
    "PDTEM": ("Packaging others materials", "PACKAGING"),
    "PDTEP": ("Packaging Components PLASTIC", "PACKAGING"),
    "PDTET": ("Etiquettes", "PACKAGING"),
    "PDTFL": ("Flexibles", "PACKAGING"),
    "PDTNV": ("Non valuated products", "PACKAGING"),
    "PDTOU": ("Outercase", "PACKAGING"),
    "GAPRD": ("Gamme Production", "Raw Material"),
    "PDTBC": ("Bulks", "Raw Material"),
    "PDTMA": ("RAW MATERIALS ACTIV INGREDIENT", "Raw Material"),
    "PDTMH": ("RM & SOL high price volatility", "Raw Material"),
    "PDTMP": ("Matières Premières", "Raw Material"),
    "PDTPF": ("Finish Goods batch number", "Raw Material"),
    "PDTSF": ("SEMI-FINISHED PRODUCTs", "Raw Material"),
    "PDTSO": ("SOLVENTS", "Raw Material"),
    "PDTRC": ("Produits recyclables", "TO BE EXCLUDED"),
    "PDTTI": ("Produits tiers non valorisés", "TO BE EXCLUDED"),
    "REFAC": ("sold overheads", "TO BE EXCLUDED"),
    "RFCSV": ("reinvoicing services/ customer", "TO BE EXCLUDED"),
    "ZGAPI": ("Reserved Variance Items (PUR)", "TO BE EXCLUDED"),
    "PDTSP": ("PARTIEL TOLLING FEE", "TOLLING"),
    "PDTST": ("FULL SERVICE TOLLING FEE", "TOLLING"),
    "PDTSX": ("FULL TOLLING FEE", "TOLLING"),
    "PDTAR": ("Trading goods", "Trading Goods"),
}

# Catégories de composants (référentiel Categorie_Match)
CATEGORY_CODES = {
    "MP": "Matières Premières",
    "PF": "Trading goods",
    "PDTMP": "Matières Premières",
    "PDTPF": "Finish Goods batch number",
    "PDTEL": "LOGISTICS components",
    "PDTSO": "SOLVENTS",
    "PDTFL": "Flexibles",
    "PDTCO": "Containers",
    "PDTEC": "Packaging Components CARTON",
    "PDTOU": "Outercase",
    "PDTBO": "Bottles",
    "PDTCL": "Closures",
    "PDTEP": "Packaging Components PLASTIC",
    "PDTDI": "Composants displays",
    "PDTAR": "Trading goods",
    "PDTBC": "Bulks",
    "PDTET": "Etiquettes",
    "PDTAC": "Accessories",
    "PDTMA": "RAW MATERIALS ACTIV INGREDIENT",
    "PDTDE": "Composants Devices",
    "PDTSF": "SEMI-FINISHED PRODUCTs",
    "PDTTI": "Produits tiers non valorisés",
    "PDTST": "FULL SERVICE TOLLING FEE",
    "PDTSP": "PARTIEL TOLLING FEE",
    "PDTSX": "FULL TOLLING FEE",
    "PDTEM": "Packaging others materials",
    "PDTDP": "SKU display",
}

# Description des flags de validation pour l'onglet Flags du livrable (source : Clé de détermination)
FLAG_DESCRIPTIONS = {
    "PRODUIT INTROUVABLE": (
        "La référence est introuvable dans la Master Base (MB_Products ou MB_BOM). "
        "Vérifier la saisie de la référence fournisseur LM auprès de SBM."
    ),
    "FOURNISSEUR INTROUVABLE": (
        "Aucun fournisseur rattaché : la colonne BPSNAM/BPSNUM est absente de la MB_Products "
        "et aucune ligne d'achat CK_MaterialPurchase n'existe pour ce composant. "
        "Le calcul reste possible, mais les trajets fret ne peuvent pas être rattachés."
    ),
    "QTE NULLE (oubli de suppression ?)": (
        "La quantité UVQTY dans la nomenclature (BOM alternative 1) est nulle ou vide. "
        "Probable oubli de suppression de la ligne : à vérifier avec SBM."
    ),
    "POIDS > 120% ITEM WEIGHT": (
        "Le poids brut du composant (Gross Weight) dépasse 120 % de son poids unitaire (Net Weight). "
        "Le packaging primaire est peut-être compté en double : à vérifier."
    ),
    "PAS DE MATIERE ASSOCIEE": (
        "Aucune matière associée au composant : ZCODMAT2 vide dans la MB_Products "
        "et aucun RawMat_SubFamily dans la matrice de consolidation CK. "
        "Sans matière, aucun facteur d'émission ne peut être rattaché."
    ),
    "PAS DE FE ASSOCIE A LA MATIERE": (
        "Aucun facteur d'émission trouvé pour la matière du composant "
        "(ni dans CK_MaterialPurchase, ni dans CK_EF_Packaging). "
        "Le PCF du produit est incomplet : rechercher le FE en priorité."
    ),
    "PAS DE BOM": (
        "Le produit n'a pas de nomenclature (aucune ligne BOM alternative 1 active). "
        "Le PCF ne peut pas être calculé pour ce produit."
    ),
    "PAS DE TRAJET FRET": (
        "Aucun trajet fret trouvé pour le couple composant/fournisseur dans le fichier "
        "Freight (onglets CK_FreightConsolidation et Freight_Raw_In_*). "
        "Freight GHG non calculé."
    ),
}

# ---------------------------------------------------------------------------
# Structure du Fichier de collecte (règles v0.7, onglet Clé de détermination)
# ---------------------------------------------------------------------------

# Mise en page du livrable : en-têtes en ligne 6 (tableau démarré en A6),
# volets figés en C7, filtres automatiques sur la ligne d'en-tête.
# Hauteur de la ligne d'en-tête : 145 px (= 108,75 points Excel).
# Largeur des colonnes GROUPING : 71 px (= 9,43 unités de largeur Excel).
# Lignes 1 à 4 groupées (Plans > Grouper) et repliées à l'ouverture du fichier.
COLLECTE_HEADER_ROW = 6
COLLECTE_FREEZE_PANES = "C7"
COLLECTE_HEADER_ROW_HEIGHT = 108.75
COLLECTE_TOP_GROUP_ROWS = (1, 4)

# Couleurs des blocs : reproduisent les codes de remplissage de la colonne A de la
# Clé de détermination v0.7 (thème Office : accent1=156082, accent3=196B24,
# accent4=0F9ED5) avec les tints exacts du fichier (accent1 -0.5 pour les champs
# PCF du produit ; accent3 -0.25/0/0.4/0.6/0.8 pour les blocs composants ;
# accent4 0.4 pour le fret). Le bloc
# « Impact Production SBM et Packaging » n'a pas encore de colonnes définies.
BLOCK_COLORS: dict[str, str] = {
    "Produit PCF": "FF0A3041",
    "Produit": "FF156082",
    "Composants": "FF13501B",
    "Description du composant": "FF196B24",
    "PCF fournisseur": "FF75A67C",
    "Impact matière fournisseur": "FFA3C4A7",
    "Impact fabrication fournisseur": "FFD1E1D3",
    "Impact Appro Transport": "FF6FC5E6",
    "Flag": "FFFFC000",
}

# Colonnes marqueurs GROUPING (type GROUPING de la Clé v0.7) : insérées comme des
# colonnes à part entière à la position qu'elles occupent dans la Clé, en-tête en
# texte vertical (rotation vers le le haut) avec renvoi à la ligne automatique.
# Deux marqueurs portent deux fois le même nom dans la Clé (rows 41 et 75) : le
# second occurrence est suffixée «  (2) » en interne et affichée avec son nom exact
# dans Excel (EXCEL_HEADER_MAP).
# Couleurs de police des en-têtes : blanc sur les blocs foncés (produit bleu,
# composants vert foncé), noir sur les blocs clairs (comme la colonne B de la
# Clé de détermination v0.7).
BLOCK_FONT_COLORS: dict[str, str] = {
    "Produit PCF": "FFFFFFFF",
    "Produit": "FFFFFFFF",
    "Composants": "FFFFFFFF",
    "Description du composant": "FFFFFFFF",
    "PCF fournisseur": "FF000000",
    "Impact matière fournisseur": "FF000000",
    "Impact fabrication fournisseur": "FF000000",
    "Impact Appro Transport": "FF000000",
    "Flag": "FF000000",
}

# Colonne Flag : fond orange (toute la colonne, en-tête et données).
FLAG_FILL = "FFFFC000"

# Onglet Synthèse : une ligne par produit, décomposition du PCF par phase
# (catégories carbon du Fichier de collecte). Use et End of Life ne sont pas
# encore alimentées (périmètre cradle-to-gate) mais leurs colonnes existent.
SYNTHESE_PHASES: list[tuple[str, str | None]] = [
    ("Raw Material", "Raw Material"),
    ("Packaging", "PACKAGING"),
    ("Transformation", None),
    ("Freight", None),
    ("Use", None),
    ("End of Life", None),
]

GROUPING_MARKER_COLUMNS = [
    "Product Details",
    "Component details",
    "Component Material",
    "Component Material (2)",
    "RM GHG details",
    "Transformation Details",
    "Transformation Details (2)",
]
GROUPING_MARKER_WIDTH = 9.43

# Blocs de colonnes du Fichier de collecte : (nom du bloc, colonnes du bloc).
# Les noms suivent la colonne B « Nom du champ » de la Clé de détermination v0.7 :
# PCF Value / Unité / DQR Product / PDS Product sont remontés dans le bloc Produit
# (juste après Product Designation), RM PDS Activity Data est déplacé dans le bloc
# Composants (juste après Quantity) et le bloc transformation passe à 12 champs
# (FE du procédé + FE de l'énergie).
# Les lignes de type GROUPING sont insérées comme des colonnes à part entière
# (GROUPING_MARKER_COLUMNS) à la position qu'elles occupent dans la Clé ; les
# colonnes de données qui les suivent sont mises en retrait d'un niveau dans les
# groupements Excel (Plans > Grouper).
# Les colonnes absentes de la Clé v0.7 sont conservées sous leur nom historique :
# Part du composant dans le produit et Flag.
BLOCK_HEADERS: list[tuple[str, list[str]]] = [
    ("Produit", [
        "Product SKU",
        "Product Designation",
    ]),
    ("Produit PCF", [
        "PCF Value",
        "PCF Unit",
        "DQR Product",
        "PDS Product",
    ]),
    ("Produit", [
        "Product Category Code",
        "Product Category description",
        "Pack Unit Box",
        "Product Net Weight",
        "Product Gross Weight",
        "Product Supplier Code",
        "Product Supplier Name",
    ]),
    ("Product Details", ["Product Details"]),
    ("Composants", [
        "Component SKU",
        "Component Designation",
        "Category Code",
        "Category description",
        "Carbon category",
        "Pack unit box",
        "Quantity",
        "RM PDS Activity Data",
    ]),
    ("Component details", ["Component details"]),
    ("Description du composant", [
        "Net Weight",
        "Net Weight Unit",
        "Stock unit",
        "Gross Weight",
        "Gross Weight Unit",
        "Part du composant dans le produit",
        "Supplier code",
        "Supplier Name",
        "Raw Material",
        "Raw Material - Carbon Footprint",
        "UVP description",
        "Recycled %",
        "Scrap Rate",
    ]),
    ("Component Material", ["Component Material"]),
    ("PCF fournisseur", [
        "Supplier PCF value",
        "Supplier PCF Unit",
        "Supplier PDS",
        "Supplier DQR",
        "Supplier PCF source",
        "Supplier PCF external review",
    ]),
    ("Component Material", ["Component Material (2)"]),
    ("Impact matière fournisseur", [
        "RM EF Name",
        "RM EF Value",
        "RM EF Unit",
        "RM EF Source",
        "RM EF PDS",
        "RM PDS value",
        "Prod_EF_Geography",
        "RM GEO DQR",
        "RM TECH DQR",
        "RM TEMP DQR",
        "RM DQR value",
        "RM GHG",
        "RM GHG Unit",
    ]),
    ("RM GHG details", ["RM GHG details"]),
    ("Impact fabrication fournisseur", [
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
        "Transformation GHG",
        "Transformation GHG Unit",
    ]),
    ("Transformation Details", ["Transformation Details"]),
    ("Impact Appro Transport", [
        "Freight Supplier Code",
        "Freight Supplier Name",
        "Freight Route",
        "Freight Transportation Mode",
        "Freight GHG",
        "Freight GHG Unit",
    ]),
    ("Transformation Details", ["Transformation Details (2)"]),
    ("Flag", ["Flag"]),
]

# En-têtes Excel du livrable : la Clé de détermination réutilise les mêmes noms
# pour le produit et le composant ('Category Code', 'Category description',
# 'Net Weight', 'Gross Weight', 'Supplier Name') ; l'unité du PCF produit
# s'appelle « Unité » dans la Clé v0.7. En interne, les colonnes du bloc produit
# sont préfixées « Product » ; ce renommage est appliqué à l'écriture du fichier
# Excel (les doublons d'en-têtes sont volontaires, comme dans la Clé).
EXCEL_HEADER_MAP: dict[str, str] = {
    "PCF Unit": "Unité",
    "Component Material (2)": "Component Material",
    "Transformation Details (2)": "Transformation Details",
    "Product Category Code": "Category Code",
    "Product Category description": "Category description",
    "Product Net Weight": "Net Weight",
    "Product Gross Weight": "Gross Weight",
    "Product Supplier Code": "Supplier Code",
    "Product Supplier Name": "Supplier Name",
}

# Groupements de colonnes Excel (Plans > Grouper) : (première colonne, dernière colonne)
# incluses, issu des lignes GROUPING de la Clé de détermination v0.7 (le groupe va
# de la colonne nommée dans l'instruction jusqu'à la colonne précédant le marqueur).
COLUMN_GROUPS: list[tuple[str, str]] = [
    ("Product Category Code", "Product Supplier Name"),
    ("Category Code", "RM PDS Activity Data"),
    ("Net Weight", "Recycled %"),
    ("Supplier PCF value", "Supplier PCF external review"),
    ("RM EF Name", "RM DQR value"),
    ("Transformation Process Name", "Transformation Energy EF Source"),
    ("Freight Supplier Code", "Freight Transportation Mode"),
]

# Formats d'affichage des cellules numériques, issus de la colonne D « Type de
# données » de la Clé de détermination v0.7 : DECIMAL(10,3) -> 0,000,
# DECIMAL(10,5) -> 0,00000, DECIMAL(3;0) et BINARY/BINAIRE -> entier.
# Transformation GHG est typé TEXT dans la Clé v0.7 (probable coquille) : la
# valeur reste calculée et formatée comme un DECIMAL(10,3).
NUMBER_FORMATS: dict[str, str] = {
    "PCF Value": "0.000",
    "DQR Product": "0.000",
    "PDS Product": "0.000",
    "Pack Unit Box": "0.000",
    "Product Net Weight": "0.000",
    "Product Gross Weight": "0.000",
    "Pack unit box": "0.000",
    "Quantity": "0.000",
    "RM PDS Activity Data": "0",
    "Net Weight": "0.000",
    "Gross Weight": "0.000",
    "Part du composant dans le produit": "0.000",
    "Recycled %": "0",
    "Supplier PCF value": "0.00000",
    "Supplier PDS": "0.000",
    "Supplier DQR": "0.000",
    "Supplier PCF external review": "0",
    "RM EF Value": "0.00000",
    "RM EF PDS": "0",
    "RM PDS value": "0.000",
    "RM GEO DQR": "0.000",
    "RM TECH DQR": "0.000",
    "RM TEMP DQR": "0.000",
    "RM DQR value": "0.000",
    "RM GHG": "0.000",
    "Transformation Process EF Value": "0.00000",
    "Transformation Energy EF Value": "0.00000",
    "Transformation GHG": "0.000",
    "Freight GHG": "0.000",
    "Scrap Rate": "0.00000",
}

# Fret : structure des onglets Freight_Raw_In_* du fichier SBM Freight.
FREIGHT_SHEETS: dict[str, dict] = {
    "IRIS": {
        "sheet": "Freight_Raw_In_IRIS",
        "header_row": 13,
        "component_col": "Référence Article",
        "supplier_col": "Code Fournisseur / Facturé",
    },
    "LS Eu": {
        "sheet": "Freight_Raw_In_LSEur",
        "header_row": 13,
        "component_col": "Product",
        "supplier_col": "Supplier",
    },
    "Fert": {
        "sheet": "Freight_Raw_In_Fert",
        "header_row": 14,
        "component_col": "Référence Article",
        "supplier_col": "Code Fournisseur / Facturé",
    },
}

FREIGHT_CK_SHEET = "CK_FreightConsolidation"
FREIGHT_CK_HEADER_ROW = 12
FREIGHT_GHG_PERUNIT_COL = "GHG_perunit (kgCO2e/kg)"
FREIGHT_UNIQUEKEY_COL = "UniqueKey"
FREIGHT_ADDRESSKEY_COL = "ADDRESS_KEY"
FREIGHT_MODE_COL = "Transportation_Mode"
FREIGHT_FRET_TYPE_COL = "FRET TYPE"
FREIGHT_FRET_TYPE_PREFIX = "Inbound freight (suppliers) - "

UNIT_GHG_TRANSFORMATION = "kgCO2e/composant"
UNIT_GHG_TRANSPORT = "kgCO2e/composant"

# Colonnes de type texte : forcées en object pour éviter les conflits de dtype
# pandas quand une unité (chaîne) est assignée après création (colonne d'abord vide).
OBJECT_COLUMNS = {
    "Transformation GHG Unit",
    "Freight GHG Unit",
    "RM GHG Unit",
    "PCF Unit",
    "Net Weight Unit",
    "Gross Weight Unit",
    "Stock unit",
    "Transformation Process Name",
    "Transformation Process EF Name",
    "Transformation Process EF Unit",
    "Transformation EF Source",
    "Transformation Energy Name",
    "Transformation Energy EF Name",
    "Transformation Energy EF Unit",
    "Transformation Energy EF Source",
}

# Fichier de saisie de transformation (rempli par le responsable production SBM).
SAISIE_TRANSFORMATION_SHEET = "Saisie transformation"
SAISIE_GUIDE_SHEET = "Guide"
SAISIE_COLUMNS = [
    "Product SKU",
    "Component SKU",
    "Supplier Name",
    "Transformation Process Name",
    "Transformation Process EF Value",
    "Transformation Process EF Unit",
    "Transformation EF Source",
    "Transformation Energy Name",
    "Transformation Energy EF Value",
    "Transformation Energy EF Unit",
    "Transformation Energy EF Source",
    "Scrap Rate",
]
SAISIE_GUIDE_ROWS: list[tuple[str, str]] = [
    ("Product SKU", "Référence du produit fini (pré-rempli par l'outil, à conserver telle quelle)."),
    ("Component SKU", "Référence du composant (pré-rempli par l'outil)."),
    ("Supplier Name", "Fournisseur du composant (pré-rempli par l'outil, à corriger si nécessaire)."),
    (
        "Transformation Process Name",
        "Procédé de transformation appliqué au composant chez le fournisseur "
        "(ex. : injection, extrusion, soufflage, mixage, formulation). "
        "Laisser vide si le composant n'est pas transformé.",
    ),
    (
        "Transformation Process EF Value",
        "Facteur d'émission du procédé de transformation en kgCO2e par kg de "
        "composant transformé (ex. : ICV ecoinvent du procédé). Si vous ne le "
        "connaissez pas, laissez-le vide.",
    ),
    (
        "Transformation Process EF Unit",
        "Unité du facteur d'émission du procédé (ex. : kgCO2e/kg).",
    ),
    (
        "Transformation EF Source",
        "D'où vient le facteur d'émission du procédé (ex. : fiche technique "
        "fournisseur, ICV ecoinvent, estimation interne).",
    ),
    (
        "Transformation Energy Name",
        "Source d'énergie utilisée pour le procédé (ex. : électricité, gaz "
        "naturel, vapeur). Laisser vide si inconnue.",
    ),
    (
        "Transformation Energy EF Value",
        "Facteur d'émission de l'énergie en kgCO2e par unité d'énergie "
        "(ex. : kgCO2e/kWh).",
    ),
    (
        "Transformation Energy EF Unit",
        "Unité du facteur d'émission de l'énergie (ex. : kgCO2e/kWh).",
    ),
    (
        "Transformation Energy EF Source",
        "D'où vient le facteur d'émission de l'énergie (ex. : ICV ecoinvent, "
        "base ADEME, mix électrique pays).",
    ),
    (
        "Scrap Rate",
        "Taux de perte / rebuts du composant lors de sa mise en œuvre, en "
        "décimal (ex. : 0,03 pour 3 %). Multiplie le poids effectif du composant "
        "dans le calcul du GHG Transformation. Laisser vide si inconnu "
        "(aucun facteur appliqué par défaut).",
    ),
]

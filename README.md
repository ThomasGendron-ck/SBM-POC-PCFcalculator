# SBM PCF Calculation

Outil de calcul des PCF (Product Carbon Footprint) des produits SBM à partir
des fichiers Excel du Bilan Carbone (Masterbase, BOM, matières et facteurs
d'émission, fret amont) — ici sur un extrait de références Leroy Merlin (LM).

## Pipeline unifié `sbm-pcf`

Chaque étape est appelable séparément. Par défaut les données proviennent de la
session en cache (`--work-dir`) : rapide en développement. `--reload` recharge
tous les fichiers Excel sources.

```bash
sbm-pcf load      --input ./input --work-dir ./pcf_workspace     # 1. charger les sources
sbm-pcf collect   --work-dir ./pcf_workspace --output collection.xlsx   # 2a. fichier de collecte
# ... fichier rempli par SBM / fournisseurs ...
sbm-pcf activity  --work-dir ./pcf_workspace --collection collection_filled.xlsx  # 2b. données d'activité manquantes
sbm-pcf ef-match  --work-dir ./pcf_workspace                     # 2c. matching FE manquants (ecoinvent)
sbm-pcf compute   --work-dir ./pcf_workspace                     # 2d. calcul des PCF + flags HIGH/MEDIUM/LOW
sbm-pcf report    --work-dir ./pcf_workspace --output PCF_LM.pdf [--selection produits.xlsx]  # 2e. rapport PDF
```

### Structure du package (`src/sbm_pcf_calculation/`)

- `sources.py` — chargement des inputs (liste produits, Masterbase, BOM, matières
  et FE, fret, LCIA ecoinvent) + `PcfSession` (conteneur de données)
- `cache.py` — sauvegarde/rechargement de la session (parquet, sans relire les Excel)
- `collection.py` / `datashare.py` + `prefill.py` — fichier de collecte LM (spec v0.95/0.96)
- `activity.py` — évaluation des données d'activité manquantes (fichier rempli)
- `EF_Matching.py` — matching des facteurs d'émission manquants avec ecoinvent
- `pcf_calc.py` — calcul des PCF par produit, DQR/PDS, flags de qualité
- `pdf_report.py` — rapport PDF (WeasyPrint, gabarit HTML/CSS modifiable :
  `src/sbm_pcf_calculation/templates/pcf_report.html`) ; `load_pcf_selection()`
  charge un fichier listant les PCF à générer (xlsx/csv)
- `sbm_pcf_cli.py` — ligne de commande `sbm-pcf`

### Commandes historiques (conservées)

- `sbm-pcf-collecte` — fichier de collecte complet avec calcul PCF (v0.3–v0.74)
- `sbm-pcf-collection` — génération du fichier de collecte LM (ex `pcf-datashare`)
- `sbm-pcf-load` — extraction POC initiale

---

# PCF Extraction — Extraction et analyse des composants SBM

## Génération v0.93 du fichier de collecte LM (pcf-datashare)

La commande `pcf-datashare` génère le fichier de collecte LM
(« SBM - PCF - LM references - Data collection ») selon les spécifications v0.93 :

```bash
pcf-datashare \
  --input "input/SBM - PCF - LM references - Data collection (v0.91).xlsx" \
  --output "data/output/SBM - PCF - LM references - Data collection v0.93.xlsx"
```

- **Layout v0.93** (Spec_CollectionFile) : Product (50 colonnes) et Component
  (55 colonnes) — ajout de `Supplier PCF framework`, `Supplier PCF scope`,
  `Supplier PCF declared unit`, `Transformation Process declared unit`,
  `Transformation Process location`, `Transformation Energy Type`,
  `Transformation Energy Consumption`, `Transformation Process Comment` ;
  champs Energy avant champs Process dans le bloc Transformation
- **Formules Excel explicites** : `Transformation DQR value` (moyenne GEO/TECH/TEMP),
  `Transformation GHG` (Process EF × Net Weight × (1 + Scrap Rate)),
  `Data validation flag` (règles de validation de la spec)
- **Listes déroulantes** : framework (PACT, TfS, ISO14067, other), scope
  (Cradle-to-Gate, Cradle-to-Grave), external review (Yes, No), Energy Type,
  DQR 1-5
- **Guides au format validé SBM** : les onglets `UserGuide` et `DQR_Guide` sont
  repris tels quels du template `templates/Data_collection_template.xlsx`
  (issu du fichier v0.92 corrigé par SBM : contenu, couleurs de cellules et
  polices, hauteurs de lignes, largeurs de colonnes, fusions de cellules,
  bordures) ; ordre des onglets UserGuide, Product, Component, DQR_Guide
- Les données saisies du fichier d'entrée sont recopiées vers le nouveau
  layout (renommage `Transformation Energy Name` → `Transformation Energy Type`)

Outil d'extraction et d'analyse des composants pour calculer les facteurs d'émission (PCF — Product Carbon Footprint) à partir des fichiers Excel SBM (Bilan Carbone FY24-25).

Le pipeline produit des rapports détaillés pour prioriser les recherches de données manquantes :
- **Produits** : référentiel des produits analysés
- **Relations Produit-Composant** : nomenclature enrichie des émissions par phase
- **Stats FE manquants** : taux de complétude par phase (Production, Usage, Fin de Vie, Fret amont)
- **Composants uniques** : liste dédoublonnée avec contribution et nombre de produits concernés
- **Priorités à investiguer** : composants triés par impact (FE manquants × nb de produits)

## Structure du projet

```
pcf-extraction/
├── input/                  # 5 fichiers Excel sources (non versionnés)
├── data/
│   ├── raw/                  # Anciens fichiers SBM (v0.2, non versionnés)
│   ├── interim/              # Données intermédiaires
│   └── output/               # Rapports générés (non versionnés)
├── src/pcf_extraction/
│   ├── __init__.py
│   ├── cli.py                # Lignes de commande pcf-extract / pcf-collecte
│   ├── collect.py            # Pipeline « Fichier de collecte » PCF (v0.3)
│   ├── config.py             # Règles d'extraction et catégorisation SAGE
│   ├── extract.py            # Pipeline d'extraction v0.2
│   ├── io_sbm.py             # Lecture des fichiers SBM
│   └── report.py             # Rapports et statistiques
├── tests/
│   ├── test_collect.py       # Règles DQR + non-régression SORHOY15
│   └── test_report.py        # Tests unitaires
├── docs/
├── pyproject.toml
└── README.md
```

## Installation

```bash
cd pcf-extraction
python -m venv .venv
source .venv/bin/activate   # Windows : .venv\Scripts\activate
pip install -e ".[dev]"
```

## Utilisation

```bash
pcf-extract \
  --material "data/raw/SBM_Material_Packaging.xlsx" \
  --freight "data/raw/SBM_Freight.xlsx" \
  --sample "data/raw/Echantillon - Produits à analyser.xlsx" \
  --output "data/output/Extraction_Composants.xlsx"
```

L'argument `--sample` est optionnel : sans échantillon, le pipeline traite l'intégralité du référentiel.

## Génération du Fichier de collecte (règles v0.7)

Le pipeline `pcf-collecte` produit le livrable au format client, défini par le
fichier « POC Calculateur - règles de calcul - v0.7.xlsx » (onglets Fichier de
collecte et Clé de détermination, colonne B « Nom du champ », colonne D
« Type de données ») :

```bash
pcf-collecte \
  --input "input" \
  --output "data/output/Fichier de collecte.xlsx" \
  --lcia "input/Cut-off Cumulative LCIA v3.12.xlsx"
```

Toutes les données source doivent être placées dans le dossier `input/` :

| Fichier (motif) | Contenu |
|---|---|
| `Référencement LM*.xlsx` (onglet Export) | Liste des produits LM à traiter |
| `Masterbase_Product*.xlsx` (onglet MASTERBASE Products, en-tête ligne 3) | Référentiel produits : poids, matières, BPSNAM/BPSNUM |
| `Masterbase_BOM*.xlsx` (onglet MASTERBASE BOM, en-tête ligne 4) | Nomenclatures (BOMALT=1, USESTA_0=2) |
| `*Material and Packaging*.xlsx` | FE matières : CK_MaterialPurchase (prioritaire), CK_EF_Packaging (secours) |
| `*Freight*.xlsx` | FE fret (pas encore utilisé dans le Fichier de collecte) |

Les motifs tolèrent les mises à jour mensuelles (dates dans les noms) ; si
plusieurs fichiers correspondent à un motif, le plus récent (tri
lexicographique) est retenu.

Onglets générés :
- **Fichier de collecte** : une ligne par composant de BOM (alternative 1, active),
  avec facteurs d'émission, scores DQR (GEO/TECH/TEMP), PDS, GHG et part du
  composant dans le produit. Les colonnes suivent les blocs de la Clé de
  détermination v0.4 (colonne B « Nom du champ ») : couleurs de remplissage des en-têtes par bloc et
  groupements de colonnes (Plans > Grouper) reproduits du fichier de règles.
  Le bloc fret (Impact Appro Transport) est alimenté par les onglets
  Freight_Raw_In_* (rattachement composant + fournisseur -> UniqueKey) et
  CK_FreightConsolidation (ADDRESS_KEY = trajet, Transportation_Mode) :
  Freight GHG = Net Weight x Quantity x GHG_perunit (kgCO2e/kg)
- **Stats** : produits traités, PCF calculés, lignes sans FE
- **Flags** : distribution des anomalies (fournisseur introuvable, quantité
  nulle, poids incohérent, matière ou FE manquant)
- **FE manquants par matière** et **Priorités à investiguer** : priorisation des
  recherches de données manquantes
- **Matching ecoinvent** (si `--lcia` fourni) : proposition de dataset ecoinvent
  et de FE GWP100 pour chaque composant unique sans FE (voir ci-dessous)

### Saisie des données de transformation (v0.5)

Les colonnes du bloc « Impact fabrication fournisseur » (Transformation Process
Name, Transformation Energy Value/Unit, Transformation EF, Transformation EF
Source) correspondent à des données que SBM doit fournir. L'option `--saisie`
génère un classeur à remplir par un responsable des opérations/production :

- **onglet Guide** : description de chaque colonne et comment la remplir ;
- **onglet Saisie transformation** : une ligne par couple (produit, composant),
  pré-remplie avec le fournisseur, à compléter pour les composants transformés.

Une fois rempli, le fichier se réinjecte dans le calcul :

```bash
pcf-collecte \
  --input "input" \
  --output "data/output/Fichier de collecte.xlsx" \
  --transformation "input/Saisie transformation.xlsx"
```

Transformation GHG = Transformation EF x Quantity x Net Weight. Le flag
« PAS DE TRAJET FRET » signale les composants sans trajet inbound trouvé.

### Matching ecoinvent des FE manquants (v0.4.1)

Le module `src/pcf_extraction/ecoinvent.py` propose, pour chaque composant unique
sans facteur d'émission, un dataset ecoinvent (cut-off v3.12) et sa valeur
**GWP100 EF v3.1 (kg CO2-Eq/kg)** :

1. **Règles de correspondance** (table `MATCH_RULES`, ordre = priorité, première
   qui matche gagne) sur désignation + catégorie SAGE + description + matière ;
2. **Géographie en cascade** : pays du fournisseur (déduit du code BPSNUM, ex.
   EFR06853 → FR) si ecoinvent le couvre, sinon RER → RoW → GLO ;
3. **Produit principal** : dans un dataset multi-produits, sélection du produit
   dont le nom est dans l'activité (ex. « zinc » et non le co-produit « ammonium
   sulfate ») et préférence pour les datasets `market for ...`.

L'indicateur est la colonne « EF v3.1 / climate change / global warming
potential (GWP100) / kg CO2-Eq » (colonne AG), divisé par le Reference Product
Amount pour obtenir un FE par kg.

Sur les données actuelles : **105/105 composants sans FE matchés**. Résultats
notables : eaux → `market for tap water` RoW 0.0012 ; PP RER 2.3136 ; PE-HD RER
2.3692 ; caisses carton → `corrugated board box production` RER 1.0060 ; NPK →
`market for NPK (15-15-15) fertiliser` RER 1.0724 ; zinc RoW 2.9008 ; cuivre
GLO 3.6509 ; bentonite RoW 0.0423 ; soufre RoW 0.1581.

#### Limites et proxies à valider

- **Proxies matière** : films hydrosolubles PVOH → film PE-LD ; BHT → toluène ;
  colorant → chimie organique non spécifiée ; lécithine → huile de soja ;
  chélateur → EDTA ; biocides/pesticides → `market for pesticide, unspecified`
  (GLO 10.3524, très conservateur).
- **SOSIRYL** → `injection moulding` : FE de transformation plastique, pas d'une
  matière.
- **BULK SUPPLIED BY SUBCONTRACTOR** → proxy NPK.
- 13 composants sans code fournisseur → cascade par défaut RER/RoW/GLO.
- Les FE proposés restent à valider par SBM avant intégration dans le calcul
  (l'onglet est une aide à la priorisation, pas une source de FE officielle).

### Règles de calcul implémentées

- **Quantity** = `UVQTY` de la BOM alternative 1 (statut actif)
- **RM GHG** = Net Weight × Quantity × RM EF Value (kgCO2e/composant)
- **PCF Value** (produit) = Σ RM GHG des composants
- **Part du composant** = (Net Weight × Quantity) / Σ(Net Weight × Quantity)
- **DQR Product** = moyenne pondérée des DQR composants par leur part
- **RM GEO DQR** : vide/0 → 5 ; GLO → 4 ; RoW → 3 ; RER → 2 ; sinon → 1
- **RM TECH DQR** : source vide → 5 ; contient PROXY → 3 ; contient CK → 1 ; sinon → 2
- **RM TEMP DQR** : 1 si FE associé, 5 sinon
- **RM DQR value** = (RM GEO + RM TECH + RM TEMP) / 3
- **RM EF PDS** = 0 (données secondaires) ; **RM PDS value** = RM PDS Activity Data × RM EF PDS
- **FE** : CK_MaterialPurchase en priorité (FE produit > 0), sinon CK_EF_Packaging
  par (matière, description de catégorie) avec part recyclé si ZRECYCLE/Taux Recyclé > 0
- **Fournisseur produit/composant** : BPSNAM/BPSNUM de la Masterbase produits,
  fallback CK_MaterialPurchase (SUPPLIER_NAME/SUPPLIER_CODE) si vide
- **Carbon category** : lookup Category Code → référentiel Catégorisation Cmpt
- **PCF fournisseur** : si `Supplier PCF value` renseigné, le FE calculé est
  écarté (EF = 0, RM PDS value = 0, RM DQR value = 0)

### Écarts constatés avec l'exemple SORHOY15

Le pipeline reproduit exactement la structure et les scores DQR de l'exemple
(DQR Product = 1.5824, parts et GHG par composant identiques), mais le PCF
diffère légèrement (0.6865 vs 0.6567) car trois cellules de l'exemple manuel
contredisent les données sources :

| Composant | Exemple manuel | Données CK | Impact GHG |
|---|---|---|---|
| 407929 FILM PEBD | matière COR, FE 1.390 | matière LDPE, FE 2.090 | +0.000118 |
| 555350FR BOXGF | FE vierge 0.390 malgré 100 % recyclé | FE recyclé 0.670 | +0.029680 |
| 555356FR LABST | FE « PP » 2.000 pour matière PAP | FE PAP 2.930 | +0.000003 |

Le pipeline suit les données sources (CK) de façon déterministe ; ces trois points
sont à valider avec SBM avant industrialisation.

## Tests

```bash
python -m pytest tests/ -v
```

## Sources de données et règles d'extraction

Les règles sont centralisées dans `src/pcf_extraction/config.py`, transposées de l'onglet `CalculPCF` du fichier SBM Material and Packaging :

| Section | Fichier | Onglet | Colonnes clés |
|---|---|---|---|
| Produit Final | Material & Packaging | `MasterBase_Products` (en-tête ligne 12) | SKU, SKU Designation, Segments |
| Relation Produit-Composant | Material & Packaging | `MasterBase_BOM` (en-tête ligne 12) | ITMREF → CPNITMREF, BOMQTY |
| Composant (émissions) | Material & Packaging | `CK_MaterialPurchase` (en-tête ligne 12) | PRODUCT, Qty_Prod/Use/EoL_GHG |
| Fret amont | Freight | `Freight_Raw_In_IRIS` (l.13), `Freight_Raw_In_Fert` (l.14), `Freight_Raw_In_LSEur` (l.13) | Référence Article / Product, GHG_perunit (kgCO2e/kg) × Weigh Final (in t) |

**Calcul de l'émission fret** : `GHG_perunit` est un facteur kgCO2e/kg ; l'émission par ligne d'achat = facteur × poids total, puis sommée par article.

## Traçabilité

- Les fichiers sources ne sont jamais modifiés.
- Chaque champ extrait est tracé vers son fichier/onglet/colne d'origine via `config.py`.
- Les incertitudes de mapping (colonnes manquantes, doublons) sont conservées en `NaN` et comptées dans les stats de complétude plutôt que silencieusement ignorées.

## Roadmap POC

- [x] v0 : script monolithique (ExtraireComposants_v1) — démonstration client
- [x] v0.2 : refonte modulaire + tests + traçabilité (ce dépôt)
- [x] v0.3 : Fichier de collecte au format client (règles DQR/PDS/GHG, flags de validation, priorisation des FE manquants)
- [x] v0.4 : dossier input/ centralisé (5 sources, motifs datés), fournisseurs depuis BPSNAM/BPSNUM (Masterbase)
- [x] v0.4.1 : matching ecoinvent des 105 composants sans FE (GWP100 EF v3.1, pays fournisseur → RER/RoW/GLO, onglet « Matching ecoinvent »)
- [x] v0.5 : blocs transformation + fret (règles v0.3), couleurs et groupements de colonnes, fichier de saisie transformation, GHG Transport depuis les onglets Freight
- [x] v0.6 : colonnes alignées sur la Clé de détermination v0.4 (noms anglais, colonne B ; lignes GROUPING traitées comme marqueurs de regroupement et non comme colonnes)
- [x] v0.7 : colonnes et mise en page alignées sur la Clé de détermination v0.7 — PCF Value / Unité / DQR Product / PDS Product remontés dans le bloc Produit, RM PDS Activity Data déplacé dans le bloc Composants, bloc transformation à 12 champs (FE procédé + FE énergie), tableau démarré en A6 (en-têtes ligne 6), volets figés en C7, filtres automatiques en ligne 6, formats nombre issus du « Type de données » (DECIMAL(10,3) → 0,000 ; DECIMAL(10,5) → 0,00000 ; DECIMAL(3;0)/BINARY → entier)
- [x] v0.71 : lignes GROUPING de la Clé ajoutées comme colonnes à part entière (texte vertical, renvoi à la ligne, largeur 3,5, couleur du bloc), nouvelles couleurs de la Clé v0.7 (accent1 bleu pour le produit, accent3 vert pour les blocs composants, accent4 bleu pour le fret), groupement transformation étendu (Transformation Process Name → Transformation Energy EF Source) et auto-ajustement des largeurs de colonnes au contenu
- [x] v0.72 : couleurs de police de la Clé (blanc sur blocs foncés, noir sur blocs clairs), largeur des colonnes GROUPING à 71 px, hauteur de ligne d'en-tête à 145 px, lignes 1-4 groupées et repliées à l'ouverture, colonne Flag en orange auto-ajustée, onglet Priorités enrichi des reco ecoinvent + colonne « A valider (OUI/NON) », onglet Synthèse (1 ligne/produit : PCF/DQR/PDS puis décomposition Raw Material / Packaging / Transformation / Freight / Use / End of Life avec PDS Activity Data, PDS EF, PDS Total, DQR, GHG Value, % PCF total, groupe PDS Activity Data → DQR)
- [x] v0.73 : Synthèse en 1er onglet, décimales `#,##0.000` (séparateur de milliers) et pourcentages `0.00%` dans la Synthèse, en-tête Synthèse en retour à la ligne auto (hauteur 60 px), réinjection des 105 FE ecoinvent validés SBM dans le calcul (2 passes : matching puis recalcul complet avec source « EcoInvent 3.12 cut-off (validé SBM) » ; plus aucun flag PAS DE FE, DQR produit moyen 2,43 → 1,50)
- [x] v0.74 : périmètre BOM élargi sur instruction SBM (toutes les alternatives `BOMALT` sauf 2 et 9, statut actif `USESTA_0 = 2`, déduplication produit/composant sur la plus petite alternative) — 212/217 produits avec PCF calculé (vs 112) ; nouvelle colonne **Scrap Rate** (taux de perte/rebuts par composant, bloc Description du composant + fichier de saisie transformation) appliquée au calcul : Transformation GHG = FE procédé × Quantity × Net Weight × (1 + Scrap Rate)
- [ ] Validation des 3 écarts SORHOY15 avec SBM (matières/FE packaging)
- [ ] Ajout d'autres sources de données (électricité, déchets, etc.)
- [ ] Ajout d'autres sources de données (électricité, déchets, etc.)
- [ ] Calcul du PCF complet par produit (agrégation pondérée par la BOM)

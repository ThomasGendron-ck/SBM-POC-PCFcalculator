"""Ligne de commande : pcf-collecte."""
import argparse
import sys
from pathlib import Path

from .collect import build_collecte, build_fe_overrides, resolve_inputs
from .ecoinvent import load_lcia_gwp, match_missing_fe
from .report import write_collecte_report
from .saisie import write_saisie_transformation

def main_collecte(argv=None) -> int:
    parser = argparse.ArgumentParser(
        prog="pcf-collecte",
        description="Génération du Fichier de collecte PCF (format livraison client)",
    )
    parser.add_argument(
        "--input",
        required=True,
        help="Dossier contenant les fichiers source (Référencement LM, Masterbase_Product*, Masterbase_BOM*, Material and Packaging, Freight)",
    )
    parser.add_argument("--output", required=True, help="Chemin du classeur de sortie (.xlsx)")
    parser.add_argument("--lcia", default=None, help="Fichier Cut-off Cumulative LCIA ecoinvent (.xlsx) : ajoute l'onglet Matching ecoinvent (FE manquants -> ICV GWP100 EF v3.1)")
    parser.add_argument("--transformation", default=None, help="Fichier de saisie de transformation rempli par SBM (.xlsx) : alimente les colonnes du bloc Impact fabrication fournisseur")
    parser.add_argument("--saisie", default=None, help="Chemin du fichier de saisie de transformation à générer (.xlsx), pré-rempli avec les couples produit/composant")
    args = parser.parse_args(argv)

    if not Path(args.input).is_dir():
        print(f"Erreur : dossier input introuvable : {args.input}", file=sys.stderr)
        return 1
    if args.lcia and not Path(args.lcia).is_file():
        print(f"Erreur : fichier lcia introuvable : {args.lcia}", file=sys.stderr)
        return 1

    if args.transformation and not Path(args.transformation).is_file():
        print(f"Erreur : fichier transformation introuvable : {args.transformation}", file=sys.stderr)
        return 1
    matching = None
    lcia_path = args.lcia
    if lcia_path is None:
        resolved = resolve_inputs(args.input)
        if "lcia" in resolved:
            lcia_path = str(resolved["lcia"])
    if lcia_path:
        print("Construction du Fichier de collecte (passe 1 : identification des composants sans FE)...")
        collecte = build_collecte(args.input, transformation=args.transformation)
        matching = match_missing_fe(collecte, load_lcia_gwp(lcia_path))
        n_match = int((matching["Statut"] == "MATCHÉ").sum())
        print(f"Composants sans FE matchés : {n_match}/{len(matching)}")
        fe_overrides = build_fe_overrides(matching)
        if fe_overrides:
            print(f"Réinjection des FE ecoinvent validés : {len(fe_overrides)} composants")
            print("Recalcul complet (passe 2 : FE ecoinvent complétés)...")
            collecte = build_collecte(
                args.input,
                transformation=args.transformation,
                fe_overrides=fe_overrides,
            )
    else:
        print("Construction du Fichier de collecte...")
        collecte = build_collecte(args.input, transformation=args.transformation)

    n_produits = collecte["Product SKU"].nunique()
    pcf = collecte.groupby("Product SKU")["PCF Value"].first()
    print(f"Produits traités : {n_produits}")
    print(f"Produits avec PCF calculé : {int(pcf.notna().sum())}")
    print(f"Lignes composants : {len(collecte)}")


    if args.saisie:
        saisie_path = Path(args.saisie)
        saisie_path.parent.mkdir(parents=True, exist_ok=True)
        print(f"Génération du fichier de saisie transformation : {saisie_path}")
        write_saisie_transformation(collecte, str(saisie_path))
    output = Path(args.output)
    output.parent.mkdir(parents=True, exist_ok=True)
    print(f"Écriture du fichier : {output}")
    write_collecte_report(collecte, str(output), matching=matching)
    n_trajets = int(collecte["Freight GHG"].notna().sum())
    print(f"Lignes avec trajet fret : {n_trajets}")
    print("Terminé.")
    return 0


if __name__ == "__main__":
    sys.exit(main())

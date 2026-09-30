"""Ligne de commande : pcf-datashare (fichier de collecte LM, spec v0.95)."""

import argparse
import sys
from pathlib import Path

from .datashare import DEFAULT_TEMPLATE, apply_v093


def main_datashare(argv=None) -> int:
    parser = argparse.ArgumentParser(
        prog="pcf-datashare",
        description="Génération du fichier de collecte LM selon les spécifications v0.95 "
        "(layout des onglets Product/Component, formules, listes déroulantes, "
        "guides UserGuide/DQR_Guide au format validé SBM). "
        "Deux modes : --mode full (régénération complète depuis les données "
        "source) ou --mode quick (régénération rapide du format depuis un "
        "ancien fichier de collecte).",
    )
    parser.add_argument(
        "--mode",
        choices=["full", "quick"],
        default="full",
        help="full : régénération complète à partir des données source "
        "(--spec ou --lm, + --material). quick : régénération rapide du "
        "format/contenu à partir de l'ancien fichier de collecte (--input)",
    )
    parser.add_argument(
        "--input",
        default=None,
        help="[mode quick] Fichier de collecte LM existant (.xlsx) dont on "
        "reprend le contenu : SBM_-_PCF_-_LM_references_-_Data_collection_vX.XX",
    )
    parser.add_argument(
        "--output",
        required=True,
        help="Chemin du classeur de sortie généré (.xlsx), "
        "ex : SBM_-_PCF_-_LM_references_-_Data_collection_v0.95.xlsx",
    )
    parser.add_argument(
        "--template",
        default=None,
        help=f"Template des guides (défaut : {DEFAULT_TEMPLATE})",
    )
    parser.add_argument(
        "--spec",
        default=None,
        help="Spécifications v0.95 (.xlsx) : layout + onglet « Produits LM » "
        "pour les références LM",
    )
    parser.add_argument(
        "--lm",
        default=None,
        help="Fichier « Référencement LM » (.xlsx, onglet Export) : source des "
        "références LM, alternative à l'onglet « Produits LM » de la spec",
    )
    parser.add_argument(
        "--material",
        default=None,
        help="Fichier « Material and Packaging - ExtractPourPCF » (.xlsx) : "
        "Masterbase (MB Product, MB BOM, fournisseurs, catégories) pour le "
        "pré-remplissage des produits et composants",
    )
    args = parser.parse_args(argv)

    spec_source = args.spec or args.lm
    if args.mode == "quick":
        if not args.input:
            print(
                "Erreur : le mode quick nécessite --input (ancien fichier de "
                "collecte).",
                file=sys.stderr,
            )
            return 1
        if not Path(args.input).is_file():
            print(f"Erreur : fichier input introuvable : {args.input}", file=sys.stderr)
            return 1
        input_path = args.input
    else:
        if not spec_source or not args.material:
            print(
                "Erreur : le mode full nécessite les données source : "
                "--spec ou --lm (références LM) ET --material (Masterbase).",
                file=sys.stderr,
            )
            return 1
        if args.input:
            print(
                "Avertissement : --input ignoré en mode full (régénération "
                "complète depuis les sources).",
                file=sys.stderr,
            )
        input_path = None

    output = Path(args.output)
    output.parent.mkdir(parents=True, exist_ok=True)
    if args.mode == "full":
        print(f"Régénération complète (données source) -> {output}")
    else:
        print(f"Régénération rapide ({input_path}) -> {output}")
    counts = apply_v093(
        output,
        input_path=input_path,
        template_path=args.template,
        spec_path=spec_source,
        material_path=args.material,
    )
    if counts is not None:
        print(f"Pré-remplissage : {counts[0]} produits, {counts[1]} composants.")
    print("Terminé (layout v0.95, formules, listes déroulantes, guides SBM).")
    return 0


if __name__ == "__main__":
    sys.exit(main_datashare())

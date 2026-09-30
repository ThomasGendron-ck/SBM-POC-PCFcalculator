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
        "guides UserGuide/DQR_Guide au format validé SBM)",
    )
    parser.add_argument(
        "--input",
        required=True,
        help="Fichier de collecte LM v0.9x (.xlsx) : SBM - PCF - LM references - Data collection",
    )
    parser.add_argument(
        "--output",
        required=True,
        help="Chemin du classeur de sortie généré (.xlsx)",
    )
    parser.add_argument(
        "--template",
        default=None,
        help=f"Template des guides (défaut : {DEFAULT_TEMPLATE})",
    )
    parser.add_argument(
        "--spec",
        default=None,
        help="Spécifications v0.95 (.xlsx) : pré-remplissage des produits LM",
    )
    parser.add_argument(
        "--material",
        default=None,
        help="Fichier « Material and Packaging - ExtractPourPCF » (.xlsx) : "
        "pré-remplissage des produits et composants",
    )
    args = parser.parse_args(argv)
    if not Path(args.input).is_file():
        print(f"Erreur : fichier input introuvable : {args.input}", file=sys.stderr)
        return 1
    output = Path(args.output)
    output.parent.mkdir(parents=True, exist_ok=True)
    print(f"Génération selon la spec v0.95 : {args.input} -> {output}")
    counts = apply_v093(
        args.input,
        output,
        template_path=args.template,
        spec_path=args.spec,
        material_path=args.material,
    )
    if counts is not None:
        print(f"Pré-remplissage : {counts[0]} produits, {counts[1]} composants.")
    print("Terminé (layout v0.95, formules, listes déroulantes, guides SBM).")
    return 0


if __name__ == "__main__":
    sys.exit(main_datashare())

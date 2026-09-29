"""Ligne de commande : pcf-datashare (fichier de collecte LM, spec v0.93)."""

import argparse
import sys
from pathlib import Path

from .datashare import DEFAULT_TEMPLATE, apply_v093


def main_datashare(argv=None) -> int:
    parser = argparse.ArgumentParser(
        prog="pcf-datashare",
        description="Génération du fichier de collecte LM selon les spécifications v0.93 "
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
    args = parser.parse_args(argv)
    if not Path(args.input).is_file():
        print(f"Erreur : fichier input introuvable : {args.input}", file=sys.stderr)
        return 1
    output = Path(args.output)
    output.parent.mkdir(parents=True, exist_ok=True)
    print(f"Génération selon la spec v0.93 : {args.input} -> {output}")
    apply_v093(args.input, output, template_path=args.template)
    print("Terminé (layout v0.93, formules, listes déroulantes, guides SBM).")
    return 0


if __name__ == "__main__":
    sys.exit(main_datashare())

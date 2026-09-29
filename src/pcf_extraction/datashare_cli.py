"""Ligne de commande : pcf-datashare (fichier de collecte LM, spec v0.92)."""

import argparse
import sys
from pathlib import Path

from .datashare import apply_v092


def main_datashare(argv=None) -> int:
    parser = argparse.ArgumentParser(
        prog="pcf-datashare",
        description="Mise à jour du fichier de collecte LM vers les spécifications v0.92 "
        "(couleurs par bloc, onglet DQR_Guide, UserGuide enrichi)",
    )
    parser.add_argument(
        "--input",
        required=True,
        help="Fichier de collecte LM v0.9x (.xlsx) : SBM - PCF - LM references - Data collection",
    )
    parser.add_argument(
        "--output",
        required=True,
        help="Chemin du classeur de sortie mis à jour (.xlsx)",
    )
    args = parser.parse_args(argv)

    if not Path(args.input).is_file():
        print(f"Erreur : fichier input introuvable : {args.input}", file=sys.stderr)
        return 1

    output = Path(args.output)
    output.parent.mkdir(parents=True, exist_ok=True)
    print(f"Mise à jour vers la spec v0.92 : {args.input} -> {output}")
    apply_v092(args.input, output)
    print("Terminé (couleurs v0.92, onglet DQR_Guide, UserGuide enrichi).")
    return 0


if __name__ == "__main__":
    sys.exit(main_datashare())

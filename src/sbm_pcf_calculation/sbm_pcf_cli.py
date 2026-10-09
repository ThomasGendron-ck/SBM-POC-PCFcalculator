"""Unified SBM PCF calculation command line.

One command per pipeline step, each with the same two options:
- full source load: --reload
- fast recalculation from the cached session (default)

Workflow:
  sbm-pcf load      --input ./input --work-dir ./pcf_workspace
  sbm-pcf collect   --work-dir ... --output collection.xlsx [--mode full|quick]
  sbm-pcf activity  --work-dir ... --collection collection_filled.xlsx
  sbm-pcf ef-match  --work-dir ... [--lcia ...]
  sbm-pcf compute   --work-dir ... [--transformation ...]
  sbm-pcf report    --work-dir ... --output PCF_LM.pdf [--selection ...]
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path


def _work_dir(args) -> Path:
    work_dir = Path(args.work_dir)
    work_dir.mkdir(parents=True, exist_ok=True)
    return work_dir


def _load_session(args):
    from .cache import load_session
    from .sources import load_all_sources

    if args.reload:
        kwargs = {"input_dir": args.input, "work_dir": args.work_dir}
        for key, attr in [("lcia", "lcia_path"), ("full_product_extract", "full_product_extract"),
                          ("sample", "sample_path")]:
            value = getattr(args, attr, None)
            if value:
                kwargs[key] = value
        return load_all_sources(**kwargs)
    return load_session(args.work_dir)


def cmd_load(args) -> int:
    from .sources import load_all_sources

    if not Path(args.input).is_dir():
        print(f"Error: input directory not found: {args.input}", file=sys.stderr)
        return 1
    print(f"Loading all sources from {args.input} ...")
    session = load_all_sources(
        args.input,
        args.work_dir,
        sample_path=args.sample,
        lcia_path=args.lcia,
        full_product_extract=args.full_product_extract,
    )
    print(f"Products to compute: {len(session.products)}")
    print(f"BOM lines: {len(session.bom)}")
    print(f"Components: {len(session.component_database)}")
    print("Session saved. Done.")
    return 0


def cmd_collect(args) -> int:
    session = _load_session(args)
    from .datashare import apply_v093

    print(f"Generating collection file -> {args.output}")
    counts = apply_v093(args.output, input_path=args.input_file, template_path=None,
                         spec_path=args.spec, material_path=args.material,
                         mb_product_path=args.mb_product)
    if counts is not None:
        print(f"Pre-filled: {counts[0]} products, {counts[1]} components.")
    print("Done.")
    return 0


def cmd_activity(args) -> int:
    session = _load_session(args)
    from .activity import activity_data_report, assess_missing_activity_data, import_filled_collection

    filled = import_filled_collection(args.collection)
    session.filled_collection = filled
    missing = assess_missing_activity_data(filled)
    to_complete = activity_data_report(filled)
    print(f"Rows read: {len(filled)}")
    print(missing.to_string(index=False))
    out = Path(args.output or "missing_activity_data.xlsx")
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        missing.to_excel(writer, sheet_name="Summary", index=False)
        to_complete.to_excel(writer, sheet_name="To complete", index=False)
    from .cache import save_session

    session.metadata["work_dir"] = args.work_dir
    save_session(session, args.work_dir)
    print(f"Report written: {out}")
    return 0


def cmd_ef_match(args) -> int:
    import pandas as pd

    session = _load_session(args)
    from .EF_Matching import run_ef_matching

    matching = run_ef_matching(session, args.lcia)
    n = int((matching["Statut"] == "MATCHÉ").sum()) if not matching.empty else 0
    print(f"Matched: {n}/{len(matching)}")
    out = Path(args.output or "ef_matching.xlsx")
    if session.component_results is not None:
        from .ef_matching_writer import write_ef_matching as write_ef_matching_full
        n_rows = write_ef_matching_full(str(out), session.component_results, matching=matching)
        print(f"Composants sans FE listés (spec v0.97) : {n_rows}")
    else:
        matching.to_excel(out, index=False)
    validated = matching[matching["Statut"] == "MATCHÉ"].copy()
    validated["A valider (OUI/NON)"] = "OUI"
    from .EF_Matching import apply_validated_factors

    apply_validated_factors(session, validated)
    session.ef_matching = matching
    from .cache import save_session

    session.metadata["work_dir"] = args.work_dir
    save_session(session, args.work_dir)
    print(f"Matching written: {out} (validated factors stored in session)")
    return 0


def cmd_compute(args) -> int:
    session = _load_session(args)
    session.metadata["work_dir"] = args.work_dir
    from .pcf_calc import run_pcf_calculation

    session = run_pcf_calculation(session, transformation_path=args.transformation)
    results = session.product_results
    pcf = results["PCF Value"] if "PCF Value" in results.columns else None
    print(f"Products computed: {len(results)}")
    if pcf is not None:
        print(f"Products with a PCF value: {int(pcf.notna().sum())}")
    out = Path(args.output or "pcf_results.xlsx")
    from .results_writer import write_pcf_results

    write_pcf_results(out, session.component_results, results, session.ef_matching)
    print(f"Results written: {out}")
    if args.check_baseline:
        from .regression import check_regression, save_baseline

        if args.check_baseline == "save":
            path = save_baseline(results)
            print(f"Baseline saved: {path}")
        else:
            report = check_regression(results, args.check_baseline)
            if report.empty:
                print("Non-regression check passed: results match the baseline.")
            else:
                print("PCF RESULTS REGRESSED against the baseline:", file=sys.stderr)
                print(report.to_string(index=False), file=sys.stderr)
                return 2
    return 0


def cmd_report(args) -> int:
    session = _load_session(args)
    if session.product_results is None:
        print("Error: no computed results in session: run the compute step first.", file=sys.stderr)
        return 1
    from .pdf_report import generate_pdf_report, generate_pdf_report_from_list, load_pcf_selection

    selected = load_pcf_selection(args.selection) if args.selection else None
    generate_pdf_report(session.product_results, selected, args.output, args.template)
    print(f"PDF report written: {args.output}")
    return 0


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(
        prog="sbm-pcf",
        description="SBM PCF calculation pipeline: load sources, generate the data "
        "collection file, assess missing activity data, match missing emission "
        "factors with ecoinvent, compute PCF per product, generate the PDF report.",
    )
    parser.add_argument("--work-dir", default="./pcf_workspace", help="Session work directory (cache)")
    parser.add_argument("--reload", action="store_true", help="Full source reload instead of cached session")
    sub = parser.add_subparsers(dest="command", required=True)

    def _add_global_options(sub_parser):
        sub_parser.add_argument("--work-dir", default=argparse.SUPPRESS,
                                help="Session work directory (cache)")
        sub_parser.add_argument("--reload", action="store_true", default=argparse.SUPPRESS,
                                help="Full source reload instead of cached session")

    load = sub.add_parser("load", help="Load all source files into the cached session")
    load.add_argument("--input", required=True, help="Input directory with the SBM source files")
    load.add_argument("--sample", default=None, help="Product list file (default: LM referencing file)")
    load.add_argument("--lcia", default=None, help="Ecoinvent Cut-off Cumulative LCIA file")
    load.add_argument("--full-product-extract", default=None, help="Complete Masterbase products extract")
    _add_global_options(load)
    load.set_defaults(func=cmd_load)

    collect = sub.add_parser("collect", help="Generate the data collection workbook")
    collect.add_argument("--output", required=True, help="Output collection workbook (.xlsx)")
    collect.add_argument("--input-file", default=None, help="[quick mode] previously filled collection file")
    collect.add_argument("--spec", default=None, help="Specifications file (v0.95)")
    collect.add_argument("--material", default=None, help="Material and Packaging ExtractPourPCF file")
    collect.add_argument("--mb-product", default=None, help="Complete Masterbase products extract")
    _add_global_options(collect)
    collect.set_defaults(func=cmd_collect)

    activity = sub.add_parser("activity", help="Assess missing activity data on a filled collection file")
    activity.add_argument("--collection", required=True, help="Filled collection workbook (.xlsx)")
    activity.add_argument("--output", default=None, help="Missing activity data report (.xlsx)")
    _add_global_options(activity)
    activity.set_defaults(func=cmd_activity)

    ef_match = sub.add_parser("ef-match", help="Match missing emission factors with ecoinvent")
    ef_match.add_argument("--lcia", default=None, help="LCIA file (if not in the cached session)")
    ef_match.add_argument("--output", default=None, help="Matching output (.xlsx)")
    _add_global_options(ef_match)
    ef_match.set_defaults(func=cmd_ef_match)

    compute = sub.add_parser("compute", help="Compute PCF per product with quality flags")
    compute.add_argument("--transformation", default=None, help="Filled transformation input file")
    compute.add_argument("--output", default=None, help="Results output (.xlsx)")
    compute.add_argument(
        "--check-baseline",
        default=None,
        metavar="BASELINE",
        help="Non-regression check against a baseline CSV. Use the special value "
        "'save' to write the current results as the new baseline "
        "(baselines/pcf_results_baseline.csv), or a path to an existing "
        "baseline file. The command fails (exit code 2) if results drifted.",
    )
    _add_global_options(compute)
    compute.set_defaults(func=cmd_compute)

    report = sub.add_parser("report", help="Generate the PCF PDF report")
    report.add_argument("--output", default="SBM_PCF_report.pdf", help="Output PDF path")
    report.add_argument("--selection", default=None, help="File listing the PCFs to generate (xlsx/csv)")
    report.add_argument("--template", default=None, help="Custom HTML template for the report")
    _add_global_options(report)
    report.set_defaults(func=cmd_report)

    import argparse as _ap

    args = parser.parse_args(argv)
    if not hasattr(args, "work_dir"):
        args.work_dir = "./pcf_workspace"
    if not hasattr(args, "reload"):
        args.reload = False
    _work_dir(args)
    try:
        return args.func(args)
    except (FileNotFoundError, NotADirectoryError) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1
    except (ValueError, KeyError) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    import pandas as pd  # noqa: F401 (used in cmd_activity)
    sys.exit(main())

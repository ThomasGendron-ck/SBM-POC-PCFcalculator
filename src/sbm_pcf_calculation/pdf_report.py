"""PCF-formatted PDF report (one page per product).

Renders an HTML template with WeasyPrint, so layout and design adjustments are
made in the template (HTML/CSS) rather than in Python code.
"""
from __future__ import annotations

from pathlib import Path

import pandas as pd

DEFAULT_TEMPLATE = Path(__file__).resolve().parent / "templates" / "pcf_report.html"

FLAG_COLORS = {
    "HIGH": "#C00000",
    "MEDIUM": "#ED7D31",
    "LOW": "#548235",
}


def load_pcf_selection(selection_path: str | Path) -> list[str]:
    """Load a file listing the PCFs to generate (xlsx or csv).

    Accepts any sheet with a recognizable SKU column (auto-detected)."""
    path = Path(selection_path)
    if path.suffix.lower() == ".csv":
        df = pd.read_csv(path)
    else:
        from .io_sbm import load_sample_products

        return load_sample_products(str(path))["SKU"].tolist()
    df.columns = [str(c).strip() for c in df.columns]
    sku_col = next(
        (c for c in df.columns if "sku" in str(c).lower() or "product" in str(c).lower() or "id" in str(c).lower()),
        df.columns[0],
    )
    return df[sku_col].astype(str).str.strip().drop_duplicates().tolist()


def _phase_rows(product: pd.Series) -> str:
    phases = [
        ("Raw Material", "Raw Material GHG (kgCO2e)"),
        ("Packaging", "Packaging GHG (kgCO2e)"),
        ("Transformation", "Transformation GHG (kgCO2e)"),
        ("Freight", "Freight GHG (kgCO2e)"),
    ]
    rows = []
    for label, col in phases:
        if col in product.index and pd.notna(product[col]):
            rows.append(
                f"<tr><td>{label}</td><td class='num'>{product[col]:.5f}</td>"
                f"<td class='num'>{100 * product[col] / max(product['PCF Value'], 1e-12):.1f}%</td></tr>"
            )
    return "".join(rows)


def _flags_html(product: pd.Series) -> str:
    flag_col = "Quality flag"
    if flag_col not in product.index or pd.isna(product[flag_col]) or not str(product[flag_col]).strip():
        return "<p class='ok'>No quality flag detected.</p>"
    items = [f"<li>{f}</li>" for f in str(product[flag_col]).split(" | ")]
    return f"<ul class='flags'>{''.join(items)}</ul>"


def render_report_html(product_results: pd.DataFrame, selected_skus: list[str] | None = None,
                       template_path: str | Path | None = None) -> str:
    """Render the HTML report for the selected SKUs (all products if None)."""
    template = Path(template_path) if template_path else DEFAULT_TEMPLATE
    if template.is_file():
        page = template.read_text(encoding="utf-8")
    else:
        page = "<html><body>{pages}</body></html>"
    results = product_results
    if selected_skus is not None:
        results = product_results[product_results["Product SKU"].isin(selected_skus)]
    pages_html = []
    for _, product in results.iterrows():
        if pd.isna(product.get("PCF Value")):
            continue
        page_html = (
            f"<div class='page'>"
            f"<h1>{product['Product SKU']} - {product.get('Product Designation', '')}</h1>"
            f"<p class='pcf'>PCF Value: <b>{product['PCF Value']:.5f} kg CO2e</b></p>"
            f"<p>DQR: {product.get('DQR Product', '')} - PDS: {product.get('PDS Product', '')}</p>"
            f"<table><thead><tr><th>Phase</th><th>GHG (kgCO2e)</th><th>% of PCF</th></tr></thead>"
            f"<tbody>{_phase_rows(product)}</tbody></table>"
            f"<h2>Quality flags</h2>{_flags_html(product)}"
            f"</div>"
        )
        pages_html.append(page_html)
    return page.replace("{pages}", "\n".join(pages_html))


def generate_pdf_report(product_results: pd.DataFrame,
                        selected_skus: list[str] | None = None,
                        output_path: str | Path = "SBM_PCF_report.pdf",
                        template_path: str | Path | None = None) -> None:
    """One page per selected SKU: total PCF value + unit, DQR, PDS, phase
    decomposition, quality flags. selected_skus=None -> all products."""
    try:
        from weasyprint import HTML
    except ImportError as exc:  # pragma: no cover
        raise ImportError(
            "WeasyPrint is required for the PDF report: pip install weasyprint"
        ) from exc
    html = render_report_html(product_results, selected_skus, template_path)
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)
    HTML(string=html).write_pdf(str(output))


def generate_pdf_report_from_list(product_results: pd.DataFrame,
                                  selection_path: str | Path,
                                  output_path: str | Path = "SBM_PCF_report.pdf",
                                  template_path: str | Path | None = None) -> None:
    """Convenience wrapper: load the PCF selection file, then generate the
    report for those SKUs only."""
    generate_pdf_report(product_results, load_pcf_selection(selection_path), output_path, template_path)

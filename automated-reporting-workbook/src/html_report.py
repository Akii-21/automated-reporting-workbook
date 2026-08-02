"""
html_report.py
Builds a standalone, single-file HTML report (charts embedded as base64 PNGs
via matplotlib) — the "script version" deliverable that doesn't require
Excel to view, and is easy to export to PDF (Ctrl/Cmd+P -> Save as PDF) or
attach in an email/Slack update.
"""
import base64
import io
from datetime import datetime

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

from . import kpi_engine as kpi

plt.rcParams["font.family"] = "DejaVu Sans"
NAVY = "#1F3864"
ACCENT = "#2E75B6"
PALETTE = ["#2E75B6", "#5B9BD5", "#8FAADC", "#B4C7E7", "#1F3864", "#70AD47"]


def _fig_to_base64(fig) -> str:
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150, bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return base64.b64encode(buf.read()).decode("utf-8")


def _bar_chart(df, x_col, y_col, title):
    fig, ax = plt.subplots(figsize=(6, 4))
    ax.bar(df[x_col], df[y_col], color=ACCENT)
    ax.set_title(title, color=NAVY, fontweight="bold")
    ax.spines[["top", "right"]].set_visible(False)
    plt.xticks(rotation=30, ha="right")
    fig.tight_layout()
    return _fig_to_base64(fig)


def _pie_chart(df, label_col, val_col, title):
    fig, ax = plt.subplots(figsize=(5, 4))
    ax.pie(df[val_col], labels=df[label_col], autopct="%1.0f%%", colors=PALETTE,
           wedgeprops={"edgecolor": "white"})
    ax.set_title(title, color=NAVY, fontweight="bold")
    fig.tight_layout()
    return _fig_to_base64(fig)


def _line_chart(df, x_col, y_col, title):
    fig, ax = plt.subplots(figsize=(10, 4))
    ax.plot(df[x_col], df[y_col], marker="o", color=ACCENT, linewidth=2)
    ax.set_title(title, color=NAVY, fontweight="bold")
    ax.spines[["top", "right"]].set_visible(False)
    ax.fill_between(range(len(df)), df[y_col], alpha=0.08, color=ACCENT)
    plt.xticks(rotation=30, ha="right")
    fig.tight_layout()
    return _fig_to_base64(fig)


def _fmt_money(v, symbol="$"):
    return f"{symbol}{v:,.0f}"


def build_html_report(df, cfg, output_path: str):
    symbol = cfg.get("currency_symbol", "$")
    kpis = kpi.summary_kpis(df)
    cats_df = kpi.category_summary(df)
    reg_df = kpi.region_summary(df)
    trend_df = kpi.monthly_trend(df)
    top_df = kpi.top_products(df, cfg.get("top_n", 5))

    bar_b64 = _bar_chart(cats_df, "Category", "Sales", "Sales by Category")
    pie_b64 = _pie_chart(reg_df, "Region", "Sales", "Revenue Share by Region")
    line_b64 = _line_chart(trend_df, "Month", "Sales", "Monthly Sales Trend")

    kpi_cards = ""
    for label, value in kpis.items():
        display = f"{value:.1f}%" if "%" in label else _fmt_money(value, symbol)
        kpi_cards += f"""
        <div class="kpi-card">
          <div class="kpi-label">{label}</div>
          <div class="kpi-value">{display}</div>
        </div>"""

    top_rows = ""
    for _, r in top_df.iterrows():
        profit_cell = f"<td>{_fmt_money(r['Profit'], symbol)}</td>" if "Profit" in top_df.columns else ""
        top_rows += f"<tr><td>{r['Product']}</td><td>{_fmt_money(r['Sales'], symbol)}</td>{profit_cell}</tr>"

    profit_header = "<th>Profit</th>" if "Profit" in top_df.columns else ""

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>{cfg.get('report_title', 'Sales Performance Report')}</title>
<style>
  body {{ font-family: Arial, Helvetica, sans-serif; background: #F4F6F8; color: #222; margin: 0; padding: 0; }}
  .wrap {{ max-width: 1000px; margin: 0 auto; padding: 32px 24px 64px; }}
  h1 {{ color: {NAVY}; margin-bottom: 4px; }}
  .subtitle {{ color: #777; font-size: 13px; margin-top: 0; margin-bottom: 28px; }}
  .kpi-row {{ display: flex; gap: 16px; flex-wrap: wrap; margin-bottom: 32px; }}
  .kpi-card {{ background: white; border-radius: 10px; padding: 16px 20px; box-shadow: 0 1px 3px rgba(0,0,0,0.08); flex: 1; min-width: 150px; }}
  .kpi-label {{ font-size: 11px; font-weight: bold; color: #888; letter-spacing: 0.5px; }}
  .kpi-value {{ font-size: 26px; font-weight: bold; color: {NAVY}; margin-top: 4px; }}
  .chart-card {{ background: white; border-radius: 10px; padding: 16px; margin-bottom: 24px; box-shadow: 0 1px 3px rgba(0,0,0,0.08); text-align: center; }}
  .chart-row {{ display: flex; gap: 16px; flex-wrap: wrap; }}
  .chart-row .chart-card {{ flex: 1; min-width: 320px; }}
  img {{ max-width: 100%; height: auto; }}
  table {{ width: 100%; border-collapse: collapse; background: white; border-radius: 10px; overflow: hidden; }}
  th {{ background: {NAVY}; color: white; text-align: left; padding: 10px 12px; font-size: 13px; }}
  td {{ padding: 9px 12px; border-bottom: 1px solid #eee; font-size: 13px; }}
  tr:last-child td {{ border-bottom: none; }}
  .section-title {{ color: {NAVY}; margin: 32px 0 12px; font-size: 16px; }}
  footer {{ color: #999; font-size: 11px; margin-top: 40px; text-align: center; }}
</style>
</head>
<body>
<div class="wrap">
  <h1>{cfg.get('report_title', 'Sales Performance Report')}</h1>
  <p class="subtitle">Generated {datetime.now().strftime('%d %b %Y, %H:%M')} &middot; {len(df):,} order rows</p>

  <div class="kpi-row">{kpi_cards}</div>

  <div class="chart-card">
    <img src="data:image/png;base64,{line_b64}" alt="Monthly sales trend">
  </div>

  <div class="chart-row">
    <div class="chart-card"><img src="data:image/png;base64,{bar_b64}" alt="Sales by category"></div>
    <div class="chart-card"><img src="data:image/png;base64,{pie_b64}" alt="Revenue share by region"></div>
  </div>

  <div class="section-title">Top {cfg.get('top_n', 5)} Products by Sales</div>
  <table>
    <tr><th>Product</th><th>Sales</th>{profit_header}</tr>
    {top_rows}
  </table>

  <footer>Generated automatically by the Automated Reporting Workbook pipeline &middot; source: Raw Data ({len(df):,} rows)</footer>
</div>
</body>
</html>"""

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)
    return output_path

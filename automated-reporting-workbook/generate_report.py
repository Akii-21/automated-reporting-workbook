#!/usr/bin/env python3
"""
generate_report.py
Entry point for the Automated Reporting Workbook.

Usage:
    python generate_report.py --input data/sample_sales_data.csv
    python generate_report.py --input data/my_export.csv --config config.yaml --outdir reports

Reads any sales/orders CSV (as long as config.yaml's column mapping matches
its headers), and produces two report artifacts in --outdir:
  - sales_report.xlsx  (multi-sheet workbook, live formulas, native charts)
  - sales_report.html  (single-file HTML dashboard, easy to share or print-to-PDF)
"""
import argparse
import sys
import time
from pathlib import Path

from src.data_loader import load_config, load_data
from src.excel_report import build_excel_report
from src.html_report import build_html_report


def parse_args():
    p = argparse.ArgumentParser(description="Generate an automated sales report from a CSV.")
    p.add_argument("--input", "-i", default="data/sample_sales_data.csv", help="Path to input CSV")
    p.add_argument("--config", "-c", default="config.yaml", help="Path to config.yaml")
    p.add_argument("--outdir", "-o", default="reports", help="Directory to write report files into")
    p.add_argument("--skip-excel", action="store_true", help="Skip generating the .xlsx workbook")
    p.add_argument("--skip-html", action="store_true", help="Skip generating the .html report")
    return p.parse_args()


def main():
    args = parse_args()
    t0 = time.time()

    print(f"[1/4] Loading config from {args.config} ...")
    cfg = load_config(args.config)

    print(f"[2/4] Loading + validating data from {args.input} ...")
    df = load_data(args.input, cfg)
    print(f"      -> {len(df):,} rows, {df['date'].min().date()} to {df['date'].max().date()}")

    outdir = Path(args.outdir)
    outdir.mkdir(parents=True, exist_ok=True)

    if not args.skip_excel:
        print("[3/4] Building Excel workbook (live formulas + native charts) ...")
        xlsx_path = outdir / cfg["output"]["excel_filename"]
        build_excel_report(df, cfg, str(xlsx_path))
        print(f"      -> {xlsx_path}")
    else:
        print("[3/4] Skipped Excel workbook.")

    if not args.skip_html:
        print("[4/4] Building HTML dashboard ...")
        html_path = outdir / cfg["output"]["html_filename"]
        build_html_report(df, cfg, str(html_path))
        print(f"      -> {html_path}")
    else:
        print("[4/4] Skipped HTML report.")

    print(f"Done in {time.time() - t0:.1f}s")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"ERROR: {e}", file=sys.stderr)
        sys.exit(1)

"""
data_loader.py
Loads a raw CSV, validates it against the column mapping in config.yaml,
and returns a clean, typed pandas DataFrame with standardized column names.
"""
from pathlib import Path
import pandas as pd
import yaml


REQUIRED_FIELDS = ["date", "category", "region", "sales"]


def load_config(config_path: str = "config.yaml") -> dict:
    with open(config_path, "r") as f:
        return yaml.safe_load(f)


def load_data(csv_path: str, config: dict) -> pd.DataFrame:
    """Load the raw CSV and rename columns to the pipeline's internal names."""
    path = Path(csv_path)
    if not path.exists():
        raise FileNotFoundError(f"Input file not found: {csv_path}")

    df = pd.read_csv(path)
    col_map = config["columns"]

    # Only keep + rename columns that exist in both the config and the file
    missing = [f for f in REQUIRED_FIELDS if col_map.get(f) not in df.columns]
    if missing:
        raise ValueError(
            f"Missing required column(s) for: {missing}. "
            f"Check the 'columns' mapping in config.yaml against your CSV headers: "
            f"{list(df.columns)}"
        )

    rename_map = {v: k for k, v in col_map.items() if v in df.columns}
    df = df.rename(columns=rename_map)

    df["date"] = pd.to_datetime(df["date"], errors="coerce")
    n_before = len(df)
    df = df.dropna(subset=["date", "sales"]).copy()
    n_dropped = n_before - len(df)
    if n_dropped:
        print(f"[data_loader] Dropped {n_dropped} row(s) with invalid date/sales values.")

    df["sales"] = pd.to_numeric(df["sales"], errors="coerce").fillna(0)
    if "profit" in df.columns:
        df["profit"] = pd.to_numeric(df["profit"], errors="coerce").fillna(0)
    else:
        df["profit"] = pd.NA
    if "quantity" in df.columns:
        df["quantity"] = pd.to_numeric(df["quantity"], errors="coerce").fillna(0)

    df["year_month"] = df["date"].dt.to_period("M").astype(str)
    df = df.sort_values("date").reset_index(drop=True)

    return df

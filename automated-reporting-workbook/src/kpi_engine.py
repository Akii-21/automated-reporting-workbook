"""
kpi_engine.py
Pure-pandas aggregation layer. Every function returns a small DataFrame or
dict that both the Excel builder and the HTML builder consume — keeping the
"what to calculate" logic in one place, separate from "how to render it".
"""
import pandas as pd


def has_profit(df: pd.DataFrame) -> bool:
    return df["profit"].notna().any()


def summary_kpis(df: pd.DataFrame) -> dict:
    total_sales = df["sales"].sum()
    total_orders = df["order_id"].nunique() if "order_id" in df.columns else len(df)
    avg_order_value = total_sales / total_orders if total_orders else 0

    kpis = {
        "Total Revenue": total_sales,
        "Total Orders": total_orders,
        "Average Order Value": avg_order_value,
    }
    if has_profit(df):
        total_profit = df["profit"].sum()
        kpis["Total Profit"] = total_profit
        kpis["Profit Margin %"] = (total_profit / total_sales * 100) if total_sales else 0
    return kpis


def monthly_trend(df: pd.DataFrame) -> pd.DataFrame:
    agg = {"sales": "sum"}
    if has_profit(df):
        agg["profit"] = "sum"
    out = df.groupby("year_month", as_index=False).agg(agg)
    out = out.sort_values("year_month")
    out = out.rename(columns={"sales": "Sales", "profit": "Profit", "year_month": "Month"})
    return out


def category_summary(df: pd.DataFrame) -> pd.DataFrame:
    agg = {"sales": "sum"}
    if has_profit(df):
        agg["profit"] = "sum"
    out = df.groupby("category", as_index=False).agg(agg)
    out = out.sort_values("sales", ascending=False)
    out = out.rename(columns={"sales": "Sales", "profit": "Profit", "category": "Category"})
    return out


def region_summary(df: pd.DataFrame) -> pd.DataFrame:
    agg = {"sales": "sum"}
    if has_profit(df):
        agg["profit"] = "sum"
    out = df.groupby("region", as_index=False).agg(agg)
    out = out.sort_values("sales", ascending=False)
    out = out.rename(columns={"sales": "Sales", "profit": "Profit", "region": "Region"})
    return out


def top_products(df: pd.DataFrame, n: int = 5) -> pd.DataFrame:
    key = "sub_category" if "sub_category" in df.columns else "category"
    agg = {"sales": "sum"}
    if has_profit(df):
        agg["profit"] = "sum"
    out = df.groupby(key, as_index=False).agg(agg)
    out = out.sort_values("sales", ascending=False).head(n)
    out = out.rename(columns={"sales": "Sales", "profit": "Profit", key: "Product"})
    return out

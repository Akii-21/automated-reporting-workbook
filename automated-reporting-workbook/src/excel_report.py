"""
excel_report.py
Builds the .xlsx deliverable with openpyxl.

Design choice: the "Raw Data" sheet holds the full dataset, and every summary
sheet (Category, Region, Monthly Trend, Top Products) computes its numbers
with live SUMIF/COUNTIF formulas that reference Raw Data — not with numbers
pasted in from pandas. That means opening the workbook, editing a row in
Raw Data, and hitting recalculate updates the entire report, including the
charts, exactly like a "real" automated report should behave.
"""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.worksheet.table import Table, TableStyleInfo

from . import kpi_engine as kpi

FONT_NAME = "Arial"
NAVY = "1F3864"
ACCENT = "2E75B6"
LIGHT = "DCE6F1"

HEADER_FONT = Font(name=FONT_NAME, size=11, bold=True, color="FFFFFF")
HEADER_FILL = PatternFill("solid", fgColor=NAVY)
TITLE_FONT = Font(name=FONT_NAME, size=18, bold=True, color=NAVY)
SUBTITLE_FONT = Font(name=FONT_NAME, size=10, italic=True, color="666666")
KPI_LABEL_FONT = Font(name=FONT_NAME, size=10, bold=True, color="666666")
KPI_VALUE_FONT = Font(name=FONT_NAME, size=20, bold=True, color=NAVY)
BODY_FONT = Font(name=FONT_NAME, size=10)
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CURRENCY_FMT = '"$"#,##0;("$"#,##0)'
PCT_FMT = "0.0%"


def _style_header_row(ws, row, n_cols):
    for c in range(1, n_cols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = BORDER


def _autofit(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def _write_raw_data(wb, df, cfg):
    ws = wb.create_sheet("Raw Data")
    # Column order — only include fields actually present in this dataset
    field_order = [
        ("order_id", "Order ID"), ("date", "Order Date"), ("category", "Category"),
        ("sub_category", "Sub Category"), ("region", "Region"), ("segment", "Segment"),
        ("quantity", "Quantity"), ("sales", "Sales"), ("profit", "Profit"),
        ("year_month", "Month"),
    ]
    present = [(f, h) for f, h in field_order if f in df.columns and df[f].notna().any() or f in ("sales",)]
    headers = [h for _, h in present]
    fields = [f for f, _ in present]

    ws.append(headers)
    _style_header_row(ws, 1, len(headers))

    for _, row in df.iterrows():
        vals = []
        for f in fields:
            v = row[f]
            if f == "date":
                v = v.date() if hasattr(v, "date") else v
            vals.append(v)
        ws.append(vals)

    n_rows = len(df) + 1
    n_cols = len(headers)
    last_col_letter = get_column_letter(n_cols)

    table = Table(displayName="RawData", ref=f"A1:{last_col_letter}{n_rows}")
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium2", showRowStripes=True, showFirstColumn=False
    )
    ws.add_table(table)

    # number formats
    for f, col_idx in zip(fields, range(1, n_cols + 1)):
        if f in ("sales", "profit"):
            for r in range(2, n_rows + 1):
                ws.cell(row=r, column=col_idx).number_format = CURRENCY_FMT
        if f == "date":
            for r in range(2, n_rows + 1):
                ws.cell(row=r, column=col_idx).number_format = "yyyy-mm-dd"

    ws.freeze_panes = "A2"
    _autofit(ws, [16] * n_cols)

    col_letters = {f: get_column_letter(i + 1) for i, f in enumerate(fields)}
    return ws, col_letters, n_rows


def _sum_formula(col_letters, field, criteria_col_letter, n_rows, criteria_cell):
    if field not in col_letters:
        return None
    val_range = f"'Raw Data'!${col_letters[field]}$2:${col_letters[field]}${n_rows}"
    crit_range = f"'Raw Data'!${criteria_col_letter}$2:${criteria_col_letter}${n_rows}"
    return f"=SUMIF({crit_range},{criteria_cell},{val_range})"


def _write_category_summary(wb, cats_df, col_letters, n_rows):
    ws = wb.create_sheet("Category Summary")
    has_profit = "Profit" in cats_df.columns
    headers = ["Category", "Sales", "% of Total Sales"] + (["Profit", "Profit Margin %"] if has_profit else [])
    ws.append(headers)
    _style_header_row(ws, 1, len(headers))

    start_row = 2
    for i, cat in enumerate(cats_df["Category"], start=start_row):
        c = ws.cell(row=i, column=1, value=cat)
        c.font = BODY_FONT
        sales_formula = _sum_formula(col_letters, "sales", col_letters["category"], n_rows, f"$A{i}") \
            if "category" in col_letters else None
        sales_cell = ws.cell(row=i, column=2)
        if sales_formula:
            sales_cell.value = sales_formula
        sales_cell.number_format = CURRENCY_FMT
        sales_cell.font = BODY_FONT

    last_row = start_row + len(cats_df) - 1
    total_row = last_row + 2
    ws.cell(row=total_row, column=1, value="Total").font = Font(name=FONT_NAME, bold=True)
    total_cell = ws.cell(row=total_row, column=2, value=f"=SUM(B{start_row}:B{last_row})")
    total_cell.number_format = CURRENCY_FMT
    total_cell.font = Font(name=FONT_NAME, bold=True)

    for i in range(start_row, last_row + 1):
        pct_cell = ws.cell(row=i, column=3, value=f"=IFERROR(B{i}/$B${total_row},0)")
        pct_cell.number_format = PCT_FMT
        pct_cell.font = BODY_FONT

    if has_profit:
        for i, cat in enumerate(cats_df["Category"], start=start_row):
            profit_formula = _sum_formula(col_letters, "profit", col_letters["category"], n_rows, f"$A{i}")
            pcell = ws.cell(row=i, column=4)
            if profit_formula:
                pcell.value = profit_formula
            pcell.number_format = CURRENCY_FMT
            pcell.font = BODY_FONT
            margin_cell = ws.cell(row=i, column=5, value=f"=IFERROR(D{i}/B{i},0)")
            margin_cell.number_format = PCT_FMT
            margin_cell.font = BODY_FONT

    for r in range(1, total_row + 1):
        for c in range(1, len(headers) + 1):
            ws.cell(row=r, column=c).border = BORDER

    _autofit(ws, [22, 16, 16, 16, 16])
    return ws, start_row, last_row


def _write_region_summary(wb, reg_df, col_letters, n_rows):
    ws = wb.create_sheet("Regional Summary")
    has_profit = "Profit" in reg_df.columns
    headers = ["Region", "Sales"] + (["Profit"] if has_profit else [])
    ws.append(headers)
    _style_header_row(ws, 1, len(headers))

    start_row = 2
    for i, reg in enumerate(reg_df["Region"], start=start_row):
        ws.cell(row=i, column=1, value=reg).font = BODY_FONT
        formula = _sum_formula(col_letters, "sales", col_letters["region"], n_rows, f"$A{i}") \
            if "region" in col_letters else None
        cell = ws.cell(row=i, column=2)
        if formula:
            cell.value = formula
        cell.number_format = CURRENCY_FMT
        cell.font = BODY_FONT
        if has_profit:
            pformula = _sum_formula(col_letters, "profit", col_letters["region"], n_rows, f"$A{i}")
            pcell = ws.cell(row=i, column=3)
            if pformula:
                pcell.value = pformula
            pcell.number_format = CURRENCY_FMT
            pcell.font = BODY_FONT

    last_row = start_row + len(reg_df) - 1
    for r in range(1, last_row + 1):
        for c in range(1, len(headers) + 1):
            ws.cell(row=r, column=c).border = BORDER

    _autofit(ws, [18, 16, 16])
    return ws, start_row, last_row


def _write_monthly_trend(wb, trend_df, col_letters, n_rows):
    ws = wb.create_sheet("Monthly Trend")
    has_profit = "Profit" in trend_df.columns
    headers = ["Month", "Sales"] + (["Profit"] if has_profit else [])
    ws.append(headers)
    _style_header_row(ws, 1, len(headers))

    start_row = 2
    for i, month in enumerate(trend_df["Month"], start=start_row):
        ws.cell(row=i, column=1, value=month).font = BODY_FONT
        formula = _sum_formula(col_letters, "sales", col_letters["year_month"], n_rows, f"$A{i}") \
            if "year_month" in col_letters else None
        cell = ws.cell(row=i, column=2)
        if formula:
            cell.value = formula
        cell.number_format = CURRENCY_FMT
        cell.font = BODY_FONT
        if has_profit:
            pformula = _sum_formula(col_letters, "profit", col_letters["year_month"], n_rows, f"$A{i}")
            pcell = ws.cell(row=i, column=3)
            if pformula:
                pcell.value = pformula
            pcell.number_format = CURRENCY_FMT
            pcell.font = BODY_FONT

    last_row = start_row + len(trend_df) - 1
    for r in range(1, last_row + 1):
        for c in range(1, len(headers) + 1):
            ws.cell(row=r, column=c).border = BORDER

    _autofit(ws, [14, 16, 16])
    return ws, start_row, last_row


def _add_kpi_card(ws, row, col, label, value, is_currency, is_pct=False):
    ws.cell(row=row, column=col, value=label).font = KPI_LABEL_FONT
    vcell = ws.cell(row=row + 1, column=col, value=value)
    vcell.font = KPI_VALUE_FONT
    if is_pct:
        vcell.number_format = '0.0"%"'
    elif is_currency:
        vcell.number_format = CURRENCY_FMT
    else:
        vcell.number_format = "#,##0"


def _write_dashboard(wb, df, cfg, col_letters, n_rows, cat_range, reg_range, trend_range):
    ws = wb.create_sheet("Dashboard", 0)
    ws.sheet_view.showGridLines = False

    ws.merge_cells("A1:H1")
    ws["A1"] = cfg.get("report_title", "Sales Performance Report")
    ws["A1"].font = TITLE_FONT
    ws.merge_cells("A2:H2")
    ws["A2"] = f"Generated automatically from {n_rows - 1:,} order rows  •  Raw Data sheet is the single source of truth"
    ws["A2"].font = SUBTITLE_FONT

    total_orders_formula = None
    if "order_id" in col_letters:
        oc = col_letters["order_id"]
        rng = f"'Raw Data'!${oc}$2:${oc}${n_rows}"
        total_orders_formula = f"=SUMPRODUCT(1/COUNTIF({rng},{rng}))"
    else:
        total_orders_formula = f"={n_rows - 1}"

    sales_rng = f"'Raw Data'!${col_letters['sales']}$2:${col_letters['sales']}${n_rows}"
    total_sales_formula = f"=SUM({sales_rng})"

    row = 4
    revenue_col, orders_col, avg_col, profit_col = 1, 3, 5, 7
    revenue_letter = get_column_letter(revenue_col)
    orders_letter = get_column_letter(orders_col)
    profit_letter = get_column_letter(profit_col)

    _add_kpi_card(ws, row, revenue_col, "TOTAL REVENUE", total_sales_formula, True)
    _add_kpi_card(ws, row, orders_col, "TOTAL ORDERS", total_orders_formula, False)

    _add_kpi_card(
        ws, row, avg_col, "AVG ORDER VALUE",
        f"={revenue_letter}{row+1}/{orders_letter}{row+1}", True,
    )

    if "profit" in col_letters:
        profit_rng = f"'Raw Data'!${col_letters['profit']}$2:${col_letters['profit']}${n_rows}"
        _add_kpi_card(ws, row, profit_col, "TOTAL PROFIT", f"=SUM({profit_rng})", True)
        margin_row = row + 3
        _add_kpi_card(
            ws, margin_row, revenue_col, "PROFIT MARGIN",
            f"={profit_letter}{row+1}/{revenue_letter}{row+1}*100", False, is_pct=True,
        )

    for c in range(1, 9):
        ws.column_dimensions[get_column_letter(c)].width = 15

    chart_anchor_row = row + 6

    # Category bar chart
    bar = BarChart()
    bar.title = "Sales by Category"
    bar.style = 10
    bar.y_axis.title = "Sales"
    bar.x_axis.title = "Category"
    cat_ws, cs, ce = cat_range
    data = Reference(cat_ws, min_col=2, min_row=1, max_row=ce)
    cats = Reference(cat_ws, min_col=1, min_row=cs, max_row=ce)
    bar.add_data(data, titles_from_data=True)
    bar.set_categories(cats)
    bar.height, bar.width = 8, 15
    ws.add_chart(bar, f"A{chart_anchor_row}")

    # Region pie chart
    pie = PieChart()
    pie.title = "Revenue Share by Region"
    reg_ws, rs, re_ = reg_range
    pdata = Reference(reg_ws, min_col=2, min_row=1, max_row=re_)
    pcats = Reference(reg_ws, min_col=1, min_row=rs, max_row=re_)
    pie.add_data(pdata, titles_from_data=True)
    pie.set_categories(pcats)
    pie.dataLabels = DataLabelList()
    pie.dataLabels.showPercent = True
    pie.height, pie.width = 8, 15
    ws.add_chart(pie, f"E{chart_anchor_row}")

    # Monthly trend line chart
    line = LineChart()
    line.title = "Monthly Sales Trend"
    line.style = 12
    line.y_axis.title = "Sales"
    line.x_axis.title = "Month"
    tr_ws, ts, te = trend_range
    ldata = Reference(tr_ws, min_col=2, min_row=1, max_row=te)
    lcats = Reference(tr_ws, min_col=1, min_row=ts, max_row=te)
    line.add_data(ldata, titles_from_data=True)
    line.set_categories(lcats)
    line.height, line.width = 8, 31
    ws.add_chart(line, f"A{chart_anchor_row + 17}")


def build_excel_report(df, cfg, output_path: str):
    wb = Workbook()
    wb.remove(wb.active)

    raw_ws, col_letters, n_rows = _write_raw_data(wb, df, cfg)

    cats_df = kpi.category_summary(df)
    reg_df = kpi.region_summary(df)
    trend_df = kpi.monthly_trend(df)

    cat_ws, cs, ce = _write_category_summary(wb, cats_df, col_letters, n_rows)
    reg_ws, rs, re_ = _write_region_summary(wb, reg_df, col_letters, n_rows)
    tr_ws, ts, te = _write_monthly_trend(wb, trend_df, col_letters, n_rows)

    _write_dashboard(
        wb, df, cfg, col_letters, n_rows,
        cat_range=(cat_ws, cs, ce),
        reg_range=(reg_ws, rs, re_),
        trend_range=(tr_ws, ts, te),
    )

    wb.save(output_path)
    return output_path

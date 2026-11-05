# -*- coding: utf-8 -*-
"""
Water Step 05 - Market Concentration Analysis

Input:
    Water_Panel_Data.xlsx, sheet: Panel_Data
Output:
    Water_Step_05_Market_Concentration.xlsx

Place this script in the same folder as the input workbook.
"""

from pathlib import Path
import re
import math
import warnings

import numpy as np
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.chart import LineChart, Reference
from openpyxl.utils import get_column_letter


# ============================================================
# 1. USER SETTINGS
# ============================================================

INPUT_FILE = "Energydrinks_Panel_Data.xlsx"
INPUT_SHEET = "Panel_Data"
OUTPUT_FILE = "Energydrinks_Step_05_Market_Concentration.xlsx"

# A missing SubCategory value is retained in an explicit group.
UNKNOWN_CATEGORY_LABEL = "Unknown"

# If True, negative sales are retained in the validation report but
# excluded from positive-sales market-share calculations.
EXCLUDE_NEGATIVE_BRAND_PERIODS_FROM_MARKET = True

# Set to True to create Excel line charts where sufficient data exist.
CREATE_EXCEL_CHARTS = True


# ============================================================
# 2. GENERAL HELPERS
# ============================================================

def normalize_name(value):
    """Normalize a column name for robust matching."""
    return re.sub(r"[^a-z0-9]", "", str(value).strip().lower())


def find_column(columns, candidates, required=True):
    """Find a column using normalized candidate names."""
    lookup = {normalize_name(col): col for col in columns}
    for candidate in candidates:
        key = normalize_name(candidate)
        if key in lookup:
            return lookup[key]
    if required:
        raise ValueError(
            "Required column not found. Expected one of: "
            + ", ".join(candidates)
            + f"\nAvailable columns: {list(columns)}"
        )
    return None


def safe_sheet_name(name, used_names):
    """Return a valid, unique Excel worksheet name."""
    cleaned = re.sub(r"[\[\]:*?/\\]", "_", str(name)).strip("'")
    cleaned = cleaned[:31] or "Sheet"
    candidate = cleaned
    counter = 1
    while candidate.lower() in used_names:
        suffix = f"_{counter}"
        candidate = cleaned[:31 - len(suffix)] + suffix
        counter += 1
    used_names.add(candidate.lower())
    return candidate


def sum_min_count(series):
    """Sum observed values while preserving all-missing groups as NaN."""
    return series.sum(min_count=1)


def normalize_category(series):
    """Convert missing or blank category values into an explicit label."""
    result = series.astype("string").str.strip()
    result = result.mask(result.isna() | result.eq(""), UNKNOWN_CATEGORY_LABEL)
    return result.astype(str)


def parse_month_number(df, month_col, date_text_col=None):
    """Return month numbers from MonthNo, or derive them from DateText."""
    month = pd.to_numeric(df[month_col], errors="coerce") if month_col else pd.Series(np.nan, index=df.index)

    if month.isna().any() and date_text_col:
        parsed = pd.to_datetime(df[date_text_col], errors="coerce", dayfirst=False)
        month = month.fillna(parsed.dt.month)

    invalid = month.isna() | ~month.between(1, 12)
    if invalid.any():
        examples = df.loc[invalid, [c for c in [month_col, date_text_col] if c]].head(10)
        raise ValueError(
            "Some rows have missing or invalid month numbers. "
            "Please correct MonthNo/DateText before analysis.\n"
            + examples.to_string(index=False)
        )
    return month.astype(int)


def add_period_columns(df):
    """Add sortable monthly period labels."""
    result = df.copy()
    # Build a monthly datetime first, then convert it to a monthly Period.
    # PeriodIndex does not accept separate year= and month= arguments.
    monthly_dates = pd.to_datetime({
        "year": result["Year"].astype(int),
        "month": result["MonthNo"].astype(int),
        "day": 1
    }, errors="coerce")

    if monthly_dates.isna().any():
        bad_rows = result.loc[monthly_dates.isna(), ["Year", "MonthNo"]].head(10)
        raise ValueError(
            "Could not construct monthly periods for some Year/MonthNo values:\n"
            + bad_rows.to_string(index=False)
        )

    result["Period"] = monthly_dates.dt.to_period("M")
    result["Month"] = result["Period"].astype(str)
    return result


# ============================================================
# 3. MARKET CONCENTRATION INDICATORS
# ============================================================

def calculate_market_metrics(sales_values):
    """
    Calculate concentration measures from positive observed brand sales.
    Market shares are represented as proportions between 0 and 1.
    """
    values = pd.to_numeric(pd.Series(sales_values), errors="coerce")
    values = values.replace([np.inf, -np.inf], np.nan).dropna()
    values = values[values > 0]

    if values.empty:
        return {
            "Market_Size_Observed": np.nan,
            "Active_Brands": 0,
            "CR3": np.nan, "CR4": np.nan, "CR5": np.nan, "CR10": np.nan,
            "HHI_0_1": np.nan, "HHI_0_10000": np.nan,
            "Normalized_HHI": np.nan,
            "Gini": np.nan,
            "Normalized_Gini": np.nan,
            "Entropy": np.nan,
            "Effective_Brands_HHI": np.nan,
            "Effective_Brands_Entropy": np.nan,
        }

    market_size = float(values.sum())
    shares = np.sort((values / market_size).to_numpy())[::-1]
    n = len(shares)

    def cr(k):
        return float(shares[:k].sum())

    hhi = float(np.square(shares).sum())
    normalized_hhi = (
        (hhi - (1.0 / n)) / (1.0 - (1.0 / n))
        if n > 1 else np.nan
    )
    entropy = float(-np.sum(shares * np.log(shares)))

    # Gini coefficient for the distribution of positive brand market shares.
    ascending = np.sort(shares)
    if n > 1:
        ranks = np.arange(1, n + 1, dtype=float)
        gini = float(
            (2.0 * np.sum(ranks * ascending) / (n * np.sum(ascending)))
            - ((n + 1.0) / n)
        )
    else:
        gini = 0.0

    normalized_gini = (
        gini / ((n - 1.0) / n)
        if n > 1 else np.nan
    )

    return {
        "Market_Size_Observed": market_size,
        "Active_Brands": int(n),
        "CR3": cr(3),
        "CR4": cr(4),
        "CR5": cr(5),
        "CR10": cr(10),
        "HHI_0_1": hhi,
        "HHI_0_10000": hhi * 10000.0,
        "Normalized_HHI": normalized_hhi,
        "Gini": gini,
        "Normalized_Gini": normalized_gini,
        "Entropy": entropy,
        "Effective_Brands_HHI": (1.0 / hhi) if hhi > 0 else np.nan,
        "Effective_Brands_Entropy": math.exp(entropy),
    }


# ============================================================
# 4. PREPARE INPUT DATA
# ============================================================

def load_and_prepare(input_path):
    raw = pd.read_excel(input_path, sheet_name=INPUT_SHEET, engine="openpyxl")
    raw.columns = [str(c).strip() for c in raw.columns]

    caption_col = find_column(raw.columns, ["Element Caption"])
    brand_col = find_column(raw.columns, ["Brand"])
    year_col = find_column(raw.columns, ["Year"])
    month_col = find_column(raw.columns, ["MonthNo", "Month No", "Month Number"], required=False)
    date_text_col = find_column(raw.columns, ["DateText", "Date Text"], required=False)
    sales_col = find_column(raw.columns, ["Sales"])
    subcategory_cols = [
        col for col in raw.columns
        if normalize_name(col).startswith("subcategory")
    ]

    if not subcategory_cols:
        warnings.warn("No SubCategory columns were found. Only overall market analysis will be produced.")

    raw["_Source_Row_Number"] = np.arange(2, len(raw) + 2)

    caption = raw[caption_col].astype("string").str.strip().str.casefold()
    brand_text = raw[brand_col].astype("string").str.strip().str.casefold()

    is_brand_caption = caption.eq("brand").fillna(False)
    is_total_brand_caption = caption.eq("total brand").fillna(False)
    is_total_brand_name = brand_text.eq("total brand").fillna(False)
    keep = is_brand_caption & ~is_total_brand_name

    # Total Brand rows are retained separately for an independent reconciliation.
    total_brand_rows = raw.loc[is_total_brand_caption].copy()
    excluded = raw.loc[~keep].copy()
    data = raw.loc[keep].copy()

    data["Brand"] = data[brand_col].astype("string").str.strip()
    data["Brand"] = data["Brand"].replace("", pd.NA)
    data["Year"] = pd.to_numeric(data[year_col], errors="coerce")
    if data["Year"].isna().any():
        bad = data.loc[data["Year"].isna(), ["_Source_Row_Number", year_col]].head(10)
        raise ValueError("Invalid or missing Year values found:\n" + bad.to_string(index=False))
    data["Year"] = data["Year"].astype(int)

    data["MonthNo"] = parse_month_number(data, month_col, date_text_col)
    data["Sales"] = pd.to_numeric(data[sales_col], errors="coerce")

    # Keep rows with missing Brand visible in validation, but they cannot be
    # assigned to a brand-level market calculation.
    missing_brand_rows = data[data["Brand"].isna()].copy()
    data = data[data["Brand"].notna()].copy()

    for col in subcategory_cols:
        data[col] = normalize_category(data[col])

    data = add_period_columns(data)

    validation = {
        "Input_Rows": int(len(raw)),
        "Brand_Caption_Rows": int(is_brand_caption.sum()),
        "Excluded_Non_Brand_or_TotalBrand_Rows": int((~keep).sum()),
        "Brand_Rows_With_Missing_Brand_Name": int(len(missing_brand_rows)),
        "Brand_Rows_Used_For_Brand_Analysis": int(len(data)),
        "Missing_Sales_Rows": int(data["Sales"].isna().sum()),
        "Zero_Sales_Rows": int(data["Sales"].eq(0).sum()),
        "Negative_Sales_Rows": int(data["Sales"].lt(0).sum()),
        "Positive_Sales_Rows": int(data["Sales"].gt(0).sum()),
        "SubCategory_Column_Count": int(len(subcategory_cols)),
        "SubCategory_Columns": ", ".join(subcategory_cols) if subcategory_cols else "(none)",
        "Years_In_Data": ", ".join(map(str, sorted(data["Year"].unique()))),
        "Months_In_Data": int(data[["Year", "MonthNo"]].drop_duplicates().shape[0]),
    }

    # Duplicate source records at the overall Brand-Month level.
    dup_counts = data.groupby(["Brand", "Year", "MonthNo"], dropna=False).size()
    validation["Brand_Month_Groups_With_Multiple_Source_Rows"] = int((dup_counts > 1).sum())
    validation["Excess_Source_Rows_In_Brand_Month_Groups"] = int((dup_counts - 1).clip(lower=0).sum())

    # Reconcile each brand-month aggregate against its corresponding Total Brand record.
    # In this source structure, Total Brand is stored in Element Caption and the
    # actual brand name is stored in Brand (and/or Element Description).
    reconciliation = pd.DataFrame()
    if not total_brand_rows.empty:
        total_brand_rows["_Brand_Key"] = (
            total_brand_rows[brand_col].astype("string").str.strip()
        )
        total_brand_rows["_Year_Key"] = pd.to_numeric(
            total_brand_rows[year_col], errors="coerce"
        )
        total_brand_rows["_Month_Key"] = parse_month_number(
            total_brand_rows, month_col, date_text_col
        )
        total_brand_rows["_Total_Sales"] = pd.to_numeric(
            total_brand_rows[sales_col], errors="coerce"
        )
        total_brand_agg = total_brand_rows.groupby(
            ["_Brand_Key", "_Year_Key", "_Month_Key"], dropna=False
        )["_Total_Sales"].agg(sum_min_count).reset_index()

        brand_agg = data.groupby(
            ["Brand", "Year", "MonthNo"], dropna=False
        )["Sales"].agg(sum_min_count).reset_index()
        brand_agg = brand_agg.rename(columns={
            "Brand": "_Brand_Key",
            "Year": "_Year_Key",
            "MonthNo": "_Month_Key",
            "Sales": "_Brand_Sales"
        })

        reconciliation = total_brand_agg.merge(
            brand_agg,
            on=["_Brand_Key", "_Year_Key", "_Month_Key"],
            how="outer",
            indicator=True
        )
        reconciliation["Difference"] = (
            reconciliation["_Brand_Sales"] - reconciliation["_Total_Sales"]
        )
        comparable = (
            reconciliation["_Brand_Sales"].notna()
            & reconciliation["_Total_Sales"].notna()
        )
        reconciliation["Match_Status"] = "NOT_COMPARABLE"
        reconciliation.loc[comparable, "Match_Status"] = np.where(
            np.isclose(
                reconciliation.loc[comparable, "_Brand_Sales"],
                reconciliation.loc[comparable, "_Total_Sales"],
                rtol=1e-8, atol=1e-6
            ),
            "PASS",
            "MISMATCH"
        )
        reconciliation.loc[reconciliation["_merge"] == "left_only", "Match_Status"] = "TOTAL_ONLY"
        reconciliation.loc[reconciliation["_merge"] == "right_only", "Match_Status"] = "BRAND_ONLY"

        validation["Total_Brand_Rows"] = int(len(total_brand_rows))
        validation["Total_Brand_Month_Groups"] = int(len(total_brand_agg))
        validation["Brand_Total_Reconciliation_Pass"] = int(
            reconciliation["Match_Status"].eq("PASS").sum()
        )
        validation["Brand_Total_Reconciliation_Mismatch"] = int(
            reconciliation["Match_Status"].eq("MISMATCH").sum()
        )
        validation["Brand_Total_Reconciliation_Total_Only"] = int(
            reconciliation["Match_Status"].eq("TOTAL_ONLY").sum()
        )
        validation["Brand_Total_Reconciliation_Brand_Only"] = int(
            reconciliation["Match_Status"].eq("BRAND_ONLY").sum()
        )
        validation["Brand_Total_Reconciliation_Not_Comparable"] = int(
            reconciliation["Match_Status"].eq("NOT_COMPARABLE").sum()
        )
    else:
        validation["Total_Brand_Rows"] = 0
        validation["Brand_Total_Reconciliation_Status"] = (
            "NOT AVAILABLE: no rows with Element Caption = Total Brand"
        )

    negative_rows = data[data["Sales"] < 0].copy()
    return (
        data, subcategory_cols, validation, excluded, missing_brand_rows,
        negative_rows, reconciliation
    )


# ============================================================
# 5. AGGREGATE MONTHLY AND ANNUAL BRAND SALES
# ============================================================

def aggregate_brand_month(data, dimensions):
    """
    Aggregate source rows to one record per market-dimension / brand / month.
    Missing source sales remain missing if every source row in the group is missing.
    """
    keys = dimensions + ["Brand", "Year", "MonthNo", "Period", "Month"]
    grouped = data.groupby(keys, dropna=False, as_index=False).agg(
        Sales=("Sales", sum_min_count),
        Source_Rows=("Sales", "size"),
        Missing_Source_Rows=("Sales", lambda x: int(x.isna().sum())),
        Negative_Source_Rows=("Sales", lambda x: int(x.lt(0).sum())),
        Positive_Source_Rows=("Sales", lambda x: int(x.gt(0).sum())),
    )
    return grouped


def aggregate_annual(monthly_brand):
    """Aggregate monthly brand sales into annual observed sales."""
    keys = [c for c in monthly_brand.columns if c not in {
        "Sales", "Source_Rows", "Missing_Source_Rows",
        "Negative_Source_Rows", "Positive_Source_Rows",
        "Month", "MonthNo", "Period", "Year"
    }]
    annual = monthly_brand.groupby(keys + ["Year"], dropna=False, as_index=False).agg(
        Annual_Sales=("Sales", sum_min_count),
        Observed_Months=("Sales", lambda x: int(x.notna().sum())),
        Missing_Months=("Sales", lambda x: int(x.isna().sum())),
        Source_Rows=("Source_Rows", "sum"),
        Missing_Source_Rows=("Missing_Source_Rows", "sum"),
        Negative_Source_Rows=("Negative_Source_Rows", "sum"),
        Positive_Source_Rows=("Positive_Source_Rows", "sum"),
    )
    annual["Unrepresented_Months"] = (
        12 - annual["Observed_Months"] - annual["Missing_Months"]
    ).clip(lower=0)
    annual["Coverage_Rate"] = annual["Observed_Months"] / 12.0
    return annual


# ============================================================
# 6. CALCULATE MONTHLY / ANNUAL MARKET TABLES
# ============================================================

def calculate_period_table(brand_data, period_type, dimensions):
    """
    Calculate market indicators for each period and market segment.
    Missing values are not imputed. Available positive brand sales are used.
    """
    if period_type == "Monthly":
        sales_col = "Sales"
        period_cols = dimensions + ["Year", "MonthNo", "Month"]
        period_label = "Month"
    else:
        sales_col = "Annual_Sales"
        period_cols = dimensions + ["Year"]
        period_label = "Year"

    rows = []
    group_columns = dimensions + ([ "Year", "MonthNo", "Month" ] if period_type == "Monthly" else ["Year"])

    for group_key, group in brand_data.groupby(group_columns, dropna=False, sort=True):
        if not isinstance(group_key, tuple):
            group_key = (group_key,)
        key_values = dict(zip(group_columns, group_key))

        sales = pd.to_numeric(group[sales_col], errors="coerce")
        observed_brands = int(sales.notna().sum())
        missing_brands = int(sales.isna().sum())
        positive = sales[sales > 0]
        negative_brands = int(sales.lt(0).sum())

        metrics = calculate_market_metrics(positive)
        row = dict(key_values)
        row.update(metrics)
        row["Observed_Brands_With_Sales_Data"] = observed_brands
        row["Brands_With_Missing_Aggregated_Sales"] = missing_brands
        row["Negative_Brand_Periods"] = negative_brands
        row["Brand_Data_Coverage_Rate"] = (
            observed_brands / (observed_brands + missing_brands)
            if (observed_brands + missing_brands) > 0 else np.nan
        )

        if period_type == "Annual":
            row["Observed_Brand_Months"] = int(group["Observed_Months"].sum())
            row["Explicit_Missing_Brand_Months"] = int(group["Missing_Months"].sum())
            row["Unrepresented_Brand_Months"] = int(group["Unrepresented_Months"].sum())
            total_brand_months = (
                row["Observed_Brand_Months"]
                + row["Explicit_Missing_Brand_Months"]
                + row["Unrepresented_Brand_Months"]
            )
            row["Brand_Month_Coverage_Rate"] = (
                row["Observed_Brand_Months"] / total_brand_months
                if total_brand_months > 0 else np.nan
            )

        row["Period_Type"] = period_type
        rows.append(row)

    result = pd.DataFrame(rows)
    if not result.empty:
        sort_cols = dimensions + (["Year", "MonthNo"] if period_type == "Monthly" else ["Year"])
        result = result.sort_values(sort_cols).reset_index(drop=True)
    return result


def create_brand_ranking(monthly_brand, annual_brand, dimensions=None):
    """Create monthly and annual brand ranking tables."""
    dimensions = dimensions or []

    monthly = monthly_brand.copy()
    monthly["Market_Size_Observed"] = monthly.groupby(
        dimensions + ["Year", "MonthNo"], dropna=False
    )["Sales"].transform(lambda x: x[x > 0].sum())
    monthly["Market_Share"] = np.where(
        (monthly["Sales"] > 0) & (monthly["Market_Size_Observed"] > 0),
        monthly["Sales"] / monthly["Market_Size_Observed"],
        np.nan
    )
    monthly["Rank"] = monthly.groupby(
        dimensions + ["Year", "MonthNo"], dropna=False
    )["Market_Share"].rank(method="min", ascending=False)
    monthly["Active"] = monthly["Sales"] > 0

    annual = aggregate_annual(monthly_brand)
    annual["Market_Size_Observed"] = annual.groupby(
        dimensions + ["Year"], dropna=False
    )["Annual_Sales"].transform(lambda x: x[x > 0].sum())
    annual["Market_Share"] = np.where(
        (annual["Annual_Sales"] > 0) & (annual["Market_Size_Observed"] > 0),
        annual["Annual_Sales"] / annual["Market_Size_Observed"],
        np.nan
    )
    annual["Rank"] = annual.groupby(
        dimensions + ["Year"], dropna=False
    )["Market_Share"].rank(method="min", ascending=False)
    annual["Active"] = annual["Annual_Sales"] > 0

    # Annual rank change: positive values indicate an improvement in rank.
    rank_group_cols = dimensions + ["Brand"]
    annual = annual.sort_values(rank_group_cols + ["Year"]).reset_index(drop=True)
    annual["Previous_Year_Rank"] = annual.groupby(
        rank_group_cols, dropna=False
    )["Rank"].shift(1)
    annual["Rank_Change"] = (
        annual["Previous_Year_Rank"] - annual["Rank"]
    )

    # Monthly rank change is calculated only against the immediately preceding
    # calendar month; gaps are not treated as consecutive observations.
    monthly = monthly.sort_values(
        dimensions + ["Brand", "Year", "MonthNo"]
    ).reset_index(drop=True)
    monthly["_Period_Ordinal"] = monthly["Year"] * 12 + monthly["MonthNo"]
    monthly["_Previous_Rank"] = monthly.groupby(
        rank_group_cols, dropna=False
    )["Rank"].shift(1)
    monthly["_Previous_Ordinal"] = monthly.groupby(
        rank_group_cols, dropna=False
    )["_Period_Ordinal"].shift(1)
    monthly["Previous_Month_Rank"] = monthly["_Previous_Rank"].where(
        monthly["_Period_Ordinal"] - monthly["_Previous_Ordinal"] == 1
    )
    monthly["Monthly_Rank_Change"] = (
        monthly["Previous_Month_Rank"] - monthly["Rank"]
    )
    monthly = monthly.drop(
        columns=["_Period_Ordinal", "_Previous_Rank", "_Previous_Ordinal"]
    )

    return monthly, annual


def add_changes(metrics_df, dimensions, time_col, month_col=None):
    """Add first differences of selected concentration indicators."""
    result = metrics_df.copy()
    sort_cols = dimensions + [time_col] + ([month_col] if month_col else [])
    result = result.sort_values(sort_cols).reset_index(drop=True)

    indicators = [
        "CR3", "CR4", "CR5", "CR10", "HHI_0_1",
        "HHI_0_10000", "Normalized_HHI", "Gini", "Normalized_Gini", "Entropy",
        "Active_Brands", "Effective_Brands_HHI", "Effective_Brands_Entropy"
    ]
    group_cols = dimensions if dimensions else None
    for col in indicators:
        if col in result.columns:
            result[f"Delta_{col}"] = (
                result.groupby(dimensions, dropna=False)[col].diff()
                if dimensions else result[col].diff()
            )
    return result


# ============================================================
# 7. EXCEL OUTPUT FORMATTING AND CHARTS
# ============================================================

def write_dataframe(writer, df, sheet_name):
    """Write a DataFrame and apply basic worksheet formatting."""
    df.to_excel(writer, sheet_name=sheet_name, index=False)
    ws = writer.sheets[sheet_name]
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    header_fill = PatternFill("solid", fgColor="1F4E78")
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for column_cells in ws.columns:
        max_length = 0
        column_letter = get_column_letter(column_cells[0].column)
        for cell in column_cells[: min(len(column_cells), 500)]:
            value = "" if cell.value is None else str(cell.value)
            max_length = max(max_length, len(value))
        ws.column_dimensions[column_letter].width = min(max(max_length + 2, 12), 32)

    return ws


def add_line_chart(ws, title, y_title, categories_col, data_cols, anchor):
    """Add an Excel line chart using columns already present in the worksheet."""
    if ws.max_row < 3:
        return
    chart = LineChart()
    chart.title = title
    chart.style = 13
    chart.y_axis.title = y_title
    chart.x_axis.title = "Period"
    chart.height = 8
    chart.width = 16

    for col_idx in data_cols:
        data = Reference(ws, min_col=col_idx, max_col=col_idx, min_row=1, max_row=ws.max_row)
        chart.add_data(data, titles_from_data=True)

    cats = Reference(ws, min_col=categories_col, min_row=2, max_row=ws.max_row)
    chart.set_categories(cats)
    ws.add_chart(chart, anchor)


def make_readme():
    rows = [
        ["Item", "Definition / Method"],
        ["Purpose", "Market concentration analysis for the overall market and each SubCategory dimension independently."],
        ["Input", f"{INPUT_FILE} | Sheet: {INPUT_SHEET}"],
        ["Brand filter", "Only rows where Element Caption equals Brand are included. Total Brand rows are excluded."],
        ["Overall market", "Brand sales are aggregated across all SubCategory values before calculating overall market shares."],
        ["SubCategory analysis", "Each SubCategory column is analyzed independently. SubCategory columns are not treated as a hierarchy."],
        ["Missing sales", "Missing values are not converted to zero. A brand-period with all source values missing remains missing after aggregation."],
        ["Observed-data approach", "Indicators are calculated from observed positive brand sales. Coverage and missing-brand counts are reported alongside the indicators."],
        ["Annual sales", "Annual brand sales are the sum of observed monthly sales. Observed and missing months are reported separately."],
        ["CR3 / CR4 / CR5 / CR10", "Sum of the market shares of the 3, 4, 5, or 10 largest active brands. If fewer brands exist, all active brands are included."],
        ["HHI", "Sum of squared brand market shares. Reported on 0-1 and 0-10,000 scales."],
        ["Normalized HHI", "(HHI - 1/N) / (1 - 1/N), where N is the number of active brands. Undefined for N=1."],
        ["Gini", "Gini coefficient of the distribution of positive active-brand market shares."],
        ["Normalized Gini", "Gini divided by its finite-sample maximum ((N-1)/N), i.e. Gini * N/(N-1). Undefined for N=1."],
        ["Rank changes", "Previous-period rank minus current rank. Positive values indicate rank improvement; annual change compares consecutive represented years."],
        ["SubCategory market share", "Observed positive sales in a SubCategory divided by observed positive sales in the overall market for the same month/year."],
        ["Brand vs Total Brand validation", "For each brand-year-month, the sum of Brand-caption sales is compared with the corresponding Total Brand-caption sales. Missing/unmatched groups are separately classified."],
        ["Entropy", "Shannon entropy using natural logarithms: -sum(share * ln(share))."],
        ["Effective brands - HHI", "1 / HHI."],
        ["Effective brands - Entropy", "exp(Entropy)."],
        ["Negative sales", "Negative observations are reported in Validation. They are not treated as active positive sales for concentration indicators."],
        ["Interpretation warning", "Indicators based on observed sales may not represent true market concentration if missingness is systematic."],
        ["Partial years", "Annual results are based on observed months; they must not be interpreted as complete-year sales when coverage is below 12 months."],
        ["Coverage", "Monthly Brand_Data_Coverage_Rate is the share of represented brand-month aggregates with observed sales. Annual Coverage_Rate is Observed_Months / 12. Missing_Months counts explicit all-missing months; Unrepresented_Months counts months with no brand-month record in the panel."],
    ]
    return pd.DataFrame(rows[1:], columns=rows[0])


# ============================================================
# 8. MAIN
# ============================================================

def main():
    base_dir = Path(__file__).resolve().parent
    input_path = base_dir / INPUT_FILE
    output_path = base_dir / OUTPUT_FILE

    if not input_path.exists():
        raise FileNotFoundError(
            f"Input file not found:\n{input_path}\n"
            "Place the Excel input file in the same folder as this Python script."
        )

    print("Loading and validating input data...")
    (
        data, subcategory_cols, validation_dict, excluded, missing_brand_rows,
        negative_rows, reconciliation
    ) = load_and_prepare(input_path)

    used_sheet_names = set()
    output_tables = {}

    # Overall market: aggregate across all SubCategory dimensions.
    overall_monthly_brand = aggregate_brand_month(data, dimensions=[])
    overall_annual_brand = aggregate_annual(overall_monthly_brand)

    overall_monthly = calculate_period_table(overall_monthly_brand, "Monthly", dimensions=[])
    overall_annual = calculate_period_table(overall_annual_brand, "Annual", dimensions=[])

    overall_monthly = add_changes(overall_monthly, [], "Year", "MonthNo")
    overall_annual = add_changes(overall_annual, [], "Year")

    output_tables["Overall_Monthly"] = overall_monthly
    output_tables["Overall_Annual"] = overall_annual

    # Rankings for the overall market.
    rank_monthly, rank_annual = create_brand_ranking(
        overall_monthly_brand, overall_annual_brand, dimensions=[]
    )
    output_tables["Brand_Ranking_Monthly"] = rank_monthly
    output_tables["Brand_Ranking_Annual"] = rank_annual

    # Each SubCategory column is analyzed independently.
    for index, subcat_col in enumerate(subcategory_cols, start=1):
        monthly_brand = aggregate_brand_month(data, dimensions=[subcat_col])
        annual_brand = aggregate_annual(monthly_brand)

        monthly_metrics = calculate_period_table(monthly_brand, "Monthly", dimensions=[subcat_col])
        annual_metrics = calculate_period_table(annual_brand, "Annual", dimensions=[subcat_col])

        monthly_metrics = add_changes(monthly_metrics, [subcat_col], "Year", "MonthNo")
        annual_metrics = add_changes(annual_metrics, [subcat_col], "Year")

        # Segment size as a share of the corresponding overall market.
        overall_monthly_size = overall_monthly[
            ["Year", "MonthNo", "Market_Size_Observed"]
        ].rename(columns={
            "Market_Size_Observed": "Overall_Market_Size_Observed"
        })
        overall_annual_size = overall_annual[
            ["Year", "Market_Size_Observed"]
        ].rename(columns={
            "Market_Size_Observed": "Overall_Market_Size_Observed"
        })

        monthly_metrics = monthly_metrics.merge(
            overall_monthly_size, on=["Year", "MonthNo"], how="left"
        )
        monthly_metrics["Share_of_Overall_Market"] = np.where(
            monthly_metrics["Overall_Market_Size_Observed"] > 0,
            monthly_metrics["Market_Size_Observed"]
            / monthly_metrics["Overall_Market_Size_Observed"],
            np.nan
        )

        annual_metrics = annual_metrics.merge(
            overall_annual_size, on=["Year"], how="left"
        )
        annual_metrics["Share_of_Overall_Market"] = np.where(
            annual_metrics["Overall_Market_Size_Observed"] > 0,
            annual_metrics["Market_Size_Observed"]
            / annual_metrics["Overall_Market_Size_Observed"],
            np.nan
        )

        monthly_rank, annual_rank = create_brand_ranking(
            monthly_brand, annual_brand, dimensions=[subcat_col]
        )

        output_tables[f"SubCat{index}_Monthly"] = monthly_metrics
        output_tables[f"SubCat{index}_Annual"] = annual_metrics
        output_tables[f"SubCat{index}_Brand_Monthly"] = monthly_rank
        output_tables[f"SubCat{index}_Brand_Annual"] = annual_rank

    # Validation summary.
    validation_rows = [{"Check": key, "Result": value} for key, value in validation_dict.items()]
    validation_rows.extend([
        {"Check": "Negative_Sales_Policy", "Result": "Reported; negative values are not included in positive-sales concentration calculations."},
        {"Check": "Overall_Monthly_Metric_Periods", "Result": int(len(overall_monthly))},
        {"Check": "Overall_Annual_Metric_Periods", "Result": int(len(overall_annual))},
        {"Check": "Brand_Total_Reconciliation_Detail_Sheet", "Result": (
            "Created" if not reconciliation.empty else "Not created: Total Brand rows not found"
        )},
    ])

    # Check that source-level total observed sales reconcile with overall aggregated sales.
    source_observed_total = data["Sales"].sum(min_count=1)
    aggregated_observed_total = overall_monthly_brand["Sales"].sum(min_count=1)
    validation_rows.append({
        "Check": "Observed_Sales_Reconciliation_Source_vs_Overall_Aggregate",
        "Result": (
            "PASS" if (
                (pd.isna(source_observed_total) and pd.isna(aggregated_observed_total))
                or np.isclose(source_observed_total, aggregated_observed_total, equal_nan=True)
            ) else f"CHECK: source={source_observed_total}, aggregate={aggregated_observed_total}"
        )
    })

    validation_df = pd.DataFrame(validation_rows)

    # Preserve row-level exceptions in dedicated validation sheets.
    excluded_report = excluded.copy()
    missing_brand_report = missing_brand_rows.copy()
    negative_report = negative_rows.copy()

    # Workbook writing.
    print("Calculating indicators and writing Excel output...")
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        readme = make_readme()
        write_dataframe(writer, readme, safe_sheet_name("README", used_sheet_names))
        write_dataframe(writer, validation_df, safe_sheet_name("Validation", used_sheet_names))

        for table_name, table in output_tables.items():
            sheet = safe_sheet_name(table_name, used_sheet_names)
            write_dataframe(writer, table, sheet)

        if not excluded_report.empty:
            write_dataframe(writer, excluded_report, safe_sheet_name("Excluded_Rows", used_sheet_names))
        if not missing_brand_report.empty:
            write_dataframe(writer, missing_brand_report, safe_sheet_name("Missing_Brand_Rows", used_sheet_names))
        if not negative_report.empty:
            write_dataframe(writer, negative_report, safe_sheet_name("Negative_Sales_Rows", used_sheet_names))
        if not reconciliation.empty:
            write_dataframe(
                writer, reconciliation,
                safe_sheet_name("Brand_Total_Reconciliation", used_sheet_names)
            )

    # Add charts to the overall annual metrics sheet.
    if CREATE_EXCEL_CHARTS and not overall_annual.empty:
        wb = load_workbook(output_path)
        ws = wb["Overall_Annual"]

        header_map = {cell.value: cell.column for cell in ws[1]}
        year_col = header_map.get("Year")
        if year_col:
            chart_columns = [
                header_map[name] for name in ["CR3", "CR5", "CR10"]
                if name in header_map
            ]
            if chart_columns:
                add_line_chart(ws, "Annual Market Concentration - CRn", "Market share", year_col, chart_columns, "A20")

            hhi_cols = [
                header_map[name] for name in ["HHI_0_1", "Normalized_HHI"]
                if name in header_map
            ]
            if hhi_cols:
                add_line_chart(ws, "Annual HHI Indicators", "Index", year_col, hhi_cols, "J20")

            other_cols = [
                header_map[name] for name in [
                    "Gini", "Normalized_Gini", "Entropy", "Active_Brands"
                ] if name in header_map
            ]
            if other_cols:
                add_line_chart(ws, "Annual Market Structure Indicators", "Value", year_col, other_cols, "A36")

        # Add independent annual charts for every SubCategory dimension.
        for index, _subcat_col in enumerate(subcategory_cols, start=1):
            sheet_name = f"SubCat{index}_Annual"
            if sheet_name not in wb.sheetnames:
                continue
            sub_ws = wb[sheet_name]
            sub_headers = {cell.value: cell.column for cell in sub_ws[1]}
            sub_year_idx = sub_headers.get("Year")
            if not sub_year_idx or sub_ws.max_row < 3:
                continue

            cr_cols = [
                sub_headers[name] for name in ["CR3", "CR5", "CR10"]
                if name in sub_headers
            ]
            if cr_cols:
                add_line_chart(
                    sub_ws, f"SubCategory {index} - CRn",
                    "Market share", sub_year_idx, cr_cols, "A20"
                )

            hhi_cols = [
                sub_headers[name] for name in ["HHI_0_10000"]
                if name in sub_headers
            ]
            if hhi_cols:
                add_line_chart(
                    sub_ws, f"SubCategory {index} - HHI",
                    "HHI (0-10,000)", sub_year_idx, hhi_cols, "J20"
                )

            inequality_cols = [
                sub_headers[name] for name in ["Gini", "Normalized_Gini"]
                if name in sub_headers
            ]
            if inequality_cols:
                add_line_chart(
                    sub_ws, f"SubCategory {index} - Inequality",
                    "Index", sub_year_idx, inequality_cols, "A36"
                )

            share_col = sub_headers.get("Share_of_Overall_Market")
            if share_col:
                add_line_chart(
                    sub_ws, f"SubCategory {index} - Share of Overall Market",
                    "Market share", sub_year_idx, [share_col], "J36"
                )

        wb.save(output_path)

    print("\nAnalysis completed successfully.")
    print(f"Output file:\n{output_path}")
    print(f"SubCategory dimensions analyzed: {len(subcategory_cols)}")
    print("Note: Review the Validation sheet and interpret partial-year results cautiously.")


if __name__ == "__main__":
    main()

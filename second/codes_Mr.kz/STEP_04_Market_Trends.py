# ============================================================
# STEP 04 - MARKET TRENDS ANALYSIS
# FMCG Bottled Water / Drinking Water Market
#
# Input:
#   Step_02_Panel_Data.xlsx
#   Sheet: Panel_Data
#
# Output:
#   Step_04_Market_Trends.xlsx
#
# Required packages:
#   pandas, numpy, openpyxl
#
# Run this script from VS Code. Keep the input workbook in the
# same folder as this Python file.
# ============================================================

from pathlib import Path
from itertools import product
import re
import warnings

import numpy as np
import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


# ============================================================
# 1. CONFIGURATION
# ============================================================

BASE_DIR = Path(__file__).resolve().parent

INPUT_FILE = BASE_DIR / "Energydrinks_Panel_Data.xlsx"
INPUT_SHEET = "Panel_Data"

OUTPUT_FILE = BASE_DIR / "Energydrinks_Step_04_Market_Trends.xlsx"

PREFERRED_BASE_YEAR = 2016
MONTHS_PER_YEAR = 12
MIN_YEARS_FOR_SEASONALITY = 3
EXCEL_SHEET_LIMIT = 31

# Brand lifecycle and trend classification rules.
ENTRY_CONFIRMATION_MONTHS = 12
EXIT_CONFIRMATION_MONTHS = 12
STABLE_GROWTH_THRESHOLD_PCT = 5.0

# Brand lifecycle and trend classification rules.
# Entry/exit are confirmed only after 12 consecutive complete zero-sales months.
ENTRY_CONFIRMATION_MONTHS = 12
EXIT_CONFIRMATION_MONTHS = 12
# Annual brand trend is compared over the latest two consecutive complete
# calendar years shared by the whole market (not a different period per brand).
STABLE_GROWTH_THRESHOLD_PCT = 5.0

# A brand is considered within its observed lifecycle proxy from
# its first to its last month with a source record.
# Missing months inside this interval are treated as data gaps.
# Months outside this interval are not assumed to be zero sales.
#
# This is an operational proxy, not proof of actual market entry
# or exit dates. Confirm entry/exit dates with independent evidence.

# Incomplete market coverage rule:
# market sales are published only when every brand whose observed
# record window includes that month has complete sales data.
# Partial missing rows retain the sum of observed values separately,
# but the month is flagged incomplete and Market_Sales is blank.

# Seasonality is calculated whenever at least one complete year exists.
# If fewer than MIN_YEARS_FOR_SEASONALITY years are available,
# the estimate is retained but flagged as limited evidence.

# Negative sales are retained and reported. They are not silently
# removed or converted to zero.

# ============================================================
# 2. GENERAL UTILITIES
# ============================================================

def clean_text(value):
    """Trim text values and preserve true missing values."""
    if pd.isna(value):
        return np.nan
    value = str(value).strip()
    return np.nan if value == "" else value


def safe_sheet_name(name, used_names):
    """Return a valid, unique Excel worksheet name."""
    name = re.sub(r"[\[\]:*?/\\]", "_", str(name))
    name = name[:EXCEL_SHEET_LIMIT] or "Sheet"
    candidate = name
    counter = 1
    while candidate in used_names:
        suffix = f"_{counter}"
        candidate = name[:EXCEL_SHEET_LIMIT - len(suffix)] + suffix
        counter += 1
    used_names.add(candidate)
    return candidate


def safe_divide(numerator, denominator):
    """Divide Series while excluding missing and zero denominators."""
    numerator = pd.to_numeric(numerator, errors="coerce")
    denominator = pd.to_numeric(denominator, errors="coerce")
    result = numerator / denominator
    return result.where(denominator.notna() & denominator.ne(0))


def sum_min_count(series):
    """Sum observed values while keeping all-missing groups missing."""
    return series.sum(min_count=1)


def as_period_month(series):
    """Convert supported monthly values to monthly Period values."""
    if isinstance(series.dtype, pd.PeriodDtype):
        return series.dt.asfreq("M")
    return pd.to_datetime(series, errors="coerce").dt.to_period("M")


def get_complete_years(data, value_col, complete_col="CompleteCoverage"):
    """Identify years with 12 calendar months and 12 complete months."""
    temp = data.copy()
    temp["Month"] = as_period_month(temp["Month"])
    temp = temp.loc[temp["Month"].notna()].copy()
    if temp.empty:
        return pd.DataFrame(
            columns=["Year", "Months_Observed", "Complete_Months", "CompleteYear"]
        )

    temp["Year"] = temp["Month"].dt.year
    result = (
        temp.groupby("Year", as_index=False)
        .agg(
            Months_Observed=("Month", "nunique"),
            Complete_Months=(complete_col, "sum"),
            Nonmissing_Months=(value_col, "count"),
        )
    )
    result["CompleteYear"] = (
        result["Months_Observed"].eq(MONTHS_PER_YEAR)
        & result["Complete_Months"].eq(MONTHS_PER_YEAR)
        & result["Nonmissing_Months"].eq(MONTHS_PER_YEAR)
    )
    return result


def choose_base_year(data, value_col, preferred_year=PREFERRED_BASE_YEAR):
    """Choose preferred complete base year, otherwise first complete year."""
    year_status = get_complete_years(data, value_col)
    valid_years = year_status.loc[year_status["CompleteYear"], "Year"].tolist()

    if not valid_years:
        return np.nan, "No complete year available"

    if preferred_year in valid_years:
        return preferred_year, f"Preferred base year {preferred_year} is complete"

    selected_year = min(valid_years)
    return selected_year, (
        f"{preferred_year} is not complete; first complete year "
        f"({selected_year}) selected"
    )


def add_calendar_lags(data, value_col, complete_col="CompleteCoverage"):
    """Calculate calendar MoM and YoY growth using exact calendar lags."""
    result = data.copy()
    result["Month"] = as_period_month(result["Month"])
    result = result.sort_values("Month").reset_index(drop=True)

    lookup = result.set_index("Month")
    previous_month = result["Month"] - 1
    previous_year_month = result["Month"] - 12

    result["Previous_Month_Sales"] = previous_month.map(lookup[value_col])
    result["Previous_Month_Complete"] = (
        previous_month.map(lookup[complete_col]).fillna(False).astype(bool)
    )
    result["Previous_Year_Month_Sales"] = previous_year_month.map(lookup[value_col])
    result["Previous_Year_Month_Complete"] = (
        previous_year_month.map(lookup[complete_col]).fillna(False).astype(bool)
    )

    current_sales = pd.to_numeric(result[value_col], errors="coerce")
    prev_mom = pd.to_numeric(result["Previous_Month_Sales"], errors="coerce")
    prev_yoy = pd.to_numeric(result["Previous_Year_Month_Sales"], errors="coerce")
    current_complete = result[complete_col].fillna(False).astype(bool)

    valid_mom = (
        current_complete & result["Previous_Month_Complete"]
        & current_sales.notna() & prev_mom.notna() & prev_mom.ne(0)
    )
    valid_yoy = (
        current_complete & result["Previous_Year_Month_Complete"]
        & current_sales.notna() & prev_yoy.notna() & prev_yoy.ne(0)
    )

    result["MoM_Growth_%"] = np.nan
    result["YoY_Growth_%"] = np.nan
    result.loc[valid_mom, "MoM_Growth_%"] = (
        (current_sales[valid_mom] / prev_mom[valid_mom]) - 1
    ) * 100
    result.loc[valid_yoy, "YoY_Growth_%"] = (
        (current_sales[valid_yoy] / prev_yoy[valid_yoy]) - 1
    ) * 100

    result["MoM_Status"] = "Not Calculated"
    result["YoY_Status"] = "Not Calculated"

    result.loc[~current_complete, ["MoM_Status", "YoY_Status"]] = (
        "Current Month Incomplete"
    )
    result.loc[
        current_complete & ~result["Previous_Month_Complete"], "MoM_Status"
    ] = "Previous Calendar Month Missing or Incomplete"
    result.loc[
        current_complete & ~result["Previous_Year_Month_Complete"], "YoY_Status"
    ] = "Previous Year Month Missing or Incomplete"
    result.loc[
        current_complete & result["Previous_Month_Complete"] & prev_mom.eq(0),
        "MoM_Status"
    ] = "Previous Month Zero"
    result.loc[
        current_complete & result["Previous_Year_Month_Complete"] & prev_yoy.eq(0),
        "YoY_Status"
    ] = "Previous Year Month Zero"
    result.loc[valid_mom, "MoM_Status"] = "Calculated"
    result.loc[valid_yoy, "YoY_Status"] = "Calculated"

    return result


def add_ytd_yoy(data, value_col, complete_col="CompleteCoverage"):
    """Calculate comparable Jan-to-current-month YTD YoY values."""
    result = data.copy()
    result["Month"] = as_period_month(result["Month"])
    result = result.sort_values("Month").reset_index(drop=True)

    result["YTD_Current_Sales"] = np.nan
    result["YTD_Previous_Year_Sales"] = np.nan
    result["YTD_YoY_Growth_%"] = np.nan
    result["YTD_Status"] = "Not Calculated"

    lookup = result.set_index("Month")
    for idx, row in result.iterrows():
        month = row["Month"]
        if pd.isna(month):
            continue

        current_periods = pd.period_range(
            start=pd.Period(year=month.year, month=1, freq="M"),
            end=month,
            freq="M",
        )
        previous_periods = current_periods - 12

        current_rows = lookup.reindex(current_periods)
        previous_rows = lookup.reindex(previous_periods)

        current_ok = (
            len(current_rows) == month.month
            and current_rows[complete_col].fillna(False).astype(bool).all()
            and current_rows[value_col].notna().all()
        )
        previous_ok = (
            len(previous_rows) == month.month
            and previous_rows[complete_col].fillna(False).astype(bool).all()
            and previous_rows[value_col].notna().all()
        )

        if not current_ok:
            result.at[idx, "YTD_Status"] = "Current YTD Incomplete"
            continue
        if not previous_ok:
            result.at[idx, "YTD_Status"] = "Previous YTD Incomplete or Missing"
            continue

        current_total = current_rows[value_col].sum(min_count=12 if month.month == 12 else month.month)
        previous_total = previous_rows[value_col].sum(min_count=12 if month.month == 12 else month.month)

        result.at[idx, "YTD_Current_Sales"] = current_total
        result.at[idx, "YTD_Previous_Year_Sales"] = previous_total

        if pd.isna(previous_total) or previous_total == 0:
            result.at[idx, "YTD_Status"] = "Previous YTD Zero or Missing"
        else:
            result.at[idx, "YTD_YoY_Growth_%"] = (
                (current_total / previous_total) - 1
            ) * 100
            result.at[idx, "YTD_Status"] = "Calculated"

    return result


def add_monthly_trend_index(data, value_col, base_year):
    """Index monthly sales to the average monthly sales in base year."""
    result = data.copy()
    result["Month"] = as_period_month(result["Month"])
    result["Year"] = result["Month"].dt.year
    result["Trend_Index"] = np.nan
    result["Trend_Base_Year"] = base_year

    if pd.isna(base_year):
        return result

    base_rows = result.loc[
        result["Year"].eq(int(base_year))
        & result["CompleteCoverage"].fillna(False)
        & result[value_col].notna()
    ]
    if len(base_rows) != 12 or base_rows["Month"].nunique() != 12:
        return result

    base_value = base_rows[value_col].mean()
    if pd.notna(base_value) and base_value != 0:
        valid = result["CompleteCoverage"].fillna(False) & result[value_col].notna()
        result.loc[valid, "Trend_Index"] = (
            result.loc[valid, value_col] / base_value
        ) * 100
    return result


def calculate_annual_trends(monthly_data, value_col, base_year=None):
    """Calculate annual totals and growth only for complete consecutive years."""
    temp = monthly_data.copy()
    temp["Month"] = as_period_month(temp["Month"])
    temp = temp.loc[temp["Month"].notna()].copy()
    temp["Year"] = temp["Month"].dt.year

    annual = (
        temp.groupby("Year", as_index=False)
        .agg(
            Observed_Annual_Sales=(value_col, sum_min_count),
            Months_Observed=("Month", "nunique"),
            Complete_Months=("CompleteCoverage", "sum"),
            Nonmissing_Months=(value_col, "count"),
        )
    )
    annual["CompleteYear"] = (
        annual["Months_Observed"].eq(12)
        & annual["Complete_Months"].eq(12)
        & annual["Nonmissing_Months"].eq(12)
    )
    annual["Annual_Sales"] = annual["Observed_Annual_Sales"].where(
        annual["CompleteYear"]
    )
    annual = annual.sort_values("Year").reset_index(drop=True)

    previous = annual[["Year", "Annual_Sales", "CompleteYear"]].copy()
    previous["Year"] = previous["Year"] + 1
    previous = previous.rename(
        columns={
            "Annual_Sales": "Previous_Year_Sales",
            "CompleteYear": "Previous_Year_Complete",
        }
    )
    annual = annual.merge(previous, on="Year", how="left")
    annual["Previous_Year_Complete"] = (
        annual["Previous_Year_Complete"].fillna(False).astype(bool)
    )

    valid = (
        annual["CompleteYear"]
        & annual["Previous_Year_Complete"]
        & annual["Annual_Sales"].notna()
        & annual["Previous_Year_Sales"].notna()
        & annual["Previous_Year_Sales"].ne(0)
    )
    annual["Annual_Growth_%"] = np.nan
    annual.loc[valid, "Annual_Growth_%"] = (
        annual.loc[valid, "Annual_Sales"]
        / annual.loc[valid, "Previous_Year_Sales"] - 1
    ) * 100

    annual["Annual_Growth_Status"] = "Not Calculated"
    annual.loc[~annual["CompleteYear"], "Annual_Growth_Status"] = (
        "Current Year Incomplete"
    )
    annual.loc[
        annual["CompleteYear"] & ~annual["Previous_Year_Complete"],
        "Annual_Growth_Status"
    ] = "Previous Year Incomplete or Missing"
    annual.loc[
        annual["CompleteYear"]
        & annual["Previous_Year_Complete"]
        & annual["Previous_Year_Sales"].eq(0),
        "Annual_Growth_Status"
    ] = "Previous Year Zero"
    annual.loc[valid, "Annual_Growth_Status"] = "Calculated"

    if base_year is None:
        complete = annual.loc[annual["CompleteYear"] & annual["Annual_Sales"].notna()]
        base_year = int(complete["Year"].min()) if not complete.empty else np.nan

    annual["Annual_Trend_Index"] = np.nan
    annual["Annual_Base_Year"] = base_year
    if pd.notna(base_year):
        base_values = annual.loc[
            annual["Year"].eq(int(base_year)), "Annual_Sales"
        ]
        if not base_values.empty and pd.notna(base_values.iloc[0]) and base_values.iloc[0] != 0:
            valid_index = annual["Annual_Sales"].notna()
            annual.loc[valid_index, "Annual_Trend_Index"] = (
                annual.loc[valid_index, "Annual_Sales"] / base_values.iloc[0]
            ) * 100

    return annual


def calculate_seasonality(monthly_data, value_col, group_cols=None):
    """Normalize each month by its own year's mean monthly sales."""
    if group_cols is None:
        group_cols = []

    temp = monthly_data.copy()
    temp["Month"] = as_period_month(temp["Month"])
    temp = temp.loc[temp["Month"].notna()].copy()
    temp["Year"] = temp["Month"].dt.year
    temp["MonthNo"] = temp["Month"].dt.month

    grouping = group_cols + ["Year"]
    annual_status = (
        temp.groupby(grouping, as_index=False)
        .agg(
            Months_Observed=("Month", "nunique"),
            Complete_Months=("CompleteCoverage", "sum"),
            Nonmissing_Months=(value_col, "count"),
        )
    )
    annual_status["CompleteYear"] = (
        annual_status["Months_Observed"].eq(12)
        & annual_status["Complete_Months"].eq(12)
        & annual_status["Nonmissing_Months"].eq(12)
    )

    temp = temp.merge(
        annual_status[group_cols + ["Year", "CompleteYear"]],
        on=group_cols + ["Year"],
        how="left",
    )
    temp = temp.loc[temp["CompleteYear"].fillna(False)].copy()

    if temp.empty:
        return pd.DataFrame(
            columns=group_cols + [
                "MonthNo", "Seasonality_Index", "Seasonality_Median",
                "Seasonality_Std", "Years_Used", "Seasonality_Evidence"
            ]
        )

    annual_means = (
        temp.groupby(group_cols + ["Year"], as_index=False)[value_col]
        .mean()
        .rename(columns={value_col: "Annual_Monthly_Mean"})
    )
    temp = temp.merge(
        annual_means,
        on=group_cols + ["Year"],
        how="left",
    )
    temp["Seasonality_Index_Year"] = (
        temp[value_col] / temp["Annual_Monthly_Mean"]
    ) * 100

    seasonality = (
        temp.groupby(group_cols + ["MonthNo"], as_index=False)
        .agg(
            Seasonality_Index=("Seasonality_Index_Year", "mean"),
            Seasonality_Median=("Seasonality_Index_Year", "median"),
            Seasonality_Std=("Seasonality_Index_Year", "std"),
            Years_Used=("Year", "nunique"),
        )
    )
    seasonality["Seasonality_Evidence"] = np.where(
        seasonality["Years_Used"] >= MIN_YEARS_FOR_SEASONALITY,
        "At least 3 complete years",
        "Limited evidence: fewer than 3 complete years",
    )
    return seasonality


# ============================================================
# 3. BUILD MONTHLY ENTITY PANEL
# ============================================================

def build_entity_month_panel(raw, entity_cols, all_months):
    """
    Aggregate source rows to entity-month, then create a monthly grid.
    Missing months inside the observed record window are data gaps.
    Outside-window months are not imputed as zero.
    """
    observed = (
        raw.groupby(entity_cols + ["Month"], as_index=False, dropna=False)
        .agg(
            Observed_Sales=("Sales", sum_min_count),
            Source_Rows=("Sales", "size"),
            Missing_Sales_Rows=("Sales_Is_Missing", "sum"),
            Negative_Sales_Rows=("Sales_Is_Negative", "sum"),
        )
    )
    observed["Has_Record"] = True

    entity_values = observed[entity_cols].drop_duplicates().copy()
    month_frame = pd.DataFrame({"Month": all_months})
    grid = entity_values.merge(month_frame, how="cross")
    panel = grid.merge(observed, on=entity_cols + ["Month"], how="left")

    panel["Has_Record"] = panel["Has_Record"].fillna(False).astype(bool)
    for col in ["Source_Rows", "Missing_Sales_Rows", "Negative_Sales_Rows"]:
        panel[col] = panel[col].fillna(0).astype(int)

    panel["Observed_Sales"] = pd.to_numeric(
        panel["Observed_Sales"], errors="coerce"
    )

    bounds = (
        observed.groupby(entity_cols, as_index=False, dropna=False)
        .agg(
            First_Observed_Month=("Month", "min"),
            Last_Observed_Month=("Month", "max"),
        )
    )
    panel = panel.merge(bounds, on=entity_cols, how="left")
    panel["Within_Observed_Window"] = (
        panel["Month"].ge(panel["First_Observed_Month"])
        & panel["Month"].le(panel["Last_Observed_Month"])
    )

    panel["Record_Status"] = "No_Record_Outside_Observed_Window"
    panel.loc[
        panel["Within_Observed_Window"] & ~panel["Has_Record"],
        "Record_Status"
    ] = "No_Record_Within_Observed_Window"
    panel.loc[
        panel["Has_Record"] & panel["Missing_Sales_Rows"].gt(0)
        & panel["Observed_Sales"].notna(),
        "Record_Status"
    ] = "Partial_Missing_Sales"
    panel.loc[
        panel["Has_Record"] & panel["Observed_Sales"].isna(),
        "Record_Status"
    ] = "All_Sales_Missing"
    panel.loc[
        panel["Has_Record"] & panel["Missing_Sales_Rows"].eq(0)
        & panel["Observed_Sales"].notna(),
        "Record_Status"
    ] = "Observed"

    panel["CompleteCoverage"] = (
        panel["Has_Record"]
        & panel["Missing_Sales_Rows"].eq(0)
        & panel["Observed_Sales"].notna()
    )
    panel["Sales"] = panel["Observed_Sales"].where(
        panel["CompleteCoverage"]
    )
    panel["Active_Window"] = panel["Within_Observed_Window"]

    return panel.sort_values(entity_cols + ["Month"]).reset_index(drop=True)


def build_category_month_panel(pair_panel, dimension, all_months):
    """Aggregate active brand-category pairs to independent category-months."""
    active = pair_panel.loc[pair_panel["Active_Window"]].copy()

    if active.empty:
        return pd.DataFrame()

    category_month = (
        active.groupby([dimension, "Month"], as_index=False)
        .agg(
            Observed_Sales=("Observed_Sales", sum_min_count),
            Expected_Brand_Category_Pairs=("Brand", "size"),
            Complete_Brand_Category_Pairs=("CompleteCoverage", "sum"),
            Source_Rows=("Source_Rows", "sum"),
            Missing_Sales_Rows=("Missing_Sales_Rows", "sum"),
            Negative_Sales_Rows=("Negative_Sales_Rows", "sum"),
        )
    )
    category_month["Has_Record"] = True
    category_month["CompleteCoverage"] = (
        category_month["Expected_Brand_Category_Pairs"].gt(0)
        & category_month["Complete_Brand_Category_Pairs"].eq(
            category_month["Expected_Brand_Category_Pairs"]
        )
    )
    category_month["Category_Sales"] = category_month[
        "Observed_Sales"
    ].where(category_month["CompleteCoverage"])
    category_month["Record_Status"] = np.where(
        category_month["CompleteCoverage"],
        "Complete",
        "Incomplete: one or more active brand-category pairs incomplete",
    )

    categories = pair_panel[dimension].drop_duplicates().tolist()
    grid = pd.DataFrame({dimension: categories}).merge(
        pd.DataFrame({"Month": all_months}), how="cross"
    )
    result = grid.merge(
        category_month,
        on=[dimension, "Month"],
        how="left",
    )
    for col in [
        "Expected_Brand_Category_Pairs",
        "Complete_Brand_Category_Pairs",
        "Source_Rows",
        "Missing_Sales_Rows",
        "Negative_Sales_Rows",
    ]:
        result[col] = result[col].fillna(0).astype(int)
    result["Has_Record"] = result["Has_Record"].fillna(False).astype(bool)
    result["CompleteCoverage"] = result["CompleteCoverage"].fillna(False).astype(bool)
    result["Record_Status"] = result["Record_Status"].fillna("No_Active_Category_Record")
    result["Observed_Sales"] = pd.to_numeric(result["Observed_Sales"], errors="coerce")
    result["Category_Sales"] = pd.to_numeric(result["Category_Sales"], errors="coerce")
    return result.sort_values([dimension, "Month"]).reset_index(drop=True)


# ============================================================
# 4. LOAD AND VALIDATE INPUT DATA
# ============================================================

def main():
    if not INPUT_FILE.exists():
        raise FileNotFoundError(
            f"Input file not found: {INPUT_FILE}\n"
            "Place Step_02_Panel_Data.xlsx in the same folder as this script."
        )

    print("Loading Step 02 panel data...")
    df = pd.read_excel(
        INPUT_FILE,
        sheet_name=INPUT_SHEET,
        engine="openpyxl",
    )
    df.columns = [str(col).strip() for col in df.columns]

    required = ["Element Caption", "Brand", "Sales"]
    missing_required = [col for col in required if col not in df.columns]
    if missing_required:
        raise ValueError(f"Missing required columns: {missing_required}")

    input_rows = len(df)
    df["Element Caption"] = df["Element Caption"].apply(clean_text)
    df["Brand"] = df["Brand"].apply(clean_text)

    df = df.loc[
        df["Element Caption"].str.casefold().eq("brand").fillna(False)
    ].copy()
    df = df.loc[
        ~df["Brand"].str.casefold().eq("total brand").fillna(False)
    ].copy()

    if df.empty:
        raise ValueError(
            "No records remain after Element Caption = Brand "
            "and Total Brand filters."
        )

    subcategory_columns = [
        col for col in df.columns
        if re.fullmatch(r"SubCategory\d+", col, flags=re.IGNORECASE)
    ]
    subcategory_columns = sorted(
        subcategory_columns,
        key=lambda name: int(re.search(r"\d+", name).group()),
    )

    missing_brand_rows = int(df["Brand"].isna().sum())
    df["Brand"] = df["Brand"].fillna("(Missing Brand)")

    for col in subcategory_columns:
        df[col] = df[col].apply(clean_text).fillna("(Missing Category)")

    df["Sales"] = pd.to_numeric(df["Sales"], errors="coerce")
    df["Sales_Is_Missing"] = df["Sales"].isna()
    df["Sales_Is_Negative"] = df["Sales"].lt(0)

    invalid_date_rows = 0
    if {"Year", "MonthNo"}.issubset(df.columns):
        year = pd.to_numeric(df["Year"], errors="coerce")
        month_no = pd.to_numeric(df["MonthNo"], errors="coerce")
        valid = (
            year.notna() & month_no.notna()
            & month_no.between(1, 12)
            & year.mod(1).eq(0)
            & month_no.mod(1).eq(0)
        )
        invalid_date_rows = int((~valid).sum())
        df = df.loc[valid].copy()
        df["Year"] = year.loc[valid].astype(int)
        df["MonthNo"] = month_no.loc[valid].astype(int)
        df["Month"] = pd.PeriodIndex(
            pd.to_datetime(
                dict(
                    year=df["Year"],
                    month=df["MonthNo"],
                    day=1,
                )
            ),
            freq="M",
        )
    elif "DateText" in df.columns:
        parsed = pd.to_datetime(df["DateText"], errors="coerce")
        invalid_date_rows = int(parsed.isna().sum())
        df["Month"] = parsed.dt.to_period("M")
        df = df.loc[df["Month"].notna()].copy()
    else:
        raise ValueError(
            "Input must contain Year and MonthNo, or DateText."
        )

    if df.empty:
        raise ValueError("No valid monthly observations remain after date validation.")

    df["Year"] = df["Month"].dt.year
    df["MonthNo"] = df["Month"].dt.month
    all_months = pd.period_range(
        start=df["Month"].min(),
        end=df["Month"].max(),
        freq="M",
    )

    print(f"Filtered source rows: {len(df):,}")
    print(f"Brands: {df['Brand'].nunique():,}")
    print(f"SubCategory dimensions: {len(subcategory_columns)}")
    print(f"Period: {all_months.min()} to {all_months.max()}")

    # ========================================================
    # 5. BRAND-MONTH PANEL AND MARKET TRENDS
    # ========================================================

    print("Building brand-month panel...")
    brand_panel = build_entity_month_panel(
        df, ["Brand"], all_months
    )
    brand_panel = brand_panel.rename(
        columns={
            "Observed_Sales": "Observed_Brand_Sales",
            "Sales": "Brand_Sales",
        }
    )

    # Market totals are based on brands whose observed record
    # windows include each month. No-record months inside a window
    # count as missing coverage; months outside the window do not.
    active_brand_rows = brand_panel.loc[brand_panel["Active_Window"]].copy()
    market_monthly = (
        active_brand_rows.groupby("Month", as_index=False)
        .agg(
            Observed_Market_Sales=("Observed_Brand_Sales", sum_min_count),
            Brands_Expected=("Brand", "size"),
            Brands_With_Record=("Has_Record", "sum"),
            Brands_Complete=("CompleteCoverage", "sum"),
            Brands_With_Partial_Missing=(
                "Record_Status",
                lambda x: x.eq("Partial_Missing_Sales").sum(),
            ),
            Brands_All_Sales_Missing=(
                "Record_Status",
                lambda x: x.eq("All_Sales_Missing").sum(),
            ),
            Brands_With_No_Record=(
                "Record_Status",
                lambda x: x.eq("No_Record_Within_Observed_Window").sum(),
            ),
            Source_Rows=("Source_Rows", "sum"),
            Missing_Sales_Rows=("Missing_Sales_Rows", "sum"),
            Negative_Sales_Rows=("Negative_Sales_Rows", "sum"),
        )
    )
    market_monthly["Record_Coverage_%"] = safe_divide(
        market_monthly["Brands_With_Record"],
        market_monthly["Brands_Expected"],
    ) * 100
    market_monthly["Complete_Coverage_%"] = safe_divide(
        market_monthly["Brands_Complete"],
        market_monthly["Brands_Expected"],
    ) * 100
    market_monthly["CompleteCoverage"] = (
        market_monthly["Brands_Expected"].gt(0)
        & market_monthly["Brands_Complete"].eq(
            market_monthly["Brands_Expected"]
        )
    )
    market_monthly["Market_Sales"] = market_monthly[
        "Observed_Market_Sales"
    ].where(market_monthly["CompleteCoverage"])
    market_monthly["Coverage_Status"] = np.where(
        market_monthly["CompleteCoverage"],
        "Complete",
        "Incomplete: at least one expected brand is incomplete",
    )

    # Add any missing calendar months explicitly.
    market_monthly = (
        pd.DataFrame({"Month": all_months})
        .merge(market_monthly, on="Month", how="left")
        .sort_values("Month")
        .reset_index(drop=True)
    )
    count_cols = [
        "Brands_Expected", "Brands_With_Record", "Brands_Complete",
        "Brands_With_Partial_Missing", "Brands_All_Sales_Missing",
        "Brands_With_No_Record", "Source_Rows", "Missing_Sales_Rows",
        "Negative_Sales_Rows",
    ]
    for col in count_cols:
        market_monthly[col] = market_monthly[col].fillna(0).astype(int)
    market_monthly["CompleteCoverage"] = (
        market_monthly["CompleteCoverage"].fillna(False).astype(bool)
    )
    market_monthly["Coverage_Status"] = market_monthly[
        "Coverage_Status"
    ].fillna("No active brand records")
    market_monthly["Observed_Market_Sales"] = pd.to_numeric(
        market_monthly["Observed_Market_Sales"], errors="coerce"
    )
    market_monthly["Market_Sales"] = pd.to_numeric(
        market_monthly["Market_Sales"], errors="coerce"
    )
    market_monthly["Record_Coverage_%"] = pd.to_numeric(
        market_monthly["Record_Coverage_%"], errors="coerce"
    )
    market_monthly["Complete_Coverage_%"] = pd.to_numeric(
        market_monthly["Complete_Coverage_%"], errors="coerce"
    )

    market_base_year, market_base_reason = choose_base_year(
        market_monthly, "Market_Sales"
    )
    market_monthly = add_calendar_lags(
        market_monthly, "Market_Sales"
    )
    market_monthly = add_ytd_yoy(
        market_monthly, "Market_Sales"
    )
    market_monthly = add_monthly_trend_index(
        market_monthly, "Market_Sales", market_base_year
    )
    market_annual = calculate_annual_trends(
        market_monthly, "Market_Sales", market_base_year
    )
    market_seasonality = calculate_seasonality(
        market_monthly, "Market_Sales"
    )

    # ========================================================
    # 6. BRAND MONTHLY AND ANNUAL TRENDS
    # ========================================================

    print("Calculating brand-level trends...")
    brand_monthly_parts = []
    brand_annual_parts = []
    brand_seasonality_parts = []

    market_lookup = market_monthly.set_index("Month")
    for brand, temp in brand_panel.groupby("Brand", dropna=False):
        temp = temp.copy()
        temp = temp.rename(
            columns={
                "Observed_Brand_Sales": "Observed_Brand_Sales",
                "Brand_Sales": "Brand_Sales",
            }
        )

        base_year, base_reason = choose_base_year(temp, "Brand_Sales")
        temp = add_calendar_lags(temp, "Brand_Sales")
        temp = add_ytd_yoy(temp, "Brand_Sales")
        temp = add_monthly_trend_index(
            temp, "Brand_Sales", base_year
        )

        temp["Market_Sales"] = temp["Month"].map(
            market_lookup["Market_Sales"]
        )
        temp["Market_Trend_Index"] = temp["Month"].map(
            market_lookup["Trend_Index"]
        )
        temp["Market_MoM_Growth_%"] = temp["Month"].map(
            market_lookup["MoM_Growth_%"]
        )
        temp["Market_YoY_Growth_%"] = temp["Month"].map(
            market_lookup["YoY_Growth_%"]
        )
        temp["Brand_Share_of_Market_%"] = safe_divide(
            temp["Brand_Sales"], temp["Market_Sales"]
        ) * 100
        temp["Brand_vs_Market_Trend_Differential"] = (
            temp["Trend_Index"] - temp["Market_Trend_Index"]
        )
        temp["Brand_vs_Market_YoY_Growth_Differential_pp"] = (
            temp["YoY_Growth_%"] - temp["Market_YoY_Growth_%"]
        )
        temp["Brand_Base_Year"] = base_year
        temp["Brand_Base_Year_Reason"] = base_reason
        brand_monthly_parts.append(temp)

        annual = calculate_annual_trends(
            temp, "Brand_Sales", base_year
        )
        annual["Brand"] = brand
        brand_annual_parts.append(annual)

        seasonality = calculate_seasonality(
            temp, "Brand_Sales"
        )
        if not seasonality.empty:
            seasonality["Brand"] = brand
            brand_seasonality_parts.append(seasonality)

    brand_monthly = pd.concat(brand_monthly_parts, ignore_index=True)
    brand_annual = (
        pd.concat(brand_annual_parts, ignore_index=True)
        if brand_annual_parts else pd.DataFrame()
    )
    brand_seasonality = (
        pd.concat(brand_seasonality_parts, ignore_index=True)
        if brand_seasonality_parts else pd.DataFrame()
    )

    # ========================================================
    # 7. INDEPENDENT SUBCATEGORY DIMENSIONS
    # ========================================================

    print("Calculating independent SubCategory dimensions...")
    subcategory_monthly_outputs = {}
    subcategory_annual_outputs = {}
    subcategory_seasonality_outputs = {}
    brand_subcategory_monthly_outputs = {}
    brand_subcategory_annual_outputs = {}
    subcategory_validation_rows = []

    for dimension in subcategory_columns:
        print(f"Processing {dimension}...")

        # Each brand-category pair is treated as an independent
        # entity. This avoids assuming that all brands belong to
        # every category in every month.
        pair_panel = build_entity_month_panel(
            df,
            ["Brand", dimension],
            all_months,
        )
        pair_panel = pair_panel.rename(
            columns={
                "Observed_Sales": "Observed_Brand_Category_Sales",
                "Sales": "Brand_Category_Sales",
            }
        )

        category_monthly = build_category_month_panel(
            pair_panel.rename(
                columns={
                    "Observed_Brand_Category_Sales": "Observed_Sales",
                    "Brand_Category_Sales": "Sales",
                }
            ),
            dimension,
            all_months,
        )

        category_parts = []
        category_annual_parts = []
        for category, temp in category_monthly.groupby(dimension, dropna=False):
            temp = temp.copy()
            base_year, base_reason = choose_base_year(
                temp, "Category_Sales"
            )
            temp = add_calendar_lags(temp, "Category_Sales")
            temp = add_ytd_yoy(temp, "Category_Sales")
            temp = add_monthly_trend_index(
                temp, "Category_Sales", base_year
            )
            temp["Market_Sales"] = temp["Month"].map(
                market_lookup["Market_Sales"]
            )
            temp["Category_Share_of_Market_%"] = safe_divide(
                temp["Category_Sales"], temp["Market_Sales"]
            ) * 100
            temp["Category_Base_Year"] = base_year
            temp["Category_Base_Year_Reason"] = base_reason
            category_parts.append(temp)

            annual = calculate_annual_trends(
                temp, "Category_Sales", base_year
            )
            annual[dimension] = category
            category_annual_parts.append(annual)

        category_monthly_output = pd.concat(
            category_parts, ignore_index=True
        ) if category_parts else pd.DataFrame()
        category_annual_output = pd.concat(
            category_annual_parts, ignore_index=True
        ) if category_annual_parts else pd.DataFrame()

        category_seasonality = calculate_seasonality(
            category_monthly_output,
            "Category_Sales",
            group_cols=[dimension],
        )

        # Brand × category monthly and annual trends.
        pair_parts = []
        pair_annual_parts = []
        for (brand, category), temp in pair_panel.groupby(
            ["Brand", dimension], dropna=False
        ):
            temp = temp.copy()
            base_year, base_reason = choose_base_year(
                temp, "Brand_Category_Sales"
            )
            temp = add_calendar_lags(
                temp, "Brand_Category_Sales"
            )
            temp = add_ytd_yoy(
                temp, "Brand_Category_Sales"
            )
            temp = add_monthly_trend_index(
                temp, "Brand_Category_Sales", base_year
            )

            category_lookup = category_monthly_output.loc[
                category_monthly_output[dimension].eq(category)
            ].set_index("Month")
            temp["Group_Sales"] = temp["Month"].map(
                category_lookup["Category_Sales"]
            )
            temp["Group_CompleteCoverage"] = temp["Month"].map(
                category_lookup["CompleteCoverage"]
            ).fillna(False).astype(bool)
            temp["Brand_Share_of_Category_%"] = safe_divide(
                temp["Brand_Category_Sales"], temp["Group_Sales"]
            ) * 100
            temp["Pair_vs_Group_Trend_Differential"] = (
                temp["Trend_Index"]
                - temp["Month"].map(category_lookup["Trend_Index"])
                if "Trend_Index" in category_lookup.columns
                else np.nan
            )
            temp["Pair_Base_Year"] = base_year
            temp["Pair_Base_Year_Reason"] = base_reason
            pair_parts.append(temp)

            annual = calculate_annual_trends(
                temp, "Brand_Category_Sales", base_year
            )
            annual["Brand"] = brand
            annual[dimension] = category
            pair_annual_parts.append(annual)

        pair_monthly_output = (
            pd.concat(pair_parts, ignore_index=True)
            if pair_parts else pd.DataFrame()
        )
        pair_annual_output = (
            pd.concat(pair_annual_parts, ignore_index=True)
            if pair_annual_parts else pd.DataFrame()
        )

        subcategory_monthly_outputs[dimension] = category_monthly_output
        subcategory_annual_outputs[dimension] = category_annual_output
        subcategory_seasonality_outputs[dimension] = category_seasonality
        brand_subcategory_monthly_outputs[dimension] = pair_monthly_output
        brand_subcategory_annual_outputs[dimension] = pair_annual_output

        subcategory_validation_rows.append(
            {
                "Dimension": dimension,
                "Number_of_Categories": int(
                    category_monthly_output[dimension].nunique()
                ) if not category_monthly_output.empty else 0,
                "Brand_Category_Pairs": int(
                    pair_panel[["Brand", dimension]].drop_duplicates().shape[0]
                ),
                "Monthly_Category_Rows": len(category_monthly_output),
                "Category_Months_Incomplete": int(
                    (~category_monthly_output["CompleteCoverage"]).sum()
                ) if not category_monthly_output.empty else 0,
                "Brand_Category_Months_Incomplete": int(
                    (~pair_panel["CompleteCoverage"] & pair_panel["Active_Window"]).sum()
                ),
                "Categories_With_At_Least_One_Complete_Year": int(
                    category_annual_output.groupby(dimension)["CompleteYear"]
                    .sum().gt(0).sum()
                ) if not category_annual_output.empty else 0,
            }
        )

    # ========================================================
    # 8. BRAND LIFECYCLE AND TREND CLASSIFICATION
    # ========================================================

    print("Classifying brand lifecycle and annual trends...")

    # Use the same latest two consecutive complete market years for every brand.
    complete_market_years = sorted(
        market_annual.loc[
            market_annual["CompleteYear"].fillna(False)
            & market_annual["Annual_Sales"].notna(),
            "Year",
        ].astype(int).unique().tolist()
    )
    comparison_years = None
    for i in range(len(complete_market_years) - 1, 0, -1):
        y_prev, y_latest = complete_market_years[i - 1], complete_market_years[i]
        if y_latest == y_prev + 1:
            comparison_years = (y_prev, y_latest)
            break

    comparison_period_label = (
        f"{comparison_years[0]}–{comparison_years[1]}"
        if comparison_years is not None else "not available"
    )

    market_annual_lookup = market_annual.set_index("Year")
    annual_lookup_by_brand = {
        brand: group.set_index("Year")
        for brand, group in brand_annual.groupby("Brand")
    }
    monthly_by_brand = {
        brand: group.sort_values("Month").copy()
        for brand, group in brand_monthly.groupby("Brand")
    }

    classification_rows = []
    for brand in sorted(set(annual_lookup_by_brand) | set(monthly_by_brand)):
        monthly = monthly_by_brand.get(brand, pd.DataFrame()).copy()

        # Lifecycle classification uses observed monthly sales only.
        if not monthly.empty:
            if isinstance(monthly["Month"].dtype, pd.PeriodDtype):
                monthly["_MonthPeriod"] = monthly["Month"].dt.asfreq("M")
            else:
                monthly["_MonthPeriod"] = pd.to_datetime(
                    monthly["Month"], errors="coerce"
                ).dt.to_period("M")
            monthly = monthly.sort_values("_MonthPeriod").reset_index(drop=True)

            sales = pd.to_numeric(monthly["Brand_Sales"], errors="coerce")
            coverage = monthly["CompleteCoverage"].fillna(False).astype(bool)
            positive_positions = np.flatnonzero((sales.gt(0) & coverage).to_numpy())
            incomplete_months = int((~coverage).sum())
            data_coverage_status = (
                "Complete within observed panel" if incomplete_months == 0
                else f"Incomplete observations: {incomplete_months} month(s)"
            )

            if len(positive_positions) == 0:
                first_positive_month = last_positive_month = np.nan
                months_before_first_positive = months_after_last_positive = np.nan
                entry_status = "No positive sales observed"
                exit_status = "Not applicable"
                sales_pattern = "No positive sales observed"
                current_status = "No positive sales observed"
            else:
                first_pos, last_pos = int(positive_positions[0]), int(positive_positions[-1])
                first_period = monthly.loc[first_pos, "_MonthPeriod"]
                last_period = monthly.loc[last_pos, "_MonthPeriod"]
                first_positive_month, last_positive_month = str(first_period), str(last_period)

                # Entry: 12 immediately preceding months must be contiguous,
                # complete and explicitly zero. Missing is never zero.
                before = monthly.iloc[:first_pos].tail(ENTRY_CONFIRMATION_MONTHS)
                before_periods = before["_MonthPeriod"].tolist()
                before_contiguous = (
                    len(before_periods) == ENTRY_CONFIRMATION_MONTHS
                    and all(before_periods[j].ordinal - before_periods[j - 1].ordinal == 1
                            for j in range(1, len(before_periods)))
                )
                entry_confirmed = (
                    len(before) == ENTRY_CONFIRMATION_MONTHS
                    and before_contiguous
                    and before["CompleteCoverage"].fillna(False).astype(bool).all()
                    and pd.to_numeric(before["Brand_Sales"], errors="coerce").eq(0).all()
                )
                months_before_first_positive = int(first_pos)
                if entry_confirmed:
                    entry_status = (
                        f"Confirmed entrant: {ENTRY_CONFIRMATION_MONTHS} complete "
                        "zero-sales months before first positive sale"
                    )
                elif first_pos < ENTRY_CONFIRMATION_MONTHS:
                    entry_status = (
                        "Pre-existing / left-censored: insufficient pre-entry "
                        "history to confirm market entry"
                    )
                else:
                    entry_status = (
                        "Entry not confirmed: preceding observations are "
                        "incomplete or do not meet the zero-sales rule"
                    )

                # Exit: require a complete, contiguous zero-sales tail of at
                # least 12 months after the last positive month.
                after = monthly.iloc[last_pos + 1:].copy()
                trailing_zero_run = 0
                for _, row in after.iterrows():
                    row_sales = pd.to_numeric(
                        pd.Series([row["Brand_Sales"]]), errors="coerce"
                    ).iloc[0]
                    row_complete = bool(row["CompleteCoverage"]) if pd.notna(
                        row["CompleteCoverage"]
                    ) else False
                    if row_complete and pd.notna(row_sales) and row_sales == 0:
                        trailing_zero_run += 1
                    else:
                        break

                months_after_last_positive = int(trailing_zero_run)
                tail_is_all_zero_complete = (
                    len(after) == trailing_zero_run
                    and all(after["_MonthPeriod"].iloc[j].ordinal
                            - after["_MonthPeriod"].iloc[j - 1].ordinal == 1
                            for j in range(1, len(after)))
                )
                exit_confirmed = (
                    trailing_zero_run >= EXIT_CONFIRMATION_MONTHS
                    and tail_is_all_zero_complete
                )

                if exit_confirmed:
                    exit_status = (
                        f"Confirmed operational exit: at least "
                        f"{EXIT_CONFIRMATION_MONTHS} consecutive complete "
                        "zero-sales months after the last positive sale"
                    )
                    current_status = "Exit confirmed"
                elif trailing_zero_run > 0:
                    exit_status = (
                        f"Not confirmed: only {trailing_zero_run} consecutive "
                        "complete zero-sales month(s) after last positive sale"
                    )
                    current_status = (
                        f"Recently inactive ({trailing_zero_run} trailing "
                        "zero-sales month(s)); exit not confirmed"
                    )
                elif last_pos == len(monthly) - 1 and bool(
                    monthly.loc[last_pos, "CompleteCoverage"]
                ):
                    exit_status = "No exit signal"
                    current_status = "Active in latest observed month"
                else:
                    exit_status = (
                        "Cannot determine: latest post-sale observations "
                        "are incomplete or unavailable"
                    )
                    current_status = "Current status uncertain"

                between = monthly.iloc[first_pos:last_pos + 1].copy()
                between_sales = pd.to_numeric(between["Brand_Sales"], errors="coerce")
                observed_zero_gap = (
                    between["CompleteCoverage"].fillna(False).astype(bool)
                    & between_sales.eq(0)
                ).any()
                missing_inside = (
                    ~between["CompleteCoverage"].fillna(False).astype(bool)
                ).any()
                if observed_zero_gap:
                    sales_pattern = (
                        "Intermittent: zero-sales month(s) between positive-sales months"
                    )
                elif missing_inside:
                    sales_pattern = (
                        "Pattern uncertain: incomplete month(s) between positive-sales months"
                    )
                else:
                    sales_pattern = (
                        "Continuous: no observed zero-sales month between positive-sales months"
                    )
        else:
            first_positive_month = last_positive_month = np.nan
            months_before_first_positive = months_after_last_positive = np.nan
            entry_status = "No monthly observations"
            exit_status = "Cannot determine"
            sales_pattern = "Insufficient monthly observations"
            current_status = "Insufficient monthly observations"
            data_coverage_status = "No monthly observations"

        # Trend: a common comparison window for all brands.
        comparison_year = latest_comparable_year = np.nan
        previous_year_sales = latest_year_sales = np.nan
        brand_growth = market_growth = share_change = np.nan

        if comparison_years is None:
            trend_label = "Insufficient data: no consecutive complete market years"
        else:
            previous_year, latest_year = comparison_years
            comparison_year, latest_comparable_year = previous_year, latest_year
            brand_annual = annual_lookup_by_brand.get(brand, pd.DataFrame())

            valid_brand_years = (
                not brand_annual.empty
                and previous_year in brand_annual.index
                and latest_year in brand_annual.index
                and bool(brand_annual.loc[previous_year, "CompleteYear"])
                and bool(brand_annual.loc[latest_year, "CompleteYear"])
                and pd.notna(brand_annual.loc[previous_year, "Annual_Sales"])
                and pd.notna(brand_annual.loc[latest_year, "Annual_Sales"])
            )
            valid_market_years = (
                previous_year in market_annual_lookup.index
                and latest_year in market_annual_lookup.index
                and bool(market_annual_lookup.loc[previous_year, "CompleteYear"])
                and bool(market_annual_lookup.loc[latest_year, "CompleteYear"])
                and pd.notna(market_annual_lookup.loc[previous_year, "Annual_Sales"])
                and pd.notna(market_annual_lookup.loc[latest_year, "Annual_Sales"])
            )

            if valid_brand_years and valid_market_years:
                previous_year_sales = float(brand_annual.loc[previous_year, "Annual_Sales"])
                latest_year_sales = float(brand_annual.loc[latest_year, "Annual_Sales"])
                market_prev_sales = float(market_annual_lookup.loc[previous_year, "Annual_Sales"])
                market_latest_sales = float(market_annual_lookup.loc[latest_year, "Annual_Sales"])
                market_growth = (
                    (market_latest_sales / market_prev_sales - 1) * 100
                    if market_prev_sales > 0 else np.nan
                )

                if previous_year_sales > 0 and latest_year_sales > 0:
                    brand_growth = (latest_year_sales / previous_year_sales - 1) * 100
                    previous_share = (
                        previous_year_sales / market_prev_sales * 100
                        if market_prev_sales > 0 else np.nan
                    )
                    latest_share = (
                        latest_year_sales / market_latest_sales * 100
                        if market_latest_sales > 0 else np.nan
                    )
                    share_change = latest_share - previous_share

                    if brand_growth > STABLE_GROWTH_THRESHOLD_PCT:
                        trend_label = "Growing"
                    elif brand_growth < -STABLE_GROWTH_THRESHOLD_PCT:
                        trend_label = "Declining"
                    else:
                        trend_label = "Stable (within threshold)"
                else:
                    trend_label = (
                        "Not comparable: zero sales in one or both comparison years"
                    )
            else:
                trend_label = (
                    "Insufficient data: incomplete or missing annual observations"
                )

        classification_rows.append({
            "Brand": brand,
            "Current_Status": current_status,
            "Entry_Status": entry_status,
            "Exit_Status": exit_status,
            "Sales_Pattern": sales_pattern,
            "First_Positive_Month": first_positive_month,
            "Last_Positive_Month": last_positive_month,
            "Months_Before_First_Positive": months_before_first_positive,
            "Trailing_Zero_Months_After_Last_Positive": months_after_last_positive,
            "Data_Coverage_Status": data_coverage_status,
            "Comparison_Year": comparison_year,
            "Latest_Comparable_Year": latest_comparable_year,
            "Brand_Sales_Comparison_Year": previous_year_sales,
            "Brand_Sales_Latest_Year": latest_year_sales,
            "Brand_Annual_Growth_%": brand_growth,
            "Market_Annual_Growth_%": market_growth,
            "Share_Change_pp": share_change,
            "Brand_Trajectory": trend_label,
            "Classification_Basis": (
                f"Entry/exit require {ENTRY_CONFIRMATION_MONTHS}/"
                f"{EXIT_CONFIRMATION_MONTHS} consecutive complete zero-sales months; "
                "missing values are never treated as zero. Trend compares all brands "
                f"over the same latest two consecutive complete market years "
                f"({comparison_period_label}); "
                f"Growing > {STABLE_GROWTH_THRESHOLD_PCT:.1f}%, Declining < "
                f"-{STABLE_GROWTH_THRESHOLD_PCT:.1f}%, otherwise Stable. "
                "Descriptive, not causal."
            ),
        })

    brand_classification = pd.DataFrame(classification_rows)

    # ========================================================
    # 9. VALIDATION AND METHODOLOGY REPORT
    # ========================================================

    print("Preparing validation report...")
    validation_rows = [
        {
            "Check": "Input file",
            "Result": INPUT_FILE.name,
            "Details": f"Worksheet: {INPUT_SHEET}",
        },
        {
            "Check": "Input rows before filters",
            "Result": input_rows,
            "Details": "Rows read from the source worksheet.",
        },
        {
            "Check": "Filtered Brand records",
            "Result": len(df),
            "Details": "Element Caption = Brand; Total Brand excluded.",
        },
        {
            "Check": "Missing Brand identity",
            "Result": missing_brand_rows,
            "Details": (
                "Retained as '(Missing Brand)' so source records are not silently dropped."
            ),
        },
        {
            "Check": "Brand universe",
            "Result": int(df["Brand"].nunique()),
            "Details": (
                "Distinct brand labels after filtering. Brand-month lifecycle "
                "windows are based on first and last observed source-record months."
            ),
        },
        {
            "Check": "Market coverage rule",
            "Result": "Observed-window strict coverage",
            "Details": (
                "For each month, every brand whose first-to-last observed record "
                "window includes that month must have a complete brand-month value. "
                "No-record months inside the window are gaps; months outside it "
                "are not automatically treated as zero or missing."
            ),
        },
        {
            "Check": "Lifecycle proxy limitation",
            "Result": "First-to-last observed source record",
            "Details": (
                "This is an operational data-coverage proxy, not verified market "
                "entry/exit timing. Confirm lifecycle dates using independent evidence."
            ),
        },
        {
            "Check": "Market base year",
            "Result": market_base_year,
            "Details": market_base_reason,
        },
        {
            "Check": "Preferred base year",
            "Result": PREFERRED_BASE_YEAR,
            "Details": (
                "2016 is selected when it is a complete year; otherwise the first "
                "complete year is selected. Monthly and annual market indices share "
                "the same base year."
            ),
        },
        {
            "Check": "SubCategory dimensions",
            "Result": len(subcategory_columns),
            "Details": ", ".join(subcategory_columns),
        },
        {
            "Check": "Rows with missing sales",
            "Result": int(df["Sales"].isna().sum()),
            "Details": "Missing sales were not converted to zero.",
        },
        {
            "Check": "Rows with negative sales",
            "Result": int(df["Sales_Is_Negative"].sum()),
            "Details": "Negative values were retained and reported.",
        },
        {
            "Check": "Invalid date rows",
            "Result": invalid_date_rows,
            "Details": (
                "Rows with invalid or unusable calendar information were excluded "
                "from time-series calculations and counted here."
            ),
        },
        {
            "Check": "Partial missing aggregation",
            "Result": "Observed nonmissing values retained",
            "Details": (
                "For duplicate entity-month rows, nonmissing sales are summed with "
                "min_count=1. If any source sales value is missing, coverage is "
                "incomplete even when an observed partial sum exists."
            ),
        },
        {
            "Check": "Zero sales treatment",
            "Result": "Retained as observed zero",
            "Details": "Zero is distinct from missing and is a valid observed value.",
        },
        {
            "Check": "Annual completeness",
            "Result": "12 complete calendar months",
            "Details": "Annual growth requires consecutive complete years.",
        },
        {
            "Check": "YTD YoY",
            "Result": "Implemented",
            "Details": (
                "January through the current month is compared with the same "
                "calendar months in the previous year; every month in both windows "
                "must be complete."
            ),
        },
        {
            "Check": "Seasonality method",
            "Result": "Within-year monthly normalization",
            "Details": (
                "Each monthly value is divided by its year's average monthly sales; "
                "normalized indices are averaged across complete years."
            ),
        },
        {
            "Check": "Seasonality evidence threshold",
            "Result": MIN_YEARS_FOR_SEASONALITY,
            "Details": (
                "Fewer than three complete years are flagged as limited evidence; "
                "the estimate is not silently suppressed."
            ),
        },
        {
            "Check": "Brand classification",
             "Result": "Lifecycle + common-year trend classification",
            "Details": (
                "Entry/exit require 12 consecutive complete zero-sales months; "
                "intermittent sales are flagged separately. Trend compares all "
                "brands over the same latest two consecutive complete market years; "
                "Growing > 5%, Declining < -5%, Stable within +/-5%. Descriptive, "
                "not causal."
            ),
        },
        {
            "Check": "Missing-data imputation",
            "Result": "None",
            "Details": (
                "No missing values are converted to zero, interpolated, or carried "
                "forward. Missingness and observed partial sums are reported separately."
            ),
        },
    ]

    validation_rows.extend(
        [
            {
                "Check": "Market months with incomplete coverage",
                "Result": int((~market_monthly["CompleteCoverage"]).sum()),
                "Details": "At least one expected brand-month is incomplete.",
            },
            {
                "Check": "Market months with partial missing source rows",
                "Result": int(
                    market_monthly["Brands_With_Partial_Missing"].gt(0).sum()
                ),
                "Details": "At least one expected brand has partial missing source sales.",
            },
            {
                "Check": "Market months with no-record gaps",
                "Result": int(
                    market_monthly["Brands_With_No_Record"].gt(0).sum()
                ),
                "Details": (
                    "At least one brand has no source record inside its observed "
                    "first-to-last record window."
                ),
            },
        ]
    )

    validation_df = pd.DataFrame(validation_rows)
    subcategory_validation_df = pd.DataFrame(subcategory_validation_rows)

    # ========================================================
    # 10. PREPARE EXCEL OUTPUT
    # ========================================================

    print("Writing Excel workbook...")
    used_sheet_names = set()

    def write_sheet(writer, data, requested_name):
        sheet = safe_sheet_name(requested_name, used_sheet_names)
        data.to_excel(writer, sheet_name=sheet, index=False)

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        write_sheet(writer, market_monthly, "Market_Monthly_Trends")
        write_sheet(writer, market_annual, "Market_Annual_Trends")
        write_sheet(writer, market_seasonality, "Market_Seasonality")
        write_sheet(writer, brand_monthly, "Brand_Monthly_Trends")
        write_sheet(writer, brand_annual, "Brand_Annual_Trends")
        write_sheet(writer, brand_seasonality, "Brand_Seasonality")
        write_sheet(writer, brand_classification, "Brand_Classification")

        for dimension in subcategory_columns:
            write_sheet(
                writer,
                subcategory_monthly_outputs[dimension],
                f"{dimension}_Trends",
            )
            write_sheet(
                writer,
                subcategory_annual_outputs[dimension],
                f"{dimension}_Annual",
            )
            write_sheet(
                writer,
                subcategory_seasonality_outputs[dimension],
                f"{dimension}_Seasonality",
            )
            write_sheet(
                writer,
                brand_subcategory_monthly_outputs[dimension],
                f"{dimension}_Brand_Trends",
            )
            write_sheet(
                writer,
                brand_subcategory_annual_outputs[dimension],
                f"{dimension}_Brand_Annual",
            )

        write_sheet(writer, validation_df, "Validation")
        write_sheet(
            writer,
            subcategory_validation_df,
            "SubCategory_Validation",
        )

        # ====================================================
        # 11. EXCEL FORMATTING
        # ====================================================

        workbook = writer.book
        header_fill = PatternFill(
            fill_type="solid",
            fgColor="1F4E78",
        )
        header_font = Font(
            bold=True,
            color="FFFFFF",
        )

        for worksheet in workbook.worksheets:
            worksheet.freeze_panes = "A2"
            worksheet.auto_filter.ref = worksheet.dimensions

            for cell in worksheet[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(
                    horizontal="center",
                    vertical="center",
                    wrap_text=True,
                )

            for column_cells in worksheet.columns:
                col_letter = get_column_letter(column_cells[0].column)
                max_length = 0
                for cell in column_cells[:5000]:
                    try:
                        max_length = max(
                            max_length,
                            len(str(cell.value)) if cell.value is not None else 0,
                        )
                    except Exception:
                        pass
                worksheet.column_dimensions[col_letter].width = min(
                    max(max_length + 2, 12),
                    40,
                )

            worksheet.row_dimensions[1].height = 34

            for col_idx, cell in enumerate(worksheet[1], start=1):
                header = str(cell.value).lower()

                if (
                    "growth_%" in header
                    or "share_of_" in header
                    or "coverage_%" in header
                    or "share_change_pp" in header
                    or "differential_pp" in header
                ):
                    number_format = "0.00"
                elif "index" in header:
                    number_format = "0.00"
                elif "sales" in header:
                    number_format = '#,##0.00;[Red]-#,##0.00'
                else:
                    continue

                for row in range(2, worksheet.max_row + 1):
                    worksheet.cell(
                        row=row,
                        column=col_idx,
                    ).number_format = number_format

    print("=" * 60)
    print("STEP 04 COMPLETED SUCCESSFULLY")
    print(f"Output file: {OUTPUT_FILE}")
    print("=" * 60)


if __name__ == "__main__":
    main()

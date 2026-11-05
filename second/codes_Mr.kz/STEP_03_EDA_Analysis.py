
# ============================================================
# STEP 03 - EXPLORATORY DATA ANALYSIS (EDA)
# FMCG Sales Panel Data
# ============================================================

from pathlib import Path
import pandas as pd
import numpy as np
import re

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


# ============================================================
# 1. SETTINGS
# ============================================================

BASE_DIR = Path(__file__).resolve().parent

INPUT_FILE = BASE_DIR / "Energydrinks_Panel_Data.xlsx"
INPUT_SHEET = "Panel_Data"

OUTPUT_FILE = BASE_DIR / "Energydrinks_Step_03_EDA_Analysis.xlsx"

# Brand records only
BRAND_CAPTION = "Brand"

# Required input columns
REQUIRED_COLUMNS = [
    "Element Caption",
    "Brand",
    "Year",
    "Month Name",
    "MonthNo",
    "Sales"
]

# Excel output sheets
SHEET_MONTHLY = "Monthly_Brand_Sales"
SHEET_ANNUAL = "Annual_Brand_Sales"
SHEET_DYNAMICS = "Annual_Market_Dynamics"
SHEET_MATRIX = "Brand_Annual_Matrix"
SHEET_VALIDATION = "Validation"


# ============================================================
# 2. READ INPUT DATA
# ============================================================

print("\n" + "=" * 60)
print("STEP 03 - EXPLORATORY DATA ANALYSIS")
print("=" * 60)

if not INPUT_FILE.exists():
    raise FileNotFoundError(
        f"Input file not found:\n{INPUT_FILE}"
    )

print("\nReading input file...")

df = pd.read_excel(
    INPUT_FILE,
    sheet_name=INPUT_SHEET,
    engine="openpyxl"
)

df.columns = df.columns.astype(str).str.strip()

print(f"Source rows: {len(df):,}")
print(f"Source columns: {len(df.columns):,}")


# ============================================================
# 3. VALIDATE INPUT STRUCTURE
# ============================================================

missing_columns = [
    col for col in REQUIRED_COLUMNS
    if col not in df.columns
]

if missing_columns:
    raise ValueError(
        "The following required columns are missing:\n"
        + ", ".join(missing_columns)
    )

# Normalize text fields
df["Element Caption"] = (
    df["Element Caption"]
    .astype("string")
    .str.strip()
)

df["Brand"] = (
    df["Brand"]
    .astype("string")
    .str.strip()
)

df["Month Name"] = (
    df["Month Name"]
    .astype("string")
    .str.strip()
)


# ============================================================
# 4. FILTER BRAND RECORDS
# ============================================================

print("\nFiltering Brand records...")

caption_mask = (
    df["Element Caption"]
    .str.casefold()
    .eq(BRAND_CAPTION.casefold())
    .fillna(False)
)

brand_df = df.loc[caption_mask].copy()

total_source_rows = len(df)
total_brand_rows = len(brand_df)

print(f"Total source rows: {total_source_rows:,}")
print(f"Brand records: {total_brand_rows:,}")
print(
    f"Excluded non-Brand records: "
    f"{total_source_rows - total_brand_rows:,}"
)

if brand_df.empty:
    raise ValueError(
        "No records with Element Caption = Brand were found."
    )


# ============================================================
# 5. CLEAN AND VALIDATE VARIABLES
# ============================================================

# Convert Year to numeric
year_numeric = pd.to_numeric(
    brand_df["Year"],
    errors="coerce"
)

invalid_year_mask = (
    year_numeric.isna()
    | (year_numeric % 1 != 0)
)

invalid_year_count = int(invalid_year_mask.sum())

# Convert MonthNo to numeric
month_numeric = pd.to_numeric(
    brand_df["MonthNo"],
    errors="coerce"
)

invalid_month_mask = (
    month_numeric.isna()
    | (month_numeric % 1 != 0)
    | ~month_numeric.between(1, 12)
)

invalid_month_count = int(invalid_month_mask.sum())

# Convert Sales to numeric
original_sales = brand_df["Sales"].copy()

brand_df["Sales"] = pd.to_numeric(
    brand_df["Sales"],
    errors="coerce"
)

invalid_sales_count = int(
    (
        original_sales.notna()
        & brand_df["Sales"].isna()
    ).sum()
)

missing_sales_count = int(
    brand_df["Sales"].isna().sum()
)

negative_sales_count = int(
    (brand_df["Sales"] < 0).sum()
)

missing_brand_count = int(
    (
        brand_df["Brand"].isna()
        | brand_df["Brand"].eq("")
    ).sum()
)

# Keep records with valid identification and time variables
valid_mask = (
    brand_df["Brand"].notna()
    & brand_df["Brand"].ne("")
    & ~invalid_year_mask
    & ~invalid_month_mask
)

clean_df = brand_df.loc[valid_mask].copy()

clean_df["Year"] = year_numeric.loc[valid_mask].astype(int)
clean_df["MonthNo"] = month_numeric.loc[valid_mask].astype(int)

if clean_df.empty:
    raise ValueError(
        "No valid Brand records remain after data validation."
    )

# Check that valid month names exist
clean_df["Month Name"] = (
    clean_df["Month Name"]
    .replace("", pd.NA)
)

# Use the first non-missing month name for each calendar month
month_lookup = (
    clean_df.dropna(subset=["Month Name"])
    .groupby("MonthNo")["Month Name"]
    .first()
    .to_dict()
)

# Fallback names for missing month labels
fallback_months = {
    1: "January",
    2: "February",
    3: "March",
    4: "April",
    5: "May",
    6: "June",
    7: "July",
    8: "August",
    9: "September",
    10: "October",
    11: "November",
    12: "December"
}

clean_df["Month Name"] = clean_df["MonthNo"].map(
    lambda x: month_lookup.get(
        x,
        fallback_months[x]
    )
)

print(f"Valid Brand records: {len(clean_df):,}")
print(f"Unique brands: {clean_df['Brand'].nunique():,}")


# ============================================================
# 6. MONTHLY BRAND SALES
# ============================================================

print("\nCalculating monthly brand sales...")

# SubCategory columns are intentionally ignored.
# All records belonging to the same Brand and month
# are aggregated together.

monthly_df = (
    clean_df
    .groupby(
        ["Brand", "Year", "MonthNo"],
        as_index=False,
        dropna=False
    )
    .agg(
        Sale=(
            "Sales",
            lambda x: x.sum(min_count=1)
        ),
        Month=(
            "Month Name",
            "first"
        )
    )
)

monthly_df = monthly_df[
    [
        "Brand",
        "Year",
        "MonthNo",
        "Month",
        "Sale"
    ]
]

monthly_df = monthly_df.sort_values(
    ["Year", "MonthNo", "Brand"]
).reset_index(drop=True)

# Add row numbers
monthly_df.insert(
    0,
    "No.",
    range(1, len(monthly_df) + 1)
)

print(f"Monthly observations: {len(monthly_df):,}")


# ============================================================
# 7. ANNUAL BRAND SALES
# ============================================================

print("\nCalculating annual brand sales...")

annual_df = (
    monthly_df
    .groupby(
        ["Brand", "Year"],
        as_index=False
    )
    .agg(
        Sale=(
            "Sale",
            lambda x: x.sum(min_count=1)
        )
    )
)

# Calculate total market size for each year.
# Market size is calculated from aggregated Brand sales,
# never from Total Brand records.

market_df = (
    annual_df
    .groupby("Year", as_index=False)
    .agg(
        Total_Market_Size=(
            "Sale",
            lambda x: x.sum(min_count=1)
        )
    )
)

annual_df = annual_df.merge(
    market_df,
    on="Year",
    how="left",
    validate="many_to_one"
)

# Market share is expressed in percentage points (0-100).
annual_df["Market Share"] = np.where(
    annual_df["Total_Market_Size"].notna()
    & (annual_df["Total_Market_Size"] != 0)
    & annual_df["Sale"].notna(),

    (
        annual_df["Sale"]
        / annual_df["Total_Market_Size"]
    ) * 100,

    np.nan
)

annual_df = annual_df[
    [
        "Brand",
        "Year",
        "Sale",
        "Total_Market_Size",
        "Market Share"
    ]
]

annual_df = annual_df.sort_values(
    ["Year", "Brand"]
).reset_index(drop=True)

annual_df.insert(
    0,
    "No.",
    range(1, len(annual_df) + 1)
)

print(f"Annual observations: {len(annual_df):,}")


# ============================================================
# 8. ANNUAL MARKET DYNAMICS
# ============================================================

print("\nCalculating annual market dynamics...")

# Calendar years between the first and last observed years.
# This ensures that years with no observations are also represented.

first_year = int(clean_df["Year"].min())
last_year = int(clean_df["Year"].max())

all_years = list(
    range(first_year, last_year + 1)
)

# Positive annual sales indicate an active brand.
positive_sales = annual_df[
    annual_df["Sale"] > 0
].copy()

# First and last positive sales year for each brand
brand_lifecycle = (
    positive_sales
    .groupby("Brand")["Year"]
    .agg(
        First_Active_Year="min",
        Last_Active_Year="max"
    )
)

# Annual sales matrix for lifecycle calculations
lifecycle_matrix = (
    positive_sales
    .pivot(
        index="Brand",
        columns="Year",
        values="Sale"
    )
    .reindex(columns=all_years)
    .fillna(0)
)

# Brand activity matrix
activity_matrix = lifecycle_matrix > 0

# Market size by year
market_sales = (
    market_df
    .set_index("Year")["Total_Market_Size"]
    .reindex(all_years)
)

dynamics_records = []

for year in all_years:

    # Active brands in the current year
    active_brands = int(
        activity_matrix[year].sum()
    )

    # New brands: first positive sales year
    new_brands = int(
        (
            brand_lifecycle["First_Active_Year"]
            == year
        ).sum()
    )

    # Discontinued brands: last positive sales year
    discontinued_brands = int(
        (
            brand_lifecycle["Last_Active_Year"]
            == year
        ).sum()
    )

    # Dormant brands:
    # Positive sales before the current year,
    # no positive sales in the current year,
    # and positive sales after the current year.

    dormant_brands = 0

    for brand in activity_matrix.index:

        previous_active = (
            activity_matrix.loc[
                brand,
                activity_matrix.columns < year
            ].any()
        )

        current_active = (
            activity_matrix.loc[
                brand,
                year
            ]
        )

        future_active = (
            activity_matrix.loc[
                brand,
                activity_matrix.columns > year
            ].any()
        )

        if (
            previous_active
            and not current_active
            and future_active
        ):
            dormant_brands += 1

    dynamics_records.append(
        {
            "Year": year,
            "Active_Brands": active_brands,
            "New_Brands": new_brands,
            "Discontinued_Brands": discontinued_brands,
            "Dormant_Brands": dormant_brands,
            "Total_Market_Sales": market_sales.loc[year]
        }
    )

dynamics_df = pd.DataFrame(
    dynamics_records
)

print("Annual market dynamics calculated.")


# ============================================================
# 9. BRAND ANNUAL MATRIX
# ============================================================

print("\nCreating annual brand matrix...")

# Include all brands appearing in valid Brand records,
# even if they never achieved positive sales.

all_brands = sorted(
    clean_df["Brand"].dropna().unique().tolist(),
    key=lambda x: str(x).casefold()
)

# Annual sales matrix
sales_matrix = (
    annual_df
    .pivot(
        index="Brand",
        columns="Year",
        values="Sale"
    )
    .reindex(
        index=all_brands,
        columns=all_years
    )
)

# Fill only brand-year combinations with no recorded row.
# Existing all-missing annual observations remain missing.
observed_brand_year = (
    annual_df
    .assign(Observed=True)
    .pivot(
        index="Brand",
        columns="Year",
        values="Observed"
    )
    .reindex(
        index=all_brands,
        columns=all_years
    )
)

sales_matrix = sales_matrix.mask(
    observed_brand_year.isna(),
    0
)

# Annual market share matrix
share_matrix = (
    annual_df
    .pivot(
        index="Brand",
        columns="Year",
        values="Market Share"
    )
    .reindex(
        index=all_brands,
        columns=all_years
    )
)

share_matrix = share_matrix.mask(
    observed_brand_year.isna(),
    0
)

# Construct final output
matrix_df = pd.DataFrame(
    {"Brand": all_brands}
)

for year in all_years:

    matrix_df[f"Sales_{year}"] = (
        sales_matrix[year].values
    )

for year in all_years:

    matrix_df[f"MarketShare_{year}"] = (
        share_matrix[year].values
    )

print(f"Brands in annual matrix: {len(matrix_df):,}")


# ============================================================
# 10. VALIDATION REPORT
# ============================================================

print("\nCreating validation report...")

# Count brands with at least one positive sale
ever_active_brands = int(
    brand_lifecycle.shape[0]
)

never_active_brands = (
    len(all_brands) - ever_active_brands
)

# Check annual market share totals
share_validation = (
    annual_df
    .groupby("Year")["Market Share"]
    .sum(min_count=1)
)

validation_records = [
    {
        "Check": "Input file exists",
        "Result": "PASS",
        "Value": str(INPUT_FILE.name)
    },
    {
        "Check": "Input sheet exists and was read",
        "Result": "PASS",
        "Value": INPUT_SHEET
    },
    {
        "Check": "Total source rows",
        "Result": "INFO",
        "Value": total_source_rows
    },
    {
        "Check": "Brand source rows",
        "Result": "INFO",
        "Value": total_brand_rows
    },
    {
        "Check": "Excluded non-Brand rows",
        "Result": "INFO",
        "Value": total_source_rows - total_brand_rows
    },
    {
        "Check": "Valid Brand records",
        "Result": "INFO",
        "Value": len(clean_df)
    },
    {
        "Check": "Missing Brand records",
        "Result": "WARNING" if missing_brand_count else "PASS",
        "Value": missing_brand_count
    },
    {
        "Check": "Invalid Year records",
        "Result": "WARNING" if invalid_year_count else "PASS",
        "Value": invalid_year_count
    },
    {
        "Check": "Invalid MonthNo records",
        "Result": "WARNING" if invalid_month_count else "PASS",
        "Value": invalid_month_count
    },
    {
        "Check": "Missing Sales records",
        "Result": "INFO",
        "Value": missing_sales_count
    },
    {
        "Check": "Invalid numeric Sales values",
        "Result": "WARNING" if invalid_sales_count else "PASS",
        "Value": invalid_sales_count
    },
    {
        "Check": "Negative Sales records",
        "Result": "WARNING" if negative_sales_count else "PASS",
        "Value": negative_sales_count
    },
    {
        "Check": "Unique brands",
        "Result": "INFO",
        "Value": len(all_brands)
    },
    {
        "Check": "Brands with positive sales",
        "Result": "INFO",
        "Value": ever_active_brands
    },
    {
        "Check": "Brands without positive sales",
        "Result": "INFO",
        "Value": never_active_brands
    },
    {
        "Check": "First year",
        "Result": "INFO",
        "Value": first_year
    },
    {
        "Check": "Last year",
        "Result": "INFO",
        "Value": last_year
    },
    {
        "Check": "Monthly output rows",
        "Result": "INFO",
        "Value": len(monthly_df)
    },
    {
        "Check": "Annual output rows",
        "Result": "INFO",
        "Value": len(annual_df)
    },
    {
        "Check": "Dynamics years",
        "Result": (
            "PASS"
            if len(dynamics_df) == len(all_years)
            else "WARNING"
        ),
        "Value": len(dynamics_df)
    },
    {
        "Check": "Annual matrix brands",
        "Result": (
            "PASS"
            if len(matrix_df) == len(all_brands)
            else "FAIL"
        ),
        "Value": len(matrix_df)
    },
    {
        "Check": "Monthly aggregation row check",
        "Result": (
            "PASS"
            if len(monthly_df)
            <= len(clean_df)
            else "FAIL"
        ),
        "Value": len(monthly_df)
    }
]

# Add annual market share validation
for year, share in share_validation.items():

    if pd.notna(share):

        validation_records.append(
            {
                "Check": f"Market share total - {year}",
                "Result": (
                    "PASS"
                    if np.isclose(share, 100, atol=0.01)
                    else "WARNING"
                ),
                "Value": round(float(share), 4)
            }
        )

validation_df = pd.DataFrame(
    validation_records
)


# ============================================================
# 11. PREPARE EXCEL OUTPUT
# ============================================================

print("\nPreparing Excel output...")

output_monthly = monthly_df[
    [
        "No.",
        "Brand",
        "Year",
        "Month",
        "Sale"
    ]
].copy()

output_annual = annual_df[
    [
        "No.",
        "Brand",
        "Year",
        "Sale",
        "Total_Market_Size",
        "Market Share"
    ]
].copy()

# Rename column for readability
output_annual = output_annual.rename(
    columns={
        "Total_Market_Size": "Total Market Size"
    }
)

# ============================================================
# 12. EXPORT TO EXCEL
# ============================================================

print("\nExporting Excel workbook...")

with pd.ExcelWriter(
    OUTPUT_FILE,
    engine="openpyxl"
) as writer:

    output_monthly.to_excel(
        writer,
        sheet_name=SHEET_MONTHLY,
        index=False
    )

    output_annual.to_excel(
        writer,
        sheet_name=SHEET_ANNUAL,
        index=False
    )

    dynamics_df.to_excel(
        writer,
        sheet_name=SHEET_DYNAMICS,
        index=False
    )

    matrix_df.to_excel(
        writer,
        sheet_name=SHEET_MATRIX,
        index=False
    )

    validation_df.to_excel(
        writer,
        sheet_name=SHEET_VALIDATION,
        index=False
    )


# ============================================================
# 13. FORMAT EXCEL WORKBOOK
# ============================================================

print("Formatting Excel workbook...")

wb = load_workbook(OUTPUT_FILE)

header_fill = PatternFill(
    fill_type="solid",
    fgColor="1F4E78"
)

header_font = Font(
    bold=True,
    color="FFFFFF",
    size=10
)

for ws in wb.worksheets:

    # Freeze the header row
    ws.freeze_panes = "A2"

    # Enable filters
    ws.auto_filter.ref = ws.dimensions

    # Format headers
    for cell in ws[1]:

        cell.fill = header_fill
        cell.font = header_font

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

    ws.row_dimensions[1].height = 25

    # Align cells
    for row in ws.iter_rows(
        min_row=2
    ):
        for cell in row:

            cell.alignment = Alignment(
                vertical="center"
            )

    # Dynamic column widths
    for column_cells in ws.columns:

        column_letter = get_column_letter(
            column_cells[0].column
        )

        max_length = 0

        for cell in column_cells[:1001]:

            if cell.value is not None:

                max_length = max(
                    max_length,
                    len(str(cell.value))
                )

        ws.column_dimensions[
            column_letter
        ].width = min(
            max(max_length + 3, 12),
            35
        )

# Format sales and market share columns
for sheet_name in [
    SHEET_MONTHLY,
    SHEET_ANNUAL,
    SHEET_MATRIX
]:

    ws = wb[sheet_name]

    for row in ws.iter_rows(min_row=2):

        for cell in row:

            header = ws.cell(
                row=1,
                column=cell.column
            ).value

            if header in [
                "Sale",
                "Total Market Size"
            ] or (
                isinstance(header, str)
                and header.startswith("Sales_")
            ):

                cell.number_format = '#,##0.00'

            elif header == "Market Share" or (
                isinstance(header, str)
                and header.startswith("MarketShare_")
            ):

                # Values are stored as percentage points (0-100).
                cell.number_format = '0.00"%"'

# Format dynamics sales
ws = wb[SHEET_DYNAMICS]

for row in ws.iter_rows(min_row=2):

    ws.cell(
        row=row[0].row,
        column=6
    ).number_format = '#,##0.00'

# Format validation
ws = wb[SHEET_VALIDATION]

ws.column_dimensions["A"].width = 42
ws.column_dimensions["B"].width = 15
ws.column_dimensions["C"].width = 25

for row in ws.iter_rows(min_row=2):

    status_cell = row[1]

    if status_cell.value == "FAIL":

        status_cell.fill = PatternFill(
            fill_type="solid",
            fgColor="FFC7CE"
        )

    elif status_cell.value == "WARNING":

        status_cell.fill = PatternFill(
            fill_type="solid",
            fgColor="FFF2CC"
        )

    elif status_cell.value == "PASS":

        status_cell.fill = PatternFill(
            fill_type="solid",
            fgColor="E2F0D9"
        )

# Save formatting
wb.save(OUTPUT_FILE)


# ============================================================
# 14. FINAL REPORT
# ============================================================

print("\n" + "=" * 60)
print("STEP 03 COMPLETED SUCCESSFULLY")
print("=" * 60)

print(f"\nOutput file:\n{OUTPUT_FILE}")

print("\nOutput sheets:")
print(f"1. {SHEET_MONTHLY}")
print(f"2. {SHEET_ANNUAL}")
print(f"3. {SHEET_DYNAMICS}")
print(f"4. {SHEET_MATRIX}")
print(f"5. {SHEET_VALIDATION}")

print("\nSummary:")
print(f"Total brands: {len(all_brands):,}")
print(f"Brands with positive sales: {ever_active_brands:,}")
print(f"Monthly observations: {len(monthly_df):,}")
print(f"Annual observations: {len(annual_df):,}")
print(f"Analysis period: {first_year} - {last_year}")

print("\n" + "=" * 60)
print("END OF STEP 03")
print("=" * 60)

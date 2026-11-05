
import re
import pandas as pd
from pathlib import Path


# ============================================================
# 1. USER SETTINGS
# ============================================================

BASE_DIR = Path(__file__).resolve().parent

INPUT_FILE = BASE_DIR / "Energydrinks.xlsb"
OUTPUT_FILE = BASE_DIR / "Energydrinks_Step_02.xlsx"

SHEET_NAME = "Step 00"

# Fixed identifier columns
ID_COLUMNS = [
    "Element Caption",
    "Element Description",
    "Brand",
    "Index Name"
]

# Dynamic classification column prefix
CLASSIFICATION_PREFIX = "SubCategory"


# ============================================================
# 2. CHECK INPUT FILE
# ============================================================

if not INPUT_FILE.exists():
    raise FileNotFoundError(
        f"Input file not found: {INPUT_FILE}"
    )


# ============================================================
# 3. READ SOURCE DATA
# ============================================================

print("\nReading source data...")

df = pd.read_excel(
    INPUT_FILE,
    sheet_name=SHEET_NAME,
    engine="pyxlsb"
)

# Clean column names
df.columns = df.columns.astype(str).str.strip()

source_row_count = len(df)

print(f"Source rows: {source_row_count:,}")
print(f"Source columns: {len(df.columns):,}")


# ============================================================
# 4. CHECK REQUIRED IDENTIFIER COLUMNS
# ============================================================

missing_id_columns = [
    col for col in ID_COLUMNS
    if col not in df.columns
]

if missing_id_columns:
    raise ValueError(
        f"Missing required columns: {missing_id_columns}"
    )


# ============================================================
# 5. IDENTIFY CLASSIFICATION COLUMNS DYNAMICALLY
# ============================================================

classification_columns = [
    col for col in df.columns
    if re.fullmatch(
        rf"{re.escape(CLASSIFICATION_PREFIX)}\d+",
        col,
        flags=re.IGNORECASE
    )
]

# Sort classification columns by their numeric suffix
classification_columns = sorted(
    classification_columns,
    key=lambda col: int(
        re.search(r"(\d+)$", col).group(1)
    )
)

print("\nCLASSIFICATION STRUCTURE")
print("--------------------------------")

if classification_columns:
    for col in classification_columns:
        print(f"Detected: {col}")
else:
    print("No classification columns detected.")

print(
    f"Total classification columns: "
    f"{len(classification_columns)}"
)


# ============================================================
# 6. IDENTIFY MONTHLY DATE COLUMNS
# ============================================================

def parse_month_header(value):
    """
    Parse a monthly column header into pandas Period.
    Return None if the header is not a monthly date.
    """

    if pd.isna(value):
        return None

    if isinstance(value, pd.Timestamp):
        return value.to_period("M")

    text = str(value).strip()

    # Exclude Year and MAT columns
    if re.match(
        r"^(Year|MAT)\b",
        text,
        flags=re.IGNORECASE
    ):
        return None

    formats = [
        "%b'%y",
        "%b-%y",
        "%b %y",
        "%B %y",
        "%b'%Y",
        "%b-%Y",
        "%b %Y",
        "%B %Y",
        "%Y-%m",
        "%Y/%m",
        "%m/%Y",
        "%m-%Y"
    ]

    for date_format in formats:
        try:
            parsed_date = pd.to_datetime(
                text,
                format=date_format
            )
            return parsed_date.to_period("M")

        except (ValueError, TypeError):
            continue

    # Fallback for other recognizable date formats
    try:
        parsed_date = pd.to_datetime(
            text,
            errors="coerce"
        )

        if pd.notna(parsed_date):
            return parsed_date.to_period("M")

    except (ValueError, TypeError):
        pass

    return None


monthly_dates = {}

for col in df.columns:
    parsed_period = parse_month_header(col)

    if parsed_period is not None:
        monthly_dates[col] = parsed_period


if not monthly_dates:
    raise ValueError(
        "No monthly date columns were identified. "
        "Please check the source headers."
    )


# Sort monthly columns chronologically
monthly_columns = sorted(
    monthly_dates.keys(),
    key=lambda col: monthly_dates[col]
)

monthly_column_count = len(monthly_columns)

print("\nMONTHLY TIME STRUCTURE")
print("--------------------------------")
print(f"Monthly columns: {monthly_column_count:,}")
print(f"First month: {monthly_dates[monthly_columns[0]]}")
print(f"Last month: {monthly_dates[monthly_columns[-1]]}")


# ============================================================
# 7. VALIDATE MONTHLY PERIODS
# ============================================================

period_values = [
    monthly_dates[col]
    for col in monthly_columns
]

# Check duplicate monthly periods
if len(period_values) != len(set(period_values)):
    raise ValueError(
        "Duplicate monthly periods detected."
    )

# Check missing months in the time sequence
expected_periods = pd.period_range(
    start=min(period_values),
    end=max(period_values),
    freq="M"
)

missing_periods = expected_periods.difference(
    pd.PeriodIndex(period_values, freq="M")
)

if len(missing_periods) > 0:
    print(
        "WARNING: Missing months detected:",
        ", ".join(str(p) for p in missing_periods)
    )
else:
    print("Monthly time sequence is continuous.")


# ============================================================
# 8. CALCULATE EXPECTED PANEL ROW COUNT
# ============================================================

expected_panel_rows = (
    source_row_count * monthly_column_count
)

print("\nEXPECTED PANEL SIZE")
print("--------------------------------")
print(f"Source rows: {source_row_count:,}")
print(f"Monthly columns: {monthly_column_count:,}")
print(f"Expected panel rows: {expected_panel_rows:,}")


# ============================================================
# 9. ADD TEMPORARY SOURCE ROW ID
# ============================================================

df = df.copy()

df["_Source_Row_ID"] = range(
    1,
    len(df) + 1
)


# ============================================================
# 10. WIDE TO LONG TRANSFORMATION
# ============================================================

print("\nTransforming data from Wide to Long format...")

# Keep fixed identifiers and all dynamic classifications
all_id_columns = (
    ID_COLUMNS
    + classification_columns
    + ["_Source_Row_ID"]
)

panel_df = df.melt(
    id_vars=all_id_columns,
    value_vars=monthly_columns,
    var_name="DateText",
    value_name="Sales"
)


# ============================================================
# 11. CREATE TIME VARIABLES
# ============================================================

panel_df["_Period"] = pd.PeriodIndex(
    panel_df["DateText"].map(monthly_dates),
    freq="M"
)

if panel_df["_Period"].isna().any():
    raise ValueError(
        "Some monthly dates could not be parsed."
    )

panel_df["Year"] = panel_df["_Period"].dt.year

panel_df["MonthNo"] = panel_df["_Period"].dt.month

# Fixed English month names
MONTH_NAMES = {
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

panel_df["Month Name"] = panel_df["MonthNo"].map(
    MONTH_NAMES
)


# ============================================================
# 12. VALIDATE TIME VARIABLES
# ============================================================

time_columns = [
    "Year",
    "MonthNo",
    "Month Name"
]

for col in time_columns:
    if panel_df[col].isna().any():
        raise ValueError(
            f"Missing values detected in {col}."
        )


# ============================================================
# 13. VALIDATE PANEL ROW COUNT
# ============================================================

actual_panel_rows = len(panel_df)

print("\nPANEL ROW COUNT VALIDATION")
print("--------------------------------")
print(f"Expected rows: {expected_panel_rows:,}")
print(f"Actual rows:   {actual_panel_rows:,}")

if actual_panel_rows == expected_panel_rows:
    print("VALIDATION PASSED: Row counts match.")
else:
    raise ValueError(
        f"Row count mismatch. "
        f"Expected {expected_panel_rows:,}, "
        f"but generated {actual_panel_rows:,}."
    )


# ============================================================
# 14. VALIDATE OBSERVATIONS PER SOURCE ROW
# ============================================================

observations_per_row = (
    panel_df.groupby("_Source_Row_ID")
    .size()
)

invalid_rows = observations_per_row[
    observations_per_row != monthly_column_count
]

if len(invalid_rows) == 0:
    print(
        "VALIDATION PASSED: Every source row has "
        "all monthly observations."
    )
else:
    raise ValueError(
        f"{len(invalid_rows)} source rows have "
        "an incorrect number of observations."
    )


# ============================================================
# 15. VALIDATE MONTHLY OBSERVATIONS
# ============================================================

actual_unique_months = panel_df["_Period"].nunique()

if actual_unique_months != monthly_column_count:
    raise ValueError(
        "Unique monthly count does not match source."
    )


# ============================================================
# 16. SORT PANEL DATA
# ============================================================

panel_df = panel_df.sort_values(
    by=ID_COLUMNS + classification_columns + [
        "Year",
        "MonthNo"
    ],
    ascending=True
).reset_index(drop=True)


# ============================================================
# 17. CREATE FINAL PANEL COLUMN ORDER
# ============================================================

FINAL_COLUMNS = (
    ID_COLUMNS
    + classification_columns
    + [
        "DateText",
        "Year",
        "Month Name",
        "MonthNo",
        "Sales"
    ]
)

panel_output = panel_df[FINAL_COLUMNS].copy()


# ============================================================
# 18. CREATE SUMMARY REPORT
# ============================================================

summary_data = {
    "Metric": [
        "Source Rows",
        "Source Columns",
        "Classification Columns",
        "Monthly Columns",
        "Expected Panel Rows",
        "Actual Panel Rows",
        "Total Output Columns",
        "Unique Months",
        "First Month",
        "Last Month",
        "Unique Brands",
        "Unique Panel Units",
        "Missing Sales",
        "Zero Sales",
        "Non-Missing Sales",
        "Duplicate Source-Date Combinations"
    ],

    "Value": [
        source_row_count,
        len(df.columns) - 1,
        len(classification_columns),
        monthly_column_count,
        expected_panel_rows,
        actual_panel_rows,
        len(FINAL_COLUMNS),
        actual_unique_months,
        str(min(period_values)),
        str(max(period_values)),
        panel_output["Brand"].nunique(),
        panel_output[
            ID_COLUMNS + classification_columns
        ].drop_duplicates().shape[0],
        panel_output["Sales"].isna().sum(),
        (panel_output["Sales"] == 0).sum(),
        panel_output["Sales"].notna().sum(),
        panel_df.duplicated(
            subset=["_Source_Row_ID", "DateText"]
        ).sum()
    ]
}

summary_df = pd.DataFrame(summary_data)


# ============================================================
# 19. CREATE MONTHLY SUMMARY
# ============================================================

monthly_summary = (
    panel_output.groupby(
        ["Year", "MonthNo", "Month Name"],
        dropna=False
    )
    .agg(
        Total_Observations=("Sales", "size"),
        Non_Missing_Sales=("Sales", "count"),
        Missing_Sales=("Sales", lambda x: x.isna().sum()),
        Zero_Sales=("Sales", lambda x: (x == 0).sum())
    )
    .reset_index()
    .sort_values(["Year", "MonthNo"])
)


# ============================================================
# 20. CREATE VALIDATION REPORT
# ============================================================

validation_df = pd.DataFrame({
    "Validation": [
        "Expected and actual row counts match",
        "Every source row has all monthly observations",
        "Monthly periods are unique",
        "All monthly dates were parsed",
        "All Year values are available",
        "All MonthNo values are available",
        "All Month Name values are available"
    ],

    "Status": [
        actual_panel_rows == expected_panel_rows,
        len(invalid_rows) == 0,
        len(period_values) == len(set(period_values)),
        not panel_df["_Period"].isna().any(),
        not panel_df["Year"].isna().any(),
        not panel_df["MonthNo"].isna().any(),
        not panel_df["Month Name"].isna().any()
    ]
})


# ============================================================
# 21. EXPORT RESULTS TO EXCEL
# ============================================================

print("\nExporting results to Excel...")

with pd.ExcelWriter(
    OUTPUT_FILE,
    engine="openpyxl"
) as writer:

    panel_output.to_excel(
        writer,
        sheet_name="Panel_Data",
        index=False
    )

    summary_df.to_excel(
        writer,
        sheet_name="Summary",
        index=False
    )

    monthly_summary.to_excel(
        writer,
        sheet_name="Monthly_Summary",
        index=False
    )

    validation_df.to_excel(
        writer,
        sheet_name="Validation",
        index=False
    )


# ============================================================
# 22. COMPLETION MESSAGE
# ============================================================

print("\n========================================")
print("STEP 02 COMPLETED SUCCESSFULLY")
print("========================================")

print(f"Output file: {OUTPUT_FILE}")
print(f"Total panel rows: {actual_panel_rows:,}")
print(f"Total output columns: {len(FINAL_COLUMNS)}")
print(f"Classification columns: {len(classification_columns)}")
print(f"Unique months: {actual_unique_months:,}")

print("\nAll critical validation checks passed.")

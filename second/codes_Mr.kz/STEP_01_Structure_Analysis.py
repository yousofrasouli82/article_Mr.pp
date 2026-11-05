
import pandas as pd
from pathlib import Path
import re

# =====================================================
# 1. PROJECT SETTINGS
# =====================================================

BASE_DIR = Path(__file__).resolve().parent

INPUT_FILE = BASE_DIR / "Energydrinks.xlsb"
OUTPUT_FILE = BASE_DIR / "Energydrinks_Step_01.xlsx"

SHEET_NAME = "Step 00"

BRAND_COLUMN = "Brand"
RECORD_TYPE_COLUMN = "Element Caption"

# Automatically detect all classification columns
CLASSIFICATION_PREFIX = "SubCategory"


# =====================================================
# 2. INPUT VALIDATION
# =====================================================

if not INPUT_FILE.exists():
    raise FileNotFoundError(
        f"Input file not found:\n{INPUT_FILE}"
    )


# =====================================================
# 3. READ EXCEL DATA
# =====================================================

df = pd.read_excel(
    INPUT_FILE,
    sheet_name=SHEET_NAME,
    engine="pyxlsb"
)

df.columns = df.columns.astype(str).str.strip()

required_columns = [
    BRAND_COLUMN,
    RECORD_TYPE_COLUMN
]

missing_columns = [
    col for col in required_columns
    if col not in df.columns
]

if missing_columns:
    raise ValueError(
        f"Missing required columns: {missing_columns}"
    )


# =====================================================
# 4. AUTOMATIC CLASSIFICATION DETECTION
# =====================================================

classification_columns = [
    col for col in df.columns
    if re.fullmatch(
        rf"{re.escape(CLASSIFICATION_PREFIX)}\d+",
        col,
        flags=re.IGNORECASE
    )
]

def classification_sort_key(column):
    match = re.search(r"(\d+)$", column)
    return int(match.group(1)) if match else 0


classification_columns = sorted(
    classification_columns,
    key=classification_sort_key
)

if not classification_columns:
    raise ValueError(
        "No SubCategory columns detected in the input file."
    )


# =====================================================
# 5. CLEAN DATA
# =====================================================

clean_columns = (
    classification_columns
    + [BRAND_COLUMN, RECORD_TYPE_COLUMN]
)

for col in clean_columns:

    df[col] = (
        df[col]
        .astype("string")
        .str.strip()
        .replace("", pd.NA)
    )


# =====================================================
# 6. IDENTIFY BRAND RECORDS
# =====================================================

record_type = (
    df[RECORD_TYPE_COLUMN]
    .str.strip()
    .str.casefold()
)

brand_detail = df[
    record_type == "brand"
].copy()

brand_total = df[
    record_type == "total brand"
].copy()

if brand_detail.empty:
    raise ValueError(
        "No Brand records found."
    )

brand_detail = brand_detail.dropna(
    subset=[BRAND_COLUMN]
)


# =====================================================
# 7. CLASSIFICATION LEVEL SUMMARY
# =====================================================

classification_reports = []

for col in classification_columns:

    temp = (
        brand_detail
        .dropna(subset=[col])
        .groupby(col)
        .agg(
            Record_Count=(BRAND_COLUMN, "size"),
            Unique_Brands=(BRAND_COLUMN, "nunique")
        )
        .reset_index()
    )

    temp.insert(
        0,
        "Classification_Level",
        col
    )

    temp = temp.rename(
        columns={col: "Category"}
    )

    classification_reports.append(temp)


classification_report = pd.concat(
    classification_reports,
    ignore_index=True
)


# =====================================================
# 8. ALL CLASSIFICATION COMBINATIONS
# =====================================================

category_combinations = (
    brand_detail
    .groupby(
        classification_columns,
        dropna=False
    )
    .agg(
        Record_Count=(BRAND_COLUMN, "size"),
        Unique_Brands=(BRAND_COLUMN, "nunique")
    )
    .reset_index()
)


# =====================================================
# 9. BRAND CLASSIFICATION SUMMARY
# =====================================================

brand_summary = (
    brand_detail
    .groupby(BRAND_COLUMN)
    .size()
    .reset_index(name="Record_Count")
)

for col in classification_columns:

    counts = (
        brand_detail
        .groupby(BRAND_COLUMN)[col]
        .nunique()
        .reset_index(name=f"{col}_Count")
    )

    membership = (
        brand_detail
        .groupby(BRAND_COLUMN)[col]
        .apply(
            lambda x: " | ".join(
                sorted(x.dropna().unique())
            )
        )
        .reset_index(name=f"{col}_List")
    )

    brand_summary = brand_summary.merge(
        counts,
        on=BRAND_COLUMN,
        how="left"
    )

    brand_summary = brand_summary.merge(
        membership,
        on=BRAND_COLUMN,
        how="left"
    )


brand_summary["Multiple_Classification"] = (
    brand_summary["Record_Count"] > 1
)

multiple_classification_brands = brand_summary[
    brand_summary["Multiple_Classification"]
].copy()


# =====================================================
# 10. OVERLAP MATRICES
# =====================================================

overlap_matrices = {}

for col in classification_columns:

    valid_data = brand_detail.dropna(
        subset=[col]
    )

    membership_matrix = pd.crosstab(
        valid_data[BRAND_COLUMN],
        valid_data[col]
    )

    membership_matrix = (
        membership_matrix > 0
    ).astype(int)

    overlap_matrix = (
        membership_matrix.T
        @ membership_matrix
    )

    overlap_matrices[col] = overlap_matrix


# =====================================================
# 11. OVERLAP DETAILS
# =====================================================

overlap_records = []

for col in classification_columns:

    valid_data = brand_detail.dropna(
        subset=[col]
    )

    for brand, group in valid_data.groupby(BRAND_COLUMN):

        categories = sorted(
            group[col].dropna().unique()
        )

        if len(categories) > 1:

            for i in range(len(categories)):

                for j in range(i + 1, len(categories)):

                    overlap_records.append({

                        "Classification_Level": col,
                        "Brand": brand,
                        "Category_1": categories[i],
                        "Category_2": categories[j]

                    })


overlap_details = pd.DataFrame(
    overlap_records,
    columns=[
        "Classification_Level",
        "Brand",
        "Category_1",
        "Category_2"
    ]
)


# =====================================================
# 12. GENERAL SUMMARY
# =====================================================

summary_records = [

    ("Total Rows", len(df)),

    ("Total Columns", len(df.columns)),

    ("Unique Brands", brand_detail[BRAND_COLUMN].nunique()),

    ("Total Brand Records", len(brand_total)),

    ("Brand Detail Records", len(brand_detail)),

    (
        "Number of Classification Levels",
        len(classification_columns)
    ),

    (
        "Brands with Multiple Records",
        len(multiple_classification_brands)
    )

]

for col in classification_columns:

    summary_records.append((
        f"Unique Categories in {col}",
        brand_detail[col].nunique()
    ))


summary = pd.DataFrame(
    summary_records,
    columns=["Indicator", "Value"]
)


# =====================================================
# 13. CLASSIFICATION STRUCTURE
# =====================================================

structure_records = []

for col in classification_columns:

    categories = sorted(
        brand_detail[col].dropna().unique()
    )

    for category in categories:

        structure_records.append({

            "Classification_Level": col,
            "Category": category,

            "Unique_Brands": brand_detail.loc[
                brand_detail[col] == category,
                BRAND_COLUMN
            ].nunique(),

            "Record_Count": len(
                brand_detail[
                    brand_detail[col] == category
                ]
            )

        })


classification_structure = pd.DataFrame(
    structure_records
)


# =====================================================
# 14. EXPORT TO EXCEL
# =====================================================

with pd.ExcelWriter(
    OUTPUT_FILE,
    engine="openpyxl"
) as writer:

    summary.to_excel(
        writer,
        sheet_name="Summary",
        index=False
    )

    classification_report.to_excel(
        writer,
        sheet_name="Classification",
        index=False
    )

    classification_structure.to_excel(
        writer,
        sheet_name="Structure",
        index=False
    )

    category_combinations.to_excel(
        writer,
        sheet_name="Combinations",
        index=False
    )

    brand_summary.to_excel(
        writer,
        sheet_name="Brand_Detail",
        index=False
    )

    multiple_classification_brands.to_excel(
        writer,
        sheet_name="Multiple_Brands",
        index=False
    )

    overlap_details.to_excel(
        writer,
        sheet_name="Overlap_Details",
        index=False
    )

    for col, matrix in overlap_matrices.items():

        sheet_name = f"Overlap_{col}"

        matrix.to_excel(
            writer,
            sheet_name=sheet_name
        )


# =====================================================
# 15. FINAL REPORT
# =====================================================

print("\n" + "=" * 60)

print("STEP 01 - STRUCTURE ANALYSIS COMPLETED")

print("=" * 60)

print(
    f"Unique Brands: "
    f"{brand_detail[BRAND_COLUMN].nunique()}"
)

print(
    f"Classification Levels: "
    f"{len(classification_columns)}"
)

for col in classification_columns:

    print(
        f"{col}: "
        f"{brand_detail[col].nunique()} categories"
    )

print(
    f"Multiple Classification Brands: "
    f"{len(multiple_classification_brands)}"
)

print(
    f"\nOutput File:\n{OUTPUT_FILE.resolve()}"
)

print("=" * 60)

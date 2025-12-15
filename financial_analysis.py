from pathlib import Path
import pandas as pd

# ==========================================
# CONFIGURATION SECTION
# ==========================================
# Define your Excel file name here. It must be in the same folder as this script.
FILE_NAME = "Data_Financiera.xlsx"

# Name of the worksheet inside the Excel file
SHEET_NAME = "DataBase Combined"

# Define which year you want to analyze
ANALYSIS_YEAR = 2025

COL_DATE = "Date"
COL_AMOUNT = "Amount"
COL_MAIN_CATEGORY = "Grandparent"
COL_SUB_CATEGORY = "Parent"
COL_CLASS = "Class"
COL_SOURCE = "Source"

# Specific filters
FILTER_MAIN_CATEGORY = "Income Statement"  # Main (top-level) category to filter
FILTER_ZOOM = "2 COGS"                      # Specific category (zoom-in)


def process_data():
    # 1. Locate the file
    base_dir = Path(__file__).resolve().parent
    file_path = base_dir / FILE_NAME

    print(f"--- Starting process ---")
    print(f"Looking for file at: {file_path}")

    if not file_path.exists():
        print(f"ERROR: I couldn't find the file '{FILE_NAME}'. Please check that it's in the folder.")
        return

    # 2. Load data
    try:
        print("Loading Excel... this may take a bit if the file is large.")
        df = pd.read_excel(file_path, sheet_name=SHEET_NAME)
    except Exception as e:
        print(f"An error occurred while reading the Excel file: {e}")
        return

    df[COL_DATE] = pd.to_datetime(df[COL_DATE], errors="coerce")

    print(f"Filtering data for year {ANALYSIS_YEAR} and category '{FILTER_MAIN_CATEGORY}'...")

    filtered_df = df[
        (df[COL_MAIN_CATEGORY] == FILTER_MAIN_CATEGORY)
        & (df[COL_DATE].dt.year == ANALYSIS_YEAR)
    ].copy()

    if filtered_df.empty:
        print("Warning: No data found with those filters.")
        return

    summary = (
        filtered_df
        .groupby([COL_SUB_CATEGORY, COL_CLASS], dropna=False)[COL_AMOUNT]
        .sum()
        .reset_index()
        .sort_values([COL_SUB_CATEGORY, COL_CLASS])
    )

    summary_csv_name = f"report_{ANALYSIS_YEAR}_general_summary.csv"
    summary.to_csv(summary_csv_name, index=False)
    print(f"-> Done! Saved the general summary to: {summary_csv_name}")

    print(f"Generating detailed report for: '{FILTER_ZOOM}'...")

    zoom_df = filtered_df[filtered_df[COL_SUB_CATEGORY] == FILTER_ZOOM].copy()

    if not zoom_df.empty:
        zoom_summary = (
            zoom_df
            .groupby([COL_SOURCE, COL_CLASS], dropna=False)[COL_AMOUNT]
            .sum()
            .reset_index()
            .sort_values([COL_SOURCE, COL_CLASS])
        )

        zoom_csv_name = f"report_{ANALYSIS_YEAR}_detail_{FILTER_ZOOM.replace(' ', '_')}.csv"
        zoom_summary.to_csv(zoom_csv_name, index=False)
        print(f"-> Done! Saved the specific detail report to: {zoom_csv_name}")
    else:
        print(f"No data found for category '{FILTER_ZOOM}', so that report was not generated.")


if __name__ == "__main__":
    process_data()

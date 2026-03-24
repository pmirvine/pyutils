#!/usr/bin/env python3
"""
Excel to JSON Converter
=======================

Converts an Excel spreadsheet (.xlsx or .xls) into a JSON file.
Each row becomes a JSON object, with the column headings used as the keys.

Features:
    • --date-format to control datetime output (e.g. dd/mm/yyyy)
    • --nan-as-empty to turn NaN/NaT values into empty strings ("") instead of null
    • --db-friendly-keys to convert column headings to lowercase snake_case
      (e.g. 'Scenario - Case Type' → 'scenario_case_type')

Dependencies:
    pip install pandas openpyxl

Usage examples:
    python excel-to-json.py data.xlsx
    python excel-to-json.py data.xlsx --nan-as-empty
    python excel-to-json.py data.xlsx --date-format "%d/%m/%Y"
    python excel-to-json.py data.xlsx --db-friendly-keys
    python excel-to-json.py data.xlsx --date-format "%d/%m/%Y" --nan-as-empty --db-friendly-keys
"""

import pandas as pd
import argparse
import json
import sys
import re
from pathlib import Path


def to_db_friendly(name: str) -> str:
    """
    Convert a column name to lowercase snake_case (DB-friendly).
    Example: 'Scenario - Case Type' → 'scenario_case_type'
    """
    if not name:
        return "col"
    # Lowercase and strip whitespace
    name = str(name).strip().lower()
    # Replace any sequence of non-alphanumeric characters with a single underscore
    name = re.sub(r'[^a-z0-9]+', '_', name)
    # Collapse multiple underscores and strip leading/trailing ones
    name = re.sub(r'_+', '_', name).strip('_')
    return name or "col"


def convert_excel_to_json(
    input_path: str,
    output_path: str,
    sheet: str | int = 0,
    indent: int = 4,
    date_format: str | None = None,
    nan_as_empty: bool = False,
    db_friendly_keys: bool = False,
) -> None:
    """
    Read Excel file and write JSON, with optional date formatting, NaN handling,
    and DB-friendly column name conversion.
    """
    try:
        # Read the Excel file (pandas handles .xlsx, .xls, etc.)
        print(f"📖 Reading Excel file: {input_path} (sheet: {sheet})")
        df = pd.read_excel(input_path, sheet_name=sheet, engine="openpyxl")

        # Optional: convert column names to DB-friendly snake_case
        if db_friendly_keys:
            original_cols = df.columns.tolist()
            df.columns = [to_db_friendly(col) for col in original_cols]
            print("🔧 Converted column names to DB-friendly snake_case format")

        # Optional custom date formatting
        if date_format:
            # Find all columns that pandas recognized as datetime
            datetime_cols = df.select_dtypes(include=["datetime64"]).columns.tolist()
            if datetime_cols:
                print(f"📅 Applying custom date format '{date_format}' to {len(datetime_cols)} column(s): {datetime_cols}")
                for col in datetime_cols:
                    df[col] = df[col].dt.strftime(date_format)
            else:
                print("📅 No datetime columns found to format.")

        # Optional: replace NaN/NaT with empty strings instead of null
        if nan_as_empty:
            print("🔄 Replacing all NaN/NaT values with empty strings ('')")
            df = df.fillna("")

        # Convert to list of dictionaries (column headers become JSON keys)
        records = df.to_dict(orient="records")

        # Write to JSON file
        print(f"💾 Writing JSON to: {output_path}")
        with open(output_path, "w", encoding="utf-8") as f:
            json.dump(records, f, indent=indent, ensure_ascii=False)

        print(f"✅ Success! Converted {len(records)} rows with {len(df.columns)} columns.")

    except FileNotFoundError:
        print(f"❌ Error: Input file not found: {input_path}")
        sys.exit(1)
    except Exception as e:
        print(f"❌ Error: {e}")
        sys.exit(1)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Convert Excel spreadsheet to JSON (column headings become JSON keys)."
    )
    parser.add_argument(
        "input_file",
        help="Path to the input Excel file (.xlsx or .xls)",
    )
    parser.add_argument(
        "-o", "--output",
        help="Path to output JSON file (default: <input_file>.json)",
    )
    parser.add_argument(
        "-s", "--sheet",
        default=0,
        help="Sheet name or zero-based index to read (default: first sheet)",
    )
    parser.add_argument(
        "-i", "--indent",
        type=int,
        default=4,
        help="Indentation level for pretty-printing JSON (default: 4)",
    )
    parser.add_argument(
        "--date-format",
        "-df",
        help="Custom strftime format for datetime columns "
             "(e.g. '%%d/%%m/%%Y' for dd/mm/yyyy).",
    )
    parser.add_argument(
        "--nan-as-empty",
        action="store_true",
        help="Replace NaN/NaT values with empty strings '' instead of null in the JSON output",
    )
    parser.add_argument(
        "--db-friendly-keys",
        action="store_true",
        help="Convert column headings to lowercase snake_case "
             "(DB-friendly, e.g. 'Scenario - Case Type' → 'scenario_case_type')",
    )

    args = parser.parse_args()

    # If no output file is given, use the same name but with .json extension
    if args.output is None:
        input_path = Path(args.input_file)
        args.output = str(input_path.with_suffix(".json"))

    convert_excel_to_json(
        args.input_file,
        args.output,
        args.sheet,
        args.indent,
        args.date_format,
        args.nan_as_empty,
        args.db_friendly_keys,
    )


if __name__ == "__main__":
    main()
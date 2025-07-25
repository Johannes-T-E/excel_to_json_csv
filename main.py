#!/usr/bin/env python3
"""
Script to export all sheets of an Excel (.xlsx) file into separate CSV or JSON files.
Usage:
    python excel_exporter.py input_file.xlsx -o output_directory -f [csv|json]
"""
import pandas as pd
import os
import argparse

def main():
    parser = argparse.ArgumentParser(
        description="Export Excel sheets to separate CSV or JSON files."
    )
    parser.add_argument(
        "input_file",
        help="Path to the input .xlsx file"
    )
    parser.add_argument(
        "-o", "--output_dir",
        default=".",
        help="Directory where output files will be saved"
    )
    parser.add_argument(
        "-f", "--format",
        choices=["csv", "json"],
        default="csv",
        help="Output format for the exported sheets"
    )
    args = parser.parse_args()

    input_path = args.input_file
    output_dir = args.output_dir
    output_format = args.format

    # Validate input file
    if not os.path.isfile(input_path):
        print(f"Error: File '{input_path}' does not exist.")
        return

    # Prepare output directory
    os.makedirs(output_dir, exist_ok=True)

    # Load Excel file
    excel = pd.ExcelFile(input_path)

    # Process each sheet
    for sheet_name in excel.sheet_names:
        df = pd.read_excel(excel, sheet_name=sheet_name)
        # Create a safe filename
        safe_name = sheet_name.strip().replace(" ", "_")
        output_file = os.path.join(
            output_dir,
            f"{safe_name}.{output_format}"
        )

        # Export based on chosen format
        if output_format == "csv":
            df.to_csv(output_file, index=False)
        else:
            df.to_json(
                output_file,
                orient="records",
                indent=4
            )
            # Post-process to remove escaped slashes
            with open(output_file, "r", encoding="utf-8") as f:
                content = f.read()
            content = content.replace("\\/", "/")
            with open(output_file, "w", encoding="utf-8") as f:
                f.write(content)

        print(f"Exported sheet '{sheet_name}' to {output_file}")

if __name__ == "__main__":
    main()
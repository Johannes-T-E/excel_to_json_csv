# Excel to CSV/JSON Exporter

This project provides a simple script to export all sheets of an Excel (.xlsx) file into separate CSV or JSON files.

---

## Step 1: Clone or Download the Project

Download or clone this repository to your local machine:

```bash
git clone https://github.com/Johannes-T-E/excel_to_json_csv.git
```

Then enter the cloned directory:

```bash
cd excel_to_json_csv
```

---

## Step 2: Set Up a Python Virtual Environment (Recommended)

### Create the virtual environment

```bash
python -m venv .venv
```

### Activate the virtual environment

- **On Windows:**
  ```bash
  .venv\Scripts\activate
  ```
- **On macOS/Linux:**
  ```bash
  source .venv/bin/activate
  ```

---

## Step 3: Install Dependencies

You can install all required dependencies in one step using the provided requirements file:

```bash
pip install -r requirements.txt
```

---

## Step 4: Prepare Your Excel File

Place your `.xlsx` file in the project directory, or note its path for use in the next step.

---

## Step 5: Export Sheets to CSV or JSON

Run the script from the command line, specifying your input file, output directory, and desired format:

```bash
python main.py input_file.xlsx -o output_directory -f [csv|json]
```

- `input_file.xlsx`: Path to the Excel file you want to export
- `-o output_directory` (optional): Directory to save the exported files (default: current directory)
- `-f [csv|json]` (optional): Output format for the exported sheets (default: csv)

### Quick Examples

Export all sheets to CSV files in a dedicated results folder:
```bash
python main.py test_file.xlsx -o test_file_results/csv -f csv
```

Export all sheets to JSON files in a dedicated results folder:
```bash
python main.py test_file.xlsx -o test_file_results/json -f json
```

---

## Output
- Each sheet will be saved as a separate file named after the sheet (spaces replaced with underscores).
- Files will have the extension `.csv` or `.json` depending on the chosen format.

---

## Features
- Exports every sheet in an Excel file as a separate CSV or JSON file
- Automatically creates the output directory if it does not exist
- Handles sheet names with spaces by converting them to underscores in filenames

---

## License
MIT 
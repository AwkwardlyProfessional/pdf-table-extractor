# PDF Table Extraction Tool

## Overview
This tool extracts tables from PDF files and saves them into an Excel file, maintaining their structure. It uses `pdfplumber` to handle both bordered and borderless tables, including irregularly shaped ones.

## Features
- Extracts tables from PDFs accurately.
- Handles tables with and without borders.
- Supports merged cells and multi-line cells.
- Saves extracted tables into an Excel file.
- Each table is stored in a separate sheet in the Excel file.

## Installation
Ensure you have Python installed, then install the required dependencies:
```sh
pip install pdfplumber pandas openpyxl
```

## Usage
### 1. Place Your PDF File
Ensure your PDF file is inside the `input_pdfs/` folder (or specify the correct path).

### 2. Run the Script
Run the script using the command:
```sh
python extract_tables.py
```

### 3. Output
- Extracted tables will be saved in an Excel file (`.xlsx`) in the same directory as the PDF.
- Each table is stored on a separate sheet inside the Excel file.

## Script Explanation
1. **Extract tables from a PDF** using `pdfplumber`.
2. **Convert extracted tables into Pandas DataFrames**.
3. **Save tables in an Excel file** with multiple sheets.
4. **Handles edge cases**, including tables with merged or missing cells.

## File Structure
```
📂 project-folder/
 ├── 📂 input_pdfs/
 │    ├── test3.pdf  # Place your PDFs here
 │
 ├── extract_tables.py  # Main script to run
 ├── test3.xlsx  # Output file after extraction
```

## Example Code
```python
import pdfplumber
import pandas as pd
import os

def extract_tables_from_pdf(pdf_path):
    tables = []
    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages):
            extracted_table = page.extract_table()
            if extracted_table:
                df = pd.DataFrame(extracted_table[1:], columns=extracted_table[0])
                tables.append(df)
    return tables

def save_tables_to_excel(tables, output_excel_path):
    with pd.ExcelWriter(output_excel_path, engine="openpyxl") as writer:
        for i, df in enumerate(tables):
            df.to_excel(writer, index=False, sheet_name=f"Table_{i+1}")
    print(f"✅ Tables saved to {output_excel_path}")

# Usage
pdf_path = "input_pdfs/test3.pdf"
output_excel_path = os.path.splitext(pdf_path)[0] + ".xlsx"
tables = extract_tables_from_pdf(pdf_path)
if tables:
    save_tables_to_excel(tables, output_excel_path)
else:
    print("⚠️ No tables detected.")
```

## Notes
- If no tables are detected, the script will print `⚠️ No tables detected.`
- Works best with structured tables but may need fine-tuning for highly complex layouts.

## Future Enhancements
- Improve handling of **irregular tables**.
- Add **command-line arguments** for dynamic file selection.

## License
This project is open-source and available for modification as needed.


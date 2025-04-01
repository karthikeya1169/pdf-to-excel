# PDF to Excel Table Extraction

This Python script extracts tables from the first two pages of a PDF document and converts them into an Excel file. It uses the `pdfplumber` library to extract the tables and `openpyxl` to format and save the tables in Excel.

## Features

- Extracts all tables from the first two pages of a PDF.
- Automatically adjusts column widths for better readability.
- Merges the title cell across the columns for each table.
- Titles are extracted from the page text and placed at the top of each table in Excel.
- Saves the tables in a new Excel file with one sheet per table.

## Requirements

- Python 3.12
- `pdfplumber`
- `pandas`
- `openpyxl`

You can install the required dependencies using `pip`:

```bash
pip install pdfplumber pandas openpyxl

# PDF to Excel Converter

This project contains a Python script for automatically extracting data from PDF invoices and formatting it into Excel (`.xlsx`) tables. The tool uses a CLI (Command Line Interface) for simple data input and output.

## Prerequisites and Installation

Before running the script, you need to install the required Python libraries for handling PDF and Excel files. Open your terminal and run the following command:

```bash
pip install pdfplumber pandas openpyxl
```

## How to run the script

```bash
python pdf_to_excel.py <input_pdf> <output_excel>
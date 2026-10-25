# Excel Data Merger & Chart Generator

This project contains a Python script designed to merge multiple Excel files into a single, comprehensive master file and automatically generate an interactive chart.

## 1. Installation

Before using the script, you need to install the required Python libraries. Open your terminal (Command Prompt, PowerShell, or Terminal) and run the following command:

```bash
pip install pandas openpyxl xlsxwriter
```

## 2. How to run the script (`excel_formatting.py`)

This script looks for all Excel (`.xlsx`) files in a specified folder, merges their data into a single master Excel file, and automatically generates an interactive Bar Chart showing "Total Revenue by Service".

**How to run it:**
Open your terminal, navigate to the folder where the script is located, and use the following syntax:

```bash
python excel_formatting.py <input_folder> <output_excel>
```

*Note: You can view the help menu for the script at any time by adding `-h` to the command (e.g., `python excel_formatting.py -h`).*
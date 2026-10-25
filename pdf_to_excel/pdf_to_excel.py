import pdfplumber
import pandas as pd
import argparse
import sys

def process_invoice(input_pdf, output_excel):
    """
    Extracts data from a PDF invoice and saves it to an Excel format.
    """
    print(f"Reading file: {input_pdf}")
    
    try:
        with pdfplumber.open(input_pdf) as pdf:
            first_page = pdf.pages[0]
            
            table = first_page.extract_table()
            text = first_page.extract_text()
            
        if not table:
            print("Could not find a table in the PDF file.")
            return False

        # The first row contains the columns, the rest is data
        df = pd.DataFrame(table[1:], columns=table[0])
        
        # Searching for the invoice number in the text
        # Note: Change "INVOICE NO." to "FAKTURA BR." if you are still using the old test PDF
        for line in text.split('\n'):
            if "INVOICE NO." in line or "FAKTURA BR." in line:
                df['Invoice Number'] = line.strip()
                break
                
        # Saving to Excel
        df.to_excel(output_excel, index=False)
        print(f"Success! Data saved to {output_excel}")
        return True

    except FileNotFoundError:
        print(f"Error: The file '{input_pdf}' does not exist. Check the path.")
        sys.exit(1)
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        sys.exit(1)

def main():
    parser = argparse.ArgumentParser(description="Extract data from a PDF invoice to Excel.")
    parser.add_argument("input_pdf", help="Path to the input PDF file (e.g., input.pdf)")
    parser.add_argument("output_excel", help="Name of the output Excel file (e.g., output.xlsx)")
    
    args = parser.parse_args()
    
    # Call the main function with the provided arguments
    process_invoice(args.input_pdf, args.output_excel)

if __name__ == "__main__":
    main()
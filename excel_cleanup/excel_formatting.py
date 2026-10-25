import pandas as pd
import glob
import os
import argparse
import sys

def process_data(input_folder, output_excel):
    # Find all .xlsx files in the specified folder
    path = os.path.join(input_folder, "*.xlsx")
    excel_files = glob.glob(path)
    
    if not excel_files:
        print(f"No Excel files found in the folder: {input_folder}")
        sys.exit(1)
        
    print(f"Found {len(excel_files)} files. Merging...")
    
    df_list = []
    for file in excel_files:
        try:
            df = pd.read_excel(file)
            df_list.append(df)
        except Exception as e:
            print(f"Error reading file {file}: {e}")

    # Merge all tables into one
    merged_df = pd.concat(df_list, ignore_index=True)
    
    # Ensure 'Total' column is numeric
    if 'Total' in merged_df.columns:
        merged_df['Total'] = pd.to_numeric(merged_df['Total'], errors='coerce')
    
    # Create an Excel file using the XlsxWriter engine
    with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
        
        # 1. Write the merged data to the first sheet
        merged_df.to_excel(writer, sheet_name='All Data', index=False)
        
        # 2. Generate the chart if the required columns exist
        if 'Service Description' in merged_df.columns and 'Total' in merged_df.columns:
            print("Generating Excel chart...")
            
            # Calculate total revenue per service
            revenue_df = merged_df.groupby('Service Description')['Total'].sum().reset_index()
            
            # Write these results to a secondary sheet
            revenue_df.to_excel(writer, sheet_name='Revenue Analysis', index=False)
            
            # Access the workbook and worksheets
            workbook = writer.book
            worksheet_data = writer.sheets['All Data']
            
            # Create a horizontal Bar chart
            chart = workbook.add_chart({'type': 'bar'})
            
            # Configure data series (pulling from 'Revenue Analysis' sheet)
            max_row = len(revenue_df)
            chart.add_series({
                'name':       'Total Revenue by Service',
                'categories': ['Revenue Analysis', 1, 0, max_row, 0], # 'Service Description' column
                'values':     ['Revenue Analysis', 1, 1, max_row, 1], # 'Total' column
                'fill':       {'color': '#4F81BD'} 
            })
            
            # Chart design and formatting
            chart.set_title({'name': 'Total Revenue by Service'})
            chart.set_x_axis({'name': 'Total Amount (USD)'})
            chart.set_y_axis({'name': 'Service Type'})
            chart.set_style(11) 
            chart.set_size({'width': 700, 'height': 400}) 
            
            # Insert the chart into the first sheet, starting at cell G2
            worksheet_data.insert_chart('G2', chart)
            
    print(f"Success! Data and interactive chart saved to: {output_excel}")

def main():
    parser = argparse.ArgumentParser(description="Merge Excel files and create an Excel chart.")
    parser.add_argument("input_folder", help="Path to the folder with Excel files (e.g., . for current folder)")
    parser.add_argument("output_excel", help="Name of the main output Excel file (e.g., final_report.xlsx)")
    
    args = parser.parse_args()
    process_data(args.input_folder, args.output_excel)

if __name__ == "__main__":
    main()
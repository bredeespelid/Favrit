import os
import re
import pandas as pd
import tkinter as tk
from tkinter import filedialog

# Function to extract date from the specified row in column A
def extract_date_from_row(df):
    for row in df.iloc[:, 0]:
        match = re.match(r'Periode: (\d{2}\.\d{2}\.\d{4})', str(row))
        if match:
            return pd.to_datetime(match.group(1), dayfirst=True).strftime('%d.%m.%Y')
    return None

# Initialize an empty DataFrame to store the extracted data
extracted_data = []

# Function to process each selected file
def process_file(file_path):
    try:
        # Load the Excel file
        xl = pd.ExcelFile(file_path)
        
        # Assume we always read from the first sheet (0-indexed)
        sheet = xl.sheet_names[0]
        df = xl.parse(sheet)
        
        # Extract the date from the specified row in column A
        date = extract_date_from_row(df)
        
        # Find the row index where "Tips" is located in column 1 (A)
        tips_row_index = df[df.iloc[:, 0] == 'Tips'].index[0]
        
        # Calculate the target column index (ten columns to the right of column A)
        target_column_index = 10  # This corresponds to 10 columns to the right of column A
        fee_column_index = target_column_index - 1  # One column to the left
        
        # Initialize the variable to store the last non-whitespace value
        last_value_before_whitespace = None
        
        # Traverse down the column starting from two rows below "Tips" row
        row_index = tips_row_index + 2
        
        while row_index < len(df):
            cell_value = df.iloc[row_index, target_column_index]
            if pd.isna(cell_value) or cell_value == "":
                break
            last_value_before_whitespace = cell_value
            row_index += 1
        
        # Find the row index where "25%" is located in column 1 (A)
        row_index_25 = df[df.iloc[:, 0] == '25%'].index[0]
        value_25 = df.iloc[row_index_25, target_column_index]
        
        # Find the row index where "Betalingsformidling" is located in column 1 (A)
        row_index_betalingsformidling = df[df.iloc[:, 0] == 'Betalingsformidling'].index[0]
        
        # Initialize the sum variables
        sum_betalingsformidling = 0
        sum_fee = 0
        sum_cash_without_cashdrawer = 0
        sum_vipps = 0
        
        # Traverse down the column starting from the row below "Betalingsformidling" row
        row_index = row_index_betalingsformidling + 1
        
        while row_index < len(df):
            cell_value = df.iloc[row_index, target_column_index]
            fee_value = df.iloc[row_index, fee_column_index]
            if pd.isna(cell_value) or cell_value == "":
                break
            try:
                description = df.iloc[row_index, 0].lower()
                if re.search(r'unintegrated.*cash', description):
                    sum_cash_without_cashdrawer += float(cell_value)
                elif re.search(r'unintegrated.*vipps', description):
                    sum_vipps += float(cell_value)
                elif re.search(r'unintegrated', description) and re.search(r'cash', description):
                    sum_cash_without_cashdrawer += float(cell_value)
                else:
                    sum_betalingsformidling += float(cell_value)
                sum_fee += float(abs(fee_value))
            except ValueError:
                print(f"Ignored non-numeric value in file {file_path}, cell at row {row_index + 1}, column {target_column_index + 1}")
            row_index += 1
        
        # Find the row index where "Endring i kredittbalanse" is located in column 1 (A)
        row_index_kredittbalanse = df[df.iloc[:, 0] == 'Endring i kredittbalanse'].index[0]
        value_kredittbalanse = df.iloc[row_index_kredittbalanse, target_column_index]
        
        # Append extracted data to the list
        extracted_data.append({
            'Filename': os.path.basename(file_path),
            'Date': date,
            '30012000': sum_betalingsformidling,
            '7770': sum_fee,
            '3008': value_25,
            '5991': last_value_before_whitespace,
            'Kredittbalanse': value_kredittbalanse,
            '30011000': sum_cash_without_cashdrawer,
            '30010000': sum_vipps
        })
    
    except Exception as e:
        print(f"Error processing {file_path}: {e}")

# Function to select input files
def select_input_files():
    root = tk.Tk()
    root.withdraw()
    file_paths = filedialog.askopenfilenames(title="Select Excel Files", filetypes=[("Excel files", "*.xlsx")])
    return file_paths

# Main code execution
input_files = select_input_files()
if input_files:
    for file in input_files:
        process_file(file)
    
    # Convert the list of dictionaries to a DataFrame
    final_df = pd.DataFrame(extracted_data)
    
    # Create an empty list to store transformed data
    transformed_data = []
    
    # Iterate through the DataFrame rows
    for index, row in final_df.iterrows():
        date = row['Date']
        
        # Iterate through the accounts and add non-zero positive values to the transformed DataFrame
        for account in ['30012000', '3008', '5991', '7770', 'Kredittbalanse', '30011000', '30010000']:
            amount = row[account]
            
            # Exclude rows where amount is zero or negative, and exclude 'Kredittbalanse' if amount is NaN or 0
            if (amount > 0 and account != 'Kredittbalanse') or (account == 'Kredittbalanse' and pd.notna(amount) and amount != 0):
                additional_text = ""
                if account == 'Kredittbalanse':
                    additional_text = "Endring i kredittbalanse"
                    account = '30012000'  # Change account to '30012000' for 'Kredittbalanse'
                
                # Format amount with dot as decimal separator
                formatted_amount = round(amount, 2)
                
                # Add data to transformed_data list
                transformed_data.append({
                    "Dato": date,
                    "Konto": int(account),
                    "Beløp": formatted_amount,
                    "Tilleggstekst": additional_text
                })
    
    # Create a DataFrame from the transformed data
    transformed_df = pd.DataFrame(transformed_data)
    
    # Prompt user to select output file location
    root = tk.Tk()
    root.withdraw()
    output_file = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")], title="Save as")

    if output_file:
        # Save the transformed data to an Excel file with one sheet named "Summary"
        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            transformed_df.to_excel(writer, index=False, sheet_name='Summary')
            
            # Access the XlsxWriter workbook and worksheet objects
            workbook = writer.book
            worksheet = writer.sheets['Summary']
            
            # Add a number format to the 'Beløp' column to ensure proper formatting
            number_format = workbook.add_format({'num_format': '#,##0.00'})
            worksheet.set_column('C:C', None, number_format)
        
        print(f"Data saved to {output_file}")
else:
    print("No files selected.")

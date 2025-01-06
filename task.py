import os
import re
import sys
import pdfplumber
import pandas as pd
from datetime import datetime
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows


"""Here we get the path for the pdf files considering they are in the same folder as the exe or py file.
It will get the right path depending upon if the code runs as exe file or script file."""
def get_directory_path():
    if getattr(sys, 'frozen', False):  # If running as an exe
        return os.path.dirname(os.path.abspath(sys.executable))
    else:  # If running as a python script
        return os.path.dirname(os.path.abspath(__file__))

# This function is used to extract the required values from the pdf
def extract_value_from_pdf(file_path, target_string):
    with pdfplumber.open(file_path) as pdf:
        # Extract the text from all the pages in each pdf
        all_text = ""
        for page in pdf.pages:
            all_text += page.extract_text()

    # Look for the target value
    value = None
    lines = all_text.split('\n')
    for line in lines:
        if target_string in line:
            # Extract the value associated with the target string
            value = line.split(target_string)[-1].strip()
            break

    return value

# Function to extract and normalize the date from the PDF
def extract_date_from_pdf(file_path):
    with pdfplumber.open(file_path) as pdf:
        all_text = ""
        for page in pdf.pages:
            all_text += page.extract_text()

    # First, try to extract the 'Invoice date' or 'Invoice period'
    lines = all_text.split('\n')
    for line in lines:
        # Extract 'Invoice date:' and try to get the date
        if "Invoice date:" in line:
            date = line.split("Invoice date:")[-1].strip()
            # Normalize the date format
            return normalize_date(date)

        # Regex for 'Invoice period' with a date range
        match = re.search(r'(\d{2}.\d{2}.\d{4})\s*-\s*(\d{2}.\d{2}.\d{4})', line)
        if match:
            # Return the start date from the period range
            return normalize_date(match.group(1))

    return None

# Function to normalize the date format to 'DD-MM-YYYY'
def normalize_date(date_str):
    # Try to parse different date formats
    for fmt in ("%d.%m.%Y", "%m/%d/%Y", "%Y-%m-%d", "%b %d, %Y"):
        try:
            date_obj = datetime.strptime(date_str, fmt)
            return date_obj.strftime('%d-%m-%Y')
        except ValueError:
            continue
    return date_str  # Return the original string if no format matches

# Function to handle German decimal system and currency removal
def normalize_value(value_str):

    # If the value ends with '€', we need to swap comma and period for German decimal system
    if '€' in value_str:
        value_str = value_str.replace('.', 'temp').replace(',', '.').replace('temp', ',')

    # Remove currency symbols (USD, €, etc.) and any non-numeric characters
    value_str = re.sub(r'[^\d,.-]', '', value_str)  # Remove anything that is not a number or decimal separator
    
    return value_str

# Function to process all PDF files in the directory
def process_pdf_data(directory_path):

    # Get all files in the directory
    pdf_files = [f for f in os.listdir(directory_path) if f.endswith('.pdf')]
    data = []

    # Loop through each PDF file
    for pdf_file in pdf_files:
        pdf_file_path = os.path.join(directory_path, pdf_file)
        print(f"Processing {pdf_file}...")

        # First, try to extract 'Gross Amount incl. VAT'
        gross_amount = extract_value_from_pdf(pdf_file_path, "Gross Amount incl. VAT")
        date = extract_date_from_pdf(pdf_file_path)

        if gross_amount:
            # Normalize and clean the value if it's in German decimal system
            normalized_value = normalize_value(gross_amount)
            data.append([pdf_file, date, normalized_value])
            print(f"Gross Amount incl. VAT for {pdf_file}: {normalized_value}")
        else:
            # If 'Gross Amount incl. VAT' is not found, look for 'Total'
            total = extract_value_from_pdf(pdf_file_path, "Total")
            if total:
                # Normalize and clean the value if it's in German decimal system
                normalized_value = normalize_value(total)
                data.append([pdf_file, date, normalized_value])
                print(f"Total for {pdf_file}: {normalized_value}")
            else:
                print(f"Neither Gross Amount incl. VAT nor Total found in {pdf_file}")
        print("-" * 40)

    return pd.DataFrame(data, columns=["File Name", "Date", "Value"]) if data else None

def save_data_to_files(df, directory_path):

    # Create an Excel workbook and add the data
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "Data"

    # Convert DataFrame to rows and append to Sheet 1
    for row in dataframe_to_rows(df, index=False, header=True):
        ws1.append(row)

    # Create a pivot table on Sheet 2
    ws2 = wb.create_sheet(title="Pivot Table")
    pivot_df = df.pivot_table(values='Value', index=['Date', 'File Name'], aggfunc='sum')
    
    # Write Pivot DataFrame to Sheet 2
    for row in dataframe_to_rows(pivot_df, index=True, header=True):
        ws2.append(row)

    # Save the Excel file
    output_file_excel = os.path.join(directory_path, "invoice_data.xlsx")
    wb.save(output_file_excel)
    print(f"Excel file created: {output_file_excel}")

    # Create CSV file
    output_file_csv = os.path.join(directory_path, "invoice_data.csv")
    df.to_csv(output_file_csv, index=False, sep=";", header=True)
    print(f"CSV file created: {output_file_csv}")

def main():
    # Step 1: Get the directory path where the script or executable is located
    directory_path = get_directory_path()

    # Step 2: Process all PDF files in the directory to extract data
    df = process_pdf_data(directory_path)

    # Step 3: If data is found, save it to Excel and CSV files
    if df is not None:
        save_data_to_files(df, directory_path)
    else:
        print("No data found to process.")

if __name__ == "__main__":
    main()

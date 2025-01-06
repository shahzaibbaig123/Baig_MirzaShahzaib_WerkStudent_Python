# WerkStudent_Python

## Overview
This script is designed to extract financial and date-related information from PDF files and save the results in Excel and CSV formats. It processes files in a specified directory, normalizing values and dates to a consistent format. The script also supports creating an executable file to ensure easy execution on systems without Python installed.

---

## Assumptions
1. **Column Consistency:** All extracted values, regardless of type, are stored in the same column in the Excel file.
2. **File Differences:**
    - One type of file contains "Gross Total incl. VAT" values in euros, using the European number system (comma for decimals).
    - The other type contains "Total" values in USD, using the standard number system (period for decimals).
3. **Normalization:**
    - Values in euros are converted to the standard number system.
    - Currency symbols are stripped from all extracted values.
4. **Date Handling:**
    - Single dates are directly extracted.
    - For date ranges, only the starting date is considered.
    - All dates are normalized to the format `DD-MM-YYYY`.

---

## Core Functions

### 1. `get_directory_path`
Ensures the correct file path is taken for both Python scripts and compiled executables, enabling the script to process files located in the same directory as the executable or script.

### 2. `extract_value_from_pdf`
This function extracts the target financial values ("Gross Total incl. VAT" or "Total") from the PDF files. The process involves:
- Searching for a specified target string.
- Extracting the value corresponding to the target string.

### 3. `normalize_value`
Converts extracted financial values to a standard format by:
- Replacing commas with periods (and vice versa) for the European number system.
- Stripping currency symbols.
- Ensuring consistent numerical representation.

### 4. `extract_date_from_pdf` and `normalize_date`
- **Date Extraction:** Extracts either a single date or a range of dates from the PDF.
- **Normalization:** Converts all dates to the format `DD-MM-YYYY`. For date ranges, only the starting date is considered.

### 5. `process_pdf_data`
Processes all PDF files in the specified directory by:
- Searching for the "Gross Total incl. VAT" string and extracting its value.
- If the above string is not found, searching for "Total" instead and extracting its value.
- Extracting dates and normalizing them.

### 6. `save_data_to_files`
Saves the extracted and processed data in the following formats:
- **Excel:**
  - Contains two sheets:
    - Data Sheet: Stores the extracted values and dates.
    - Pivot Table Sheet: Provides summarized data.
- **CSV:** Contains all the data, including headers, and uses a semicolon (;) as the separator.

### 7. **Creating an Executable File**
The script can be converted into an executable file using PyInstaller. This allows it to run on systems without Python installed. Ensure PyInstaller is installed by running:

```bash
pip install pyinstaller
```

The command used to create the executable is:

```bash
pyinstaller --onefile task.py
```

---

## Installation and Usage

### Prerequisites
- Python 3.7 or later (if running the script directly).
- Required Python libraries:
  - `Pdfplumber` (for PDF processing)
  - `pandas` (for data handling)
  - `openpyxl` (for Excel file creation)

Install the required libraries using:

```bash
pip install pdfplumber pandas openpyxl
```

### Steps to Run the Script
1. Place the PDF files in the same directory as the script or executable.
2. Run the script:
   - If using Python:
     ```bash
     python task.py
     ```
   - If using the executable:
     ```bash
     ./task.exe
     ```
3. The output files (`invoice_data.xlsx` and `invoice_data.csv`) will be generated in the same directory.

---

## Output Details
- **Excel File:** Contains two sheets:
  - `Data Sheet`: Stores extracted financial values and normalized dates.
  - `Pivot Table Sheet`: Provides a summary of the data.
- **CSV File:** A flat file containing the processed data.

---

## Directory Structure
```
project_directory/
|-- task.py              # Python script
|-- task.exe             # Executable file (if created)
|-- sample_file_1.pdf    # Sample PDF file (Gross Total incl. VAT in euros)
|-- sample_file_2.pdf    # Sample PDF file (Total in USD)
|-- invoice_data.xlsx          # Extracted data in Excel format
|-- invoice_data.csv           # Extracted data in CSV format
```

---

## Limitations
- The script assumes consistent formatting in the target PDFs.
- Complex PDF layouts may require additional adjustments to the extraction logic.
- Currency conversion is not handled; the script only normalizes and extracts values.


---

## Contact
For any issues or inquiries, please contact shahzaib.beyg@gmail.com.

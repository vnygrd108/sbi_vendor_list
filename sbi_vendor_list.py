from flask import Flask, request, render_template, send_file
import pandas as pd
from datetime import datetime
import os
import re

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
PROCESSED_FOLDER = "processed"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(PROCESSED_FOLDER, exist_ok=True)

# Function to clean and trim BENEFICIARY NAME
def clean_beneficiary_name(name):
    if pd.isna(name):  # Handle NaN values
        return ""
    cleaned_name = re.sub(r'[^A-Za-z0-9 ]+', '', name)  # Remove special characters
    return cleaned_name[:30]  # Trim to 30 characters

def process_vendor_list(file_path):
    print(f"Processing file: {file_path}")  # Debugging

    # Get current date
    date_str = datetime.today().strftime('%d.%m.%Y')

    # Read Excel file using openpyxl
    vendor_list_sbi = pd.read_excel(file_path, engine="openpyxl", dtype={'Bank-A/C': str})

    # Debugging: Print column names
    print("Excel Columns Found:", vendor_list_sbi.columns.tolist())

    # Expected Column Renaming
    column_mapping = {
        'UID': 'R.NO.', 
        'Amount': 'AMT', 
        'Vendor': 'BENIFICIARY NAME', 
        'Bank-A/C': 'AC NO', 
        'IFSC': 'IFSC CODE', 
        'Branch': 'Address'
    }

    # Check if all required columns exist
    missing_columns = [col for col in column_mapping.keys() if col not in vendor_list_sbi.columns]
    if missing_columns:
        print(f"Error: Missing columns {missing_columns}")
        return None, None  # Return failure

    # Rename columns
    vendor_list_sbi = vendor_list_sbi.rename(columns=column_mapping)[list(column_mapping.values())]

    # Ensure 'IFSC CODE' is treated as a string
    vendor_list_sbi['IFSC CODE'] = vendor_list_sbi['IFSC CODE'].astype(str)

    # Insert 'BENIFICIARY TYPE' column
    vendor_list_sbi.insert(2, 'BENIFICIARY TYPE', vendor_list_sbi['IFSC CODE'].apply(lambda x: 'S' if x.startswith('SBIN') else 'O'))

    # Insert additional columns
    vendor_list_sbi.insert(3, 'BENIFICIARY ACTION', 'A')
    vendor_list_sbi.insert(6, 'Befinificially Code', '')

    # Ensure 'AC NO' is string and clean up formatting
    vendor_list_sbi['AC NO'] = vendor_list_sbi['AC NO'].astype(str).str.replace(r'\.0$', '', regex=True)

    # Insert duplicate 'Address' columns
    vendor_list_sbi.insert(9, 'Address1', vendor_list_sbi['Address'])
    vendor_list_sbi.insert(10, 'Address2', "India")

    # Clean and limit the 'BENIFICIARY NAME' column
    vendor_list_sbi['BENIFICIARY NAME'] = vendor_list_sbi['BENIFICIARY NAME'].astype(str).apply(clean_beneficiary_name)

    # Create 'Formula' column
    columns_to_concat = vendor_list_sbi.loc[:, 'BENIFICIARY TYPE':'Address2'].columns
    vendor_list_sbi['Formula'] = vendor_list_sbi[columns_to_concat].astype(str).agg('|'.join, axis=1)

    # Convert 'AMT' to numeric and calculate total
    vendor_list_sbi['AMT'] = pd.to_numeric(vendor_list_sbi['AMT'], errors='coerce')
    total_amt = vendor_list_sbi['AMT'].sum()

    # Append total row
    total_row = pd.DataFrame([['TOTAL', total_amt] + [''] * (len(vendor_list_sbi.columns) - 2)], columns=vendor_list_sbi.columns)
    vendor_list_sbi = pd.concat([vendor_list_sbi, total_row], ignore_index=True)

    # Save processed file
    output_filename = f"SCHCT_{date_str}_vendor_list_SBI_Gorhe.xlsx"
    output_path = os.path.join(PROCESSED_FOLDER, output_filename)
    vendor_list_sbi.to_excel(output_path, index=False)

    print(f"File processed and saved as: {output_path}")  # Debugging
    return output_path, output_filename

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        if "file" not in request.files:
            return "No file uploaded", 400
        file = request.files["file"]
        if file.filename == "":
            return "No selected file", 400
        
        # Save uploaded file
        file_path = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(file_path)

        print(f"File uploaded: {file_path}")  # Debugging

        # Process file
        output_path, output_filename = process_vendor_list(file_path)

        if not output_path:  # If processing failed
            return "Error processing file. Check column names.", 400

        return render_template("vendor.html", download_link=output_filename)

    return render_template("vendor.html", download_link=None)

@app.route("/download/<filename>")
def download_file(filename):
    return send_file(os.path.join(PROCESSED_FOLDER, filename), as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True)



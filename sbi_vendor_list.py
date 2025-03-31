from flask import Flask, request, render_template, send_file, redirect, url_for
import pandas as pd
from datetime import datetime
import os
import re

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
PROCESSED_FOLDER = "processed"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(PROCESSED_FOLDER, exist_ok=True)

def clean_beneficiary_name(name):
    cleaned_name = re.sub(r'[^A-Za-z0-9 ]', '', str(name))  # Remove special characters
    return cleaned_name[:30]  # Trim to 30 characters

def process_vendor_list(file_path):
    date_str = datetime.today().strftime('%d.%m.%Y')

    # Read Excel file
    vendor_list_sbi = pd.read_excel(file_path, engine="xlrd", dtype={'Bank-A/C': str})

    # Rename and select columns
    vendor_list_sbi = vendor_list_sbi.rename(columns={
        'UID': 'R.NO.', 
        'Amount': 'AMT', 
        'Vendor': 'BENIFICIARY NAME', 
        'Bank-A/C': 'AC NO', 
        'IFSC': 'IFSC CODE', 
        'Branch': 'Address'
    })[['R.NO.', 'AMT', 'BENIFICIARY NAME', 'AC NO', 'IFSC CODE', 'Address']]

    # Clean and truncate 'BENIFICIARY NAME'
    vendor_list_sbi['BENIFICIARY NAME'] = vendor_list_sbi['BENIFICIARY NAME'].apply(clean_beneficiary_name)

    # Ensure 'IFSC CODE' is treated as a string
    vendor_list_sbi['IFSC CODE'] = vendor_list_sbi['IFSC CODE'].astype(str)

    # Insert additional columns
    vendor_list_sbi.insert(2, 'BENIFICIARY TYPE', vendor_list_sbi['IFSC CODE'].apply(lambda x: 'S' if x.startswith('SBIN') else 'O'))
    vendor_list_sbi.insert(3, 'BENIFICIARY ACTION', 'A')
    vendor_list_sbi.insert(6, 'Befinificially Code', '')

    # Insert additional address columns
    vendor_list_sbi.insert(9, 'Address1', vendor_list_sbi['Address'])
    vendor_list_sbi.insert(10, 'Address2', "India")

    # Convert 'AC NO' to string and remove ".0" suffix
    vendor_list_sbi['AC NO'] = vendor_list_sbi['AC NO'].astype(str).str.replace(r'\.0$', '', regex=True)

    # Create "Formula" column
    columns_to_concat = vendor_list_sbi.loc[:, 'BENIFICIARY TYPE':'Address2'].columns
    vendor_list_sbi['Formula'] = vendor_list_sbi[columns_to_concat].astype(str).agg('|'.join, axis=1)

    # Ensure 'AMT' is numeric and compute total
    vendor_list_sbi['AMT'] = pd.to_numeric(vendor_list_sbi['AMT'], errors='coerce')
    
    # Drop NaN values from DataFrame
    vendor_list_sbi = vendor_list_sbi.dropna()

    # Compute total amount
    total_amt = vendor_list_sbi['AMT'].sum()

    # Append total row
    total_row = pd.DataFrame([['TOTAL', total_amt] + [''] * (len(vendor_list_sbi.columns) - 2)], columns=vendor_list_sbi.columns)
    vendor_list_sbi = pd.concat([vendor_list_sbi, total_row], ignore_index=True)

    # Save processed file
    output_filename = f"SCHCT_{date_str}_vendor_list_SBI_Gorhe.xlsx"
    output_path = os.path.join(PROCESSED_FOLDER, output_filename)
    vendor_list_sbi.to_excel(output_path, index=False)

    return output_filename

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

        # Process file
        output_filename = process_vendor_list(file_path)

        return redirect(url_for("download_file", filename=output_filename))

    return render_template("vendor.html")

@app.route("/download/<filename>")
def download_file(filename):
    file_path = os.path.join(PROCESSED_FOLDER, filename)
    if not os.path.exists(file_path):
        return "File not found!", 404
    return send_file(file_path, as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True)

from flask import Flask, request, render_template, send_file
import pandas as pd
from datetime import datetime
import os

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
PROCESSED_FOLDER = "processed"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(PROCESSED_FOLDER, exist_ok=True)

def process_vendor_list(file_path):
    # Get current date
    date_str = datetime.today().strftime('%d.%m.%Y')

    # Read Excel file
    vendor_list_sbi = pd.read_excel(file_path, engine="xlrd", dtype={'Bank-A/C': str})

    # Rename columns
    vendor_list_sbi = vendor_list_sbi.rename(columns={
        'UID': 'R.NO.', 
        'Amount': 'AMT', 
        'Vendor': 'BENIFICIARY NAME', 
        'Bank-A/C': 'AC NO', 
        'IFSC': 'IFSC CODE', 
        'Branch': 'Address'
    })[['R.NO.', 'AMT', 'BENIFICIARY NAME', 'AC NO', 'IFSC CODE', 'Address']]

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

        # Process file
        output_path, output_filename = process_vendor_list(file_path)

        return render_template("vendor.html", download_link=output_filename)

    return render_template("vendor.html", download_link=None)

@app.route("/download/<filename>")
def download_file(filename):
    return send_file(os.path.join(PROCESSED_FOLDER, filename), as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True)

from flask import Flask, render_template, request, send_file
import pandas as pd
from datetime import datetime
import os

app = Flask(__name__)
UPLOAD_FOLDER = "uploads"
PROCESSED_FOLDER = "processed"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(PROCESSED_FOLDER, exist_ok=True)

@app.route('/', methods=['GET', 'POST'])
def upload_form():
    if request.method == 'POST':
        return process_file()
    return render_template('vendor.html')

@app.route('/process', methods=['POST'])
def process_file():
    if 'file' not in request.files:
        return "No file uploaded"
    
    file = request.files['file']
    if file.filename == '':
        return "No selected file"
    
    file_path = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(file_path)
    
    # Processing the Excel file
    vendor_list_sbi = pd.read_excel(file_path, engine="xlrd", dtype={'Bank-A/C': str})
    vendor_list_sbi = vendor_list_sbi.rename(columns={
        'UID': 'R.NO.', 
        'Amount': 'AMT', 
        'Vendor': 'BENIFICIARY NAME', 
        'Bank-A/C': 'AC NO', 
        'IFSC': 'IFSC CODE', 
        'Branch': 'Address'
    })[['R.NO.', 'AMT', 'BENIFICIARY NAME', 'AC NO', 'IFSC CODE', 'Address']]
    
    vendor_list_sbi['IFSC CODE'] = vendor_list_sbi['IFSC CODE'].astype(str)
    vendor_list_sbi.insert(2, 'BENIFICIARY TYPE', vendor_list_sbi['IFSC CODE'].apply(lambda x: 'S' if x.startswith('SBIN') else 'O'))
    vendor_list_sbi.insert(3, 'BENIFICIARY ACTION', 'A')
    vendor_list_sbi.insert(6, 'Befinificially Code', '')
    
    index_9 = min(9, len(vendor_list_sbi.columns))
    index_10 = min(10, len(vendor_list_sbi.columns) + 1)
    vendor_list_sbi.insert(index_9, 'Address1', vendor_list_sbi['Address'])
    vendor_list_sbi.insert(index_10, 'Address2', "India")
    vendor_list_sbi['AC NO'] = vendor_list_sbi['AC NO'].astype(str).str.replace(r'\.0$', '', regex=True)
    
    columns_to_concat = vendor_list_sbi.loc[:, 'BENIFICIARY TYPE':'Address2'].columns
    vendor_list_sbi['Formula'] = vendor_list_sbi[columns_to_concat].astype(str).agg('|'.join, axis=1)
    
    output_filename = f"SCHCT_{datetime.today().strftime('%d.%m.%Y')}_vendor_list_SBI_Gorhe.xlsx"
    output_path = os.path.join(PROCESSED_FOLDER, output_filename)
    vendor_list_sbi.to_excel(output_path, index=False)
    
    return send_file(output_path, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
import pdfplumber
import pandas as pd
import os
import re
from flask import Flask, request, send_file, render_template
from pdf2image import convert_from_path
import pytesseract

app = Flask(__name__)

# Updated patterns to match Extron part numbers
PATTERNS = {
    'mfg_number': re.compile(r'(Extron\s+\d+-\d+-\d+)'),
    'quantity': re.compile(r'^\d+$'),  # Match whole numbers in Qty column
    'price': re.compile(r'\$\d+[\.,]\d+')  # Match currency format
}

@app.route("/", methods=["GET", "POST"])
def upload_file():
    if request.method == "POST":
        pdf_file = request.files["file"]
        if pdf_file:
            try:
                os.makedirs("uploads", exist_ok=True)
                file_path = os.path.join("uploads", pdf_file.filename)
                pdf_file.save(file_path)

                structured_data = []
                
                # Extract tables
                tables = extract_tables(file_path)
                
                # Process main items table
                for table in tables:
                    if len(table) > 5 and 'Qty' in table[0]:  # Identify the line items table
                        headers = [cell.strip().lower() if cell else '' for cell in table[0]]
                        for row in table[1:]:
                            if len(row) >= 7:  # Ensure we have all columns
                                item = {
                                    'Manufacturer Number': '',
                                    'Description': row[4].strip() if len(row) > 4 else '',
                                    'Quantity': row[1].strip() if len(row) > 1 else '',
                                    'Price': row[5].strip().replace('$', '') if len(row) > 5 else '',
                                    'Notes': ''
                                }
                                
                                # Extract MFG number from description
                                mfg_match = PATTERNS['mfg_number'].search(item['Description'])
                                if mfg_match:
                                    item['Manufacturer Number'] = mfg_match.group(1)
                                
                                structured_data.append(item)

                # Create DataFrame
                df = pd.DataFrame(structured_data, columns=[
                    'Manufacturer Number', 
                    'Description', 
                    'Quantity', 
                    'Price', 
                    'Notes'
                ])

                # Save to Excel
                output_path = "converted.xlsx"
                df.to_excel(output_path, index=False, engine='openpyxl')

                return send_file(output_path, as_attachment=True)

            except Exception as e:
                print(f"Error: {str(e)}")
                return f"Error processing file: {str(e)}", 500
            finally:
                if os.path.exists(file_path):
                    os.remove(file_path)

    return render_template("index.html")

def extract_tables(file_path):
    tables = []
    try:
        with pdfplumber.open(file_path) as pdf:
            for page in pdf.pages:
                page_tables = page.extract_tables()
                if page_tables:
                    tables.extend(page_tables)
        return tables
    except Exception as e:
        print(f"Table extraction error: {str(e)}")
        return []

if __name__ == "__main__":
    os.makedirs("uploads", exist_ok=True)
    app.run(host="0.0.0.0", port=5000)

import pdfplumber
import pandas as pd
import os
import re
from flask import Flask, request, send_file, render_template
from pdf2image import convert_from_path
import pytesseract

app = Flask(__name__)

# Configure these patterns according to your PDF structure
PATTERNS = {
    'mfg_number': re.compile(r'(Manufacturer Number|Mfg\s*#?):?\s*(\S+)', re.IGNORECASE),
    'quantity': re.compile(r'(Quantity|Qty):?\s*(\d+)', re.IGNORECASE),
    'price': re.compile(r'(Price|Unit Price):?\s*\$?(\d+\.\d{2})', re.IGNORECASE),
    'description': re.compile(r'(Description|Item):?\s*(.+)', re.IGNORECASE)
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

                # Extract and structure data
                structured_data = []
                text_data = extract_text(file_path)
                tables = extract_tables(file_path)
                
                # Process tables first
                structured_data += process_tables(tables)
                
                # Process text lines
                structured_data += process_text(text_data)

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
                # Clean up uploaded files
                if os.path.exists(file_path):
                    os.remove(file_path)

    return render_template("index.html")

def extract_text(file_path):
    text = []
    try:
        with pdfplumber.open(file_path) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    text.extend(page_text.split('\n'))
        
        # Fallback to OCR if no text found
        if not text or all(line.strip() == '' for line in text):
            images = convert_from_path(file_path)
            for img in images:
                ocr_text = pytesseract.image_to_string(img)
                text.extend(ocr_text.split('\n'))
    except Exception as e:
        print(f"Extraction error: {str(e)}")
    return [line.strip() for line in text if line.strip()]

def extract_tables(file_path):
    tables = []
    try:
        with pdfplumber.open(file_path) as pdf:
            for page in pdf.pages:
                page_tables = page.extract_tables()
                if page_tables:
                    tables.extend(page_tables)
    except Exception as e:
        print(f"Table extraction error: {str(e)}")
    return tables

def process_text(text_lines):
    items = []
    current_item = {
        'Manufacturer Number': '',
        'Description': '',
        'Quantity': '',
        'Price': '',
        'Notes': []
    }

    for line in text_lines:
        # Try to match all patterns
        found = False
        for key, pattern in PATTERNS.items():
            match = pattern.search(line)
            if match:
                value = match.group(2).strip()
                if key == 'mfg_number':
                    # New item found
                    if current_item['Manufacturer Number']:
                        items.append(finalize_item(current_item))
                    current_item = {
                        'Manufacturer Number': value,
                        'Description': '',
                        'Quantity': '',
                        'Price': '',
                        'Notes': []
                    }
                else:
                    current_item[key.capitalize() if key != 'description' else 'Description'] = value
                found = True
                break
        
        if not found and line:
            current_item['Notes'].append(line)

    # Add the last item
    if current_item['Manufacturer Number']:
        items.append(finalize_item(current_item))
    
    return items

def process_tables(tables):
    items = []
    for table in tables:
        if not table:
            continue
            
        # Try to find header row
        headers = [cell.strip().lower() if cell else '' for cell in table[0]]
        col_mapping = {}
        
        # Map columns to our desired fields
        for idx, header in enumerate(headers):
            if 'manufacturer' in header or 'mfg' in header:
                col_mapping['mfg'] = idx
            elif 'description' in header or 'item' in header:
                col_mapping['desc'] = idx
            elif 'quantity' in header or 'qty' in header:
                col_mapping['qty'] = idx
            elif 'price' in header:
                col_mapping['price'] = idx

        # Process data rows
        for row in table[1:]:
            item = {
                'Manufacturer Number': '',
                'Description': '',
                'Quantity': '',
                'Price': '',
                'Notes': []
            }
            
            notes = []
            for idx, cell in enumerate(row):
                cell_value = str(cell).strip() if cell else ''
                if idx == col_mapping.get('mfg'):
                    item['Manufacturer Number'] = cell_value
                elif idx == col_mapping.get('desc'):
                    item['Description'] = cell_value
                elif idx == col_mapping.get('qty'):
                    item['Quantity'] = cell_value
                elif idx == col_mapping.get('price'):
                    item['Price'] = cell_value
                else:
                    notes.append(cell_value)
            
            item['Notes'] = ' | '.join(notes)
            if item['Manufacturer Number']:
                items.append(item)
    
    return items

def finalize_item(item):
    """Clean up item before adding to final list"""
    return {
        'Manufacturer Number': item['Manufacturer Number'],
        'Description': item['Description'] or ' '.join(item['Notes']),
        'Quantity': item['Quantity'],
        'Price': item['Price'],
        'Notes': '; '.join(item['Notes']) if not item['Description'] else ''
    }

if __name__ == "__main__":
    os.makedirs("uploads", exist_ok=True)
    app.run(host="0.0.0.0", port=5000)

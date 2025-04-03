import pdfplumber
import pandas as pd
import os
import re
from flask import Flask, request, send_file, render_template
from pdf2image import convert_from_path
import pytesseract

app = Flask(__name__)

# Enhanced patterns for AV equipment detection
PATTERNS = {
    'manufacturer_number': re.compile(
        r'(?:Manufacturer\s*[\#\:]?|Mfg\s*[\#\:]?|Part\s*No)\s*([A-Z0-9\-]+)',
        re.IGNORECASE
    ),
    'item_number': re.compile(
        r'(?:Item\s*[\#\:]?|SKU)\s*([A-Z0-9\-]+)',
        re.IGNORECASE
    ),
    'quantity': re.compile(
        r'(?:Quantity|Qty)\s*[\:\s]+(\d+)',
        re.IGNORECASE
    ),
    'price': re.compile(
        r'(?:Price|Each|Unit\s*Price)\s*[\:\s]+\$?(\d+[\.,]?\d{0,2})',
        re.IGNORECASE
    ),
    'description': re.compile(
        r'(?:Description|Product)\s*[\:\s]+(.+?)(?=\s*(?:Manufacturer|Qty|Price|$))',
        re.IGNORECASE | re.DOTALL
    )
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

                # Extract and process data
                text_data = extract_text(file_path)
                tables = extract_tables(file_path)
                structured_data = process_data(text_data, tables)

                # Create DataFrame with required columns
                df = pd.DataFrame(structured_data, columns=[
                    'Manufacturer Number', 
                    'Item Number', 
                    'Description', 
                    'Quantity', 
                    'Price', 
                    'Notes'
                ])

                # Generate and return Excel file
                output_path = "converted.xlsx"
                df.to_excel(output_path, index=False, engine='openpyxl')
                return send_file(output_path, as_attachment=True)

            except Exception as e:
                return f"Error processing file: {str(e)}", 500
            finally:
                if os.path.exists(file_path):
                    os.remove(file_path)
    return render_template("index.html")

def extract_text(file_path):
    text = []
    try:
        # PDF text extraction
        with pdfplumber.open(file_path) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    text.extend(page_text.split('\n'))
        
        # OCR fallback
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

def process_data(text_lines, tables):
    items = []
    current_item = {
        'Manufacturer Number': '',
        'Item Number': '',
        'Description': '',
        'Quantity': 1,
        'Price': 0.0,
        'Notes': []
    }

    # Process text lines
    for line in text_lines:
        # Manufacturer Number
        mfg_match = PATTERNS['manufacturer_number'].search(line)
        if mfg_match:
            current_item['Manufacturer Number'] = mfg_match.group(1).strip()
            line = line.replace(mfg_match.group(0), '')

        # Item Number
        item_match = PATTERNS['item_number'].search(line)
        if item_match:
            current_item['Item Number'] = item_match.group(1).strip()
            line = line.replace(item_match.group(0), '')

        # Quantity
        qty_match = PATTERNS['quantity'].search(line)
        if qty_match:
            current_item['Quantity'] = int(qty_match.group(1))
            line = line.replace(qty_match.group(0), '')

        # Price
        price_match = PATTERNS['price'].search(line)
        if price_match:
            current_item['Price'] = float(price_match.group(1).replace(',',''))
            line = line.replace(price_match.group(0), '')

        # Description
        desc_match = PATTERNS['description'].search(line)
        if desc_match and not current_item['Description']:
            current_item['Description'] = desc_match.group(1).strip()
            line = line.replace(desc_match.group(0), '')

        # Collect remaining text as notes
        if line.strip():
            current_item['Notes'].append(line.strip())

    # Process tables
    for table in tables:
        if len(table) < 2:
            continue

        headers = [cell.strip().lower() for cell in table[0] if cell]
        col_mapping = {}
        
        # Map columns based on headers
        for idx, header in enumerate(headers):
            if 'mfg' in header or 'manufacturer' in header:
                col_mapping['mfg'] = idx
            elif 'item' in header or 'sku' in header:
                col_mapping['item'] = idx
            elif 'desc' in header:
                col_mapping['desc'] = idx
            elif 'qty' in header:
                col_mapping['qty'] = idx
            elif 'price' in header:
                col_mapping['price'] = idx

        for row in table[1:]:
            item = current_item.copy()
            for col, idx in col_mapping.items():
                if idx < len(row):
                    value = str(row[idx]).strip()
                    if col == 'mfg':
                        item['Manufacturer Number'] = value
                    elif col == 'item':
                        item['Item Number'] = value
                    elif col == 'desc':
                        item['Description'] = value
                    elif col == 'qty':
                        item['Quantity'] = int(value) if value.isdigit() else 1
                    elif col == 'price':
                        item['Price'] = float(value.replace(',','')) if value else 0.0
            items.append(item)

    # Add final item if valid
    if current_item['Manufacturer Number'] or current_item['Item Number']:
        items.append(current_item)

    # Clean notes field
    for item in items:
        item['Notes'] = ' | '.join(item['Notes'])
        if not item['Description'] and item['Notes']:
            # Use first note as description if empty
            item['Description'] = item['Notes'].split('|')[0].strip()

    return items

if __name__ == "__main__":
    os.makedirs("uploads", exist_ok=True)
    app.run(host="0.0.0.0", port=5000)

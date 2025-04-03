import pdfplumber
import pandas as pd
import os
import re
from flask import Flask, request, send_file, render_template
from pdf2image import convert_from_path
import pytesseract

app = Flask(__name__)

# Enhanced AV Equipment Patterns
AV_PATTERNS = {
    'manufacturer': re.compile(
        r'(Manufacturer|Mfg|Brand)[\s\#\:]+\s*([A-Za-z]{2,})',
        re.IGNORECASE
    ),
    'mfg_number': re.compile(
        r'(?:Part\s*(?:No|Number|#)|Mfg\s*[\#\:]|Model\s*(?:No|Number|#)|Item\s*ID)[\s\:\-]+\s*([A-Z0-9\-\.\/]{4,})',
        re.IGNORECASE
    ),
    'sku': re.compile(
        r'(SKU|Stock\s*No|Retailer\s*ID)[\s\:\-]+\s*([A-Z0-9\-]{5,})',
        re.IGNORECASE
    ),
    'quantity': re.compile(
        r'(Quantity|Qty|Units)[\s\:\-]+\s*(\d+)',
        re.IGNORECASE
    ),
    'price': re.compile(
        r'(Price|Cost|Each|Unit\s*Price)[\s\:\-]+\s*\$?((?:\d{1,3}[\,\.]?)+\d{1,3}(?:[\.\,]\d{2})?)',
        re.IGNORECASE
    ),
    'description': re.compile(
        r'(Description|Product|Item)[\s\:\-]+\s*(.+?)(?=\s*(?:Manufacturer|Mfg\s*[\#\:]|Qty|Price|\d{2,}\s*[A-Z]|$))',
        re.IGNORECASE | re.DOTALL
    ),
    'upc': re.compile(
        r'\b(UPC|EAN)[\s\:\-]+\s*(\d{12,13})\b',
        re.IGNORECASE
    )
}

MANUFACTURERS = {
    'ross', 'sony', 'panasonic', 'shure', 'sennheiser', 'qsc', 'bose',
    'crestron', 'extron', 'biamp', 'dante', 'poly', 'logitech', 'yamaha',
    'jbl', 'cisco', 'aten', 'blackmagic', 'ajax', 'lumens', 'epson', 'barco'
}

def extract_text(file_path):
    text = []
    try:
        with pdfplumber.open(file_path) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    text.extend(page_text.split('\n'))
        
        # OCR Fallback
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
        'manufacturer': '',
        'mfg_number': '',
        'description': '',
        'quantity': '',
        'price': '',
        'sku': '',
        'upc': '',
        'notes': []
    }

    # Process text lines
    for line in text_lines:
        line = re.sub(r'\s+', ' ', line).strip()
        
        # Manufacturer detection
        mfg_match = AV_PATTERNS['manufacturer'].search(line)
        if mfg_match and mfg_match.group(2).lower() in MANUFACTURERS:
            if current_item['manufacturer']:
                items.append(current_item)
                current_item = current_item.copy()
                current_item.update({
                    'manufacturer': '',
                    'mfg_number': '',
                    'description': '',
                    'quantity': '',
                    'price': '',
                    'sku': '',
                    'upc': '',
                    'notes': []
                })
            current_item['manufacturer'] = mfg_match.group(2).strip()
            line = line.replace(mfg_match.group(0), '')

        # Field extraction
        for field in ['mfg_number', 'sku', 'quantity', 'price', 'upc']:
            match = AV_PATTERNS[field].search(line)
            if match:
                current_item[field] = sanitize_field(field, match.group(2))
                line = line.replace(match.group(0), '')

        # Description extraction
        desc_match = AV_PATTERNS['description'].search(line)
        if desc_match and not current_item['description']:
            current_item['description'] = desc_match.group(2).strip()
            line = line.replace(desc_match.group(0), '')

        # Collect remaining notes
        if line.strip():
            current_item['notes'].append(line.strip())

    # Process tables
    for table in tables:
        if not table or len(table) < 2:
            continue

        headers = [cell.strip().lower() for cell in table[0] if cell]
        col_map = {}
        
        # Map columns to fields
        for idx, header in enumerate(headers):
            if 'mfg' in header or 'manufacturer' in header or 'part' in header:
                col_map['mfg_number'] = idx
            elif 'desc' in header or 'item' in header or 'product' in header:
                col_map['description'] = idx
            elif 'qty' in header or 'quantity' in header:
                col_map['quantity'] = idx
            elif 'price' in header or 'cost' in header:
                col_map['price'] = idx
            elif 'sku' in header or 'stock' in header:
                col_map['sku'] = idx
            elif 'upc' in header or 'ean' in header:
                col_map['upc'] = idx

        for row in table[1:]:
            item = current_item.copy()
            for field, col_idx in col_map.items():
                if col_idx < len(row):
                    item[field] = sanitize_field(field, row[col_idx])
            if item['mfg_number']:
                items.append(item)

    if current_item['manufacturer']:
        items.append(current_item)

    # Post-process items
    for item in items:
        if not item['description'] and item['notes']:
            item['description'] = item['notes'].pop(0)
        item['notes'] = ' | '.join(item['notes'])
        
        # Clean numerical fields
        item['quantity'] = int(re.sub(r'[^\d]', '', item['quantity'])) if item['quantity'] else 1
        item['price'] = float(re.sub(r'[^\d.]', '', item['price'])) if item['price'] else 0.0

    return items

def sanitize_field(field, value):
    if field == 'quantity':
        return str(int(re.sub(r'[^\d]', '', value))) if value else '1'
    elif field == 'price':
        return re.sub(r'[^\d.]', '', value) if value else '0.00'
    elif field == 'mfg_number':
        return re.sub(r'[^A-Z0-9\-/]', '', value.upper())
    elif field == 'sku':
        return re.sub(r'[^A-Z0-9\-]', '', value.upper())
    elif field == 'upc':
        return re.sub(r'[^\d]', '', value)
    return value.strip()

@app.route("/", methods=["GET", "POST"])
def upload_file():
    if request.method == "POST":
        if 'file' not in request.files:
            return "No file uploaded", 400
            
        pdf_file = request.files['file']
        if pdf_file.filename == '':
            return "Invalid file", 400

        try:
            os.makedirs("uploads", exist_ok=True)
            file_path = os.path.join("uploads", pdf_file.filename)
            pdf_file.save(file_path)

            text_data = extract_text(file_path)
            tables = extract_tables(file_path)
            structured_data = process_data(text_data, tables)

            df = pd.DataFrame(structured_data, columns=[
                'manufacturer', 'mfg_number', 'description',
                'quantity', 'price', 'sku', 'upc', 'notes'
            ])

            output_path = "converted.xlsx"
            df.to_excel(output_path, index=False, engine='openpyxl')
            
            return render_template("index.html", url=output_path)

        except Exception as e:
            return f"Error processing file: {str(e)}", 500
        finally:
            if os.path.exists(file_path):
                os.remove(file_path)

    return render_template("index.html")

if __name__ == "__main__":
    os.makedirs("uploads", exist_ok=True)
    app.run(host="0.0.0.0", port=5000)

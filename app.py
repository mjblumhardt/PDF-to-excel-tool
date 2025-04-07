import pdfplumber
import pandas as pd
import os
import re
import logging
from flask import Flask, request, send_file, render_template
from pdf2image import convert_from_path
import pytesseract

# Configure Tesseract paths
if os.environ.get('RENDER'):
    pytesseract.pytesseract.tesseract_cmd = '/usr/bin/tesseract'
else:
    pytesseract.pytesseract.tesseract_cmd = r'/app/.apt/usr/bin/tesseract'

app = Flask(__name__)
app.logger.setLevel(logging.INFO)

MANUFACTURER_PROFILES = {
    'Ross Video': {
        'patterns': [
            r'\b([A-Z]{3,}-[A-Z0-9-]+)\b',  # Matches CUF-124, TD2S-PANEL
            r'\b([A-Z]{3,}\d+-\w+)\b'       # Matches XDS0-0001-CPS
        ],
        'keywords': ['ROSS', 'Carbonite', 'Ultrix']
    },
    'B&H': {
        'patterns': [
            r'\b(?:MFR|mfr|Manufacturer)\s*[#:]?\s*([A-Z0-9-]+)\b',
            r'\b(BH-[A-Z0-9-]+)\b'
        ],
        'keywords': ['B&H', 'BHPhoto']
    },
    'Extron': {
        'patterns': [r'\b(\d+-\d+-\d+)\b'],
        'keywords': ['Extron']
    },
    'Generic': {
        'patterns': [
            r'\b([A-Z]{2,5}\d{3,}-[A-Z0-9]+)\b',
            r'\b([A-Z]{2,5}-\d{3,}-\w+)\b'
        ]
    }
}

@app.route("/", methods=["GET", "POST"])
def upload_file():
    if request.method == "POST":
        pdf_file = request.files["file"]
        if pdf_file and pdf_file.filename.endswith('.pdf'):
            try:
                os.makedirs("uploads", exist_ok=True)
                file_path = os.path.join("uploads", pdf_file.filename)
                pdf_file.save(file_path)

                text_data = extract_text_with_fallback(file_path)
                tables = extract_tables(file_path)
                
                structured_data = []
                structured_data += process_text_lines(text_data)
                structured_data += process_tables(tables)
                
                df = clean_dataframe(structured_data)
                output_path = "converted.xlsx"
                df.to_excel(output_path, index=False, engine='openpyxl')
                
                return send_file(output_path, as_attachment=True)

            except Exception as e:
                app.logger.error(f"Processing failed: {str(e)}")
                return f"Error processing PDF: {str(e)}", 500
            finally:
                for path in [file_path, "converted.xlsx"]:
                    if path and os.path.exists(path):
                        os.remove(path)
        else:
            return "Invalid file format", 400
    return render_template("index.html")

def extract_text_with_fallback(file_path):
    try:
        text = []
        with pdfplumber.open(file_path) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    text.extend(page_text.split('\n'))
        
        if not text or all(line.strip() == '' for line in text):
            images = convert_from_path(file_path, dpi=300)
            for img in images:
                text.extend(pytesseract.image_to_string(img).split('\n'))
        
        return [re.sub(r'\s+', ' ', line).strip() for line in text if line.strip()]
    
    except Exception as e:
        app.logger.error(f"Text extraction failed: {str(e)}")
        return []

def extract_tables(file_path):
    try:
        tables = []
        with pdfplumber.open(file_path) as pdf:
            for page in pdf.pages:
                page_tables = page.extract_tables({
                    "vertical_strategy": "lines",
                    "horizontal_strategy": "lines",
                    "snap_tolerance": 4
                })
                tables.extend(page_tables)
        return tables
    except Exception as e:
        app.logger.error(f"Table extraction failed: {str(e)}")
        return []

def process_text_lines(text_lines):
    items = []
    current_item = new_item()
    
    for line in text_lines:
        line = line.strip()
        if not line:
            continue

        # Ross-specific pattern matching
        mfg_num, mfg_match = extract_manufacturer(line)
        if mfg_num:
            if current_item['Manufacturer Number']:
                items.append(current_item)
                current_item = new_item()
            current_item['Manufacturer Number'] = mfg_num
            line = line.replace(mfg_match.group(), '').strip()

        if not current_item['Quantity']:
            qty, qty_match = extract_quantity(line)
            if qty:
                current_item['Quantity'] = qty
                line = line.replace(qty_match.group(), '').strip()

        # Extract Ross pricing components
        if not current_item['List Price']:
            list_price, lp_match = extract_price(line, prefix='List Price:?')
            if list_price:
                current_item['List Price'] = list_price
                line = line.replace(lp_match.group(), '').strip()

        if not current_item['Discount']:
            discount, disc_match = extract_discount(line)
            if discount:
                current_item['Discount'] = discount
                line = line.replace(disc_match.group(), '').strip()

        if line:
            current_item['Description'].append(line)

    if current_item['Manufacturer Number']:
        items.append(current_item)
    
    return items

def process_tables(tables):
    items = []
    for table in tables:
        if len(table) < 2:
            continue

        headers = [str(cell).strip().lower() for cell in table[0]]
        col_map = {
            'product_info': detect_column(headers, ['product info', 'item']),
            'qty': detect_column(headers, ['qty', 'quantity']),
            'list_price': detect_column(headers, ['list price']),
            'disc': detect_column(headers, ['disc.', 'discount']),
            'net_unit': detect_column(headers, ['net unit']),
            'net_price': detect_column(headers, ['net price'])
        }

        for row in table[1:]:
            item = new_item()
            product_info = str(row[col_map['product_info']]).strip() if col_map['product_info'] is not None else ''
            
            # Extract Ross product codes from table cells
            mfg_match = re.match(r'^[→\s]*([A-Z-0-9]+)\b', product_info)
            if mfg_match:
                item['Manufacturer Number'] = mfg_match.group(1).strip()
                description = product_info.replace(mfg_match.group(0), '').strip()
                item['Description'] = description

            # Map Ross-specific columns
            if col_map['qty'] is not None:
                item['Quantity'] = str(row[col_map['qty']]).strip()
            if col_map['list_price'] is not None:
                item['List Price'] = str(row[col_map['list_price']]).strip()
            if col_map['disc'] is not None:
                item['Discount'] = str(row[col_map['disc']]).strip()
            if col_map['net_unit'] is not None:
                item['Net Unit'] = str(row[col_map['net_unit']]).strip()
            if col_map['net_price'] is not None:
                item['Net Price'] = str(row[col_map['net_price']]).strip()

            if item['Manufacturer Number']:
                items.append(item)
    
    return items

def clean_dataframe(raw_data):
    df = pd.DataFrame(raw_data, columns=[
        'Manufacturer Number', 
        'Description',
        'Quantity',
        'List Price',
        'Discount',
        'Net Unit',
        'Net Price',
        'Notes'
    ])
    
    # Clean numerical fields
    for col in ['List Price', 'Discount', 'Net Unit', 'Net Price']:
        df[col] = df[col].apply(lambda x: re.sub(r'[^\d.]', '', str(x)) if x else '')
    
    df['Quantity'] = df['Quantity'].apply(
        lambda x: re.sub(r'\D', '', str(x)) if x else '')
    
    # Clean description field
    df['Description'] = df['Description'].apply(
        lambda x: ' '.join(x) if isinstance(x, list) else str(x))
    
    return df.drop_duplicates().reset_index(drop=True)

def new_item():
    return {
        'Manufacturer Number': '',
        'Description': [],
        'Quantity': '',
        'List Price': '',
        'Discount': '',
        'Net Unit': '',
        'Net Price': '',
        'Notes': []
    }

def extract_manufacturer(text):
    text = text.upper()
    for profile in MANUFACTURER_PROFILES.values():
        for pattern in profile['patterns']:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                return match.group(1).strip(), match
    return None, None

def extract_quantity(text):
    match = re.search(r'\b(\d+)\s*(?:pc|ea|units?)\b', text, re.IGNORECASE)
    return (match.group(1), match) if match else (None, None)

def extract_price(text, prefix=''):
    match = re.search(rf'{prefix}\$?\s*(\d{{1,3}}(?:,\d{{3}})*\.\d{{2}})', text)
    return (match.group(1).replace(',', ''), match) if match else (None, None)

def extract_discount(text):
    match = re.search(r'\b(\d+)%\b', text)
    return (match.group(1), match) if match else (None, None)

def detect_column(headers, keywords):
    for idx, header in enumerate(headers):
        if any(kw in header for kw in keywords):
            return idx
    return None

if __name__ == "__main__":
    os.makedirs("uploads", exist_ok=True)
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))

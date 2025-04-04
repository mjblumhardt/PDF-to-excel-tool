import pdfplumber
import pandas as pd
import os
import re
from flask import Flask, request, send_file, render_template
from pdf2image import convert_from_path
import pytesseract

app = Flask(__name__)

# Configure Tesseract path for Render compatibility
pytesseract.pytesseract.tesseract_cmd = '/usr/bin/tesseract'

# Enhanced manufacturer detection system
MANUFACTURER_PROFILES = {
    'B&H': {
        'patterns': [
            r'\b(?:MFR|mfr|Manufacturer)\s*[#:]?\s*([A-Z0-9]+(?:-[A-Z0-9]+)+)\b',
            r'\b(BH-\w+-\d+)\b'
        ],
        'keywords': ['B&H', 'BHPhoto', 'MFR']
    },
    'Extron': {
        'patterns': [
            r'\b(\d+-\d+-\d+)\b',
            r'\b(Extron[-\s]\d+-\d+-\d+)\b'
        ],
        'keywords': ['Extron']
    },
    'Generic': {
        'patterns': [
            r'\b([A-Z]{2,5}\d{3,}-[A-Z0-9]+)\b',
            r'\b([A-Z]{2,5}-\d{3,}-\w+)\b',
            r'\b(\d{3,}[A-Z]{2,5}\d*)\b'
        ]
    }
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

                # Comprehensive data extraction
                text_data = enhanced_extract_text(file_path)
                tables = enhanced_extract_tables(file_path)
                
                # Multi-stage processing
                structured_data = []
                structured_data += process_text_lines(text_data)
                structured_data += process_all_tables(tables)
                structured_data += fallback_ocr_analysis(file_path, text_data)

                # Data normalization
                df = clean_and_structure_data(structured_data)

                # System cleanup
                output_path = "converted.xlsx"
                df.to_excel(output_path, index=False, engine='openpyxl')
                return send_file(output_path, as_attachment=True)

            except Exception as e:
                app.logger.error(f"Critical Error: {str(e)}")
                return f"System Error: {str(e)}", 500
            finally:
                if os.path.exists(file_path):
                    os.remove(file_path)
                if os.path.exists("converted.xlsx"):
                    os.remove("converted.xlsx")
    
    return render_template("index.html")

def enhanced_extract_text(file_path):
    text = []
    try:
        # Primary PDF text extraction
        with pdfplumber.open(file_path) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    text.extend(page_text.split('\n'))
                else:
                    # Fallback to layout analysis
                    layout_text = page.extract_text_layout()
                    if layout_text:
                        text.extend(layout_text.split('\n'))

        # OCR fallback with image preprocessing
        if not text or all(line.strip() == '' for line in text):
            images = convert_from_path(file_path, dpi=300)
            for img in images:
                ocr_text = pytesseract.image_to_string(
                    img, 
                    config='--psm 6 --oem 3 -c preserve_interword_spaces=1'
                )
                text.extend(ocr_text.split('\n'))

        # Advanced text cleaning
        return [re.sub(r'\s+', ' ', line).strip() for line in text if line.strip()]
    
    except Exception as e:
        app.logger.error(f"Text Extraction Failed: {str(e)}")
        return []

def enhanced_extract_tables(file_path):
    tables = []
    try:
        with pdfplumber.open(file_path) as pdf:
            for page in pdf.pages:
                # Extract tables with multiple strategies
                page_tables = page.extract_tables({
                    "vertical_strategy": "lines", 
                    "horizontal_strategy": "lines",
                    "explicit_vertical_lines": [],
                    "explicit_horizontal_lines": [],
                    "snap_tolerance": 4
                })
                if page_tables:
                    tables.extend(page_tables)

        return tables
    
    except Exception as e:
        app.logger.error(f"Table Extraction Failed: {str(e)}")
        return []

def process_text_lines(text_lines):
    items = []
    current_item = {
        'Manufacturer Number': '',
        'Description': [],
        'Quantity': '',
        'Price': '',
        'Notes': []
    }

    for line in text_lines:
        line = line.strip()
        if not line:
            continue

        # Manufacturer detection
        mfg_num = detect_manufacturer(line)
        if mfg_num:
            if current_item['Manufacturer Number']:
                items.append(finalize_item(current_item))
            current_item = new_item(mfg_num)
            line = remove_matched_pattern(line, mfg_num)

        # Quantity detection
        if not current_item['Quantity']:
            qty = extract_quantity(line)
            if qty:
                current_item['Quantity'] = qty
                line = remove_matched_pattern(line, qty)

        # Price detection
        if not current_item['Price']:
            price = extract_price(line)
            if price:
                current_item['Price'] = price
                line = remove_matched_pattern(line, price)

        # Description collection
        if line:
            current_item['Description'].append(line)

    if current_item['Manufacturer Number']:
        items.append(finalize_item(current_item))
    
    return items

def process_all_tables(tables):
    items = []
    for table in tables:
        if len(table) < 2:
            continue

        # Dynamic column mapping
        headers = [str(cell).strip().lower() for cell in table[0]]
        col_map = {
            'mfg': detect_column(headers, ['mfg', 'manufacturer', 'part']),
            'desc': detect_column(headers, ['desc', 'description', 'item']),
            'qty': detect_column(headers, ['qty', 'quantity']),
            'price': detect_column(headers, ['price', 'cost', 'unit'])
        }

        # Process table rows
        for row in table[1:]:
            if len(row) < len(headers):
                continue

            item = {
                'Manufacturer Number': safe_extract(row, col_map['mfg']),
                'Description': safe_extract(row, col_map['desc']),
                'Quantity': safe_extract(row, col_map['qty']),
                'Price': clean_price(safe_extract(row, col_map['price'])),
                'Notes': []
            }

            # Fallback manufacturer detection
            if not item['Manufacturer Number']:
                item['Manufacturer Number'] = detect_manufacturer(item['Description'])

            items.append(item)
    
    return items

def fallback_ocr_analysis(file_path, existing_text):
    items = []
    if len(existing_text) < 5:  # If minimal text found
        try:
            images = convert_from_path(file_path)
            for img in images:
                ocr_text = pytesseract.image_to_string(img)
                items += process_text_lines(ocr_text.split('\n'))
        except Exception as e:
            app.logger.error(f"OCR Fallback Failed: {str(e)}")
    return items

# Helper functions --------------------------------------------------

def detect_manufacturer(text):
    text = str(text).upper()
    for manufacturer, profile in MANUFACTURER_PROFILES.items():
        # Keyword matching
        if any(kw.upper() in text for kw in profile.get('keywords', [])):
            for pattern in profile['patterns']:
                match = re.search(pattern, text, re.IGNORECASE)
                if match:
                    return match.group(1).strip()
        # Generic pattern matching
        for pattern in profile['patterns']:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                return match.group(1).strip()
    return ''

def clean_and_structure_data(raw_data):
    df = pd.DataFrame(raw_data, columns=[
        'Manufacturer Number', 
        'Description', 
        'Quantity', 
        'Price', 
        'Notes'
    ])
    
    # Data cleaning pipeline
    df = df.drop_duplicates().reset_index(drop=True)
    df['Description'] = df['Description'].apply(lambda x: ' '.join(x) if isinstance(x, list) else x)
    df['Price'] = df['Price'].apply(clean_price)
    df['Quantity'] = df['Quantity'].apply(clean_quantity)
    
    return df

def clean_price(price_str):
    return re.sub(r'[^\d.]', '', str(price_str)) if price_str else ''

def clean_quantity(qty_str):
    return re.sub(r'\D', '', str(qty_str)) if qty_str else ''

def new_item(mfg_num):
    return {
        'Manufacturer Number': mfg_num,
        'Description': [],
        'Quantity': '',
        'Price': '',
        'Notes': []
    }

def finalize_item(item):
    return {
        'Manufacturer Number': item['Manufacturer Number'],
        'Description': item['Description'],
        'Quantity': item['Quantity'],
        'Price': item['Price'],
        'Notes': item['Notes']
    }

def detect_column(headers, keywords):
    for idx, header in enumerate(headers):
        if any(kw in header for kw in keywords):
            return idx
    return None

def safe_extract(row, index):
    return str(row[index]).strip() if index is not None and len(row) > index else ''

def remove_matched_pattern(text, pattern):
    return re.sub(re.escape(pattern), '', text, flags=re.IGNORECASE).strip()

if __name__ == "__main__":
    os.makedirs("uploads", exist_ok=True)
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))

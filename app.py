import pdfplumber
import pandas as pd
import os
import re
import logging
from flask import Flask, request, send_file, render_template
from pdf2image import convert_from_path
import pytesseract

if os.environ.get('RENDER'):
    pytesseract.pytesseract.tesseract_cmd = '/usr/bin/tesseract'
else:
    pytesseract.pytesseract.tesseract_cmd = r'/app/.apt/usr/bin/tesseract'  # Render-specific path
# Configure application
app = Flask(__name__)
app.logger.setLevel(logging.INFO)

# Configure Tesseract for Render deployment
pytesseract.pytesseract.tesseract_cmd = '/usr/bin/tesseract'

# Manufacturer detection profiles
MANUFACTURER_PROFILES = {
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
                # Create uploads directory
                os.makedirs("uploads", exist_ok=True)
                file_path = os.path.join("uploads", pdf_file.filename)
                pdf_file.save(file_path)

                # Extract data from PDF
                text_data = extract_text_with_fallback(file_path)
                tables = extract_tables(file_path)
                
                # Process data
                structured_data = []
                structured_data += process_text_lines(text_data)
                structured_data += process_tables(tables)
                
                # Create final dataframe
                df = clean_dataframe(structured_data)
                
                # Generate Excel file
                output_path = "converted.xlsx"
                df.to_excel(output_path, index=False, engine='openpyxl')
                
                return send_file(output_path, as_attachment=True)

            except Exception as e:
                app.logger.error(f"Processing failed: {str(e)}")
                return f"Error processing PDF: {str(e)}", 500
            finally:
                # Cleanup temporary files
                for path in [file_path, "converted.xlsx"]:
                    if path and os.path.exists(path):
                        os.remove(path)
        else:
            return "Invalid file format", 400
    return render_template("index.html")

def extract_text_with_fallback(file_path):
    """Extract text with OCR fallback"""
    try:
        text = []
        with pdfplumber.open(file_path) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    text.extend(page_text.split('\n'))
        
        # OCR fallback if no text found
        if not text or all(line.strip() == '' for line in text):
            images = convert_from_path(file_path, dpi=300)
            for img in images:
                text.extend(pytesseract.image_to_string(img).split('\n'))
        
        return [re.sub(r'\s+', ' ', line).strip() for line in text if line.strip()]
    
    except Exception as e:
        app.logger.error(f"Text extraction failed: {str(e)}")
        return []

def extract_tables(file_path):
    """Extract tables with multiple strategies"""
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
    """Process individual text lines"""
    items = []
    current_item = new_item()
    
    for line in text_lines:
        line = line.strip()
        if not line:
            continue

        # Manufacturer detection
        mfg_num, mfg_match = extract_manufacturer(line)
        if mfg_num:
            if current_item['Manufacturer Number']:
                items.append(current_item)
                current_item = new_item()
            current_item['Manufacturer Number'] = mfg_num
            line = line.replace(mfg_match.group(), '').strip()

        # Quantity extraction
        if not current_item['Quantity']:
            qty, qty_match = extract_quantity(line)
            if qty:
                current_item['Quantity'] = qty
                line = line.replace(qty_match.group(), '').strip()

        # Price extraction
        if not current_item['Price']:
            price, price_match = extract_price(line)
            if price:
                current_item['Price'] = price
                line = line.replace(price_match.group(), '').strip()

        # Description handling
        if line:
            current_item['Description'].append(line)

    if current_item['Manufacturer Number']:
        items.append(current_item)
    
    return items

def process_tables(tables):
    """Process PDF tables"""
    items = []
    for table in tables:
        if len(table) < 2:
            continue

        headers = [str(cell).strip().lower() for cell in table[0]]
        col_map = {
            'mfg': detect_column(headers, ['mfg', 'manufacturer', 'part']),
            'desc': detect_column(headers, ['desc', 'description', 'item']),
            'qty': detect_column(headers, ['qty', 'quantity']),
            'price': detect_column(headers, ['price', 'cost', 'unit'])
        }

        for row in table[1:]:
            item = new_item()
            for idx, cell in enumerate(row):
                value = str(cell).strip() if cell else ''
                
                if idx == col_map['mfg']:
                    item['Manufacturer Number'] = value
                elif idx == col_map['desc']:
                    item['Description'] = [value]
                elif idx == col_map['qty']:
                    item['Quantity'] = value
                elif idx == col_map['price']:
                    item['Price'] = value
            
            if item['Manufacturer Number']:
                items.append(item)
    
    return items

def clean_dataframe(raw_data):
    """Clean and structure final dataframe"""
    df = pd.DataFrame(raw_data, columns=[
        'Manufacturer Number', 
        'Description', 
        'Quantity', 
        'Price', 
        'Notes'
    ])
    
    # Data cleaning operations
    df['Description'] = df['Description'].apply(
        lambda x: ' '.join(x) if isinstance(x, list) else str(x)
    df['Price'] = df['Price'].apply(
        lambda x: re.sub(r'[^\d.]', '', str(x)) if x else '')
    df['Quantity'] = df['Quantity'].apply(
        lambda x: re.sub(r'\D', '', str(x)) if x else '')
    
    return df.drop_duplicates().reset_index(drop=True)

# Helper functions --------------------------------------------------

def new_item():
    return {
        'Manufacturer Number': '',
        'Description': [],
        'Quantity': '',
        'Price': '',
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

def extract_price(text):
    match = re.search(r'\$?\s*(\d{1,3}(?:,\d{3})*\.\d{2})', text)
    return (match.group(1).replace(',', ''), match) if match else (None, None)

def detect_column(headers, keywords):
    for idx, header in enumerate(headers):
        if any(kw in header for kw in keywords):
            return idx
    return None

if __name__ == "__main__":
    os.makedirs("uploads", exist_ok=True)
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))

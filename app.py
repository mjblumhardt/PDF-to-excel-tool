import pdfplumber
import pandas as pd
import os
import re
from flask import Flask, request, send_file, render_template

app = Flask(__name__)

# Manufacturer detection profiles
MANUFACTURER_PROFILES = {
    'B&H': {
        'patterns': [
            r'\b(?:MFR|mfr|Manufacturer)\s*[#:]?\s*([A-Z0-9]+(?:-[A-Z0-9]+)*\b',
            r'\b(BH-\w+-\d+)\b'
        ],
        'keywords': ['B&H', 'BHPhoto', 'MFR']
    },
    'Extron': {
        'patterns': [r'\b(\d+-\d+-\d+)\b'],
        'keywords': ['Extron']
    },
    'Generic': {
        'patterns': [
            r'\b([A-Z]{2,5}\d{3,}-[A-Z0-9]+)\b',
            r'\b(\d+[A-Z]{2,3}\d+)\b'
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

                structured_data = []
                
                with pdfplumber.open(file_path) as pdf:
                    for page in pdf.pages:
                        # Extract and process both text and tables
                        text = page.extract_text()
                        tables = page.extract_tables()
                        
                        structured_data += process_text(text)
                        structured_data += process_tables(tables)

                # Create cleaned DataFrame
                df = pd.DataFrame(structured_data, columns=[
                    'Manufacturer Number', 
                    'Description', 
                    'Quantity', 
                    'Price', 
                    'Notes'
                ]).drop_duplicates().reset_index(drop=True)

                output_path = "converted.xlsx"
                df.to_excel(output_path, index=False, engine='openpyxl')
                return send_file(output_path, as_attachment=True)

            except Exception as e:
                return f"Error: {str(e)}", 500
            finally:
                if os.path.exists(file_path):
                    os.remove(file_path)
    return render_template("index.html")

def process_text(text):
    items = []
    lines = [line.strip() for line in text.split('\n') if line.strip()]
    
    current_item = {
        'Manufacturer Number': '',
        'Description': '',
        'Quantity': '',
        'Price': '',
        'Notes': []
    }

    for line in lines:
        # Detect manufacturer number
        mfg_num = detect_manufacturer(line)
        if mfg_num:
            if current_item['Manufacturer Number']:
                items.append(current_item)
                current_item = {
                    'Manufacturer Number': mfg_num,
                    'Description': '',
                    'Quantity': '',
                    'Price': '',
                    'Notes': []
                }
            else:
                current_item['Manufacturer Number'] = mfg_num
        
        # Extract quantity and price
        if not current_item['Quantity']:
            qty_match = re.search(r'\b(\d+)\s*(?:[Pp][Cc][Ss]?|[Ee][Aa])\b', line)
            if qty_match:
                current_item['Quantity'] = qty_match.group(1)
        
        if not current_item['Price']:
            price_match = re.search(r'\$(\d+[\.,]?\d{0,2})', line)
            if price_match:
                current_item['Price'] = price_match.group(1).replace(',', '')
        
        # Collect description
        if not any([current_item['Manufacturer Number'], 
                   current_item['Quantity'], 
                   current_item['Price']]):
            current_item['Description'] += ' ' + line
    
    if current_item['Manufacturer Number']:
        items.append(current_item)
    
    return items

def process_tables(tables):
    items = []
    for table in tables:
        if len(table) > 1 and len(table[0]) > 3:
            headers = [str(cell).lower().strip() for cell in table[0]]
            col_map = {
                'mfg': None,
                'desc': None,
                'qty': None,
                'price': None
            }
            
            # Map columns dynamically
            for idx, header in enumerate(headers):
                if 'mfg' in header or 'manufacturer' in header:
                    col_map['mfg'] = idx
                elif 'desc' in header or 'item' in header:
                    col_map['desc'] = idx
                elif 'qty' in header:
                    col_map['qty'] = idx
                elif 'price' in header or 'unit' in header:
                    col_map['price'] = idx
            
            for row in table[1:]:
                if len(row) > max([v for v in col_map.values() if v is not None]):
                    item = {
                        'Manufacturer Number': str(row[col_map['mfg']).strip() if col_map['mfg'] else '',
                        'Description': str(row[col_map['desc']).strip() if col_map['desc'] else '',
                        'Quantity': str(row[col_map['qty']).strip() if col_map['qty'] else '',
                        'Price': str(row[col_map['price']).replace('$','').strip() if col_map['price'] else '',
                        'Notes': []
                    }
                    
                    # Fallback manufacturer detection
                    if not item['Manufacturer Number']:
                        item['Manufacturer Number'] = detect_manufacturer(item['Description'])
                    
                    items.append(item)
    return items

def detect_manufacturer(text):
    text = str(text).upper()
    for manufacturer, profile in MANUFACTURER_PROFILES.items():
        # Check for brand keywords first
        if any(keyword.upper() in text for keyword in profile.get('keywords', [])):
            for pattern in profile['patterns']:
                match = re.search(pattern, text, re.IGNORECASE)
                if match:
                    return match.group(1).strip()
        # Generic patterns as fallback
        else:
            for pattern in profile.get('patterns', []):
                match = re.search(pattern, text, re.IGNORECASE)
                if match:
                    return match.group(1).strip()
    return ''

if __name__ == "__main__":
    os.makedirs("uploads", exist_ok=True)
    app.run(host="0.0.0.0", port=5000)

import pdfplumber
import pandas as pd
import os
import re
from flask import Flask, request, send_file, render_template
from pdf2image import convert_from_path
import pytesseract

app = Flask(__name__)

# Enhanced pattern matching with better field detection
PATTERNS = {
    'mfg_number': re.compile(
        r'(Manufacturer\s*[\#:]?|Mfg\s*[\#:]?|Part\s*No|Item\s*ID)\s*([A-Z0-9-]+)',
        re.IGNORECASE
    ),
    'quantity': re.compile(
        r'(Quantity|Qty|Amount)[\:\s]+(\d+)', 
        re.IGNORECASE
    ),
    'price': re.compile(
        r'(Price|Unit\s*Price|Cost|Each)[\:\s]+\$?(\d+[\.,]?\d{0,2})',
        re.IGNORECASE
    ),
    'description': re.compile(
        r'(Description|Item|Product)[\:\s]+(.+?)(?=\s*(?:Manufacturer|Qty|Price|$))',
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

                # Extract data
                text_data = extract_text(file_path)
                tables = extract_tables(file_path)
                
                # Process and combine data
                structured_data = process_data(text_data, tables)

                # Create DataFrame
                df = pd.DataFrame(structured_data, columns=[
                    'Manufacturer Number', 
                    'Description', 
                    'Quantity', 
                    'Price', 
                    'Notes'
                ])

                # Debug output
                print("\n=== FINAL STRUCTURED DATA ===")
                print(df)

                # Save and return
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

        # Debug print
        print("\n=== RAW EXTRACTED TEXT ===")
        for idx, line in enumerate(text):
            print(f"Line {idx+1}: {line}")

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
        
        # Debug print
        print("\n=== RAW TABLE DATA ===")
        for idx, table in enumerate(tables):
            print(f"Table {idx+1}:")
            for row in table:
                print(row)
    
    except Exception as e:
        print(f"Table extraction error: {str(e)}")
    return tables

def process_data(text_lines, tables):
    items = []
    current_item = {
        'Manufacturer Number': '',
        'Description': '',
        'Quantity': '',
        'Price': '',
        'Notes': []
    }

    # Process text lines
    for line in text_lines:
        found = False
        
        # Manufacturer Number starts new item
        mfg_match = PATTERNS['mfg_number'].search(line)
        if mfg_match:
            if current_item['Manufacturer Number']:
                items.append(current_item)
                current_item = current_item.copy()
                current_item.update({
                    'Manufacturer Number': '',
                    'Description': '',
                    'Quantity': '',
                    'Price': '',
                    'Notes': []
                })
            current_item['Manufacturer Number'] = mfg_match.group(2).strip()
            found = True
            line = line.replace(mfg_match.group(0), '')  # Remove matched part

        # Check other patterns
        for field in ['quantity', 'price', 'description']:
            match = PATTERNS[field].search(line)
            if match:
                current_item[field.capitalize()] = match.group(2).strip()
                found = True
                line = line.replace(match.group(0), '')  # Remove matched part
                break

        # Collect remaining text as notes
        if line.strip() and not found:
            current_item['Notes'].append(line.strip())

    # Process tables
    for table in tables:
        if not table or len(table) < 2:
            continue

        headers = [str(cell).strip().lower() for cell in table[0]]
        col_map = {}
        
        # Map columns to fields
        for idx, header in enumerate(headers):
            if 'mfg' in header or 'manufacturer' in header or 'part' in header:
                col_map['mfg'] = idx
            elif 'desc' in header or 'item' in header or 'product' in header:
                col_map['desc'] = idx
            elif 'qty' in header or 'quantity' in header:
                col_map['qty'] = idx
            elif 'price' in header or 'cost' in header:
                col_map['price'] = idx

        # Process rows
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
                if idx == col_map.get('mfg'):
                    item['Manufacturer Number'] = cell_value
                elif idx == col_map.get('desc'):
                    item['Description'] = cell_value
                elif idx == col_map.get('qty'):
                    item['Quantity'] = cell_value
                elif idx == col_map.get('price'):
                    item['Price'] = cell_value
                elif cell_value:
                    notes.append(cell_value)
            
            item['Notes'] = ' | '.join(notes)
            if item['Manufacturer Number']:
                items.append(item)

    # Add final item
    if current_item['Manufacturer Number']:
        items.append(current_item)

    # Clean up descriptions and notes
    for item in items:
        if not item['Description'] and item['Notes']:
            item['Description'] = item['Notes'].pop(0)
        item['Notes'] = '; '.join(item['Notes'])
    
    return items

if __name__ == "__main__":
    os.makedirs("uploads", exist_ok=True)
    app.run(host="0.0.0.0", port=5000)

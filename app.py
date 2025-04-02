import pdfplumber
import pandas as pd
import os
from flask import Flask, request, send_file, render_template
from pdf2image import convert_from_path
import pytesseract

app = Flask(__name__)

@app.route("/", methods=["GET", "POST"])
def upload_file():
    if request.method == "POST":
        pdf_file = request.files["file"]
        if pdf_file:
            try:
                os.makedirs("uploads", exist_ok=True)
                file_path = os.path.join("uploads", pdf_file.filename)
                pdf_file.save(file_path)

                # Extract text and tables from PDF
                text_data, tables = extract_text_and_tables(file_path)
                # If no text is extracted, fall back to OCR
                if not text_data or all(line == "" for line in text_data):
                    text_data = extract_text_with_ocr(file_path)

                output_path = save_to_excel(text_data, tables)
                return send_file(output_path, as_attachment=True)

            except Exception as e:
                print(f"Error: {str(e)}")
                return f"Error processing file: {str(e)}", 500

    return render_template("index.html")

def extract_text_and_tables(file_path):
    extracted_text = []
    tables = []
    try:
        with pdfplumber.open(file_path) as pdf_doc:
            for page in pdf_doc.pages:
                # Extract text
                text = page.extract_text()
                if text:
                    extracted_text.extend(text.split("\n"))
                # Extract tables (if any)
                page_tables = page.extract_tables()
                if page_tables:
                    for t in page_tables:
                        tables.append(t)
    except Exception as e:
        return f"Error reading PDF: {str(e)}", []
    return extracted_text if extracted_text else ["No text found"], tables

def extract_text_with_ocr(file_path):
    text = []
    try:
        images = convert_from_path(file_path)
        for img in images:
            extracted_text = pytesseract.image_to_string(img)
            text.extend(extracted_text.split("\n"))
    except Exception as e:
        return [f"OCR failed: {str(e)}"]
    return text if text else ["No text found (even with OCR)."]

def save_to_excel(data, tables):
    output_path = "converted.xlsx"
    # Save text data to a DataFrame (each line as a row)
    text_df = pd.DataFrame({'Extracted Text': data})
    
    with pd.ExcelWriter(output_path) as writer:
        text_df.to_excel(writer, sheet_name='Text Data', index=False)
        
        # Process and save each extracted table to its own sheet
        for i, table in enumerate(tables):
            # Assume first row is header
            header = table[0]
            fixed_rows = []
            for row in table[1:]:
                if len(row) < len(header):
                    row = row + [''] * (len(header) - len(row))
                elif len(row) > len(header):
                    row = row[:len(header)]
                fixed_rows.append(row)
            table_df = pd.DataFrame(fixed_rows, columns=header)
            table_df.to_excel(writer, sheet_name=f'Table {i+1}', index=False)
            
    return output_path

if __name__ == "__main__":
    os.makedirs("uploads", exist_ok=True)
    app.run(host="0.0.0.0", port=5000)

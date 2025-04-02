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

                # Extract text using pdfplumber first, fallback to OCR if needed
                text_data, tables = extract_text_and_tables(file_path)
                
                # Fix: Check if the extracted text is empty (check the list, not using .strip())
                if not text_data or all(line == "" for line in text_data):
                    text_data = extract_text_with_ocr(file_path)

                output_path = save_to_excel(text_data, tables)

                return send_file(output_path, as_attachment=True)

            except Exception as e:
                print(f"Error: {str(e)}")  # Log the error
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
                    extracted_text.extend(text.split("\n"))  # Split into lines

                # Extract tables (if any)
                table = page.extract_tables()
                if table:
                    tables.append(table)
    except Exception as e:
        return f"Error reading PDF: {str(e)}", []

    return extracted_text if extracted_text else ["No text found"], tables

def extract_text_with_ocr(file_path):
    text = []
    try:
        images = convert_from_path(file_path)
        for img in images:
            extracted_text = pytesseract.image_to_string(img)
            text.extend(extracted_text.split("\n"))  # Split OCR text into lines
    except Exception as e:
        return [f"OCR failed: {str(e)}"]
    
    return text if text else ["No text found (even with OCR)."]

def save_to_excel(data, tables):
    output_path = "converted.xlsx"

    # Convert text data to DataFrame where each line is a new row
    text_df = pd.DataFrame({'Extracted Text': data})
    
    # Handle extracted tables
    with pd.ExcelWriter(output_path) as writer:
        # Save the text data
        text_df.to_excel(writer, sheet_name='Text Data', index=False)
        
        # Save any extracted tables
        for i, table in enumerate(tables):
            table_df = pd.DataFrame(table[1:], columns=table[0])  # Convert to DataFrame
            table_df.to_excel(writer, sheet_name=f'Table {i+1}', index=False)

    return output_path

if __name__ == "__main__":
    os.makedirs("uploads", exist_ok=True)
    app.run(host="0.0.0.0", port=5000)


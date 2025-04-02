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
            file_path = os.path.join("uploads", pdf_file.filename)
            pdf_file.save(file_path)

            text_data = extract_text_from_pdf(file_path)
            
            if not text_data.strip():
                text_data = extract_text_with_ocr(file_path)

            output_path = save_to_excel(text_data)

            return send_file(output_path, as_attachment=True)

    return render_template("index.html")

def extract_text_from_pdf(file_path):
    extracted_text = []
    try:
        with pdfplumber.open(file_path) as pdf_doc:
            for page in pdf_doc.pages:
                text = page.extract_text()
                if text:
                    extracted_text.append(text)
    except Exception as e:
        return f"Error reading PDF: {str(e)}"

    return "\n".join(extracted_text) if extracted_text else ""

def extract_text_with_ocr(file_path):
    text = []
    try:
        images = convert_from_path(file_path)
        for img in images:
            text.append(pytesseract.image_to_string(img))
    except Exception as e:
        return f"OCR failed: {str(e)}"
    
    return "\n".join(text) if text else "No text found (even with OCR)."

def save_to_excel(data):
    output_path = "converted.xlsx"
    df = pd.DataFrame({'Extracted Text': [data]})
    df.to_excel(output_path, index=False)
    return output_path

if __name__ == "__main__":
    os.makedirs("uploads", exist_ok=True)
    app.run(host="0.0.0.0", port=5000)

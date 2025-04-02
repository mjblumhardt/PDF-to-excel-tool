from flask import Flask, request, send_file, render_template
import pdfplumber
import pandas as pd
import os
import tempfile

app = Flask(__name__)

@app.route("/", methods=["GET", "POST"])
def upload_file():
    if request.method == "POST":
        pdf_file = request.files["file"]
        if pdf_file:
            text_data = extract_text_from_pdf(pdf_file)
            output_path = save_to_excel(text_data)
            return send_file(output_path, as_attachment=True)
    return render_template("index.html")

def extract_text_from_pdf(pdf):
    extracted_text = []
    with pdfplumber.open(pdf) as pdf_doc:
        for page in pdf_doc.pages:
            extracted_text.append(page.extract_text())
    return "\n".join(filter(None, extracted_text))

def save_to_excel(data):
    # Save the Excel file to a temporary directory
    temp_dir = tempfile.mkdtemp()  # Create a temporary directory
    output_path = os.path.join(temp_dir, "converted.xlsx")  # Set path for the new file
    df = pd.DataFrame({'Extracted Text': [data]})
    df.to_excel(output_path, index=False)
    return output_path

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)

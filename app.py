import pdfplumber
import pandas as pd
import os
import re
import logging
from flask import Flask, request, send_file, render_template
from pdf2image import convert_from_path
import pytesseract
import io # Needed for sending file data

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# --- Tesseract Configuration ---
# Ensure Tesseract path is correctly set for Render deployment vs local
# Check if running in Render environment
IS_RENDER = os.environ.get('RENDER', False)
if IS_RENDER:
    # Path in Render build environment (adjust if necessary based on Render buildpack)
    # Common paths: /usr/bin/tesseract or ensure apt-packages installs it system-wide
    pytesseract.pytesseract.tesseract_cmd = '/usr/bin/tesseract'
    logger.info("Running on Render. Tesseract path set to /usr/bin/tesseract")
else:
    # Example for local setup (adjust to your local Tesseract installation path)
    # Common paths: '/usr/local/bin/tesseract' (macOS/Linux),
    # 'C:\\Program Files\\Tesseract-OCR\\tesseract.exe' (Windows)
    # Let's try a common Linux/macOS path first
    local_tesseract_path = '/usr/local/bin/tesseract'
    if not os.path.exists(local_tesseract_path):
         local_tesseract_path = '/usr/bin/tesseract' # Try another common path
    # Add Windows path check if needed
    # elif os.name == 'nt':
    #     local_tesseract_path = 'C:\\Program Files\\Tesseract-OCR\\tesseract.exe'

    if os.path.exists(local_tesseract_path):
        pytesseract.pytesseract.tesseract_cmd = local_tesseract_path
        logger.info(f"Running locally. Tesseract path set to {local_tesseract_path}")
    else:
        logger.warning("Tesseract command not found at common local paths. OCR fallback may fail. Please set the correct path.")
        # Keep a placeholder or let it raise an error later if OCR is needed
        # pytesseract.pytesseract.tesseract_cmd = 'tesseract' # Or raise an error

app = Flask(__name__)

# Ensure the 'uploads' directory exists
UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

@app.route("/", methods=["GET", "POST"])
def upload_file():
    if request.method == "POST":
        if 'file' not in request.files:
            return "No file part", 400
        pdf_file = request.files["file"]
        if pdf_file.filename == '':
            return "No selected file", 400

        if pdf_file and pdf_file.filename.endswith('.pdf'):
            file_path = None
            output_path = None
            try:
                file_path = os.path.join(UPLOAD_FOLDER, pdf_file.filename)
                pdf_file.save(file_path)
                logger.info(f"PDF file saved to: {file_path}")

                # --- Primary Strategy: Table Extraction ---
                tables = extract_tables(file_path)
                logger.info(f"Extracted {len(tables)} tables using pdfplumber.")

                structured_data = process_tables(tables)
                logger.info(f"Processed {len(structured_data)} items from tables.")

                # --- Fallback Strategy: Text Extraction (if tables yield no data) ---
                if not structured_data:
                    logger.info("No data extracted from tables, attempting text extraction fallback.")
                    text_data = extract_text_with_fallback(file_path)
                    structured_data = process_text_lines_fallback(text_data) # Use a fallback text processor if needed
                    logger.info(f"Processed {len(structured_data)} items using text fallback.")

                if not structured_data:
                     logger.warning("No structured data could be extracted from the PDF.")
                     return "Could not extract relevant data from the PDF.", 400

                # --- Data Cleaning and Formatting ---
                df = clean_and_format_dataframe(structured_data)
                logger.info(f"DataFrame created with {len(df)} rows.")

                if df.empty:
                    logger.warning("Extracted data resulted in an empty DataFrame.")
                    return "Extracted data was empty or could not be processed.", 400

                # --- Excel Output ---
                # Save to an in-memory bytes buffer instead of a file
                excel_buffer = io.BytesIO()
                df.to_excel(excel_buffer, index=False, engine='openpyxl')
                excel_buffer.seek(0) # Rewind the buffer to the beginning

                output_filename = os.path.splitext(pdf_file.filename)[0] + ".xlsx"

                logger.info(f"Excel file '{output_filename}' generated successfully.")
                return send_file(
                    excel_buffer,
                    as_attachment=True,
                    download_name=output_filename,
                    mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )

            except Exception as e:
                logger.exception(f"Error processing PDF: {pdf_file.filename}") # Log full traceback
                return f"Error processing PDF: {str(e)}", 500
            finally:
                # Clean up the saved PDF file
                if file_path and os.path.exists(file_path):
                    try:
                        os.remove(file_path)
                        logger.info(f"Removed temporary file: {file_path}")
                    except OSError as e:
                         logger.error(f"Error removing temporary file {file_path}: {e}")

        else:
            return "Invalid file format. Please upload a PDF.", 400

    # For GET requests
    return render_template("index.html")

def extract_text_with_fallback(file_path):
    """Extracts text using pdfplumber, falls back to OCR if needed."""
    text_lines = []
    try:
        with pdfplumber.open(file_path) as pdf:
            # Check if PDF seems to contain actual text
            has_text = any(page.extract_text() for page in pdf.pages)

            if has_text:
                logger.info("Extracting text using pdfplumber.")
                for i, page in enumerate(pdf.pages):
                    page_text = page.extract_text()
                    if page_text:
                        text_lines.extend(page_text.split('\n'))
                    else:
                        logger.warning(f"pdfplumber found no text on page {i+1}.")
            else:
                 logger.warning("pdfplumber found no text in the PDF. Attempting OCR fallback.")
                 raise ValueError("No text found by pdfplumber") # Force OCR

    except Exception as e:
        logger.warning(f"pdfplumber extraction failed ({e}), trying OCR fallback.")
        try:
            # Ensure poppler path is handled if needed by pdf2image on Render
            poppler_path = None
            # if IS_RENDER:
            #     poppler_path = "/usr/bin" # Adjust if needed

            images = convert_from_path(file_path, dpi=300, poppler_path=poppler_path)
            logger.info(f"Converted PDF to {len(images)} images for OCR.")
            full_ocr_text = ""
            for i, img in enumerate(images):
                try:
                    # Preprocessing (optional, might help OCR quality)
                    # img = img.convert('L') # Convert to grayscale
                    ocr_text = pytesseract.image_to_string(img)
                    full_ocr_text += ocr_text + "\n"
                    logger.info(f"OCR successful for page {i+1}.")
                except pytesseract.TesseractNotFoundError:
                    logger.error("Tesseract executable not found. OCR failed. Please check installation and path.")
                    return [] # Cannot proceed with OCR
                except Exception as ocr_err:
                    logger.error(f"Error during OCR processing for page {i+1}: {ocr_err}")
            text_lines = full_ocr_text.split('\n')
        except Exception as ocr_fallback_err:
            logger.error(f"OCR fallback failed: {ocr_fallback_err}")
            return [] # Return empty if both methods fail

    # Clean up lines
    cleaned_lines = [re.sub(r'\s+', ' ', line).strip() for line in text_lines if line.strip()]
    return cleaned_lines


def extract_tables(file_path):
    """Extracts tables using pdfplumber."""
    all_tables = []
    try:
        with pdfplumber.open(file_path) as pdf:
            logger.info("Extracting tables with pdfplumber using text strategy.")
            for i, page in enumerate(pdf.pages):
                 # Try text-based strategy first, more flexible
                page_tables = page.extract_tables({
                    "vertical_strategy": "text",
                    "horizontal_strategy": "text",
                    "explicit_vertical_lines": page.curves + page.edges, # Use detected lines too
                    "explicit_horizontal_lines": page.curves + page.edges,
                    "snap_tolerance": 5, # Increase tolerance slightly
                    "join_tolerance": 5,
                    "intersection_tolerance": 5,
                })

                # Basic check for plausibility (e.g., more than 1 row, expected number of columns)
                plausible_tables = []
                for tbl in page_tables:
                    if tbl and len(tbl) > 1 and len(tbl[0]) > 3: # Example check: need header and >3 cols
                         # Further cleaning: replace None with empty string
                         cleaned_tbl = [[str(cell) if cell is not None else '' for cell in row] for row in tbl]
                         plausible_tables.append(cleaned_tbl)


                if plausible_tables:
                    logger.info(f"Found {len(plausible_tables)} plausible tables on page {i+1}.")
                    all_tables.extend(plausible_tables)
                else:
                    logger.warning(f"No plausible tables found on page {i+1} with text strategy.")
                    # Optional: Could add fallback to 'lines' strategy here if needed

    except Exception as e:
        logger.error(f"Error extracting tables: {e}")
    return all_tables

def detect_column(headers, keywords):
    """Finds the index of the first header containing any of the keywords."""
    for idx, header in enumerate(headers):
        header_lower = header.lower()
        if any(kw in header_lower for kw in keywords):
            return idx
    return None

def process_tables(tables):
    """Processes data extracted from tables."""
    items = []
    processed_row_count = 0

    # Define keywords for column detection
    col_keywords = {
        'line': ['#'],
        'product_info': ['product info', 'item'],
        'qty': ['qty', 'quantity'],
        'list_price': ['list price'],
        'disc': ['disc.', 'discount'],
        'net_unit': ['net unit'],
        'net_price': ['net price', 'extended price'] # Match 'Net Price' and 'Extended Price'
    }

    for table_index, table in enumerate(tables):
        if len(table) < 2: # Need at least header + 1 data row
            logger.warning(f"Skipping table {table_index+1} as it has less than 2 rows.")
            continue

        headers = [str(cell).strip() for cell in table[0]]
        logger.info(f"Processing Table {table_index+1} with headers: {headers}")

        # Find column indices using keywords
        col_map = {key: detect_column(headers, kw_list) for key, kw_list in col_keywords.items()}

        # Check if essential columns are found
        if col_map['product_info'] is None or col_map['net_unit'] is None or col_map['net_price'] is None:
             logger.warning(f"Skipping Table {table_index+1}: Missing essential columns (Product Info, Net Unit, or Net Price). Found map: {col_map}")
             continue


        for row_index, row in enumerate(table[1:]): # Skip header row
            # Basic check: Ensure row has enough columns
            if len(row) != len(headers):
                logger.warning(f"Skipping row {row_index+1} in table {table_index+1}: Column count mismatch ({len(row)} vs {len(headers)}). Row data: {row}")
                continue

            item = new_item() # Create a fresh item dictionary for each row

            # --- Extract Line Number ---
            if col_map['line'] is not None:
                item['Line'] = str(row[col_map['line']]).strip()
                # Basic check if line number looks like a number
                if not re.match(r'^\d+$', item['Line']):
                     logger.warning(f"Row {row_index+1}, Table {table_index+1}: Invalid Line '{item['Line']}'. Skipping row.")
                     continue # Skip rows where line number isn't just digits

            # --- Extract Product Info (Manufacturer Part No + Description) ---
            product_info_raw = str(row[col_map['product_info']]).strip()
            # Regex explanation:
            # ^[→\s]* : Matches optional arrow or whitespace at the start
            # ([A-Z0-9-]+) : Captures Group 1: One or more uppercase letters, numbers, or hyphens (Part Number)
            # \s+         : Matches one or more spaces separating Part No and Description
            # (.*)        : Captures Group 2: The rest of the string (Description)
            # re.IGNORECASE might be needed if part numbers can be lower case
            mfg_match = re.match(r'^[→\s]*([A-Z0-9][A-Z0-9-]+)\s+(.*)', product_info_raw) # Ensure part no starts alphanumeric

            if mfg_match:
                item['Manufacturer Number'] = mfg_match.group(1).strip()
                item['Description'] = mfg_match.group(2).strip()
            else:
                 # Fallback: Maybe only description or only part number?
                 # Check if it looks like *only* a part number
                 simple_mfg_match = re.match(r'^[→\s]*([A-Z0-9][A-Z0-9-]+)$', product_info_raw)
                 if simple_mfg_match:
                      item['Manufacturer Number'] = simple_mfg_match.group(1).strip()
                      item['Description'] = '' # No description found
                      logger.warning(f"Row {row_index+1}, Table {table_index+1}: Found Part No '{item['Manufacturer Number']}' but no description in '{product_info_raw}'.")
                 else:
                    # Assume it's all description if it doesn't match part number pattern
                    item['Manufacturer Number'] = ''
                    item['Description'] = product_info_raw
                    logger.warning(f"Row {row_index+1}, Table {table_index+1}: Could not extract Part No from '{product_info_raw}'. Treating as description.")
                    # Optionally skip rows where part number is critical and not found
                    # continue


            # --- Extract Other Columns ---
            item['Quantity'] = str(row[col_map['qty']]).strip() if col_map['qty'] is not None else ''
            item['List Price'] = str(row[col_map['list_price']]).strip() if col_map['list_price'] is not None else ''
            item['Discount'] = str(row[col_map['disc']]).strip() if col_map['disc'] is not None else ''
            item['Net Unit'] = str(row[col_map['net_unit']]).strip() if col_map['net_unit'] is not None else ''
            item['Net Price'] = str(row[col_map['net_price']]).strip() if col_map['net_price'] is not None else ''

            # --- Basic Validation ---
            # Check if Net Unit and Net Price seem valid before adding
            if item['Net Unit'] and item['Net Price']:
                 # Could add regex check for currency format here if needed
                 items.append(item)
                 processed_row_count += 1
            else:
                 logger.warning(f"Row {row_index+1}, Table {table_index+1}: Skipping row due to missing Net Unit ('{item['Net Unit']}') or Net Price ('{item['Net Price']}').")


    logger.info(f"Total processed rows appended: {processed_row_count}")
    return items

def process_text_lines_fallback(text_lines):
    """
    Fallback function to process raw text lines if table extraction fails.
    NOTE: This is likely less reliable for the given Ross quotes.
    Implement specific logic here if needed, otherwise return empty.
    """
    logger.warning("Executing process_text_lines_fallback - This may be less accurate.")
    # Placeholder: Return empty list as table extraction is preferred
    return []


def clean_and_format_dataframe(raw_data):
    """Creates, cleans, and formats the Pandas DataFrame."""
    if not raw_data:
        return pd.DataFrame() # Return empty DataFrame if no data

    # Define desired columns and order
    desired_columns = [
        'Line',
        'Description',
        'Manufacturer Number',
        'Manufacturer', # Added Manufacturer column
        'Quantity',
        'Net Unit',
        'Net Price'
        # Removed List Price, Discount, Notes as they weren't explicitly requested in final output
    ]

    df = pd.DataFrame(raw_data)

    # Add Manufacturer column (always 'ROSS' based on context)
    df['Manufacturer'] = 'ROSS'

    # Ensure all desired columns exist, adding missing ones with default values
    for col in desired_columns:
        if col not in df.columns:
            df[col] = '' # or pd.NA

    # Reorder columns
    df = df[desired_columns]

    # --- Data Cleaning and Type Conversion ---
    # Function to clean currency/numeric strings
    def clean_numeric(value):
        if isinstance(value, (int, float)):
            return value
        if isinstance(value, str):
            # Remove currency symbols, commas, percentage signs, handle arrows/spaces
            cleaned = re.sub(r'[$\s,%→]', '', value)
            # Handle potential empty strings after cleaning
            return cleaned if cleaned else None
        return None

    # Apply cleaning
    df['Line'] = pd.to_numeric(df['Line'], errors='coerce').astype('Int64') # Convert to nullable Integer
    df['Quantity'] = pd.to_numeric(df['Quantity'].apply(clean_numeric), errors='coerce').astype('Int64')
    df['Net Unit'] = pd.to_numeric(df['Net Unit'].apply(clean_numeric), errors='coerce')
    df['Net Price'] = pd.to_numeric(df['Net Price'].apply(clean_numeric), errors='coerce')

    # Clean description (ensure it's a string)
    df['Description'] = df['Description'].astype(str)
    df['Manufacturer Number'] = df['Manufacturer Number'].astype(str)

    # Drop rows where essential numeric fields couldn't be parsed (became NaT/NaN)
    # Keep rows even if Line or Quantity is missing, but drop if prices are missing.
    df.dropna(subset=['Net Unit', 'Net Price'], inplace=True)

    # Remove duplicate rows (optional, consider if needed)
    # df = df.drop_duplicates()

    # Sort by Line number (optional)
    df = df.sort_values(by='Line').reset_index(drop=True)

    logger.info(f"DataFrame cleaned. Final shape: {df.shape}")
    return df


def new_item():
    """Returns a dictionary template for a structured item."""
    # Simplified based on required output
    return {
        'Line': '',
        'Manufacturer Number': '',
        'Description': '',
        'Quantity': '',
        'List Price': '', # Keep for intermediate processing if needed
        'Discount': '',   # Keep for intermediate processing if needed
        'Net Unit': '',
        'Net Price': ''
    }

# --- Main Execution ---
if __name__ == "__main__":
    # Set port for local execution, fallback to 5000
    port = int(os.environ.get("PORT", 5001))
    # Use debug=True for local development ONLY, disable for production/Render
    app.run(host="0.0.0.0", port=port, debug=False)

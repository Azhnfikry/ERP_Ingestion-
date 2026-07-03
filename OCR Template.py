import io
import csv
import fitz  # PyMuPDF
import pytesseract
from PIL import Image

# --- Configuration ---
# Update this path if you are testing locally on Windows
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Define target dimensions for layout consistency (A4 @ 200 DPI standard)
TARGET_WIDTH = 1654
TARGET_HEIGHT = 2339

# Normalized Bounding Boxes: (left, top, right, bottom)
# Note: You will tune these exact pixel numbers based on your sample scan layout
TNB_ZONAL_MAP = {
    "page_1": {
        "account_no": (1200, 310, 1550, 360),
        "invoice_no": (1200, 260, 1550, 310),
        "total_amount": (1200, 520, 1550, 600),
        "tariff_type": (200, 360, 600, 410)
    },
    "page_2": {
        "meter_no": (100, 1680, 450, 1730),
        "consumption_kwh": (1100, 1680, 1300, 1730),
        # Table block coordinates to extract the line-item breakdowns
        "billing_table_zone": (100, 1500, 1550, 1650)
    }
}

def preprocess_and_normalize(image_bytes):
    """Opens an image and resizes it to a fixed layout template size."""
    img = Image.open(io.BytesIO(image_bytes))
    # Convert to grayscale to remove background tints and artifact noise
    img = img.convert('L') 
    # Resize to guarantee absolute coordinate accuracy
    img = img.resize((TARGET_WIDTH, TARGET_HEIGHT), Image.Resampling.LANCZOS)
    return img

def extract_zone(image, coordinates, config_flags="--psm 6"):
    """Crops a specific zone and performs fast localized Tesseract OCR."""
    cropped_zone = image.crop(coordinates)
    # Perform OCR only on the target slice
    text = pytesseract.image_to_string(cropped_zone, config=config_flags)
    return text.strip()

def tnb_ingestion_pipeline(page_1_bytes, page_2_bytes):
    """Main execution pipeline to ingest multi-page TNB bills."""
    
    # 1. Normalize images
    p1_img = preprocess_and_normalize(page_1_bytes)
    p2_img = preprocess_and_normalize(page_2_bytes)
    
    # 2. Extract Key-Value Metadata Pairs
    extracted_data = {
        "account_number": extract_zone(p1_img, TNB_ZONAL_MAP["page_1"]["account_no"], "--psm 7"),
        "invoice_number": extract_zone(p1_img, TNB_ZONAL_MAP["page_1"]["invoice_no"], "--psm 7"),
        "tariff":         extract_zone(p1_img, TNB_ZONAL_MAP["page_1"]["tariff_type"], "--psm 6"),
        "total_payable":  extract_zone(p1_img, TNB_ZONAL_MAP["page_1"]["total_amount"], "--psm 7"),
        "meter_number":   extract_zone(p2_img, TNB_ZONAL_MAP["page_2"]["meter_no"], "--psm 7"),
        "total_kwh":      extract_zone(p2_img, TNB_ZONAL_MAP["page_2"]["consumption_kwh"], "--psm 7"),
    }
    
    # 3. Extract the Tabular Block
    # For tabular data, PSM 6 tells Tesseract to treat it as a single uniform block
    table_raw_text = extract_zone(p2_img, TNB_ZONAL_MAP["page_2"]["billing_table_zone"], "--psm 6")
    
    # Simple line parsing to segregate tabular rows
    table_rows = []
    for line in table_raw_text.split('\n'):
        if line.strip():
            # Splitting by double space or tab depending on text column gap patterns
            columns = [col.strip() for col in line.split('  ') if col.strip()]
            table_rows.append(columns)
            
    extracted_data["billing_breakdown_table"] = table_rows
    
    return extracted_data

# --- Execution Simulation ---
if __name__ == "__main__":
    PDF_PATH = r"C:\Users\N O\Desktop\Aethera\Product Dev\Demo Documents\Demo - TNB Bills.pdf"

    # Convert PDF pages to images at 200 DPI
    pdf = fitz.open(PDF_PATH)
    def pdf_page_to_bytes(page):
        mat = fitz.Matrix(200 / 72, 200 / 72)  # 200 DPI
        pix = page.get_pixmap(matrix=mat)
        return pix.tobytes("png")

    p1 = pdf_page_to_bytes(pdf[0])
    p2 = pdf_page_to_bytes(pdf[1])

    result = tnb_ingestion_pipeline(p1, p2)

    OUTPUT_CSV = r"C:\Users\N O\Desktop\Aethera\Product Dev\Data Ingestion Feature\tnb_output.csv"

    with open(OUTPUT_CSV, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)

        # Key-value fields
        writer.writerow(["field", "value"])
        for key in ["account_number", "invoice_number", "tariff", "total_payable", "meter_number", "total_kwh"]:
            writer.writerow([key, result[key]])

        # Blank separator then billing table
        writer.writerow([])
        writer.writerow(["--- billing_breakdown_table ---"])
        for row in result["billing_breakdown_table"]:
            writer.writerow(row)

    print(f"Saved to {OUTPUT_CSV}")
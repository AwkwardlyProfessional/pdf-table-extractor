import fitz  # PyMuPDF
import pandas as pd
from openpyxl import Workbook

def extract_text_from_pdf(pdf_path):
    """Extracts text from a PDF and returns structured data."""
    doc = fitz.open(pdf_path)
    data = []  

    for i, page in enumerate(doc):
        text = page.get_text("text").strip()
        data.append([f"Page {i+1}", text])  # Each row contains page number and its text
    
    return data

def save_to_excel(data, output_excel_path):
    """Saves extracted text into an Excel file, correctly formatted."""
    df = pd.DataFrame(data, columns=["Page", "Text"])  # Structured data
    
    with pd.ExcelWriter(output_excel_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Extracted Text")
    
    print(f"Text extracted and saved correctly to {output_excel_path}")

# Input and output file paths
pdf_path = "input_pdfs/test6.pdf"  # Replace with actual PDF path
output_excel_path = "output2.xlsx"

# Extract text and save properly
data = extract_text_from_pdf(pdf_path)
save_to_excel(data, output_excel_path)
import os
import fitz  # PyMuPDF
import pandas as pd

def extract_tables_from_pdf(pdf_path, output_excel_path):
    doc = fitz.open(pdf_path)
    writer = pd.ExcelWriter(output_excel_path, engine='openpyxl')

    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        words = page.get_text("words")  # list of words with coordinates
        words.sort(key=lambda w: (w[1], w[0]))  # sort by y1 (top), then x0 (left)

        rows = []
        current_row = []
        last_y = None

        for w in words:
            x0, y0, x1, y1, word, *_ = w
            word = word.strip()
            if not word:
                continue
            if last_y is None:
                last_y = y0
            if abs(y0 - last_y) > 5:
                if current_row:
                    rows.append(current_row)
                current_row = [word]
                last_y = y0
            else:
                current_row.append(word)

        if current_row:
            rows.append(current_row)

        df = pd.DataFrame(rows)

        df.replace('', pd.NA, inplace=True)
        df.dropna(how='all', inplace=True)
        df = df.applymap(lambda x: str(x).strip() if pd.notna(x) else x)

        df = df.drop_duplicates()

        if not df.empty:
            sheet_name = f"Page_{page_num + 1}"
            df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
            print(f"Page {page_num + 1} table extracted.")
        else:
            print(f"⚠️ Page {page_num + 1} has no table content.")

    writer.close()
    print(f"All tables saved to: {output_excel_path}")

def process_all_pdfs(pdf_folder, output_folder):
    os.makedirs(output_folder, exist_ok=True)

    for filename in os.listdir(pdf_folder):
        if filename.lower().endswith(".pdf"):
            pdf_path = os.path.join(pdf_folder, filename)
            output_filename = os.path.splitext(filename)[0] + ".xlsx"
            output_excel_path = os.path.join(output_folder, output_filename)

            print(f"\n📄 Processing: {filename}")
            extract_tables_from_pdf(pdf_path, output_excel_path)

pdf_folder = "/Users/unorphaned/Desktop/pdf-table-extractor/input_pdfs"
output_folder = "excel_output"
process_all_pdfs(pdf_folder, output_folder)

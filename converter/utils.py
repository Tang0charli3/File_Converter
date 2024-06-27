import pdfplumber
import pandas as pd
import os

def pdf_to_excel(pdf_path):
    excel_path = os.path.splitext(pdf_path)[0] + '.xlsx'
    
    with pdfplumber.open(pdf_path) as pdf:
        with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
            for i, page in enumerate(pdf.pages):
                table = page.extract_table()
                if table:
                    df = pd.DataFrame(table[1:], columns=table[0])
                    sheet_name = f'Page_{i + 1}'
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
    print(f"Conversion completed: {excel_path}")

    return excel_path

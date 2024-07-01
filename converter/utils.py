import pdfplumber
import pandas as pd
from tempfile import NamedTemporaryFile
import os
import docx

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
    print(f"PDF Conversion completed: {excel_path}")

    return excel_path

def docx_to_excel(docx_path):
    excel_path = os.path.splitext(docx_path)[0] + '.xlsx'
    
    doc = docx.Document(docx_path)
    data = []

    for table in doc.tables:
        for row in table.rows:
            row_data = []
            for cell in row.cells:
                # Add cell text to row data
                row_data.append(cell.text)
            data.append(row_data)

    if data:
        # Adjust for merged cells (if necessary)
        max_cols = max(len(row) for row in data)
        for row in data:
            while len(row) < max_cols:
                row.append('')

        df = pd.DataFrame(data[1:], columns=data[0])
        with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Sheet1', index=False)
    print(f"DOCX Conversion completed: {excel_path}")
    return excel_path

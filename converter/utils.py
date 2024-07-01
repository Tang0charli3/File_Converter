import pdfplumber
import pandas as pd
from tempfile import NamedTemporaryFile
import os
import docx
from openpyxl.reader.excel import load_workbook
from openpyxl.workbook import Workbook
from pptx import Presentation
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer


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

def ppt_to_excel(ppt_path):
    excel_path = os.path.splitext(ppt_path)[0] + '.xlsx'

    # Load the PowerPoint presentation
    prs = Presentation(ppt_path)

    # Create a new Excel workbook and select the active worksheet
    wb = Workbook()
    ws = wb.active
    ws.title = "Tables"

    # Iterate over slides in the presentation
    for slide_index, slide in enumerate(prs.slides):
        # Iterate over shapes in the slide
        for shape in slide.shapes:
            if shape.has_table:
                table = shape.table
                # Get the table data and write it to the Excel sheet
                for row_index, row in enumerate(table.rows):
                    for col_index, cell in enumerate(row.cells):
                        cell_value = cell.text.strip()
                        ws.cell(row=row_index + 1, column=col_index + 1, value=cell_value)
                # Move to a new sheet for the next table
                ws = wb.create_sheet(title=f"Table_Slide_{slide_index + 1}")

    # Remove the default empty sheet if there is more than one sheet
    if len(wb.sheetnames) > 1:
        wb.remove(wb['Tables'])

    wb.save(excel_path)
    print(f"PPT Conversion completed: {excel_path}")

    return excel_path


def excel_to_pdf(excel_path, pdf_path):
    workbook = load_workbook(excel_path)
    sheet_names = workbook.sheetnames
    doc = SimpleDocTemplate(pdf_path, pagesize=letter)
    elements = []
    styles = getSampleStyleSheet()
    title_style = styles['Title']
    body_style = styles['BodyText']

    for sheet_name in sheet_names:
        sheet = workbook[sheet_name]
        df = pd.DataFrame(sheet.values)

        if not df.empty:
            # Check for merged cells and adjust the DataFrame accordingly
            for merge in sheet.merged_cells.ranges:
                merged_cell = merge.coord
                start_row, start_col, end_row, end_col = merge.min_row - 1, merge.min_col - 1, merge.max_row - 1, merge.max_col - 1
                merge_value = df.iloc[start_row, start_col]
                df.iloc[start_row:end_row + 1, start_col:end_col + 1] = None
                df.iloc[start_row, start_col] = merge_value

            elements.append(Paragraph(sheet_name, title_style))
            elements.append(Spacer(1, 12))  # Add space after the title

            data = [df.columns.tolist()] + df.fillna('').values.tolist()
            if data and len(data[0]) > 0:
                table = Table(data)

                # Setting table styles for better visual representation
                table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                    ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, 0), 12),
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                    ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                    ('GRID', (0, 0), (-1, -1), 1, colors.black),
                    ('ALIGN', (0, 1), (-1, -1), 'CENTER'),
                    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ]))

                elements.append(table)
                elements.append(Spacer(1, 24))  # Add space after the table

    if elements:
        doc.build(elements)
        print(f"PDF created: {pdf_path}")
    else:
        print("No valid data found in the Excel file to create a PDF.")

    return pdf_path
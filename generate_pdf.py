from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
import openpyxl
import os

# File paths
OUTPUT_FILE_PATH = os.path.join("output", "Workmen_Updated.xlsx")
PDF_OUTPUT_PATH = os.path.join("output", "Workmen_Updated.pdf")

def excel_to_pdf(excel_file_path, pdf_file_path):
    # Load the Excel file
    workbook = openpyxl.load_workbook(excel_file_path)
    sheet = workbook.active

    # Extract data from Excel
    data = [[cell.value if cell.value is not None else "" for cell in row] for row in sheet.iter_rows()]

    # Set up the PDF
    pdf = SimpleDocTemplate(pdf_file_path, pagesize=landscape(A4))
    elements = []

    # Create the table
    table = Table(data)

    # Add table styling
    style = TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 10),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 8),
        ("BACKGROUND", (0, 1), (-1, -1), colors.whitesmoke),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
    ])
    table.setStyle(style)

    # Append the table to the PDF
    elements.append(table)

    # Build the PDF
    pdf.build(elements)
    print(f"PDF saved to: {pdf_file_path}")

# Generate PDF
excel_to_pdf(OUTPUT_FILE_PATH, PDF_OUTPUT_PATH)

import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.platypus import Table, TableStyle, Paragraph
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
from pypdf import PdfReader, PdfWriter
import io
import os

# --- File paths ---
excel_file = "List of RCM details for Hexa.xlsx"
template_file = "Hexa RCM Self Invoices V1 (1) (2).pdf"
output_dir = "invoices"
os.makedirs(output_dir, exist_ok=True)

# --- Verify files ---
if not os.path.exists(template_file):
    print(f"ERROR: Template file '{template_file}' not found!")
    exit(1)
if not os.path.exists(excel_file):
    print(f"ERROR: Excel file '{excel_file}' not found!")
    exit(1)

# --- Load Excel Data ---
df = pd.read_excel(excel_file)
print(f"Loaded {len(df)} invoices from Excel")

# --- ReportLab styles ---
styles = getSampleStyleSheet()
styleN = styles['Normal']

def create_overlay(row):
    """Creates overlay for each invoice with consistent alignment and GST."""
    packet = io.BytesIO()
    c = canvas.Canvas(packet, pagesize=A4)

    # Base coordinates
    x_left = 50
    y_start = 700
    line_height = 16

    # --- Date & Invoice No (stacked vertically) ---
    c.setFont("Helvetica-Bold", 10)
    c.drawString(x_left, y_start, f"Date: {str(row['Inv date'])[:10]}")
    
    y_start -= line_height
    c.drawString(x_left, y_start, f"Invoice No: {row['Inv no']}")

    # --- Bill To / Ship To ---
    y_start -= (line_height * 2)
    c.setFont("Helvetica-Bold", 10)
    c.drawString(x_left, y_start, "Bill To / Ship To")

    # Company Info
    y_start -= line_height
    c.setFont("Helvetica-Bold", 9)
    c.drawString(x_left, y_start, str(row['Company Name']))

    y_start -= line_height
    c.setFont("Helvetica", 9)
    c.drawString(x_left, y_start, str(row['Company Address']))

    # --- GST Number ---
    y_start -= line_height
    c.setFont("Helvetica-Bold", 9)
    gst_number = str(row.get('Company GST Number', ''))
    if gst_number and gst_number.lower() != 'nan':
        c.drawString(x_left, y_start, f"GST No.: {gst_number}")

    # --- Place of Supply ---
    y_start -= (line_height * 2)
    c.setFont("Helvetica-Bold", 10)
    c.drawString(x_left, y_start, f"Place of Supply: {row['Party state']}")

    # --- Description Section ---
    description_text = (
        f"{row['Type of Service']}\n\n"
        f"(From\n{row['Party Name']}\nAddress: {row['City name']})"
    )
    description_paragraph = Paragraph(description_text.replace("\n", "<br/>"), styleN)

    # --- RCM section (proper line breaks) ---
    rcm_text = "RCM CGST<br/>RCM SGST<br/>RCM IGST"
    rcm_paragraph = Paragraph(rcm_text, styleN)

    # --- Table Setup ---
    table_data = [
        ['Description', 'HSN/SAC', 'Amount (INR)'],
        [description_paragraph, str(row['SAC CODE']), f"{row['Taxable value']:,.2f}"],
        ["", "", ""],
        [rcm_paragraph, "", f"{row['Total RCM payable']:,.2f}"],
        ["Total", "", f"{row['Total RCM payable'] + row['Taxable value']:,.2f}"]
    ]

    table = Table(table_data, colWidths=[210, 110, 120], rowHeights=[30, 90, 15, 45, 25])
    style = TableStyle([
        ('BOX', (0, 0), (-1, -1), 1, colors.black),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('FONTSIZE', (0, 0), (-1, 0), 11),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold')
    ])
    table.setStyle(style)

    # --- Consistent Table Position ---
    table.wrapOn(c, A4[0], A4[1])
    table.drawOn(c, 50, 280)  # fixed Y position for consistency

    c.save()
    packet.seek(0)
    return packet


# --- Generate Invoices ---
for idx, row in df.iterrows():
    invoice_no = str(row.get('Inv no', 'unknown')).replace('/', '-').replace('\\', '-')
    overlay_pdf_buffer = create_overlay(row)

    template_pdf = PdfReader(open(template_file, "rb"))
    overlay_pdf = PdfReader(overlay_pdf_buffer)
    
    writer = PdfWriter()
    page = template_pdf.pages[0]
    page.merge_page(overlay_pdf.pages[0])
    writer.add_page(page)

    output_path = os.path.join(output_dir, f"Invoice_{invoice_no}.pdf")
    with open(output_path, "wb") as output_stream:
        writer.write(output_stream)

    print(f"✅ Generated: {output_path}")

print(f"\n✓ Successfully generated {len(df)} invoices in '{output_dir}' directory.")

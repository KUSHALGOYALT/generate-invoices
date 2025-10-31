from flask import Flask, render_template, request, send_file, after_this_request, jsonify
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.platypus import Table, TableStyle, Paragraph
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
from pypdf import PdfReader, PdfWriter
import io
import os
import tempfile
import shutil
from datetime import datetime

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = tempfile.mkdtemp()
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

# Setup paragraph style for wrapping
styles = getSampleStyleSheet()
styleN = styles['Normal']


def create_overlay(row):
    """Creates overlay for each invoice with consistent alignment."""
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


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/generate', methods=['POST'])
def generate():
    try:
        if 'excel_file' not in request.files or 'template_file' not in request.files:
            return jsonify({'error': 'Both Excel and Template files are required'}), 400
        
        excel_file = request.files['excel_file']
        template_file = request.files['template_file']
        
        if excel_file.filename == '' or template_file.filename == '':
            return jsonify({'error': 'No file selected'}), 400
        
        # Create temporary directory for this session
        session_dir = tempfile.mkdtemp(dir=app.config['UPLOAD_FOLDER'])
        temp_excel = os.path.join(session_dir, 'excel.xlsx')
        temp_template = os.path.join(session_dir, 'template.pdf')
        
        excel_file.save(temp_excel)
        template_file.save(temp_template)
        
        # Read Excel data
        df = pd.read_excel(temp_excel)
        
        # Create output directory
        output_dir = os.path.join(session_dir, 'invoices')
        os.makedirs(output_dir, exist_ok=True)
        
        # Generate invoices
        generated_files = []
        for idx, row in df.iterrows():
            invoice_no = str(row.get('Inv no', 'unknown')).replace('/', '-').replace('\\', '-')
            overlay_pdf_buffer = create_overlay(row)

            template_pdf = PdfReader(open(temp_template, "rb"))
            overlay_pdf = PdfReader(overlay_pdf_buffer)
            
            writer = PdfWriter()
            page = template_pdf.pages[0]
            page.merge_page(overlay_pdf.pages[0])
            writer.add_page(page)

            output_path = os.path.join(output_dir, f"Invoice_{invoice_no}.pdf")
            with open(output_path, "wb") as output_stream:
                writer.write(output_stream)
            
            generated_files.append(output_path)
        
        # Create a zip file of all invoices
        import zipfile
        zip_filename = os.path.join(session_dir, 'invoices.zip')
        with zipfile.ZipFile(zip_filename, 'w') as zipf:
            for file_path in generated_files:
                zipf.write(file_path, os.path.basename(file_path))
        
        # Schedule cleanup
        @after_this_request
        def cleanup(response):
            try:
                shutil.rmtree(session_dir)
            except Exception:
                pass
            return response
        
        return send_file(zip_filename, as_attachment=True, download_name=f'invoices_{datetime.now().strftime("%Y%m%d_%H%M%S")}.zip')
    
    except Exception as e:
        return jsonify({'error': str(e)}), 500


if __name__ == '__main__':
    app.run(debug=True, port=5001)


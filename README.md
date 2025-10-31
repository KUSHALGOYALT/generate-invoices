# Invoice Generator Web Application

A Flask-based web application for automatically generating invoice PDFs from Excel data.

## Features

- ✅ Upload Excel file with invoice data
- ✅ Upload invoice template PDF
- ✅ Generate multiple PDF invoices automatically
- ✅ Download all invoices as a ZIP file
- ✅ Modern, responsive web interface
- ✅ Proper layering with pypdf (text on top of template)

## Installation

### 1. Install Dependencies

```bash
pip install -r requirements.txt
```

### 2. Run the Application

```bash
python3 app.py
```

The application will start on http://localhost:5001

### 3. Open in Browser

Navigate to http://localhost:5001 in your web browser.

## Usage

1. Upload your Excel file containing invoice data (must have columns: Inv date, Inv no, Company Name, Company Address, etc.)
2. Upload your template PDF file
3. Click "Generate Invoices"
4. Download the ZIP file containing all generated invoices

## Excel File Format

Your Excel file should contain the following columns:

- `Inv date` - Invoice date
- `Inv no` - Invoice number
- `Company Name` - Company name
- `Company Address` - Company address
- `Company GST Number` - GST number (optional)
- `Party Name` - Party name
- `City name` - City/address
- `Type of Service` - Type of service
- `SAC CODE` - SAC/HSN code
- `Party state` - Place of supply
- `Taxable value` - Taxable amount
- `Total RCM payable` - Total RCM amount

## API

### POST /generate

Generates invoices and returns a ZIP file.

**Form Data:**
- `excel_file`: Excel file with invoice data (.xlsx, .xls)
- `template_file`: PDF template file (.pdf)

**Response:**
- ZIP file containing all generated invoice PDFs

## Technologies Used

- Flask - Web framework
- pandas - Excel data processing
- ReportLab - PDF generation
- pypdf - PDF merging and manipulation
- OpenPyXL - Excel file reading

## License

MIT

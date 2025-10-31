# How to Use the Invoice Generator Web Application

## Starting the Application

1. Open a terminal in the project directory
2. Run one of these commands:
   ```bash
   # Option 1: Using Python directly
   python3 app.py
   
   # Option 2: Using the startup script
   ./run.sh
   ```

3. The application will start on **http://localhost:5001**

## Using the Web Interface

1. Open your web browser and go to **http://localhost:5001**

2. You'll see a modern interface with two file upload fields:
   - **Excel File**: Upload your invoice data file (`List of RCM details for Hexa.xlsx`)
   - **Template PDF**: Upload your template PDF file (`Hexa RCM Self Invoices V1 (1) (2).pdf`)

3. Click **"Generate Invoices"**

4. Wait for the processing to complete (you'll see a loading spinner)

5. The browser will automatically download a ZIP file containing all generated invoice PDFs

## Example Workflow

```bash
# 1. Start the server
python3 app.py

# 2. Open browser to http://localhost:5001

# 3. Upload:
#    - Excel file: List of RCM details for Hexa.xlsx
#    - Template PDF: Hexa RCM Self Invoices V1 (1) (2).pdf

# 4. Click "Generate Invoices"

# 5. Download the invoices_YYYYMMDD_HHMMSS.zip file
```

## Troubleshooting

### Port Already in Use
If port 5001 is in use, the app will automatically try a different port. Check the terminal output for the actual URL.

### Missing Dependencies
Run: `pip install -r requirements.txt`

### Files Not Uploading
Make sure your files are:
- Excel files are in .xlsx or .xls format
- PDF template is in .pdf format
- Files are not corrupted

## Stopping the Server

Press `Ctrl+C` in the terminal where the server is running.


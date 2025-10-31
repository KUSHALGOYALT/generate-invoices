#!/bin/bash
# Run the Invoice Generator Web Application

echo "Starting Invoice Generator Web Application..."
echo ""

# Activate virtual environment if it exists
if [ -d "venv" ]; then
    echo "Activating virtual environment..."
    source venv/bin/activate
fi

# Install/check dependencies
echo "Checking dependencies..."
pip install -q -r requirements.txt

# Run the application
echo "Starting server on http://localhost:5001..."
echo "Press Ctrl+C to stop the server"
echo ""
python3 app.py


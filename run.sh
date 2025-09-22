#!/bin/bash

# Superdrug ITG Invoice Generator - Startup Script

echo "=================================="
echo "Superdrug ITG Invoice Generator v3.0"
echo "=================================="

# Check if virtual environment exists
if [ ! -d "venv" ]; then
    echo "Creating virtual environment..."
    python3 -m venv venv
fi

# Activate virtual environment
echo "Activating virtual environment..."
source venv/bin/activate

# Check if dependencies are installed
echo "Checking dependencies..."
pip show streamlit > /dev/null 2>&1
if [ $? -ne 0 ]; then
    echo "Installing dependencies..."
    pip install -r requirements.txt
fi

# Run the application
echo "Starting application..."
echo "Opening browser at http://localhost:8501"
streamlit run invoice_app_v3.py

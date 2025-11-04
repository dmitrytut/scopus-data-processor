#!/bin/bash

echo "========================================"
echo " Scopus Data Processor"
echo " Khazar University"
echo "========================================"
echo ""

# Check if Python is installed
if ! command -v python3 &> /dev/null; then
    echo "ERROR: Python 3 is not installed!"
    echo "Please install Python 3.9+ from https://www.python.org/downloads/"
    echo ""
    read -p "Press Enter to exit..."
    exit 1
fi

echo "Python found: $(python3 --version)"
echo ""

# Check if dependencies are installed
echo "Checking dependencies..."
if ! python3 -c "import streamlit" &> /dev/null; then
    echo "Streamlit not found. Installing dependencies..."
    echo "This may take a few minutes on first run..."
    echo ""
    pip3 install -r requirements.txt
    echo ""
    echo "Dependencies installed successfully!"
    echo ""
else
    echo "Dependencies OK!"
    echo ""
fi

# Run the Streamlit app
echo "Starting application..."
echo "The app will open in your browser automatically."
echo ""
echo "To stop the application, press Ctrl+C in this terminal"
echo ""

streamlit run app.py

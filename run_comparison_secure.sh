#!/bin/bash

echo "========================================"
echo "Excel File Comparison Tool"
echo "========================================"
echo
echo "This tool requires a 4-digit password for security."
echo

# Navigate to the script directory
cd "$(dirname "$0")"

# Check if virtual environment exists
if [ ! -d ".venv" ]; then
    echo "❌ Virtual environment not found!"
    echo "Please run the setup script first."
    exit 1
fi

# Activate virtual environment and run Python script
source .venv/bin/activate
python test.py

echo
echo "Press Enter to exit..."
read

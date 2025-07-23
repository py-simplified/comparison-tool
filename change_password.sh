#!/bin/bash

echo "========================================"
echo "Password Change Utility"
echo "========================================"
echo
echo "Use this utility to change the 4-digit password."
echo "You will need to know the current password."
echo

# Navigate to the script directory
cd "$(dirname "$0")"

# Check if virtual environment exists
if [ ! -d ".venv" ]; then
    echo "❌ Virtual environment not found!"
    echo "Please run the setup script first."
    exit 1
fi

# Activate virtual environment and run password change utility
source .venv/bin/activate
python test.py --change-password

echo
echo "Press Enter to exit..."
read

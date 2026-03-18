#!/bin/bash

echo "========================================"
echo "Employee Timesheet Processor"
echo "Installing Dependencies..."
echo "========================================"
echo ""

# Check if Python is installed
echo "Checking Python installation..."

if command -v python3 &> /dev/null; then
    PYTHON_CMD="python3"
    PIP_CMD="pip3"
    echo "✓ Python 3 found: $(python3 --version)"
elif command -v python &> /dev/null; then
    PYTHON_CMD="python"
    PIP_CMD="pip"
    echo "✓ Python found: $(python --version)"
else
    echo "✗ ERROR: Python is not installed"
    echo ""
    echo "Please install Python from: https://www.python.org/downloads/"
    echo ""
    echo "Installation steps:"
    echo "1. Go to https://www.python.org/downloads/"
    echo "2. Download Python 3.x for macOS"
    echo "3. Run the installer"
    echo "4. Run this script again"
    echo ""
    read -p "Press Enter to exit..."
    exit 1
fi

echo ""
echo "Installing required packages..."
echo ""

# Install pandas and openpyxl
echo "Installing pandas..."
$PIP_CMD install pandas

if [ $? -ne 0 ]; then
    echo ""
    echo "✗ ERROR: Failed to install pandas"
    echo ""
    echo "Possible solutions:"
    echo "1. Try running with sudo: sudo $PIP_CMD install pandas openpyxl"
    echo "2. Use --user flag: $PIP_CMD install --user pandas openpyxl"
    echo "3. Check your internet connection"
    echo ""
    read -p "Press Enter to exit..."
    exit 1
fi

echo ""
echo "Installing openpyxl..."
$PIP_CMD install openpyxl

if [ $? -ne 0 ]; then
    echo ""
    echo "✗ ERROR: Failed to install openpyxl"
    echo ""
    echo "Possible solutions:"
    echo "1. Try running with sudo: sudo $PIP_CMD install pandas openpyxl"
    echo "2. Use --user flag: $PIP_CMD install --user pandas openpyxl"
    echo ""
    read -p "Press Enter to exit..."
    exit 1
fi

echo ""
echo "========================================"
echo "SUCCESS! All dependencies installed"
echo "========================================"
echo ""

# Verify installation
echo "Verifying installation..."
$PYTHON_CMD -c "import pandas; import openpyxl; print('✓ pandas and openpyxl are ready to use!')"

if [ $? -eq 0 ]; then
    echo ""
    echo "You can now run the timesheet processor!"
    echo ""
    echo "Usage:"
    echo "  Drag and drop your Excel file onto process_timesheet_mac.command"
    echo ""
else
    echo ""
    echo "Warning: Verification failed"
    echo "Dependencies may not be properly installed"
    echo ""
fi

read -p "Press Enter to exit..."

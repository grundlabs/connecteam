#!/bin/bash

# Employee Timesheet Processor - macOS Version
# This script processes timesheet Excel files

echo "========================================"
echo "Employee Timesheet Processor"
echo "========================================"
echo ""

# Get the directory where this script is located
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"

# Check if a file was provided (drag and drop)
if [ -z "$1" ]; then
    echo "ERROR: No file specified"
    echo "Drag and drop your Excel file onto this script"
    echo ""
    read -p "Press Enter to exit..."
    exit 1
fi

# Check if file exists
INPUT_FILE="$1"
if [ ! -f "$INPUT_FILE" ]; then
    echo "ERROR: File not found: $INPUT_FILE"
    echo ""
    read -p "Press Enter to exit..."
    exit 1
fi

echo "Processing: $(basename "$INPUT_FILE")"
echo ""

# Temporary files
TEMP_FILE="$SCRIPT_DIR/_temp.xlsx"
LOG_FILE="$SCRIPT_DIR/_process.log"

# Run Python script
if command -v python3 &> /dev/null; then
    PYTHON_CMD="python3"
elif command -v python &> /dev/null; then
    PYTHON_CMD="python"
else
    echo "ERROR: Python not found"
    echo "Please install Python 3 from https://www.python.org/downloads/"
    echo ""
    read -p "Press Enter to exit..."
    exit 1
fi

echo "Running processor..."
echo ""

# Execute Python script and capture output
$PYTHON_CMD "$SCRIPT_DIR/process_timesheets.py" "$INPUT_FILE" "$TEMP_FILE" > "$LOG_FILE" 2>&1

if [ $? -ne 0 ]; then
    echo "ERROR: Processing failed"
    echo ""
    cat "$LOG_FILE"
    rm -f "$LOG_FILE" 2>/dev/null
    echo ""
    read -p "Press Enter to exit..."
    exit 1
fi

# Display log
cat "$LOG_FILE"
echo ""

# Extract date from log
FILTER_DATE=$(grep "Filtering for date:" "$LOG_FILE" | awk '{print $4}')

# Determine output filename
if [ -n "$FILTER_DATE" ]; then
    OUTPUT_NAME="rekordok_${FILTER_DATE}.xlsx"
else
    echo "Warning: Could not extract date from log"
    OUTPUT_NAME="rekordok_$(date +%Y-%m-%d).xlsx"
fi

echo "Output filename: $OUTPUT_NAME"

# Move temp file to final output
OUTPUT_FILE="$SCRIPT_DIR/$OUTPUT_NAME"

if [ -f "$TEMP_FILE" ]; then
    mv "$TEMP_FILE" "$OUTPUT_FILE"
    
    if [ -f "$OUTPUT_FILE" ]; then
        echo ""
        echo "========================================"
        echo "SUCCESS: $OUTPUT_NAME"
        echo "========================================"
        echo "Location: $SCRIPT_DIR"
        
        # Open the output file in Excel (macOS)
        open "$OUTPUT_FILE"
    else
        echo "ERROR: Failed to create output file"
    fi
else
    echo "ERROR: Temp file not created"
fi

# Cleanup
rm -f "$LOG_FILE" 2>/dev/null

echo ""
read -p "Press Enter to exit..."

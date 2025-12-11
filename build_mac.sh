#!/bin/bash
# Mac-specific build script for sec_financials_tool
# This script builds the executable and handles Mac-specific setup

echo "Building SEC Financials Tool for macOS..."
echo ""

# Check if PyInstaller is installed
if ! command -v pyinstaller &> /dev/null; then
    echo "PyInstaller not found. Installing..."
    pip install pyinstaller
fi

# Build the executable
python build_exe.py

# Remove quarantine attributes (allows running without right-click -> Open)
if [ -f "dist/sec_financials_tool" ]; then
    echo ""
    echo "Removing quarantine attributes..."
    xattr -cr dist/sec_financials_tool
    echo "✓ Quarantine attributes removed"
    
    # Make executable (if not already)
    chmod +x dist/sec_financials_tool
    
    echo ""
    echo "Build complete! Executable: dist/sec_financials_tool"
    echo ""
    echo "To test, run:"
    echo "  ./dist/sec_financials_tool TSLA"
    echo ""
    echo "Note: If you plan to distribute this, consider code signing:"
    echo "  codesign --force --deep --sign \"Developer ID Application: Your Name\" dist/sec_financials_tool"
else
    echo "Error: Executable not found in dist/"
    exit 1
fi


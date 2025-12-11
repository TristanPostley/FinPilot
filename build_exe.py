#!/usr/bin/env python3
"""
Build script to create an executable from sec_financials_tool.py
Supports both Windows and macOS.
Run this script to build the executable: python build_exe.py
"""

import PyInstaller.__main__
import sys
import os
import platform

# Detect the platform
system = platform.system()
is_mac = system == 'Darwin'
is_windows = system == 'Windows'

# Determine executable extension and name
if is_windows:
    exe_ext = '.exe'
    exe_name = 'sec_financials_tool.exe'
elif is_mac:
    exe_ext = ''
    exe_name = 'sec_financials_tool'
else:
    exe_ext = ''
    exe_name = 'sec_financials_tool'

# PyInstaller arguments
args = [
    'sec_financials_tool.py',  # Main script
    '--name=sec_financials_tool',  # Name of the executable
    '--onefile',  # Create a single executable file
    '--console',  # Console application (shows command window)
    '--clean',  # Clean PyInstaller cache before building
    '--noconfirm',  # Overwrite output directory without asking
    # Include hidden imports that might be needed
    '--hidden-import=pandas',
    '--hidden-import=openpyxl',
    '--hidden-import=edgar',
    '--hidden-import=edgartools',
    # Collect all submodules
    '--collect-all=edgar',
    '--collect-all=edgartools',
]

# Mac-specific options
if is_mac:
    # On Mac, you might want to add code signing (optional)
    # Uncomment and set your Developer ID if you have one:
    # args.extend(['--codesign-identity', 'Developer ID Application: Your Name'])
    pass

# Run PyInstaller
PyInstaller.__main__.run(args)

print("\n" + "="*60)
print("Build complete!")
print("="*60)
print(f"Platform: {system}")
print(f"Executable location: dist/{exe_name}")
print("\nYou can now distribute the executable from the 'dist' folder.")
print("Note: The executable is standalone and includes all dependencies.")

if is_mac:
    print("\nMac-specific notes:")
    print("- The executable may need to be signed for distribution")
    print("- Users may need to right-click and select 'Open' the first time")
    print("- Or run: xattr -cr dist/sec_financials_tool")
    print("- To test: ./dist/sec_financials_tool TSLA")


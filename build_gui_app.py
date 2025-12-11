#!/usr/bin/env python3
"""
Build script to create a Mac .app bundle with GUI
Also works for Windows to create a windowed executable
"""

import PyInstaller.__main__
import sys
import os
import platform

# Detect the platform
system = platform.system()
is_mac = system == 'Darwin'
is_windows = system == 'Windows'

if not os.path.exists('sec_financials_gui.py'):
    print("Error: sec_financials_gui.py not found!")
    print("Please make sure the GUI file exists.")
    sys.exit(1)

# Determine app name and options based on platform
if is_mac:
    app_name = 'SEC Financials Tool'
    windowed_flag = '--windowed'
    bundle_id = 'com.secfinancials.tool'
    print("Building Mac .app bundle...")
elif is_windows:
    app_name = 'SEC Financials Tool'
    windowed_flag = '--windowed'
    bundle_id = None
    print("Building Windows executable...")
else:
    app_name = 'sec_financials_tool'
    windowed_flag = '--noconsole'
    bundle_id = None
    print("Building executable for", system)

# PyInstaller arguments
args = [
    'sec_financials_gui.py',  # GUI script
    f'--name={app_name}',  # App name
    windowed_flag,  # No console window (GUI app)
    '--onefile',  # Single file
    '--clean',  # Clean cache
    '--noconfirm',  # Overwrite without asking
    # Include hidden imports
    '--hidden-import=pandas',
    '--hidden-import=openpyxl',
    '--hidden-import=edgar',
    '--hidden-import=edgartools',
    '--hidden-import=tkinter',
    # Collect all submodules
    '--collect-all=edgar',
    '--collect-all=edgartools',
]

# Mac-specific options
if is_mac:
    args.extend([
        '--osx-bundle-identifier', bundle_id,
    ])

# Windows-specific options
if is_windows:
    # Windows doesn't need bundle identifier
    pass

# Run PyInstaller
try:
    PyInstaller.__main__.run(args)
    
    print("\n" + "="*60)
    print("Build complete!")
    print("="*60)
    
    if is_mac:
        print(f"App bundle: dist/{app_name}.app")
        print("\nUsers can double-click the .app to run it!")
        print("\nNote: First-time users may need to:")
        print("  1. Right-click the .app")
        print("  2. Select 'Open'")
        print("  3. Click 'Open' in the security dialog")
    elif is_windows:
        print(f"Executable: dist/{app_name}.exe")
        print("\nUsers can double-click the .exe to run it!")
    else:
        print(f"Executable: dist/{app_name}")
        print("\nUsers can run the executable to launch the GUI!")
        
except Exception as e:
    print(f"\nError during build: {e}")
    sys.exit(1)


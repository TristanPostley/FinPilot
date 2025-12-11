# Building an Executable

This guide explains how to package the SEC Financials Tool as a standalone executable for Windows and macOS.

## Prerequisites

1. Install PyInstaller:
```bash
pip install pyinstaller
```

2. Ensure all dependencies are installed:
```bash
pip install -r requirements.txt
```

## Method 1: Using the Build Script (Recommended)

Simply run the build script:

```bash
python build_exe.py
```

This will create a single executable file in the `dist` folder.

## Method 2: Using the Spec File

For more control over the build process:

```bash
pyinstaller build_exe.spec
```

## Method 3: Direct PyInstaller Command

You can also run PyInstaller directly:

```bash
pyinstaller --onefile --console --name sec_financials_tool sec_financials_tool.py
```

## Build Options

### Single File vs. Directory

- `--onefile`: Creates a single `.exe` file (larger, slower startup)
- Without `--onefile`: Creates a folder with multiple files (faster startup)

### Console vs. Windowed

- `--console`: Shows command window (current setting)
- `--windowed` or `--noconsole`: No console window (GUI apps)

### Adding an Icon

To add a custom icon to your executable:

**Windows:**
1. Create or download a `.ico` file
2. Add `--icon=icon.ico` to the PyInstaller command
3. Or edit `build_exe.spec` and set `icon='icon.ico'`

**macOS:**
1. Create or download a `.icns` file (or convert from PNG)
2. Add `--icon=icon.icns` to the PyInstaller command
3. Or edit `build_exe.spec` and set `icon='icon.icns'`

**Note:** To convert PNG to ICNS on Mac:
```bash
# Create iconset directory
mkdir icon.iconset
# Convert PNG to various sizes (requires sips command)
sips -z 16 16 icon.png --out icon.iconset/icon_16x16.png
sips -z 32 32 icon.png --out icon.iconset/icon_16x16@2x.png
# ... (add more sizes)
# Create .icns file
iconutil -c icns icon.iconset
```

## Output

After building, you'll find:
- **Windows**: `dist/sec_financials_tool.exe` - The executable file
- **macOS/Linux**: `dist/sec_financials_tool` - The executable file (no extension)
- `build/` - Temporary build files (can be deleted)
- `sec_financials_tool.spec` - Auto-generated spec file (if using direct command)

## Distribution

The executable in `dist/` is standalone and can be distributed to other machines without requiring Python installation.

### Windows Distribution

Simply distribute the `.exe` file. Users can run it directly.

### macOS Distribution

On macOS, there are a few additional considerations:

1. **Gatekeeper**: macOS may block unsigned executables. Users can:
   - Right-click the executable and select "Open" (first time only)
   - Or run: `xattr -cr dist/sec_financials_tool` to remove quarantine attributes

2. **Code Signing** (Optional, for distribution):
   - Requires an Apple Developer account ($99/year)
   - Sign with: `codesign --force --deep --sign "Developer ID Application: Your Name" dist/sec_financials_tool`
   - Or add to spec file: `codesign_identity='Developer ID Application: Your Name'`

3. **Notarization** (Optional, for App Store or wider distribution):
   - Required for macOS 10.15+ for unsigned apps
   - Submit to Apple: `xcrun notarytool submit dist/sec_financials_tool --keychain-profile "AC_PASSWORD" --wait`

4. **Testing on Mac**:
   ```bash
   ./dist/sec_financials_tool TSLA
   ```

### File Size

The executable will be relatively large (typically 50-100 MB) because it includes:
- Python interpreter
- All dependencies (pandas, openpyxl, edgartools, etc.)
- Required libraries

### Testing

Before distributing, test the executable:

**Windows:**
```bash
dist\sec_financials_tool.exe TSLA
```

**macOS/Linux:**
```bash
./dist/sec_financials_tool TSLA
```

## Troubleshooting

### Missing Modules

If you get import errors, add hidden imports to the spec file or build command:

```bash
--hidden-import=module_name
```

### Antivirus Warnings

Some antivirus software may flag PyInstaller executables as suspicious. This is a false positive. You can:
1. Sign the executable with a code signing certificate
2. Submit to antivirus vendors for whitelisting
3. Use `--key` option to encrypt the bytecode (requires PyInstaller 4.0+)

### Large File Size

To reduce file size:
- Use `--exclude-module` to exclude unused modules
- Consider using `--onedir` instead of `--onefile` (faster startup, but multiple files)

## Platform-Specific Notes

### macOS

- **Architecture**: By default, builds for your Mac's architecture (Intel or Apple Silicon)
- **Universal Binary**: To create a universal binary for both architectures:
  ```bash
  # Build on Apple Silicon Mac with Rosetta 2
  arch -x86_64 python build_exe.py  # For Intel
  # Then merge or build separately for each arch
  ```
- **Permissions**: May need Terminal/Full Disk Access permissions for file operations

### Windows

- **Antivirus**: Some antivirus software may flag PyInstaller executables (false positive)
- **Windows Defender**: May need to add exception for the executable

## Building the GUI Version (Recommended for Non-Technical Users)

A GUI version is available that provides a user-friendly interface - perfect for non-technical users!

### Building the GUI App

**For Mac (.app bundle):**
```bash
python build_gui_app.py
```

**For Windows (.exe):**
```bash
python build_gui_app.py
```

**Using the spec file:**
```bash
pyinstaller build_gui_app.spec
```

### GUI Features

- Simple form-based interface
- No command-line knowledge required
- Progress indicators
- Automatic file location opening
- Error messages with helpful suggestions

### Output

- **Mac**: `dist/SEC Financials Tool.app` - Double-clickable app bundle
- **Windows**: `dist/SEC Financials Tool.exe` - Double-clickable executable

### Mac .app Bundle Notes

The GUI version creates a proper Mac `.app` bundle that:
- Can be double-clicked to launch
- Appears in Applications folder
- Has a proper bundle identifier
- May require right-click → Open on first launch (macOS security)

See [USER_INSTRUCTIONS.md](USER_INSTRUCTIONS.md) for end-user instructions.

## Advanced: Command-Line Version

The command-line version is still available for technical users or automation:
1. Use `build_exe.py` to build the console version
2. Users run it from Terminal/Command Prompt
3. See the main README for command-line usage


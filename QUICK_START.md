# Quick Start Guide - Building for Distribution

This is a quick reference for building executables to distribute to end users.

## For Non-Technical Users (Recommended)

Build the GUI version - users can double-click to run:

```bash
# Install PyInstaller if needed
pip install pyinstaller

# Build GUI app
python build_gui_app.py
```

**Result:**
- Mac: `dist/SEC Financials Tool.app` 
- Windows: `dist/SEC Financials Tool.exe`

**Give users:** The `.app` (Mac) or `.exe` (Windows) file + [USER_INSTRUCTIONS.md](USER_INSTRUCTIONS.md)

## For Technical Users

Build the command-line version:

```bash
pip install pyinstaller
python build_exe.py
```

**Result:**
- Mac: `dist/sec_financials_tool`
- Windows: `dist/sec_financials_tool.exe`

## Testing Before Distribution

1. **Test the executable** on a clean machine (or VM) without Python installed
2. **Test with different tickers**: TSLA, AAPL, MSFT
3. **Test error handling**: Invalid ticker, network issues
4. **Verify output**: Check that Excel files are created correctly

## Distribution Checklist

- [ ] Built executable/app bundle
- [ ] Tested on clean system
- [ ] Included USER_INSTRUCTIONS.md
- [ ] (Mac) Tested first-launch Gatekeeper behavior
- [ ] (Optional) Added custom icon
- [ ] (Optional) Code signed (for wider distribution)

## File Sizes

Expect these approximate sizes:
- GUI version: 80-120 MB (includes tkinter)
- CLI version: 50-100 MB

These are normal for PyInstaller executables that bundle Python and all dependencies.


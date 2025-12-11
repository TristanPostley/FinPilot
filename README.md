# SEC Financials Tool

A Python tool that fetches company financial data from the SEC EDGAR API and creates Excel files similar to the Tesla FY 2024 Financials format.

## Features

- Fetches financial statements (Income Statement, Balance Sheet, Cash Flow Statement) from SEC EDGAR
- Generates Excel files with multiple sheets
- Supports any publicly traded company with SEC filings
- Configurable fiscal year and form type

## Installation

1. Install Python dependencies:

```bash
pip install -r requirements.txt
```

## Usage

### Basic Usage

Fetch the latest 10-K filing for a company:

```bash
python sec_financials_tool.py TSLA
```

### Advanced Usage

Fetch a specific year:

```bash
python sec_financials_tool.py AAPL --year 2023
```

Specify output file:

```bash
python sec_financials_tool.py MSFT --output Microsoft-Financials.xlsx
```

Set email for SEC API (or set `SEC_API_EMAIL` environment variable):

```bash
python sec_financials_tool.py GOOGL --email your.email@example.com
```

### Command Line Options

- `ticker` (required): Company stock ticker symbol (e.g., TSLA, AAPL, MSFT)
- `--year`: Fiscal year (default: latest available)
- `--output`, `-o`: Output Excel file path (default: auto-generated)
- `--form`: Form type (default: 10-K, can use 10-Q for quarterly)
- `--email`: Email for SEC API identification (or set `SEC_API_EMAIL` env var)

## Examples

```bash
# Get Tesla's latest annual report
python sec_financials_tool.py TSLA

# Get Apple's 2023 financials
python sec_financials_tool.py AAPL --year 2023

# Get Microsoft's quarterly report (10-Q)
python sec_financials_tool.py MSFT --form 10-Q

# Custom output filename
python sec_financials_tool.py GOOGL --output Google-2024-Financials.xlsx
```

## Output

The tool generates an Excel file with three sheets:
1. **Income Statement** - Revenue, expenses, and net income
2. **Balance Sheet** - Assets, liabilities, and equity
3. **Cash Flow Statement** - Operating, investing, and financing cash flows

The filename is auto-generated as: `{Company-Name}-FY-{Year}-Financials.xlsx`

## Requirements

- Python 3.8+
- Internet connection (to access SEC EDGAR API)
- Valid email address (for SEC API identification)

## Notes

- The SEC requires identification when accessing their API. You can provide your email via the `--email` flag or set the `SEC_API_EMAIL` environment variable.
- Some companies may have different fiscal year ends, so the "FY" year in the filename corresponds to the filing date year.
- The tool uses the `edgartools` library which provides a convenient interface to SEC EDGAR data.

## Building an Executable

To package this tool as a standalone executable for Windows or macOS, see [BUILD_INSTRUCTIONS.md](BUILD_INSTRUCTIONS.md) for detailed instructions.

### GUI Version (Recommended for Non-Technical Users)

The GUI version provides a user-friendly interface - perfect for end users who aren't comfortable with command-line tools.

**Build the GUI app:**
```bash
pip install pyinstaller
python build_gui_app.py
```

**Output:**
- **Mac**: `dist/SEC Financials Tool.app` - Double-clickable app bundle
- **Windows**: `dist/SEC Financials Tool.exe` - Double-clickable executable

See [USER_INSTRUCTIONS.md](USER_INSTRUCTIONS.md) for end-user instructions.

### Command-Line Version

For technical users or automation:

**Windows:**
```bash
pip install pyinstaller
python build_exe.py
```

**macOS:**
```bash
pip install pyinstaller
python build_exe.py
```

**Output:**
- **Windows**: `dist/sec_financials_tool.exe`
- **macOS**: `dist/sec_financials_tool`

The executables can be distributed without requiring Python installation.

## Troubleshooting

If you encounter errors:

1. **No filings found**: The company may not have filed the requested form type, or the ticker symbol may be incorrect.
2. **API rate limiting**: The SEC API may rate limit requests. Wait a few moments and try again.
3. **Missing data**: Some companies may have incomplete financial statements in their filings.


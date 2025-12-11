# SEC Financials Tool - User Instructions

Welcome! This guide will help you use the SEC Financials Tool to generate Excel reports from company financial data.

## What This Tool Does

The SEC Financials Tool fetches financial statements (Income Statement, Balance Sheet, and Cash Flow Statement) from the SEC EDGAR database and creates a formatted Excel file for any publicly traded company.

## First Time Setup (Mac Users)

If you're using a Mac, you may need to allow the app to run the first time:

1. **Locate the app**: Find "SEC Financials Tool.app" in your Downloads or wherever you saved it

2. **First launch**:
   - **Right-click** (or Control+click) on "SEC Financials Tool.app"
   - Select **"Open"** from the menu
   - Click **"Open"** in the security dialog that appears
   - (You only need to do this once - after that, you can double-click normally)

3. **If you see "App can't be opened"**:
   - Go to System Settings → Privacy & Security
   - Scroll down to find a message about the app being blocked
   - Click "Open Anyway"

## Using the Tool

### Step 1: Launch the App

- **Mac**: Double-click "SEC Financials Tool.app"
- **Windows**: Double-click "SEC Financials Tool.exe"

### Step 2: Enter Information

1. **Company Ticker** (Required):
   - Enter the stock ticker symbol (e.g., TSLA, AAPL, MSFT, GOOGL)
   - Examples:
     - TSLA = Tesla
     - AAPL = Apple
     - MSFT = Microsoft
     - GOOGL = Google/Alphabet
     - AMZN = Amazon
     - META = Meta (Facebook)

2. **Fiscal Year** (Optional):
   - Leave blank to get the latest available year
   - Or enter a specific year (e.g., 2023, 2022)

3. **Email** (Optional):
   - The SEC requires an email for API identification
   - You can leave this blank if it's already configured

4. **Save Location** (Optional):
   - Click "Browse..." to choose where to save the file
   - Or leave as "Auto-generated" to save in the same folder as the app

### Step 3: Generate the File

1. Click the **"Generate Excel File"** button
2. Wait while the tool fetches data from the SEC (this may take 30-60 seconds)
3. You'll see a progress bar and status messages

### Step 4: Open Your File

When complete, you'll see a success message asking if you want to open the file location:
- Click **"Yes"** to open the folder containing your Excel file
- Click **"No"** if you know where it is

The Excel file will be named something like: `Tesla-Inc-FY-2024-Financials.xlsx`

## Examples

### Example 1: Get Tesla's Latest Financials
1. Enter ticker: **TSLA**
2. Leave year blank
3. Click "Generate Excel File"

### Example 2: Get Apple's 2023 Financials
1. Enter ticker: **AAPL**
2. Enter year: **2023**
3. Click "Generate Excel File"

### Example 3: Get Microsoft's Latest Report
1. Enter ticker: **MSFT**
2. Leave year blank
3. Click "Generate Excel File"

## Understanding the Output

The Excel file contains three sheets:

1. **Income Statement**: Shows revenue, expenses, and net income
2. **Balance Sheet**: Shows assets, liabilities, and equity
3. **Cash Flow Statement**: Shows operating, investing, and financing cash flows

All amounts are in millions of dollars.

## Troubleshooting

### "No filings found" Error

- **Check the ticker symbol**: Make sure you entered it correctly (e.g., TSLA not TESLA)
- **Try a different year**: The company may not have filed for that specific year
- **Check if company is public**: Only publicly traded US companies file with the SEC

### "Error fetching financials" or Network Error

- **Check your internet connection**: The tool needs internet to access SEC data
- **Wait and try again**: The SEC servers may be temporarily busy
- **Try a different company**: Some companies may have incomplete data

### App Won't Open (Mac)

- **Right-click and select "Open"** (first time only)
- **Check System Settings**: Go to Privacy & Security and allow the app
- **Remove quarantine**: Open Terminal and run:
  ```bash
  xattr -cr "/path/to/SEC Financials Tool.app"
  ```

### Excel File Won't Open

- **Make sure you have Excel installed** (or compatible software like Google Sheets, LibreOffice)
- **Check the file location**: Make sure the file was created successfully
- **Try generating again**: Sometimes the file may be corrupted during creation

## Tips

- **Ticker symbols** are usually 1-5 letters (e.g., AAPL, not "Apple Inc.")
- **Latest year** is usually the most recent annual filing (10-K form)
- **File size**: Excel files are typically 50-200 KB
- **Processing time**: Usually takes 30-60 seconds per company

## Need Help?

If you encounter issues not covered here:
1. Check that you have an active internet connection
2. Verify the ticker symbol is correct
3. Try a different company to test if the issue is company-specific
4. Contact support with:
   - The exact error message
   - The ticker symbol you tried
   - Your operating system (Mac/Windows)

## Privacy Note

- Your email (if provided) is only used for SEC API identification
- No data is stored or transmitted except to the SEC EDGAR database
- All processing happens on your computer


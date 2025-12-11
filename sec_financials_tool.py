#!/usr/bin/env python3
"""
SEC Financials Tool
Fetches company financial data from SEC EDGAR API and creates Excel files
similar to the Tesla FY 2024 Financials format.
"""

import argparse
import sys
from pathlib import Path
from datetime import datetime
import pandas as pd
from edgar import Company, set_identity
import logging
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

# Suppress edgar library verbose logging
logging.getLogger("edgar").setLevel(logging.WARNING)


def get_company_financials(ticker: str, year: int = None, form_type: str = "10-K"):
    """
    Fetch financial statements for a given company ticker.
    
    Args:
        ticker: Company stock ticker symbol (e.g., 'TSLA')
        year: Fiscal year (optional, defaults to latest)
        form_type: Form type to fetch (default: '10-K' for annual reports)
    
    Returns:
        Dictionary containing income statement, balance sheet, and cash flow statement
    """
    try:
        company = Company(ticker)
        # Get filings, excluding amendments to ensure full XBRL data
        filings = company.get_filings(form=form_type, amendments=False)
        
        if filings.empty:
            raise ValueError(f"No {form_type} filings found for {ticker}")
        
        # Get the latest filing or specific year if provided
        if year:
            # Filter by year if specified
            filing = filings[filings.filing_date.dt.year == year]
            if filing.empty:
                raise ValueError(f"No {form_type} filing found for {ticker} in year {year}")
            filing = filing.iloc[0]
        else:
            filing = filings.latest()
        
        filing_date = filing.filing_date
        if hasattr(filing_date, 'date'):
            filing_date = filing_date.date()
        print(f"Fetching financials from {filing_date} filing...")
        
        # Get the filing object and extract financials
        filing_obj = filing.obj()
        financials = filing_obj.financials
        
        # Extract the three main financial statements
        income_statement = financials.income_statement()
        balance_sheet = financials.balance_sheet()
        cash_flow_statement = financials.cashflow_statement()
        
        return {
            'income_statement': income_statement,
            'balance_sheet': balance_sheet,
            'cash_flow_statement': cash_flow_statement,
            'filing_date': filing.filing_date,
            'company_name': company.name if hasattr(company, 'name') else ticker
        }
    
    except Exception as e:
        raise Exception(f"Error fetching financials for {ticker}: {str(e)}")


def format_financial_dataframe(stmt) -> pd.DataFrame:
    """
    Format financial statement to DataFrame and match Excel structure.
    Handles both Statement objects and DataFrames.
    Extracts the most relevant columns (label and date values).
    """
    if stmt is None:
        return pd.DataFrame()
    
    # Convert Statement object to DataFrame if needed
    if not isinstance(stmt, pd.DataFrame):
        if hasattr(stmt, 'to_dataframe'):
            df = stmt.to_dataframe()
        else:
            return pd.DataFrame()
    else:
        df = stmt
    
    if df.empty:
        return pd.DataFrame()
    
    # Filter out abstract/metadata rows (rows where abstract is True or label is missing)
    if 'abstract' in df.columns:
        df = df[df['abstract'] != True]
    if 'label' in df.columns:
        df = df[df['label'].notna()]
    
    # Extract relevant columns: label and date columns
    relevant_cols = []
    if 'label' in df.columns:
        relevant_cols.append('label')
    
    # Find date columns (columns that look like dates: YYYY-MM-DD format)
    import re
    date_pattern = re.compile(r'^\d{4}-\d{2}-\d{2}$')
    for col in df.columns:
        if date_pattern.match(str(col)) or (isinstance(col, str) and col.startswith('20')):
            relevant_cols.append(col)
    
    # If we found relevant columns, use them; otherwise use all columns
    if relevant_cols:
        df = df[relevant_cols].copy()
        # Rename label to Item for consistency
        if 'label' in df.columns:
            df = df.rename(columns={'label': 'Item'})
    else:
        # Fallback: reset index and use all columns
        if df.index.nlevels > 1:
            df.index = df.index.map(lambda x: ' - '.join(str(i) for i in x) if isinstance(x, tuple) else str(x))
        df = df.reset_index()
        if len(df.columns) > 0:
            df.columns.values[0] = 'Item'
    
    return df


def format_number_to_millions(value):
    """Convert a number to millions, handling None and NaN values."""
    if pd.isna(value) or value is None:
        return None
    try:
        num_value = float(value)
        if num_value == 0:
            return 0
        return num_value / 1_000_000
    except (ValueError, TypeError):
        return None


def format_date_for_header(date_str):
    """Format a date string (YYYY-MM-DD) to 'Year ended Month DD, YYYY' format."""
    try:
        from datetime import datetime
        date_obj = datetime.strptime(str(date_str), '%Y-%m-%d')
        return date_obj.strftime('Year ended %B %d, %Y')
    except:
        # Fallback: try to extract year from string
        if isinstance(date_str, str) and len(date_str) >= 4:
            year = date_str[:4]
            return f'Year ended {year}'
        return str(date_str)


def format_sheet_with_headers(writer, sheet_name, df, company_name, report_type, fiscal_year):
    """
    Format a financial statement sheet with proper headers, spacing, and number formatting.
    """
    if df.empty:
        return
    
    # Create a copy to avoid modifying original
    formatted_df = df.copy()
    
    # Convert numeric columns (date columns) to millions
    numeric_cols = []
    date_headers = []
    for col in formatted_df.columns:
        if col != 'Item':
            # Check if column is numeric or looks like a date column (YYYY-MM-DD format)
            is_numeric = pd.api.types.is_numeric_dtype(formatted_df[col])
            is_date_col = isinstance(col, str) and len(str(col)) == 10 and str(col)[4] == '-' and str(col)[7] == '-'
            
            if is_numeric or is_date_col:
                numeric_cols.append(col)
                # Format date header
                date_headers.append(format_date_for_header(col) + ' (In millions)')
                # Convert values to millions
                formatted_df[col] = formatted_df[col].apply(format_number_to_millions)
    
    # Write to Excel starting at row 4 (0-indexed, so row 4 = Excel row 4)
    # Row 1: Company name
    # Row 2: Report type
    # Row 3: Empty (spacing)
    # Row 4: Column headers
    # Row 5+: Data
    start_row = 4
    formatted_df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=start_row, header=False)
    
    # Get the workbook and worksheet for formatting
    workbook = writer.book
    worksheet = writer.sheets[sheet_name]
    
    # Add company name header (row 1, Excel row 1)
    worksheet.cell(row=1, column=1, value=company_name.upper())
    worksheet.cell(row=1, column=1).font = Font(bold=True, size=12)
    
    # Add report type header (row 2, Excel row 2)
    report_header = f"Annual Report - FY {fiscal_year} - {sheet_name}"
    worksheet.cell(row=2, column=1, value=report_header)
    worksheet.cell(row=2, column=1).font = Font(bold=True, size=11)
    
    # Row 3 is empty (spacing row) - already handled by startrow
    
    # Add column headers (row 4, Excel row 4)
    worksheet.cell(row=start_row, column=1, value='')  # Empty cell for Item column header
    for idx, header in enumerate(date_headers, start=2):
        cell = worksheet.cell(row=start_row, column=idx, value=header)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='right')
    
    # Format the data rows (starting at row 5)
    for row_idx in range(start_row + 1, start_row + 1 + len(formatted_df)):
        # Format Item column (column 1)
        item_cell = worksheet.cell(row=row_idx, column=1)
        item_cell.alignment = Alignment(horizontal='left')
        
        # Format numeric columns (right-aligned, comma formatting)
        for col_idx, col_name in enumerate(numeric_cols, start=2):
            cell = worksheet.cell(row=row_idx, column=col_idx)
            if cell.value is not None and pd.notna(cell.value):
                try:
                    # Check if value is numeric
                    num_value = float(cell.value)
                    # Format as number with commas, no decimals for whole numbers
                    if abs(num_value) >= 1:
                        cell.number_format = '#,##0'
                    else:
                        cell.number_format = '#,##0.00'
                except (ValueError, TypeError):
                    # If not numeric, leave as is
                    pass
            cell.alignment = Alignment(horizontal='right')
    
    # Add spacing rows for grouping (identify rows that should have spacing)
    # Look for rows where Item column contains a colon (indicating a section header)
    current_row = start_row + 1
    rows_to_insert = []
    for idx, row in formatted_df.iterrows():
        item_value = str(row['Item']) if pd.notna(row['Item']) else ''
        # If item ends with colon, add spacing after
        if item_value.strip().endswith(':'):
            rows_to_insert.append(current_row)
        current_row += 1
    
    # Adjust column widths
    worksheet.column_dimensions['A'].width = 50  # Item column
    for col_idx in range(2, len(numeric_cols) + 2):
        col_letter = get_column_letter(col_idx)
        worksheet.column_dimensions[col_letter].width = 25


def create_excel_file(ticker: str, output_path: str = None, year: int = None, 
                      form_type: str = "10-K", user_email: str = None):
    """
    Create an Excel file with company financial statements.
    
    Args:
        ticker: Company stock ticker symbol
        output_path: Path for output Excel file (optional)
        year: Fiscal year (optional)
        form_type: Form type to fetch (default: '10-K')
        user_email: Email for SEC API identification (optional, will prompt if not provided)
    """
    # Set identity for SEC API (required)
    if user_email:
        set_identity(user_email)
    else:
        # Try to get from environment or use a default
        import os
        email = os.getenv('SEC_API_EMAIL', 'user@example.com')
        set_identity(email)
        print(f"Using email: {email} (set SEC_API_EMAIL env var or use --email to customize)")
    
    # Fetch financial data
    print(f"Fetching financial data for {ticker}...")
    financials_data = get_company_financials(ticker, year, form_type)
    
    # Format the data
    income_df = format_financial_dataframe(financials_data['income_statement'])
    balance_df = format_financial_dataframe(financials_data['balance_sheet'])
    cash_flow_df = format_financial_dataframe(financials_data['cash_flow_statement'])
    
    # Generate output filename if not provided
    if output_path is None:
        company_name = financials_data['company_name']
        # Clean company name for filename (remove commas, periods, etc.)
        company_name = company_name.replace(',', '').replace('.', '').replace(' ', '-')
        filing_date = financials_data['filing_date']
        if filing_date:
            if hasattr(filing_date, 'year'):
                filing_year = filing_date.year
            elif hasattr(filing_date, 'date'):
                filing_year = filing_date.date().year
            else:
                filing_year = year or datetime.now().year
        else:
            filing_year = year or datetime.now().year
        output_path = f"{company_name}-FY-{filing_year}-Financials.xlsx"
    
    # Determine fiscal year for headers
    filing_date = financials_data['filing_date']
    if filing_date:
        if hasattr(filing_date, 'year'):
            fiscal_year = filing_date.year
        elif hasattr(filing_date, 'date'):
            fiscal_year = filing_date.date().year
        else:
            fiscal_year = year or datetime.now().year
    else:
        fiscal_year = year or datetime.now().year
    
    # Create Excel file with multiple sheets
    print(f"Creating Excel file: {output_path}")
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        if not income_df.empty:
            format_sheet_with_headers(
                writer, 'Income Statement', income_df,
                financials_data['company_name'], 'Income Statement', fiscal_year
            )
        
        if not balance_df.empty:
            format_sheet_with_headers(
                writer, 'Balance Sheet', balance_df,
                financials_data['company_name'], 'Balance Sheet', fiscal_year
            )
        
        if not cash_flow_df.empty:
            format_sheet_with_headers(
                writer, 'Cash Flow Statement', cash_flow_df,
                financials_data['company_name'], 'Cash Flow Statement', fiscal_year
            )
    
    print(f"✓ Successfully created {output_path}")
    print(f"  Company: {financials_data['company_name']}")
    filing_date = financials_data['filing_date']
    if filing_date and hasattr(filing_date, 'date'):
        filing_date = filing_date.date()
    print(f"  Filing Date: {filing_date if filing_date else 'N/A'}")
    
    return output_path


def main():
    parser = argparse.ArgumentParser(
        description='Fetch company financial data from SEC EDGAR API and create Excel files',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python sec_financials_tool.py TSLA
  python sec_financials_tool.py AAPL --year 2023
  python sec_financials_tool.py MSFT --output Microsoft-Financials.xlsx
  python sec_financials_tool.py GOOGL --email your.email@example.com
        """
    )
    
    parser.add_argument('ticker', help='Company stock ticker symbol (e.g., TSLA, AAPL)')
    parser.add_argument('--year', type=int, help='Fiscal year (default: latest available)')
    parser.add_argument('--output', '-o', help='Output Excel file path (default: auto-generated)')
    parser.add_argument('--form', default='10-K', help='Form type (default: 10-K)')
    parser.add_argument('--email', help='Email for SEC API identification (or set SEC_API_EMAIL env var)')
    
    args = parser.parse_args()
    
    try:
        create_excel_file(
            ticker=args.ticker.upper(),
            output_path=args.output,
            year=args.year,
            form_type=args.form,
            user_email=args.email
        )
    except KeyboardInterrupt:
        print("\nOperation cancelled by user.")
        sys.exit(1)
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == '__main__':
    main()


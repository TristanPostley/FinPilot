#!/usr/bin/env python3
"""
Test script to verify date column sorting fixes data extraction issues.
This script compares unsorted vs sorted date column extraction for Apple.
"""

import sys
from pathlib import Path
from datetime import datetime
import pandas as pd
from edgar import Company, set_identity
import os
import re

# Add parent directory to path to import sec_financials_tool
sys.path.insert(0, str(Path(__file__).parent))

from sec_financials_tool import get_company_financials, format_financial_dataframe

def parse_date_col(col_name):
    """Extract date from column name for sorting."""
    date_pattern = re.compile(r'^\d{4}-\d{2}-\d{2}$')
    try:
        if date_pattern.match(str(col_name)):
            return datetime.strptime(str(col_name), '%Y-%m-%d')
        elif isinstance(col_name, str) and len(col_name) >= 10:
            return datetime.strptime(str(col_name)[:10], '%Y-%m-%d')
    except:
        pass
    return datetime(1900, 1, 1)

def format_financial_dataframe_with_sorting(stmt) -> pd.DataFrame:
    """
    Version of format_financial_dataframe WITH date column sorting.
    This is the fixed version we want to test.
    """
    if stmt is None:
        return pd.DataFrame()
    
    if not isinstance(stmt, pd.DataFrame):
        if hasattr(stmt, 'to_dataframe'):
            df = stmt.to_dataframe()
        else:
            return pd.DataFrame()
    else:
        df = stmt
    
    if df.empty:
        return pd.DataFrame()
    
    if 'abstract' in df.columns:
        df = df[df['abstract'] != True]
    if 'label' in df.columns:
        df = df[df['label'].notna()]
    
    relevant_cols = []
    if 'label' in df.columns:
        relevant_cols.append('label')
    
    date_pattern = re.compile(r'^\d{4}-\d{2}-\d{2}$')
    date_cols = []
    for col in df.columns:
        if date_pattern.match(str(col)) or (isinstance(col, str) and col.startswith('20')):
            date_cols.append(col)
    
    # SORT DATE COLUMNS (most recent first) - THIS IS THE FIX
    date_cols.sort(key=parse_date_col, reverse=True)
    
    relevant_cols.extend(date_cols)
    
    if relevant_cols:
        df = df[relevant_cols].copy()
        if 'label' in df.columns:
            df = df.rename(columns={'label': 'Item'})
    else:
        if df.index.nlevels > 1:
            df.index = df.index.map(lambda x: ' - '.join(str(i) for i in x) if isinstance(x, tuple) else str(x))
        df = df.reset_index()
        if len(df.columns) > 0:
            df.columns.values[0] = 'Item'
    
    return df

def get_value_from_row(row, column_index=0):
    """Extract value from row at specified column index."""
    if row is None or row.empty:
        return None
    
    numeric_cols = [col for col in row.index if col != 'Item']
    valid_numeric_cols = []
    for col in numeric_cols:
        try:
            val = row[col]
            if pd.notna(val):
                float(val)
                valid_numeric_cols.append(col)
        except:
            continue
    
    if valid_numeric_cols:
        try:
            col_name = valid_numeric_cols[column_index] if column_index < len(valid_numeric_cols) else valid_numeric_cols[0]
            value = row[col_name]
            if pd.notna(value):
                return float(value)
        except:
            pass
    
    return None

def find_matching_row(df, search_terms):
    """Find row matching search terms."""
    if df.empty or 'Item' not in df.columns:
        return None
    
    for term in search_terms:
        mask = df['Item'].str.contains(term, na=False, regex=False, case=False)
        if mask.any():
            return df[mask].iloc[0]
    return None

def test_date_column_extraction(ticker="AAPL", year=None):
    """
    Test date column extraction with and without sorting.
    """
    print(f"=" * 80)
    print(f"Testing Date Column Extraction for {ticker}")
    print(f"=" * 80)
    print()
    
    # Set identity for SEC API
    email = os.getenv('SEC_API_EMAIL', 'test@example.com')
    set_identity(email)
    print(f"Using SEC API email: {email}")
    print()
    
    # Fetch financial data
    print(f"Fetching financial data for {ticker}...")
    try:
        financials_data = get_company_financials(ticker, year, "10-K")
        print(f"✓ Successfully fetched data")
        print(f"  Company: {financials_data['company_name']}")
        print(f"  Filing Date: {financials_data['filing_date']}")
        print()
    except Exception as e:
        print(f"✗ Error fetching data: {e}")
        return
    
    # Test Income Statement
    print("-" * 80)
    print("INCOME STATEMENT ANALYSIS")
    print("-" * 80)
    
    income_stmt = financials_data['income_statement']
    
    # Format WITHOUT sorting (current behavior)
    df_unsorted = format_financial_dataframe(income_stmt)
    
    # Format WITH sorting (proposed fix)
    df_sorted = format_financial_dataframe_with_sorting(income_stmt)
    
    # Get date columns
    date_cols_unsorted = [col for col in df_unsorted.columns if col != 'Item']
    date_cols_sorted = [col for col in df_sorted.columns if col != 'Item']
    
    print(f"\nDate columns found: {len(date_cols_unsorted)}")
    print(f"\nUNSORTED order (current behavior):")
    for i, col in enumerate(date_cols_unsorted):
        print(f"  [{i}] {col}")
    
    print(f"\nSORTED order (proposed fix - most recent first):")
    for i, col in enumerate(date_cols_sorted):
        print(f"  [{i}] {col}")
    
    # Test key line items
    test_items = {
        'Revenue': ['Sales', 'Revenue', 'Net Sales', 'Total Revenue', 'Revenues'],
        'Cost of Goods Sold': ['Cost of Goods Sold', 'Cost of Revenue', 'Cost of Sales', 'COGS'],
        'Net Income': ['Net Income', 'Net income', 'Net earnings']
    }
    
    print(f"\n" + "=" * 80)
    print("VALUE COMPARISON (in millions)")
    print("=" * 80)
    print(f"{'Line Item':<30} {'Unsorted[0]':<20} {'Sorted[0]':<20} {'Difference':<15}")
    print("-" * 80)
    
    for item_name, search_terms in test_items.items():
        row_unsorted = find_matching_row(df_unsorted, search_terms)
        row_sorted = find_matching_row(df_sorted, search_terms)
        
        if row_unsorted is not None:
            val_unsorted = get_value_from_row(row_unsorted, 0)
            val_unsorted_millions = val_unsorted / 1_000_000 if val_unsorted else None
        else:
            val_unsorted_millions = None
        
        if row_sorted is not None:
            val_sorted = get_value_from_row(row_sorted, 0)
            val_sorted_millions = val_sorted / 1_000_000 if val_sorted else None
        else:
            val_sorted_millions = None
        
        diff = None
        if val_unsorted_millions and val_sorted_millions:
            diff = val_sorted_millions - val_unsorted_millions
        
        unsorted_str = f"${val_unsorted_millions:,.0f}M" if val_unsorted_millions else "N/A"
        sorted_str = f"${val_sorted_millions:,.0f}M" if val_sorted_millions else "N/A"
        diff_str = f"${diff:,.0f}M" if diff is not None else "N/A"
        
        print(f"{item_name:<30} {unsorted_str:<20} {sorted_str:<20} {diff_str:<15}")
    
    # Show all values for first date column (unsorted vs sorted)
    print(f"\n" + "=" * 80)
    print("DETAILED COMPARISON - First Date Column Values")
    print("=" * 80)
    
    if date_cols_unsorted and date_cols_sorted:
        first_col_unsorted = date_cols_unsorted[0]
        first_col_sorted = date_cols_sorted[0]
        
        print(f"\nUNSORTED - First column: {first_col_unsorted}")
        print(f"SORTED - First column: {first_col_sorted}")
        
        if first_col_unsorted != first_col_sorted:
            print(f"\n⚠️  WARNING: First column differs! This could cause incorrect data extraction.")
            print(f"   The tool is currently using: {first_col_unsorted}")
            print(f"   It should be using: {first_col_sorted}")
        else:
            print(f"\n✓ First column matches (no sorting needed or already correct)")
    
    # Show sample rows with all date column values
    print(f"\n" + "=" * 80)
    print("SAMPLE ROW - All Date Column Values")
    print("=" * 80)
    
    revenue_row = find_matching_row(df_sorted, ['Sales', 'Revenue', 'Net Sales', 'Total Revenue'])
    if revenue_row is not None:
        print(f"\nLine Item: {revenue_row['Item']}")
        print(f"\nAll date column values (sorted order):")
        for col in date_cols_sorted:
            val = revenue_row[col]
            if pd.notna(val):
                val_millions = float(val) / 1_000_000
                print(f"  {col}: ${val_millions:,.0f}M")
            else:
                print(f"  {col}: N/A")
    
    print(f"\n" + "=" * 80)
    print("RECOMMENDATION")
    print("=" * 80)
    
    if date_cols_unsorted != date_cols_sorted:
        print("✓ Sorting is needed - date columns are in wrong order")
        print("  The sorted version should extract the correct (most recent) fiscal year data")
    else:
        print("⚠️  Date columns are already in correct order (or only one column exists)")
        print("  Sorting may not fix the issue - investigate other causes")
    
    print()

if __name__ == '__main__':
    import argparse
    
    parser = argparse.ArgumentParser(description='Test date column sorting for SEC financial data extraction')
    parser.add_argument('ticker', nargs='?', default='AAPL', help='Company ticker (default: AAPL)')
    parser.add_argument('--year', type=int, help='Fiscal year to test')
    
    args = parser.parse_args()
    
    test_date_column_extraction(args.ticker, args.year)
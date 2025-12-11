#!/usr/bin/env python3
"""
Example usage of the SEC Financials Tool
Demonstrates how to use the tool programmatically
"""

from sec_financials_tool import create_excel_file
import os

# Example 1: Basic usage with Tesla
print("Example 1: Fetching Tesla's latest financials...")
try:
    create_excel_file(
        ticker="TSLA",
        user_email=os.getenv('SEC_API_EMAIL', 'user@example.com')
    )
except Exception as e:
    print(f"Error: {e}")

# Example 2: Fetch specific year
print("\nExample 2: Fetching Apple's 2023 financials...")
try:
    create_excel_file(
        ticker="AAPL",
        year=2023,
        user_email=os.getenv('SEC_API_EMAIL', 'user@example.com')
    )
except Exception as e:
    print(f"Error: {e}")

# Example 3: Custom output path
print("\nExample 3: Fetching Microsoft's financials with custom output...")
try:
    create_excel_file(
        ticker="MSFT",
        output_path="Microsoft-Custom-Financials.xlsx",
        user_email=os.getenv('SEC_API_EMAIL', 'user@example.com')
    )
except Exception as e:
    print(f"Error: {e}")


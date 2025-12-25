# Template vs Output Comparison Analysis

## Executive Summary

The comparison reveals **two critical issues** preventing the Financial Statements sheet from displaying correct company data:

1. **Values are NOT converted to millions** - Case Data contains raw dollar values (~12,000x larger than expected)
2. **Financial Statements sheet uses hardcoded template data** - The "Raw Data Inputs" section (rows 82-94) contains "DOLLAR TREE 2021" instead of referencing Case Data

## Detailed Findings

### Case Data Sheet - ✅ Being Populated, But Wrong Units

| Cell | Description | Template Value | Output Value | Status |
|------|-------------|----------------|--------------|--------|
| B10 | Company Name | AMERICA ONLINE INC | Zoom Communications, Inc. | ✅ **CORRECT** |
| D10 | Ticker | AOL | ZM | ✅ **CORRECT** |
| B12 | Fiscal Year End | 1995-06-30 | 2025-12-31 | ✅ **CORRECT** |
| B13 | Sales/Revenue | 394,290 | 4,665,433,000 | ❌ **NOT IN MILLIONS** (11,832x larger) |
| B15 | R&D | -64,598 | 852,415,000 | ❌ **NOT IN MILLIONS** |
| B20 | Income Taxes | -15,169 | 305,346,000 | ❌ **NOT IN MILLIONS** |

**Issue**: Values are being written in raw dollars instead of millions. The output value for B13 is approximately **12,000 times larger** than the template value.

### Financial Statements Sheet - ❌ Using Hardcoded Template Data

The Financial Statements sheet has a **"Raw Data Inputs" section** starting at row 80 that contains hardcoded template values:

| Cell | Label | Template Value | Status |
|------|-------|----------------|--------|
| B82 | Company Name | **"DOLLAR TREE 2021"** | ❌ **HARDCODED** |
| B83 | Common Shares Outstanding | 225,100.198 | ❌ **HARDCODED** |
| B85 | Sales (Net) | 22,245,500 | ❌ **HARDCODED** |
| B86 | Cost of Goods Sold | -15,223,600 | ❌ **HARDCODED** |

**Display formulas** (rows 1-79) reference these Raw Data Input cells:
- B4 (Company Name) has formula: `=B82`
- B5 (Shares Outstanding) has formula: `=B83`
- Various financial statement line items reference B85, B86, B87, etc.

## Root Cause Analysis

### Issue 1: Missing Millions Conversion

**Location**: `sec_financials_tool.py` functions:
- `map_income_statement_to_case_data()` - Line 540
- `map_balance_sheet_to_case_data()` - Line 593
- `_get_value_from_row()` - Line 458

**Problem**: The code comment says "Values are already in millions from format_financial_dataframe", but `format_financial_dataframe()` does **NOT** convert to millions. It only structures the DataFrame.

**Fix Required**: Apply `format_number_to_millions()` to values before writing to Case Data.

### Issue 2: Financial Statements Raw Data Inputs Not Populated

**Location**: `populate_case_data_sheet()` and related functions

**Problem**: The code only populates the Case Data sheet, but the Financial Statements sheet has its own "Raw Data Inputs" section (rows 82-94) that needs to be populated with the same data.

**Two Possible Solutions**:

1. **Populate Raw Data Inputs section** (simpler, preserves template structure):
   - Add a function to populate rows 82-94 in Financial Statements sheet
   - Copy data from Case Data to Financial Statements Raw Data Inputs section

2. **Change formulas to reference Case Data** (requires template modification):
   - Update Financial Statements formulas to reference `Case_Data!B13` instead of `B85`
   - Would require template changes and is more invasive

## Recommended Fix Priority

### Priority 1: Fix Millions Conversion ⚠️ CRITICAL
- **Impact**: All financial values are incorrect (12,000x too large)
- **Effort**: Low (add conversion function call)
- **Files**: `sec_financials_tool.py` - functions `map_income_statement_to_case_data()`, `map_balance_sheet_to_case_data()`

### Priority 2: Populate Financial Statements Raw Data Inputs
- **Impact**: Financial Statements sheet shows wrong company data
- **Effort**: Medium (new function to populate rows 82-94)
- **Files**: `sec_financials_tool.py` - add `populate_financial_statements_raw_data()` function

## Row Mapping: Case Data → Financial Statements Raw Data Inputs

The Financial Statements sheet has a "Raw Data Inputs" section (rows 82-94) that needs to be populated from Case Data:

| FS Row | Label | Case Data Source | Notes |
|--------|-------|------------------|-------|
| B82 | Company Name and Ticker | B10 (Name) + " " + D10 (Ticker) | Combine two cells |
| B83 | Common Shares Outstanding | B11 | Direct copy |
| B84 | Fiscal Year End | B12 | Direct copy |
| B85 | Sales (Net) | B13 | Direct copy (after millions conversion) |
| B86 | Cost of Goods Sold | B14 | Direct copy (after millions conversion) |
| B87 | R&D Expense | B15 | Direct copy (after millions conversion) |
| B88 | SG&A Expense | B16 | Direct copy (after millions conversion) |
| B89 | Depreciation & Amortization | B17 | Direct copy (after millions conversion) |
| B90 | Interest Expense | B18 | Direct copy (after millions conversion) |
| B91 | Non-Operating Income (Loss) | B19 | Direct copy (after millions conversion) |
| B92 | Income Taxes | B20 | Direct copy (after millions conversion) |
| B93 | Noncontrolling Interest in Earnings | B21 | Direct copy (after millions conversion) |
| B94 | Other Income (Loss) | B22 | Direct copy (after millions conversion) |

## Next Steps

1. ✅ **Fix millions conversion** - Update `_get_value_from_row()` or mapping functions to convert values to millions before writing to Case Data
2. ✅ **Add function to populate Raw Data Inputs** - Create `populate_financial_statements_raw_data()` function to copy values from Case Data to Financial Statements rows 82-94
3. ✅ **Add logging** - Log successful mappings and values written for debugging
4. ✅ **Test with Zoom** - Verify both Case Data and Financial Statements show correct values


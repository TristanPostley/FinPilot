# Manual Testing Guide: Template-Based Financial Analysis

This guide walks you through manual testing of the template-based financial analysis feature.

## Prerequisites

1. Ensure you have the template file: `templates/financial_analysis_template.xlsx`
2. Ensure you have Python 3.8+ and all dependencies installed:
   ```bash
   pip install -r requirements.txt
   ```
3. Have an internet connection (for SEC API access)
4. Set your SEC API email (optional, but recommended):
   ```bash
   # Windows PowerShell
   $env:SEC_API_EMAIL = "your.email@example.com"
   
   # Windows CMD
   set SEC_API_EMAIL=your.email@example.com
   
   # Mac/Linux
   export SEC_API_EMAIL="your.email@example.com"
   ```

---

## Test 1: Backward Compatibility (No Template)

**Purpose**: Verify that existing functionality still works when template is not found.

### Steps:

1. **Temporarily rename the template** (to simulate it not existing):
   ```bash
   # Windows PowerShell
   Rename-Item -Path "templates\financial_analysis_template.xlsx" -NewName "financial_analysis_template.xlsx.backup"
   
   # Mac/Linux
   mv templates/financial_analysis_template.xlsx templates/financial_analysis_template.xlsx.backup
   ```

2. **Run the tool with a simple test**:
   ```bash
   python sec_financials_tool.py TSLA
   ```

3. **Expected Behavior**:
   - Should NOT show "Using template-based workflow" message
   - Should create a file like `Tesla-Inc-FY-2024-Financials.xlsx`
   - Should contain only 3 sheets: Income Statement, Balance Sheet, Cash Flow Statement
   - Should print: "✓ Successfully created [filename]"

4. **Verify Output**:
   - Open the generated Excel file
   - Check that it has exactly 3 sheets (Income Statement, Balance Sheet, Cash Flow Statement)
   - Verify data is present and formatted correctly
   - Check that amounts are in millions

5. **Restore the template**:
   ```bash
   # Windows PowerShell
   Rename-Item -Path "templates\financial_analysis_template.xlsx.backup" -NewName "financial_analysis_template.xlsx"
   
   # Mac/Linux
   mv templates/financial_analysis_template.xlsx.backup templates/financial_analysis_template.xlsx
   ```

**✅ Test Pass Criteria**: Tool works without template, creates 3-sheet workbook, no errors.

---

## Test 2: Template-Based Workflow (With Template, No Config)

**Purpose**: Verify template-based workflow works with default config values.

### Steps:

1. **Ensure template exists**:
   ```bash
   # Verify template exists
   Test-Path templates\financial_analysis_template.xlsx  # Windows PowerShell
   ls templates/financial_analysis_template.xlsx          # Mac/Linux
   ```

2. **Run the tool**:
   ```bash
   python sec_financials_tool.py AAPL
   ```

3. **Expected Behavior**:
   - Should print: "Using template-based workflow with template: [path]"
   - Should create a file like `Apple-Inc-FY-2024-Financials.xlsx`
   - Should contain ALL template sheets (not just 3)
   - Should print: "✓ Successfully created [filename] (template-based)"

4. **Verify Output - Case Data Sheet**:
   - Open the generated Excel file
   - Go to the **Case Data** sheet
   - Check Row 10, Column B (B10): Should contain "Apple Inc." (or similar)
   - Check Row 10, Column D (D10): Should contain "AAPL"
   - Check Row 11, Column B (B11): May contain shares outstanding (or be blank/formula)
   - Check Row 12, Column B (B12): Should contain a date (fiscal year end)
   - Check Row 13, Column B (B13): Should contain Sales/Revenue value (in millions)
   - Check Row 14, Column B (B14): Should contain COGS value
   - Check Row 15, Column B (B15): Should contain R&D Expense value (if applicable)
   - Check Row 16, Column B (B16): Should contain SG&A Expense value
   - Continue checking other Income Statement rows (17-24)
   - Check Balance Sheet rows starting around row 25 (Cash, Receivables, etc.)

5. **Verify Output - Financial Statements Sheet**:
   - Go to the **Financial Statements** sheet
   - Check that it displays formatted financial statements
   - Verify that values are calculated correctly (formulas reference Case Data)

6. **Verify Output - Analysis Sheets**:
   - Check that these sheets exist and contain formulas:
     - Credit Analysis
     - Forecasting Assumptions
     - Valuation Parameters
     - Residual Income Valuations
     - DCF Valuations
     - EPS Forecaster
   - Verify that formulas are intact (not broken/errored)
   - Check that default values from `get_default_config()` are used if no config file

**✅ Test Pass Criteria**: 
- Template-based workflow activates
- Case Data sheet populated with SEC data
- All analysis sheets present with formulas intact
- Financial Statements sheet displays correctly

---

## Test 3: Template-Based Workflow (With Config File)

**Purpose**: Verify that configuration file values are properly populated into analysis sheets.

### Steps:

1. **Check the example config file exists**:
   ```bash
   cat analysis_config.json  # Or open in editor
   ```

2. **Modify the config file** (optional - to see changes):
   ```json
   {
     "valuation_parameters": {
       "wacc": 0.12,
       "terminal_growth_rate": 0.04,
       "risk_free_rate": 0.03
     },
     "forecasting_assumptions": {
       "revenue_growth_rate": 0.20,
       "operating_margin": 0.25
     },
     "cell_mappings": {
       "Valuation Parameters": {
         "B2": "wacc",
         "B3": "terminal_growth_rate",
         "B4": "risk_free_rate"
       },
       "Forecasting Assumptions": {
         "C3": "revenue_growth_rate",
         "C4": "operating_margin"
       }
     }
   }
   ```
   **Note**: Adjust cell references (B2, B3, etc.) based on actual template structure if needed.

3. **Run the tool**:
   ```bash
   python sec_financials_tool.py MSFT
   ```

4. **Verify Config Values Were Applied**:
   - Open the generated Excel file
   - Go to **Valuation Parameters** sheet
   - Check the cells specified in `cell_mappings` (e.g., B2, B3, B4)
   - Verify they contain the values from config file (0.12, 0.04, 0.03)
   - Go to **Forecasting Assumptions** sheet
   - Check cells C3, C4 for the config values (0.20, 0.25)

5. **Test with missing config file** (should use defaults):
   ```bash
   # Temporarily rename config
   Rename-Item -Path "analysis_config.json" -NewName "analysis_config.json.backup"  # Windows
   mv analysis_config.json analysis_config.json.backup  # Mac/Linux
   
   python sec_financials_tool.py GOOGL
   
   # Restore config
   Rename-Item -Path "analysis_config.json.backup" -NewName "analysis_config.json"  # Windows
   mv analysis_config.json.backup analysis_config.json  # Mac/Linux
   ```

**✅ Test Pass Criteria**: 
- Config file values populate correctly into specified cells
- Missing config file falls back to defaults without errors

---

## Test 4: Error Handling

**Purpose**: Verify error handling works correctly.

### Test 4a: Invalid Template Path

```bash
python -c "from sec_financials_tool import create_excel_file; create_excel_file('TSLA', template_path='nonexistent_template.xlsx')"
```

**Expected**: Should raise `FileNotFoundError` with clear message about template not found.

### Test 4b: Invalid Config File (Malformed JSON)

1. Create a test invalid config:
   ```bash
   echo '{invalid json}' > test_invalid_config.json
   ```

2. Run with invalid config:
   ```bash
   python -c "from sec_financials_tool import create_excel_file; create_excel_file('TSLA', config_path='test_invalid_config.json')"
   ```

3. **Expected**: Should raise `ValueError` with clear message about JSON parsing error.

4. Clean up:
   ```bash
   Remove-Item test_invalid_config.json  # Windows
   rm test_invalid_config.json  # Mac/Linux
   ```

### Test 4c: Template Missing Required Sheets

This test requires modifying the template (NOT recommended for regular testing). Skip unless debugging template validation.

**Expected**: Should raise `ValueError` listing missing required sheets.

**✅ Test Pass Criteria**: 
- Clear, actionable error messages
- Tool fails gracefully (doesn't crash with stack trace)

---

## Test 5: CLI Integration

**Purpose**: Verify CLI still works (note: CLI doesn't expose template_path/config_path yet, but should use defaults).

### Steps:

1. **Test basic CLI usage**:
   ```bash
   python sec_financials_tool.py TSLA
   ```

2. **Test with year**:
   ```bash
   python sec_financials_tool.py AAPL --year 2023
   ```

3. **Test with custom output**:
   ```bash
   python sec_financials_tool.py MSFT --output Test-Output.xlsx
   ```

4. **Verify**: All commands should work and use template-based workflow if template exists.

**✅ Test Pass Criteria**: CLI commands work as before, template workflow activates automatically.

---

## Test 6: GUI Integration

**Purpose**: Verify GUI still works without modification.

### Steps:

1. **Run the GUI**:
   ```bash
   python sec_financials_gui.py
   ```

2. **Enter a ticker** (e.g., TSLA, AAPL)

3. **Click "Generate Excel File"**

4. **Verify**: 
   - GUI runs without errors
   - File is generated successfully
   - Template-based workflow is used if template exists

**✅ Test Pass Criteria**: GUI works without modification, generates files correctly.

---

## Test 7: Programmatic Usage

**Purpose**: Verify the function can be called programmatically with new parameters.

### Steps:

1. **Create a test script** (`test_template_integration.py`):
   ```python
   from sec_financials_tool import create_excel_file
   import os
   
   # Test with default template and config
   output1 = create_excel_file(
       ticker="TSLA",
       user_email=os.getenv('SEC_API_EMAIL', 'test@example.com')
   )
   print(f"Generated: {output1}")
   
   # Test with custom template path
   output2 = create_excel_file(
       ticker="AAPL",
       template_path="templates/financial_analysis_template.xlsx",
       user_email=os.getenv('SEC_API_EMAIL', 'test@example.com')
   )
   print(f"Generated: {output2}")
   
   # Test with custom config path
   output3 = create_excel_file(
       ticker="MSFT",
       config_path="analysis_config.json",
       user_email=os.getenv('SEC_API_EMAIL', 'test@example.com')
   )
   print(f"Generated: {output3}")
   ```

2. **Run the test**:
   ```bash
   python test_template_integration.py
   ```

3. **Verify**: All three calls succeed and generate files.

**✅ Test Pass Criteria**: Function accepts new parameters, works programmatically.

---

## Verification Checklist

After running all tests, verify:

- [ ] Backward compatibility: Works without template
- [ ] Template workflow: Activates when template exists
- [ ] Case Data population: SEC data mapped correctly to Case Data sheet
- [ ] Config file: Values populate into analysis sheets
- [ ] Error handling: Clear error messages for invalid inputs
- [ ] CLI integration: Command-line interface still works
- [ ] GUI integration: GUI still works without modification
- [ ] Formula preservation: All formulas in template remain intact
- [ ] Data accuracy: Values in Case Data match SEC data (in millions)
- [ ] Sheet order: Template sheet order is preserved

---

## Common Issues & Troubleshooting

### Issue: "Template file not found" but template exists
- **Check**: Path is relative to script location, not current directory
- **Fix**: Ensure template is at `templates/financial_analysis_template.xlsx` relative to `sec_financials_tool.py`

### Issue: Values in Case Data are wrong
- **Check**: Verify SEC data format (should be in millions after `format_financial_dataframe`)
- **Fix**: Check mapping functions match SEC label names correctly

### Issue: Config values not populating
- **Check**: Cell references in `cell_mappings` match actual template structure
- **Fix**: Inspect template to find correct cell references, update config file

### Issue: Formulas show errors (#REF!, #VALUE!, etc.)
- **Check**: Case Data sheet structure matches template expectations
- **Fix**: Verify row numbers match template formulas (e.g., `=Case_Data!B13` expects Sales in B13)

---

## Test Results Template

```
Test Date: _______________
Tester: _______________

Test 1 (Backward Compatibility): [ ] Pass [ ] Fail
  Notes: _______________________________________

Test 2 (Template-Based Workflow): [ ] Pass [ ] Fail
  Notes: _______________________________________

Test 3 (Config File): [ ] Pass [ ] Fail
  Notes: _______________________________________

Test 4 (Error Handling): [ ] Pass [ ] Fail
  Notes: _______________________________________

Test 5 (CLI Integration): [ ] Pass [ ] Fail
  Notes: _______________________________________

Test 6 (GUI Integration): [ ] Pass [ ] Fail
  Notes: _______________________________________

Test 7 (Programmatic Usage): [ ] Pass [ ] Fail
  Notes: _______________________________________

Overall: [ ] All Tests Pass [ ] Some Tests Fail

Issues Found:
1. _______________________________________
2. _______________________________________
3. _______________________________________
```













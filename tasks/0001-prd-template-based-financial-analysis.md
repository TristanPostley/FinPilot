# PRD: Template-Based Financial Analysis Workbook Generation

## 1. Introduction/Overview

### Problem Statement
The current SEC Financials Tool generates Excel files containing only the three basic financial statements (Income Statement, Balance Sheet, and Cash Flow Statement) extracted from SEC EDGAR filings. Financial analysts require comprehensive analysis capabilities including credit analysis, forecasting, valuation models (DCF, Residual Income), and EPS forecasting to perform thorough company evaluations. Currently, users must manually create these analysis sheets and link formulas to the SEC data, which is time-consuming and error-prone.

### Solution Overview
Extend the existing `create_excel_file()` function to generate comprehensive financial analysis workbooks by using a pre-built Excel template that contains all analysis sheets with formulas and structure already in place. The system will populate the template with SEC data and user-provided assumptions from a configuration file, eliminating manual work and ensuring consistent analysis models.

### Goal
Enable users to generate complete financial analysis workbooks in a single operation by combining SEC EDGAR data extraction with a template-based approach, automatically populating both raw financial data and calculated analysis sheets.

---

## 2. Goals

1. **Extend Core Functionality**: Add six additional analysis sheets (Credit Analysis, Forecasting Assumptions, Valuation Parameters, Residual Income Valuations, DCF Valuations, EPS Forecaster) to the output workbook while preserving existing functionality.

2. **Maintain Backward Compatibility**: Ensure existing code, CLI usage, and GUI continue to work without modification. The template-based approach should enhance output without breaking existing workflows.

3. **Enable Configuration-Driven Analysis**: Allow users to provide analysis assumptions (WACC, growth rates, valuation parameters) via a structured configuration file (JSON or YAML) instead of requiring manual Excel editing.

4. **Ensure Data Integrity**: Validate template structure and SEC data compatibility before population to prevent errors and provide clear feedback when issues are detected.

5. **Preserve Template Integrity**: Maintain all formulas, formatting, and structure defined in the Excel template, ensuring calculations work correctly after data population.

---

## 3. User Stories

### US1: Comprehensive Financial Analysis
**As a** financial analyst  
**I want** to generate complete financial analysis workbooks with credit analysis, valuations, and forecasting  
**So that** I can perform comprehensive company analysis without manually building analysis models in Excel.

### US2: Configuration-Based Assumptions
**As a** user  
**I want** to configure analysis assumptions (WACC, growth rates, discount rates) via a configuration file  
**So that** I don't need to manually edit Excel cells and can version-control my analysis parameters.

### US3: Template Maintainability
**As a** developer  
**I want** the template-based approach where formulas live in Excel  
**So that** financial modeling experts can maintain and update formulas without modifying Python code.

### US4: Automated Data Population
**As a** financial analyst  
**I want** SEC data automatically populated into the correct sheets of my analysis template  
**So that** I can immediately review calculated valuations and forecasts without data entry errors.

### US5: Validation and Error Handling
**As a** user  
**I want** the system to validate the template structure before populating data  
**So that** I receive clear error messages if something is wrong, rather than generating a broken workbook.

---

## 4. Functional Requirements

### FR1: Template Loading
The system must load an Excel template file located at `templates/financial_analysis_template.xlsx` using the `openpyxl` library before populating any data.

### FR2: Template File Location
The template file must be stored in the `templates/` directory within the project repository root. The system must use this fixed path as the default template location.

### FR3: Template Structure Requirements
The template must contain the following sheets (exact names must match):
- **Required SEC Data Sheet**: `Case Data` (contains raw financial data that formulas reference)
- **Required Analysis Sheets**: `Credit Analysis`, `Forecasting Assumptions`, `Valuation Parameters`, `Residual Income Valuations`, `DCF Valuations`, `EPS Forecaster`
- **Required Display Sheet**: `Financial Statements` (contains formulas that reference Case Data sheet)
- Additional sheets (Intro, Ratio Analysis, Cash Flow Analysis, Model Summary, Master) may exist in the template and will be preserved as-is.

### FR4: SEC Data Population
The system must populate the `Case Data` sheet in the template with financial data fetched from SEC EDGAR. The Case Data sheet has a specific structure where:
- Row 10: Company Name, Ticker, and SIC
- Row 11: Common Shares Outstanding
- Row 12: Fiscal Year End dates (starting in column B)
- Row 13+: Financial statement line items (Sales, Cost of Goods Sold, R&D Expense, etc.) with historical data in columns B-F
- The system must map SEC data (Income Statement, Balance Sheet, Cash Flow Statement) to the appropriate rows and columns in the Case Data sheet

### FR5: Sheet Order Preservation
The system must maintain the sheet order as defined in the template file. Sheets should not be reordered during the population process.

### FR6: Configuration File Support
The system must read user-provided constants and assumptions from a JSON configuration file. The configuration file must support mapping parameter names to specific cells/sheets in the workbook. If no configuration file is provided, the system must use sensible default values (e.g., WACC = 0.10, growth rate = 0.03).

### FR7: Configuration Parameter Population
The system must populate constants from the configuration file into the appropriate cells in the analysis sheets (Credit Analysis, Forecasting Assumptions, Valuation Parameters, etc.) as specified in the config file structure.

### FR8: Template Structure Validation
Before populating data, the system must validate that the template contains all required sheet names (Case Data, Financial Statements, and all six analysis sheets: Credit Analysis, Forecasting Assumptions, Valuation Parameters, Residual Income Valuations, DCF Valuations, EPS Forecaster). If any required sheet is missing, the system must raise a clear error message indicating which sheets are missing.

### FR9: Data Compatibility Validation
The system must validate that the Case Data sheet exists and has the expected structure (rows 10-12 contain headers for Company Name, Shares Outstanding, and Fiscal Year End). The validation should check that the sheet exists and is writable before attempting population.

### FR10: Fixed Cell Reference Support
The system must support templates that use fixed cell references in formulas (e.g., `=Case_Data!B85`, `=Case_Data!C90`). The system will not modify these formulas; it will only populate the referenced data cells in the Case Data sheet. The Financial Statements sheet and other analysis sheets contain formulas that reference the Case Data sheet, and these formulas must be preserved.

### FR11: Template File Missing Handling
If the template file does not exist at the expected location, the system must raise a clear error message indicating the missing template path and providing guidance on how to resolve the issue.

### FR12: Formula Preservation
The system must preserve all existing formulas, formatting, styling, and structure in all analysis sheets. Only data cells (as specified by the config file) and SEC data sheets should be modified.

### FR13: Configuration File Parameter
The `create_excel_file()` function must accept an optional `config_path` parameter that specifies the path to the configuration file. If not provided, the system should look for a default configuration file (e.g., `analysis_config.json` in the project root) or proceed with default values if the file doesn't exist.

### FR14: Template Path Parameter
The `create_excel_file()` function must accept an optional `template_path` parameter that overrides the default template location. This allows users to specify custom templates if needed.

### FR15: Backward Compatibility
If no template is found or if template loading fails, the system must fall back to the existing behavior (creating a new workbook with only Income Statement, Balance Sheet, and Cash Flow Statement sheets) to maintain backward compatibility.

### FR16: Error Messaging
All error messages related to template loading, validation, or configuration file reading must be clear, actionable, and include specific details about what went wrong and how to fix it.

---

## 5. Non-Goals (Out of Scope)

1. **Dynamic Template Creation**: The system will not create or modify the Excel template structure programmatically. The template must be pre-built and maintained separately.

2. **Dynamic Formula Generation**: The system will not generate Excel formulas dynamically in Python code. All formulas must exist in the template file.

3. **Formula Modification**: The system will not modify existing formulas based on runtime conditions or data. Formulas are static in the template.

4. **GUI Configuration Interface**: The GUI version will not be modified to include configuration file selection or editing. Configuration files must be edited externally.

5. **Template Editing Tools**: The system will not provide tools to create or edit templates. Templates must be created and maintained using Excel or other external tools.

6. **Multi-Template Support**: Initial implementation will support a single default template. Multiple template variants are out of scope for this feature.

7. **Named Range Creation**: The system will not create Excel named ranges programmatically. Templates should use fixed cell references or pre-defined named ranges.

8. **Real-time Data Validation**: The system will not validate the correctness of financial calculations or formulas, only that the template structure is compatible.

9. **Template Versioning**: The system will not track or manage multiple versions of templates. Users must manage template versions externally.

10. **Automated Template Updates**: The system will not automatically update or migrate templates between versions.

---

## 6. Design Considerations

### Template Structure
- The template file (`templates/financial_analysis_template.xlsx`) must be an Excel workbook (.xlsx format) compatible with openpyxl.
- Template must contain all required sheets listed in FR3.
- Template sheets should use consistent formatting and styling that will be preserved.
- The **Case Data sheet** is the primary data input location where SEC financial data will be populated.
- The **Financial Statements sheet** contains formulas that reference the Case Data sheet using fixed cell references (e.g., `=Case_Data!B85`).
- Analysis sheets contain formulas that reference both the Case Data sheet and the Financial Statements sheet to perform calculations.

### Configuration File Format
The configuration file format will be **JSON only** (initial implementation). JSON is chosen because:
- Native Python support via `json` module (no additional dependencies)
- Wide familiarity among developers
- Easy to edit and validate
- YAML support may be added in future iterations if user feedback indicates it's needed

Example structure:
```json
{
  "valuation_parameters": {
    "wacc": 0.10,
    "terminal_growth_rate": 0.03,
    "risk_free_rate": 0.025
  },
  "forecasting_assumptions": {
    "revenue_growth_rate": 0.15,
    "operating_margin": 0.20
  },
  "cell_mappings": {
    "Valuation Parameters": {
      "B2": "wacc",
      "B3": "terminal_growth_rate"
    },
    "Forecasting Assumptions": {
      "C3": "revenue_growth_rate",
      "C4": "operating_margin"
    }
  }
}
```

### Data Flow
1. User calls `create_excel_file()` with ticker and optional config path
2. System loads template using `openpyxl.load_workbook()`
3. System validates template structure (checks for required sheets, especially Case Data sheet)
4. System fetches SEC data using existing `get_company_financials()` function
5. System maps SEC data (Income Statement, Balance Sheet, Cash Flow Statement) to the Case Data sheet structure, populating appropriate rows and columns (e.g., Sales data goes to row 13, COGS to row 14, etc.)
6. System reads configuration file (if provided, otherwise uses defaults)
7. System populates analysis sheet constants from config file (e.g., Valuation Parameters sheet)
8. System saves the populated workbook
9. Formulas in Financial Statements and analysis sheets automatically recalculate based on populated Case Data

### Integration Points
- **Existing Code**: Extends `create_excel_file()` function in `sec_financials_tool.py`
- **CLI Integration**: Works automatically with existing CLI interface (no changes needed)
- **GUI Integration**: Works automatically with existing GUI (uses `create_excel_file()` function)
- **Dependencies**: Uses existing `openpyxl` library (already in requirements.txt). No additional dependencies needed for JSON support (uses Python's built-in `json` module).

### Formatting and Styling
- All formatting, colors, fonts, and styling defined in the template must be preserved.
- SEC data sheets will use the existing `format_sheet_with_headers()` function which applies consistent formatting.
- Analysis sheets will retain all original formatting from the template.

---

## 7. Technical Considerations

### Function Signature Changes
The `create_excel_file()` function signature should be extended to:
```python
def create_excel_file(
    ticker: str, 
    output_path: str = None, 
    year: int = None, 
    form_type: str = "10-K", 
    user_email: str = None,
    template_path: str = None,
    config_path: str = None
):
```

### Implementation Approach
1. **Template Loading**: Replace `pd.ExcelWriter()` initialization with `openpyxl.load_workbook()` to load the template.
2. **Data Population**: Create new function `populate_case_data_sheet(workbook, financials_data: dict)` that maps SEC data to the Case Data sheet structure. This function must:
   - Map Income Statement data to appropriate rows (Sales → row 13, COGS → row 14, etc.)
   - Map Balance Sheet data to appropriate rows (starting around row 25+)
   - Map Cash Flow Statement data to appropriate rows
   - Handle column mapping for historical years (columns B-F for 5 years of data)
   - Populate company name, ticker, shares outstanding, and fiscal year end dates
3. **Config File Reading**: Create new function `load_analysis_config(config_path: str = None) -> dict` to read and parse JSON config files, with default values if file is missing.
4. **Config Population**: Create new function `populate_config_values(workbook, config: dict)` to write config values to specified cells in analysis sheets (e.g., Valuation Parameters sheet).
5. **Validation**: Create new function `validate_template(workbook) -> bool` that checks for required sheets (especially Case Data) and validates the Case Data sheet structure (rows 10-12 have expected headers).

### Dependencies
- **openpyxl**: Already in requirements.txt (version >=3.1.0)
- **json**: Built-in Python module (no additional dependency needed for JSON config support)
- **PyYAML**: Not required for initial implementation. May be added in future if YAML support is requested.

### File Structure
```
FinPilot/
├── templates/
│   └── financial_analysis_template.xlsx  # Template file
├── analysis_config.json                   # Example/default config file
├── sec_financials_tool.py                 # Modified to support templates
└── requirements.txt                       # May add PyYAML
```

### Error Handling Strategy
- **Template Not Found**: Raise `FileNotFoundError` with message: "Template file not found at {path}. Please ensure the template exists or provide a custom template_path."
- **Missing Required Sheets**: Raise `ValueError` with message: "Template is missing required sheets: {list}. Please ensure the template contains all required sheets."
- **Config File Errors**: Raise `ValueError` with message: "Error reading configuration file: {error}. Please check the file format and try again."
- **Data Population Errors**: Raise `RuntimeError` with message: "Error populating SEC data into template: {error}."

### Backward Compatibility Strategy
- Check if template file exists before attempting to load it.
- If template_path is None and default template doesn't exist, fall back to existing behavior (create new workbook).
- This ensures existing code continues to work if template is not present.

### Testing Considerations
- Unit tests should mock template loading and config file reading.
- Integration tests should use a test template file with all required sheets.
- Tests should verify that formulas are preserved after data population.
- Tests should verify error handling for missing templates and invalid config files.

---

## 8. Success Metrics

1. **Functional Success**:
   - Generated workbooks successfully populate the Case Data sheet with SEC financial data, mapping Income Statement, Balance Sheet, and Cash Flow Statement data to the correct rows and columns.
   - All analysis sheets (Credit Analysis, Forecasting Assumptions, Valuation Parameters, Residual Income Valuations, DCF Valuations, EPS Forecaster) maintain their formulas and calculate correctly after data population.
   - The Financial Statements sheet displays correctly formatted financial statements using formulas that reference the populated Case Data sheet.

2. **Configuration Success**:
   - Configuration file parameters are correctly populated into the specified cells in analysis sheets as defined in the config file structure.
   - Users can successfully use configuration files to set analysis assumptions without manual Excel editing.

3. **Validation Success**:
   - Template validation catches structural issues (missing sheets) before data population, preventing broken workbooks.
   - Clear error messages guide users to resolve template or configuration issues.

4. **Compatibility Success**:
   - Existing code, CLI usage, and GUI continue to work without modification.
   - System gracefully handles missing templates by falling back to existing behavior.

5. **User Experience Success**:
   - Users can generate complete financial analysis workbooks with a single command.
   - The time to generate analysis workbooks is comparable to generating basic financial statements (within 10-20% overhead).

---

## 9. Open Questions (Resolved)

1. **Configuration File Format Preference**: ✅ **Decision: JSON only**  
   Start with JSON only, add YAML support if user feedback indicates it's needed. JSON has native Python support via the `json` module, requiring no additional dependencies.

2. **Configuration File Requirement**: ✅ **Decision: Optional with defaults**  
   Make the configuration file optional with sensible defaults (e.g., WACC = 0.10, growth rate = 0.03) that users can override. If no config file is provided, the system will proceed with default values.

3. **Template Structure Verification**: ✅ **Decision: Adapt to existing template structure**  
   The actual template file (`templates/financial_analysis_template.xlsx`) uses a `Case Data` sheet for raw financial data input, and a `Financial Statements` sheet that contains formulas referencing the Case Data sheet. The system will populate the Case Data sheet with SEC data, mapping Income Statement, Balance Sheet, and Cash Flow Statement data to the appropriate rows and columns. All requirements have been updated to reflect this structure.

4. **Config File Location**: ✅ **Decision: Project root with override**  
   Look for `analysis_config.json` in the project root by default, but allow override via the `config_path` parameter in the function signature.

5. **Template Version Compatibility**: ✅ **Decision: Sheet name validation for initial version**  
   For initial implementation, rely on sheet name validation to ensure template compatibility. Add versioning in a future iteration if needed.

6. **Cell Mapping Complexity**: ✅ **Decision: Single cell values initially**  
   Start with single cell value mappings in the configuration file. Add array/range support if user feedback indicates it's needed.

---

## Appendix: Template Structure Reference

Based on examination of the existing `templates/financial_analysis_template.xlsx` file, it contains the following 13 sheets:

1. **Intro** - Introduction/instructions sheet
2. **Financial Statements** - Display sheet containing formulas that reference Case Data sheet
3. **Ratio Analysis** - Financial ratio calculations
4. **Cash Flow Analysis** - Cash flow analysis calculations
5. **Credit Analysis** - Credit analysis and metrics
6. **Forecasting Assumptions** - User-input assumptions for forecasting
7. **Valuation Parameters** - Valuation model parameters (WACC, growth rates, etc.)
8. **Residual Income Valuations** - Residual income valuation model
9. **DCF Valuations** - Discounted Cash Flow valuation model
10. **EPS Forecaster** - Earnings per share forecasting
11. **Model Summary** - Summary of valuation results
12. **Case Data** - **Primary data input sheet** containing raw financial data (this is where SEC data will be populated)
13. **Master** - Master control/reference sheet

### Case Data Sheet Structure

The Case Data sheet uses a specific row-based structure:
- **Row 10**: Company Name, Ticker, and SIC code (cells A10, B10, D10, F10)
- **Row 11**: Common Shares Outstanding (cell B11)
- **Row 12**: Fiscal Year End dates starting in column B (B12, C12, D12, E12, F12 for historical years)
- **Row 13+**: Financial statement line items:
  - Row 13: Sales (Net)
  - Row 14: Cost of Goods Sold
  - Row 15: R&D Expense
  - Row 16: SG&A Expense
  - Row 17: Depreciation & Amortization
  - Row 18: Interest Expense
  - Row 19: Non-Operating Income (Loss)
  - Row 20: Income Taxes
  - Row 21: Noncontrolling Interest in Earnings
  - Row 22: Other Income (Loss)
  - Row 23: Ext. Items & Disc. Ops.
  - Row 24: Preferred Dividends
  - Row 25+: Balance Sheet items (Operating Cash, Receivables, Inventories, etc.)

Historical data is stored in columns B-F (5 years of historical data). The Financial Statements sheet contains formulas like `=Case_Data!B85` that reference specific rows in the Case Data sheet to display formatted financial statements.



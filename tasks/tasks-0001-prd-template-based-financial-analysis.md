# Tasks: Template-Based Financial Analysis Workbook Generation

Based on PRD: `0001-prd-template-based-financial-analysis.md`

## Relevant Files

- `sec_financials_tool.py` - Main module containing `create_excel_file()` function and supporting functions. Modified to support template loading, Case Data population, config file integration, and backward compatibility.
- `analysis_config.json` - Example/default configuration file template that demonstrates the expected JSON structure for analysis parameters (created).
- `requirements.txt` - Dependency file (no changes needed - uses built-in json module and existing openpyxl).

### Notes

- All code changes will be in `sec_financials_tool.py` - no new Python modules needed for initial implementation.
- The template file `templates/financial_analysis_template.xlsx` already exists and should not be modified by the code.
- Configuration file support uses Python's built-in `json` module, so no new dependencies required.
- Testing can be done manually or with simple test scripts (no formal test framework specified in project structure).

## Tasks

- [x] 1.0 Template Loading and Validation Infrastructure
  - [x] 1.1 Create `load_template(template_path: str)` function that uses `openpyxl.load_workbook()` to load the template file, handling FileNotFoundError with clear error message if template doesn't exist
  - [x] 1.2 Create `validate_template_structure(workbook)` function that checks for required sheets: Case Data, Financial Statements, Credit Analysis, Forecasting Assumptions, Valuation Parameters, Residual Income Valuations, DCF Valuations, EPS Forecaster
  - [x] 1.3 Add validation in `validate_template_structure()` to verify Case Data sheet has expected structure (rows 10-12 contain headers for Company Name, Shares Outstanding, and Fiscal Year End dates)
  - [x] 1.4 Ensure validation raises ValueError with clear error messages listing missing sheets or structural issues
  - [x] 1.5 Add helper function `get_default_template_path()` that returns the default path `templates/financial_analysis_template.xlsx` relative to project root

- [x] 2.0 Configuration File Support (JSON Loading and Defaults)
  - [x] 2.1 Create `get_default_config()` function that returns a dictionary with sensible default values (e.g., WACC=0.10, terminal_growth_rate=0.03, risk_free_rate=0.025)
  - [x] 2.2 Create `load_analysis_config(config_path: str = None) -> dict` function that:
    - Takes optional config_path parameter
    - If config_path is None, looks for `analysis_config.json` in project root
    - If config file exists, reads and parses JSON using built-in `json` module
    - If config file doesn't exist, returns default config from `get_default_config()`
    - Handles JSON parsing errors with clear ValueError messages
  - [x] 2.3 Add error handling for invalid JSON syntax and file read permissions
  - [x] 2.4 Create example `analysis_config.json` file in project root with sample structure showing cell mappings (optional, can be documentation/example)

- [x] 3.0 Case Data Sheet Population (SEC Data Mapping)
  - [x] 3.1 Create helper function `populate_case_data_metadata(workbook, financials_data: dict)` that populates:
    - Row 10: Company name (cell B10), Ticker (cell D10), SIC (cell F10) if available
    - Row 11: Common Shares Outstanding (cell B11) - may need to extract from balance sheet or use placeholder
    - Row 12: Fiscal Year End dates in columns B-F based on available data from financials_data
  - [x] 3.2 Create helper function `map_income_statement_to_case_data(income_df, workbook)` that:
    - Maps Income Statement line items to appropriate Case Data rows (Sales → row 13, COGS → row 14, R&D → row 15, SG&A → row 16, etc.)
    - Handles column mapping for historical years (columns B-F for 5 years of data)
    - Converts values to match Case Data format (thousands, not millions if template expects thousands)
  - [x] 3.3 Create helper function `map_balance_sheet_to_case_data(balance_df, workbook)` that:
    - Maps Balance Sheet items starting around row 25+ (Operating Cash, Receivables, Inventories, PP&E, etc.)
    - Maps to appropriate rows in Case Data sheet based on template structure
    - Handles column mapping for historical years
  - [x] 3.4 Create helper function `map_cash_flow_to_case_data(cash_flow_df, workbook)` that maps Cash Flow Statement items to appropriate rows (Note: Cash Flow items may not have dedicated rows in Case Data - template may calculate from Income Statement/Balance Sheet. Implemented as placeholder for future extension.)
  - [x] 3.5 Create main function `populate_case_data_sheet(workbook, financials_data: dict)` that:
    - Orchestrates all the mapping functions above
    - Takes the workbook object and financials_data dictionary (from `get_company_financials()`)
    - Calls `format_financial_dataframe()` to prepare dataframes for each statement
    - Calls metadata population and all mapping functions in correct order
    - Handles empty dataframes gracefully (skip if empty)
  - [x] 3.6 Handle data conversion - verify if template expects values in thousands vs millions (using millions to match existing format_financial_dataframe behavior, which converts values to millions)

- [x] 4.0 Configuration Parameter Population
  - [x] 4.1 Create `populate_config_values(workbook, config: dict)` function that:
    - Takes workbook and config dictionary
    - Iterates through `cell_mappings` section of config (if present)
    - For each sheet name and cell reference in mappings, looks up the parameter value
    - Writes the value to the specified cell in the specified sheet
    - Handles missing sheets or invalid cell references gracefully
  - [x] 4.2 Add support for direct parameter-to-cell mapping structure from config file (implemented via cell_mappings structure)
  - [x] 4.3 Ensure function preserves existing cell formatting and only modifies values (not formulas) (using openpyxl cell assignment which preserves formatting)
  - [x] 4.4 Add error handling for invalid sheet names or cell references with informative messages (invalid sheets skipped silently, invalid cell refs handled with try/except)

- [x] 5.0 Integration and Backward Compatibility
  - [x] 5.1 Update `create_excel_file()` function signature to add optional parameters: `template_path: str = None` and `config_path: str = None`
  - [x] 5.2 Add logic at start of `create_excel_file()` to check if template exists:
    - Use `template_path` if provided, otherwise use default from `get_default_template_path()`
    - Check if template file exists using `os.path.exists()`
    - If template exists, proceed with template-based workflow
    - If template doesn't exist, fall back to existing behavior (create new workbook with pd.ExcelWriter)
  - [x] 5.3 When template exists, replace `pd.ExcelWriter()` workflow with:
    - Load template using `load_template()`
    - Validate template using `validate_template_structure()`
    - Fetch SEC data (existing code)
    - Populate Case Data sheet using `populate_case_data_sheet()`
    - Load config using `load_analysis_config(config_path)`
    - Populate config values using `populate_config_values()`
    - Save workbook using `workbook.save(output_path)` instead of ExcelWriter context manager
  - [x] 5.4 Ensure backward compatibility: when template not found, use existing code path (pd.ExcelWriter with format_sheet_with_headers)
  - [x] 5.5 Update function docstring to document new parameters and template-based behavior
  - [x] 5.6 Update print statements to indicate when template-based mode is being used vs. standard mode
  - [ ] 5.7 Test that existing CLI and GUI continue to work without modification (they call create_excel_file() which now handles both paths) - Manual testing required


# Buffett Indicator + Shiller CAPE + Berkshire Cash Dashboard

A Python script that downloads public market and macroeconomic data, builds an interactive dashboard, and exports an audit-friendly Excel workbook.

This project combines three long-horizon valuation/liquidity views into one place:

- **Buffett Indicator**: total U.S. stock market capitalization relative to U.S. nominal GDP
- **Shiller CAPE**: cyclically adjusted valuation for the S&P 500
- **Berkshire Hathaway liquidity**: cash on hand and total assets history

The script generates:

- `buffett_dashboard.html` — an interactive multi-panel Plotly dashboard
- `buffett_dashboard.xlsx` — an Excel workbook with raw, processed, merged, and audit sheets

---

## Features

- Downloads the latest public data from multiple sources
- Standardizes and merges monthly historical series
- Builds a **three-panel interactive Plotly dashboard** with a shared time axis
- Exports a structured **Excel workbook** for review and auditability
- Supports **legend-based on/off toggling** for individual series and grouped overlays
- Computes **Buffett trend lines** and **standard deviation bands on a log scale** so long-run reference bands remain proportional
- Keeps **progress messages visible during execution** for easier monitoring during longer downloads

---

## Data Sources

The script pulls data from the following public sources:

- **FRED** (Federal Reserve Economic Data)
  - U.S. nominal GDP
  - Wilshire 5000 Total Market Index
- **Yahoo Finance**
  - S&P 500 monthly history
- **Robert Shiller workbook**
  - S&P history and CAPE data
- **CompaniesMarketCap**
  - Berkshire Hathaway cash on hand
  - Berkshire Hathaway total assets

> **Note:** In this version, Berkshire liquidity data is sourced **only from CompaniesMarketCap** to avoid mixing incompatible liquidity definitions.

---

## Requirements

### Python

- Python **3.10+** recommended

### Python packages

Install the dependencies below before running the script:

```bash
pip install pandas numpy plotly requests yfinance openpyxl xlrd
```

Depending on your environment, you may already have some of these installed.

---

## File

- `buffett_dashboard.py`

---

## Usage

Run the script from the command line:

```bash
python buffett_dashboard.py --user-agent "Your Name your.email@company.com"
```

### Available arguments

```bash
python buffett_dashboard.py \
  --start 1950-01-01 \
  --output-dir . \
  --html-name buffett_dashboard.html \
  --excel-name buffett_dashboard.xlsx \
  --user-agent "Your Name your.email@company.com"
```

#### Arguments

- `--start`
  - Start date for displayed/exported data
  - Default: `1950-01-01`
- `--output-dir`
  - Directory where output files will be saved
  - Default: current directory (`.`)
- `--html-name`
  - Output HTML dashboard filename
  - Default: `buffett_dashboard.html`
- `--excel-name`
  - Output Excel workbook filename
  - Default: `buffett_dashboard.xlsx`
- `--user-agent`
  - Descriptive HTTP user-agent with contact information
  - You should provide a real name/email or organizational contact string when making requests to public data providers

---

## Outputs

### 1) Interactive HTML dashboard

The script writes:

- `buffett_dashboard.html`

This dashboard contains three linked panels:

1. **Buffett Indicator**
2. **Shiller CAPE / S&P history**
3. **Berkshire Hathaway liquidity metrics**

The legend acts as the **series toggle control**:

- Clicking a legend item turns a series on/off
- Grouped items such as valuation bands toggle together

### 2) Excel workbook

The script writes:

- `buffett_dashboard.xlsx`

The workbook includes:

- Raw/native-frequency sheets
- A merged monthly sheet with a unified monthly date spine
- Dedicated **Berkshire audit tabs**
- Dedicated **S&P-related sheets** for traceability/review

---

## How it works

At a high level, the script:

1. Creates an HTTP session with a polite user-agent
2. Downloads public historical data from FRED, Yahoo Finance, Robert Shiller, and CompaniesMarketCap
3. Cleans and standardizes the series
4. Converts date fields into a consistent monthly structure where needed
5. Builds derived metrics such as:
   - Buffett Indicator
   - CAPE mean and ±1 standard deviation bands
   - Berkshire history tables for plotting/export
6. Builds an interactive multi-panel Plotly figure
7. Saves both the HTML dashboard and Excel workbook

---

## Key implementation details

- **FRED CSV downloads** are used for macro and market-level benchmark series
- **Yahoo Finance** history is used for S&P 500 market data
- The script includes logic to **discover and parse the latest Robert Shiller workbook**
- The script includes a **best-effort parser** for historical data embedded in CompaniesMarketCap pages
- Dates are normalized to **month-end timestamps** for easier comparison and merging
- The merged Excel export is designed to be **human-readable** and suitable for inspection

---

## Notes and assumptions

- Berkshire liquidity data in this version comes **only from CompaniesMarketCap**
- Earlier mixed-source logic was removed to avoid combining incompatible liquidity definitions
- Buffett trend and standard deviation reference bands are computed on a **log scale**
- Progress messages are printed during execution to make long-running steps easier to follow

---

## Example

```bash
python buffett_dashboard.py \
  --start 1980-01-01 \
  --output-dir output \
  --html-name us_valuation_dashboard.html \
  --excel-name us_valuation_dashboard.xlsx \
  --user-agent "Jane Doe jane.doe@company.com"
```

---

## Troubleshooting

### The script fails while downloading data

Possible causes:

- Missing or blocked internet access
- Temporary source website changes
- Public source rate limiting or anti-bot protection
- Invalid or generic user-agent string

Try:

- Re-running the script later
- Supplying a more descriptive `--user-agent`
- Verifying that the upstream public URLs are reachable from your network

### The Excel file is created but some series are missing

Possible causes:

- Upstream data source format changes
- A source returned incomplete history
- Parsing rules no longer match the source layout

Try reviewing the raw and audit sheets in the workbook to identify where the series became incomplete.

### Plot opens but some lines do not appear

Possible causes:

- The series has missing values for the selected period
- The series was toggled off in the legend
- A source did not return data for that metric

---

## Limitations

- The script depends on **third-party public websites and file formats**, which may change without notice
- Yahoo Finance / Shiller / CompaniesMarketCap parsing can break if upstream page or workbook structure changes
- Public financial datasets may be revised over time
- This is a data aggregation/visualization tool, not investment advice

---

## Project structure (single-file version)

```text
buffett_dashboard.py
README.md
```

The current implementation is organized inside one script with sections for:

- constants
- data containers
- helper utilities
- download functions
- metric builders
- visualization
- export helpers
- main orchestration

---

## Suggested future improvements

If you plan to extend the script, useful next steps could include:

- Pinning dependency versions in a `requirements.txt`
- Adding automated tests for parser functions
- Caching downloads locally to reduce repeated network calls
- Splitting the single script into modules (`data/`, `metrics/`, `viz/`, `export/`)
- Adding CLI logging levels or a `--quiet` mode
- Adding optional static image export for reporting workflows

---

## License / usage notice

This script uses data from public third-party sources. Be sure to review the terms of use of each upstream provider before using the outputs in production, commercial, or redistributed workflows.

---

## Quick start

```bash
pip install pandas numpy plotly requests yfinance openpyxl xlrd
python buffett_dashboard.py --user-agent "Your Name your.email@company.com"
```

After the script completes, open:

- `buffett_dashboard.html`
- `buffett_dashboard.xlsx`

================================================================================
          MUTUAL FUND MONTHLY FACTSHEET AUTOMATION SYSTEM
================================================================================

1. QUICK START (RUN IN A SINGLE COMMAND)
----------------------------------------
To run the entire pipeline end-to-end (download new PDFs, validate them,
and create the date-stamped Excel report), just run ONE command:

   In Windows (Command Prompt or PowerShell):
       run

   Or using Python:
       python main.py

That's it! Everything runs automatically with zero configuration.


2. WHAT THIS SYSTEM DOES
------------------------
Every month when Indian AMCs release new factsheets:
- Automatically creates a new dated folder (e.g. "2026-10-12\").
- Checks ALL previous download folders and "download_registry.json".
- Skips ANY file that was already downloaded earlier (ZERO duplicates).
- Downloads ONLY newly released factsheet PDFs for that month.
- Deeply verifies each PDF with PyMuPDF (ensures valid header, pages > 0,
  unencrypted, and uncorrupted).
- Extracts all structured financial data from the PDFs into a beautifully
  formatted Excel workbook inside the date folder.


3. OUTPUT STRUCTURE
-------------------
e:\mutual_fund\
│
├── run.bat                         <-- Double-click to run everything
├── main.py                         <-- Single-command python entry point
├── download_registry.json          <-- Catalog of all 16,000+ downloaded PDFs
│
├── 2026-10-10\                     <-- Historical baseline folder
│   ├── <AMC Name>\
│   │   └── Factsheets\*.pdf
│   └── Mutual_Fund_Data_2026-10-10.xlsx
│
├── 2026-10-12\                     <-- Monthly run folder (New factsheets only)
│   ├── <AMC Name>\
│   │   └── Factsheets\*.pdf
│   └── Mutual_Fund_Data_2026-10-12.xlsx  <-- Generated Excel Report (1.08 MB)
│
└── ... future dated folders (YYYY-MM-DD)


4. EXCEL REPORT DETAILS (Mutual_Fund_Data_YYYY-MM-DD.xlsx)
----------------------------------------------------------
Every generated Excel workbook contains 6 standard, formatted sheets:

  Sheet 1: Scheme Details & Features
           - 25 standard columns: Fund Manager, AUM, NAV Growth, NAV IDCW,
             TER, Category, Benchmark, Exit Load, Volatility, Sharpe, etc.
  Sheet 2: Portfolio Holdings
           - Complete stock/debt holdings with % of NAV and issuer names.
  Sheet 3: Sector Allocation
           - Breakdown of portfolio across industry sectors.
  Sheet 4: Lump Sum Performance
           - Returns (CAGR) for 1Y, 3Y, 5Y, 10Y, and Since Inception.
  Sheet 5: SIP Performance
           - Monthly SIP returns and XIRR performance.
  Sheet 6: IDCW History
           - Record dates, distribution amounts, and NAV history.

All sheets include corporate styling (Navy Blue headers, frozen panes,
borders, auto-adjusted column widths, and grid lines).


5. COMMAND LINE CHEAT SHEET
---------------------------
Command                                   What it does
--------------------------------------------------------------------------------
run                                       Runs full pipeline for today's date
python main.py                            Runs full pipeline for today's date
python main.py --date 2026-10-12          Runs full pipeline for a specific date
python main.py --excel-only               Extracts Excel only (skips download)
python main.py --download-only            Downloads & validates only (no Excel)
python main.py --all                      Generates Excel for all dated folders


6. PIPELINE STEPS UNDER THE HOOD
--------------------------------
When you run `python main.py`, it executes 4 steps in order:

  Step 1: download_factsheets.py
          Scans all 57+ AMC official portals and AMFI online directory.
          Uses DownloadRegistry to immediately skip any existing PDF.

  Step 2: fix_missing_factsheets.py
          Aggregator fallback that backfills any AMC with broken portals.

  Step 3: strict_corporate_validation.py
          Opens every downloaded PDF with PyMuPDF to guarantee 100%
          structural integrity and deletes any corrupted file.

  Step 4: extract_to_excel.py
          Reads all PDFs in the dated folder and generates the
          "Mutual_Fund_Data_YYYY-MM-DD.xlsx" workbook.


7. SYSTEM REQUIREMENTS & DEPENDENCIES
-------------------------------------
- Python 3.10+
- Required packages:
    pip install requests beautifulsoup4 pymupdf pandas openpyxl urllib3
================================================================================

# 📊 India Mutual Fund Factsheet Automation & ETL Pipeline

An automated, corporate-grade Python pipeline designed to incrementally download, validate, and extract financial data from factsheets across all **57+ Indian Mutual Fund Asset Management Companies (AMCs)** into standardized, executive Excel workbooks.

---

## 🚀 Quick Start (Single Command)

Execute the complete 4-step pipeline end-to-end with just **ONE** command:

```bash
python main.py
```

*On Windows, you can also simply run:*
```cmd
run
```

---

## 🌟 Key Architecture & Capabilities

### 1. 🔄 Monthly Incremental Downloader (Zero Duplicates)
- **Automatic Date-Based Partitioning**: Creates a clean date folder (`YYYY-MM-DD/`) for every monthly run.
- **Persistent Global Registry**: Cross-checks available online documents against `download_registry.json` (tracking 16,000+ files) and all prior historical folders (`2026-10-10/`, etc.).
- **Strictly Zero Duplicates**: Downloads **only newly launched** monthly factsheet PDFs; all existing files are skipped in $O(1)$ time.
- **Noise Filtering**: Intelligently rejects non-factsheet files (KYC forms, addendums, notices, transaction slips, and annual reports).

### 2. 🛡️ Aggregator Fallback Engine
- Scans central aggregator directories to automatically backfill factsheet coverage for AMCs with broken or dynamic portals, without creating redundant downloads.

### 3. 🔍 Strict Corporate PDF Integrity Validation
- Audits every downloaded PDF using **PyMuPDF (`fitz`)**.
- Verifies `%PDF` magic byte headers, unencrypted access, and valid page counts (> 0).
- Automatically purges corrupted, truncated, or password-locked files.

### 4. 📈 Standardized Excel Extraction Engine
- Parses factsheet PDFs across all AMCs into a date-stamped workbook:
  `Mutual_Fund_Data_YYYY-MM-DD.xlsx`
- Guarantees **6 executive-formatted sheets**:
  1. **Scheme Details & Features**: 25 standardized columns (Fund Managers, AUM/AAUM, NAVs, TER, Benchmark, Risk Ratios, Sharpe, Beta, Exit Load).
  2. **Portfolio Holdings**: Equity & debt holdings, company/issuer names, industries, and % of NAV.
  3. **Sector Allocation**: Portfolio breakdown across industry sectors.
  4. **Lump Sum Performance**: CAGR returns for 1Y, 3Y, 5Y, 10Y, and Since Inception.
  5. **SIP Performance**: Systematic Investment Plan returns and XIRR performance.
  6. **IDCW History**: Record dates, distribution quantum, and prevailing NAVs.
- **Executive Visual Styling**: Navy Blue (`#1F4E78`) headers, frozen header panes, thin cell borders, auto-fitted column widths, and gridlines.

---

## 📂 Output Structure

```
e:\mutual_fund\
│
├── main.py                         # Unified single-command runner
├── run.py                          # Single-command runner alias
├── run.bat                         # Windows batch double-click runner
├── run_pipeline.py                 # Master pipeline coordinator (Steps 1–4)
│
├── download_factsheets.py          # Step 1: Incremental factsheet downloader
├── fix_missing_factsheets.py       # Step 2: Aggregator fallback engine
├── strict_corporate_validation.py  # Step 3: PyMuPDF PDF integrity validator
├── extract_to_excel.py             # Step 4: Multi-sheet Excel extraction engine
│
├── download_registry.py            # Deduplication engine
├── download_registry.json          # Persistent registry catalog
├── full_string.txt                 # AMFI offline fallback database
│
├── 2026-10-10/                     # Baseline historical archive (19,500+ PDFs)
│   ├── <AMC Name>/Factsheets/*.pdf
│   └── Mutual_Fund_Data_2026-10-10.xlsx
│
└── 2026-10-12/                     # Monthly incremental folder (New PDFs only)
    ├── <AMC Name>/Factsheets/*.pdf
    └── Mutual_Fund_Data_2026-10-12.xlsx
```

---

## ⚙️ Command Line Options

| Command | Description |
| :--- | :--- |
| `python main.py` | Full pipeline for today's date (Default) |
| `python main.py --date 2026-10-12` | Full pipeline targeting a specific dated folder |
| `python main.py --excel-only` | Extracts Excel immediately (skips downloads) |
| `python main.py --download-only` | Runs download and validation only |
| `python main.py --all` | Extracts Excel across all historical date folders |

---

## 📦 Requirements & Installation

- Python 3.10+

Install dependencies:
```bash
pip install requests beautifulsoup4 pymupdf pandas openpyxl urllib3
```

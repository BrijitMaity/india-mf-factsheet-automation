r"""
=========================================================================
MUTUAL FUND FACTSHEET DATA EXTRACTION ENGINE
=========================================================================
Extracts structured data from all factsheet PDFs in a given dated folder
and produces a comprehensive, formatted Excel workbook.

Usage:
    python extract_to_excel.py                  # Uses today's date folder
    python extract_to_excel.py --date 2026-10-12 # Specific date folder
    python extract_to_excel.py --all             # Process ALL dated folders

Output:
    e:\mutual_fund\<YYYY-MM-DD>\Mutual_Fund_Data_<YYYY-MM-DD>.xlsx

This script is designed to be the Step 4 in the pipeline, running AFTER
download, fallback, and validation steps.
=========================================================================
"""

import fitz
import re
import os
import sys
import datetime
import argparse
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

sys.stdout.reconfigure(encoding='utf-8')

WORKSPACE = r"e:\mutual_fund"

# -----------------------------------------------------------------------
# 1. FIND BEST FACTSHEET PER AMC IN A GIVEN DATED FOLDER
# -----------------------------------------------------------------------
MONTH_ORDER = {
    'january': 1, 'jan': 1, 'february': 2, 'feb': 2, 'march': 3, 'mar': 3,
    'april': 4, 'apr': 4, 'may': 5, 'june': 6, 'jun': 6, 'july': 7, 'jul': 7,
    'august': 8, 'aug': 8, 'september': 9, 'sep': 9, 'sept': 9,
    'october': 10, 'oct': 10, 'november': 11, 'nov': 11, 'december': 12, 'dec': 12
}

def extract_date_score(filename):
    """Score a filename by the year/month embedded in it."""
    fname = filename.lower()
    year_match = re.findall(r'(201\d|202\d)', fname)
    year = int(year_match[-1]) if year_match else 200
    month = 0
    for m_name, m_val in MONTH_ORDER.items():
        if m_name in fname:
            if m_val > month:
                month = m_val
    return (year, month)

def find_factsheets_in_folder(date_folder):
    """Discover the best factsheet PDF for each AMC inside a dated folder."""
    result = {}
    if not os.path.isdir(date_folder):
        return result

    for amc in sorted(os.listdir(date_folder)):
        amc_dir = os.path.join(date_folder, amc)
        if not os.path.isdir(amc_dir):
            continue

        fs_dir = os.path.join(amc_dir, 'Factsheets')
        if not os.path.exists(fs_dir):
            continue

        pdf_files = [f for f in os.listdir(fs_dir) if f.lower().endswith('.pdf')]
        if not pdf_files:
            continue

        scored_files = []
        for f in pdf_files:
            score = extract_date_score(f)
            fpath = os.path.join(fs_dir, f)
            mtime = os.path.getmtime(fpath)
            size = os.path.getsize(fpath)
            scored_files.append((score, size > 1024 * 1024, mtime, f, fpath))

        scored_files.sort(key=lambda x: (x[0], x[1], x[2]), reverse=True)
        # Include ALL PDFs for the AMC, not just the best one
        for entry in scored_files:
            if amc not in result:
                result[amc] = []
            result[amc].append(entry)

    return result


# -----------------------------------------------------------------------
# 2. PRE-COMPILED REGEX PATTERNS
# -----------------------------------------------------------------------
RE_DATE = re.compile(
    r'(?:as on|dated|data as on|factsheet)\s*[:\-]?\s*'
    r'([A-Za-z]+\s+\d{1,2},?\s*\d{4}|\d{1,2}[/-][A-Za-z]+[/-]\d{4}|\d{1,2}[/-]\d{1,2}[/-]\d{4})',
    re.IGNORECASE
)
RE_PCT = re.compile(r'^[+-]?\d+(\.\d+)?%?$')


# -----------------------------------------------------------------------
# 3. SPECIALIZED MIRAE ASSET PARSER
# -----------------------------------------------------------------------
def parse_mirae_asset_pdf(amc_name, pdf_path, data):
    """Parse Mirae Asset's unique multi-page factsheet layout."""
    doc = fitz.open(pdf_path)
    fname = os.path.basename(pdf_path)

    for pno in range(len(doc)):
        page = doc[pno]
        text = page.get_text()
        blocks = page.get_text("blocks")

        if "Fund Information" in text and ("Portfolio Top 10 Holdings" in text or "Fund Performance" in text or "Sector Allocation" in text):
            # Extract scheme title
            b0_text = blocks[0][4].strip()
            lines_b0 = [l.strip() for l in b0_text.split('\n') if l.strip()]
            scheme_title = " ".join(lines_b0)
            if not ("MIRAE ASSET" in scheme_title.upper() and "FUND" in scheme_title.upper()):
                for b in blocks[:3]:
                    if "MIRAE ASSET" in b[4].upper():
                        scheme_title = " ".join([l.strip() for l in b[4].split('\n') if l.strip()])
                        break

            category = blocks[1][4].strip() if len(blocks) > 1 and "(" in blocks[1][4] else "-"
            date_match = re.search(r'Monthly Factsheet as on\s*([^\n]+)', text, re.I)
            factsheet_date = date_match.group(1).strip() if date_match else "-"

            fund_mgr = "-"
            allot_date = "-"
            benchmark = "-"
            net_aum = "-"
            exit_load = "-"
            plan_avail = "Regular Plan and Direct Plan"
            min_inv_parts = []

            nav_dir_growth = "-"
            nav_reg_growth = "-"
            nav_dir_idcw = "-"
            nav_reg_idcw = "-"

            volatility = "-"
            beta = "-"
            rsquared = "-"
            sharpe = "-"
            turnover = "-"
            ter_reg = "-"
            ter_dir = "-"

            for i, b in enumerate(blocks):
                txt = b[4].strip()
                if "Fund Managers :" in txt:
                    if i + 1 < len(blocks):
                        mgr_cand = blocks[i+1][4].strip().replace('\n', ' ')
                        if not any(bad in mgr_cand.upper() for bad in ["SIP", "PERFORMANCE", "RATIOS", "ALLOTMENT"]):
                            fund_mgr = mgr_cand
                elif "Allotment Date :" in txt:
                    parts = txt.split('\n')
                    if len(parts) > 1: allot_date = parts[1].strip()
                elif "Benchmark :" in txt:
                    parts = txt.split('\n')
                    if len(parts) > 1: benchmark = parts[1].strip()
                elif "Net AUM" in txt:
                    parts = txt.split('\n')
                    if len(parts) > 1: net_aum = parts[1].strip()
                elif "Exit Load :" in txt:
                    parts = txt.split('\n')
                    if len(parts) > 1: exit_load = parts[1].strip()
                elif "Minimum Investment Amount" in txt:
                    for k in range(i+1, min(i+5, len(blocks))):
                        if "Net Asset Value" in blocks[k][4]:
                            break
                        min_inv_parts.append(blocks[k][4].strip())
                elif "Growth" in txt and "₹" in txt:
                    vals = re.findall(r'₹\s*([0-9]+\.[0-9]+)', txt)
                    if len(vals) >= 2:
                        nav_dir_growth, nav_reg_growth = vals[0], vals[1]
                elif "IDCW" in txt and "₹" in txt:
                    vals = re.findall(r'₹\s*([0-9]+\.[0-9]+)', txt)
                    if len(vals) >= 2:
                        nav_dir_idcw, nav_reg_idcw = vals[0], vals[1]
                elif "Volatility" in txt:
                    m = re.search(r'([0-9]+\.[0-9]+%)', txt)
                    if m: volatility = m.group(1)
                elif "Beta" in txt:
                    m = re.search(r'([0-9]+\.[0-9]+)', txt)
                    if m: beta = m.group(1)
                elif "R Squared" in txt:
                    m = re.search(r'([0-9]+\.[0-9]+)', txt)
                    if m: rsquared = m.group(1)
                elif "Sharpe Ratio" in txt:
                    m = re.search(r'([0-9]+\.[0-9]+)', txt)
                    if m: sharpe = m.group(1)
                elif "Portfolio Turnover" in txt:
                    m = re.search(r'([0-9]+\.[0-9]+(?:\s+times)?)', txt)
                    if m: turnover = m.group(1)
                elif "Regular Plan" in txt and "%" in txt:
                    m = re.search(r'([0-9]+\.[0-9]+%)', txt)
                    if m: ter_reg = m.group(1)
                elif "Direct Plan" in txt and "%" in txt:
                    m = re.search(r'([0-9]+\.[0-9]+%)', txt)
                    if m: ter_dir = m.group(1)

            if fund_mgr == "-":
                m_fm = re.search(r'Fund Managers\s*:\s*(Mr\.\s+[^\n\(]+)', text)
                if m_fm: fund_mgr = m_fm.group(1).strip()

            data['features'].append({
                'Company Name': amc_name,
                'Scheme Name': scheme_title,
                'Factsheet Date': factsheet_date,
                'Scheme Category': category,
                'Fund Manager(s)': fund_mgr,
                'Date of Allotment': allot_date,
                'Benchmark': benchmark,
                'Net AUM (₹ Cr)': net_aum,
                'Monthly Average AUM / AAUM (₹ Cr)': net_aum,
                'Plan Available': plan_avail,
                'Minimum Investment Amount': " ".join(min_inv_parts) if min_inv_parts else "₹5,000/-",
                'NAV Direct Growth (₹)': nav_dir_growth,
                'NAV Regular Growth (₹)': nav_reg_growth,
                'NAV Direct IDCW (₹)': nav_dir_idcw,
                'NAV Regular IDCW (₹)': nav_reg_idcw,
                'Volatility / Std Dev (%)': volatility,
                'Beta': beta,
                'R-Squared': rsquared,
                'Sharpe Ratio': sharpe,
                'Portfolio Turnover Ratio': turnover,
                'TER Regular Plan (%)': ter_reg,
                'TER Direct Plan (%)': ter_dir,
                'Exit Load': exit_load,
                'Factsheet Source File': fname,
                'PDF Page': pno + 1
            })

            # Holdings
            for b in blocks:
                if 270 <= b[0] <= 290 and 200 <= b[1] <= 420:
                    lines = [l.strip() for l in b[4].split('\n') if l.strip()]
                    if len(lines) == 2 and re.match(r'^\d+(\.\d+)?%$', lines[1]):
                        data['holdings'].append({
                            'Company Name': amc_name,
                            'Scheme Name': scheme_title,
                            'Factsheet Date': factsheet_date,
                            'Asset Class / Category': 'Equity Shares',
                            'Company / Issuer Name': lines[0],
                            'Industry / Rating': '-',
                            'Portfolio % / % of NAV': lines[1],
                            'Benchmark %': '-',
                            'Over/Underweight (+/- %)': '-',
                            '% Exposure of Derivative': '0.00%'
                        })

            # Sectors
            sec_names = []
            sec_pcts = []
            for b in blocks:
                if 550 <= b[0] <= 690 and 220 <= b[1] <= 390:
                    lines = [l.strip() for l in b[4].split('\n') if l.strip()]
                    for l in lines:
                        if not re.match(r'^\d+(\.\d+)?%$', l) and l not in ["0%", "20%", "40%", "60%"]:
                            sec_names.append((b[1], l))
                if 700 <= b[0] <= 860 and 220 <= b[1] <= 390:
                    lines = [l.strip() for l in b[4].split('\n') if l.strip()]
                    for l in lines:
                        if re.match(r'^\d+(\.\d+)?%$', l) and l not in ["0%", "20%", "40%", "60%"]:
                            sec_pcts.append((b[1], l))
            sec_names.sort(key=lambda x: x[0])
            sec_pcts.sort(key=lambda x: x[0])
            for sn, sp in zip(sec_names, sec_pcts):
                data['sectors'].append({
                    'Company Name': amc_name,
                    'Scheme Name': scheme_title,
                    'Factsheet Date': factsheet_date,
                    'Sector / Industry': sn[1],
                    '% Allocation': sp[1]
                })

            # Lump Sum Performance
            for b in blocks:
                lines = [l.strip() for l in b[4].split('\n') if l.strip()]
                if lines and any(p in lines[0] for p in ["Last 1 Year", "Last 3 Years", "Last 5 Years", "Last 10 Years", "Last 15 Years", "Since Inception"]):
                    pcts = [l for l in lines[1:] if re.match(r'^[+-]?\d+(\.\d+)?%$', l)]
                    if len(pcts) >= 3:
                        data['lumpsum'].append({
                            'Company Name': amc_name,
                            'Scheme Name': scheme_title,
                            'Factsheet Date': factsheet_date,
                            'Period / Horizon': lines[0],
                            'Scheme CAGR (%)': pcts[0],
                            'Scheme Benchmark Returns (%)': pcts[1],
                            'Additional Benchmark Returns (%)': pcts[2],
                            'Value of ₹10,000 Invested - Scheme (₹)': '-',
                            'Value of ₹10,000 Invested - Benchmark (₹)': '-',
                            'Value of ₹10,000 Invested - Addl Benchmark (₹)': '-',
                            'NAV as on Date': nav_reg_growth if nav_reg_growth != "-" else nav_dir_growth,
                            'Scheme Benchmark Name': benchmark,
                            'Additional Benchmark Name': 'BSE Sensex (TRI)'
                        })

            # SIP Performance
            horizons_p = []
            inv_amounts = []
            mkt_values = []
            fund_returns = []
            bench_returns = []
            add_bench_returns = []

            for b in blocks:
                lines = [l.strip() for l in b[4].split('\n') if l.strip()]
                if len(lines) > 1:
                    if lines[0] == "Period" and any("Inception" in l for l in lines):
                        horizons_p = lines[1:]
                    elif "Total Amount Invested" in lines[0]:
                        inv_amounts = lines[1:]
                    elif "Mkt Value as on" in lines[0]:
                        mkt_values = lines[1:]
                    elif "Fund Return" in lines[0]:
                        fund_returns = lines[1:]
                    elif "Benchmark Return" in lines[0]:
                        bench_returns = lines[1:]
                    elif "Add. Benchmark Return" in lines[0]:
                        add_bench_returns = lines[1:]

            if horizons_p and inv_amounts:
                for h_idx, h in enumerate(horizons_p):
                    data['sip'].append({
                        'Company Name': amc_name,
                        'Scheme Name': scheme_title,
                        'Factsheet Date': factsheet_date,
                        'SIP Horizon / Period': h,
                        'Total Amount Invested (₹)': inv_amounts[h_idx] if h_idx < len(inv_amounts) else "-",
                        'Market Value as on Date (₹)': mkt_values[h_idx] if h_idx < len(mkt_values) else "-",
                        'Fund Return % (XIRR)': fund_returns[h_idx] if h_idx < len(fund_returns) else "-",
                        'Benchmark Return %': bench_returns[h_idx] if h_idx < len(bench_returns) else "-",
                        'Additional Benchmark Return %': add_bench_returns[h_idx] if h_idx < len(add_bench_returns) else "-"
                    })

    # IDCW History (tail pages)
    for pno in range(max(0, len(doc) - 30), len(doc)):
        text = doc[pno].get_text()
        if "IDCW" in text or "DIVIDEND" in text or "Record Date" in text:
            blocks = doc[pno].get_text("blocks")
            for b in blocks:
                lines = [l.strip() for l in b[4].split('\n') if l.strip()]
                full_line = " ".join(lines)
                date_m = re.search(r'(\d{2}-[A-Za-z]{3}-\d{4}|\d{2}/\d{2}/\d{4})', full_line)
                if date_m:
                    rec_date = date_m.group(1)
                    nums = re.findall(r'(\d+\.\d+)', full_line)
                    if len(nums) >= 2:
                        data['idcw'].append({
                            'Company Name': amc_name,
                            'Scheme Name': 'Mirae Asset Equity Schemes',
                            'Option / Plan': 'Regular Plan' if len(nums) < 5 else 'Direct Plan',
                            'Record Date': rec_date,
                            'Quantum / Distribution (₹/unit)': nums[0],
                            'Face Value (₹/unit)': nums[1] if len(nums) > 1 and "10" in nums[1] else "10.00",
                            'NAV Cum IDCW / Prevailing (₹)': nums[2] if len(nums) > 2 else "-",
                            'NAV Ex IDCW (₹)': nums[3] if len(nums) > 3 else "-",
                            'Direct Plan Quantum (₹/unit)': nums[4] if len(nums) > 4 else "-",
                            'Direct Plan NAV (₹)': nums[5] if len(nums) > 5 else "-"
                        })
    doc.close()


# -----------------------------------------------------------------------
# 4. GENERAL PARSER FOR ALL OTHER AMCs
# -----------------------------------------------------------------------
def parse_pdf_general(amc_name, pdf_path, data):
    """Universal parser for all non-Mirae factsheet PDFs."""
    doc = fitz.open(pdf_path)
    fname = os.path.basename(pdf_path)
    total_pages = len(doc)

    current_scheme = amc_name + " Scheme"
    current_category = "Equity"

    for pno in range(total_pages):
        page = doc[pno]
        page_text = page.get_text()
        if len(page_text) < 60:
            continue

        blocks = page.get_text("blocks")

        # --- Scheme Title Detection ---
        d = page.get_text("dict")
        prominent = []
        for b in d.get("blocks", []):
            if "lines" in b and b.get("bbox", [0, 0, 0, 0])[1] < 220:
                for line in b["lines"]:
                    for span in line["spans"]:
                        stxt = span["text"].strip()
                        ssize = span["size"]
                        if len(stxt) > 3 and ssize >= 11:
                            prominent.append((ssize, b["bbox"][1], stxt))

        prominent.sort(key=lambda x: (x[0], -x[1]), reverse=True)
        for cand in prominent:
            t = cand[2]
            if any(k in t.upper() for k in [" FUND", " ETF", " SCHEME", " FOF", " OPPORTUNITIES", " ADVANTAGE", " BLUECHIP", " EMERGING"]):
                if not any(bad in t.upper() for bad in ["MUTUAL FUND", "FACTSHEET", "CONTENTS", "READ", "PORTFOLIO ASPECTS", "DISCLAIMER"]):
                    current_scheme = t
                    break

        # --- Factsheet Date ---
        m_dt = RE_DATE.search(page_text[:400])
        factsheet_date = m_dt.group(1).strip() if m_dt else "-"

        # --- Scheme Details & Features ---
        has_features = any(k in page_text for k in ["AUM", "Net Asset Value", "NAV", "Date of Allotment", "Allotment Date", "Inception Date"])
        if has_features and len(page_text) > 150:
            nav_reg = None
            nav_dir = None
            nav_idcw_reg = None
            nav_idcw_dir = None

            m_nav_rg = re.search(r'(?:Regular|Reg)\s*(?:Plan)?\s*(?:-|:)?\s*Growth\s*[:\-₹]?\s*([0-9,]+\.[0-9]+)', page_text, re.IGNORECASE)
            if m_nav_rg: nav_reg = m_nav_rg.group(1)

            m_nav_dg = re.search(r'(?:Direct|Dir)\s*(?:Plan)?\s*(?:-|:)?\s*Growth\s*[:\-₹]?\s*([0-9,]+\.[0-9]+)', page_text, re.IGNORECASE)
            if m_nav_dg: nav_dir = m_nav_dg.group(1)

            m_nav_rid = re.search(r'(?:Regular|Reg)\s*(?:Plan)?\s*(?:-|:)?\s*IDCW\s*[:\-₹]?\s*([0-9,]+\.[0-9]+)', page_text, re.IGNORECASE)
            if m_nav_rid: nav_idcw_reg = m_nav_rid.group(1)

            m_nav_did = re.search(r'(?:Direct|Dir)\s*(?:Plan)?\s*(?:-|:)?\s*IDCW\s*[:\-₹]?\s*([0-9,]+\.[0-9]+)', page_text, re.IGNORECASE)
            if m_nav_did: nav_idcw_dir = m_nav_did.group(1)

            aum_me = None
            m_aum = re.search(r'(?:Month End AUM|AUM as on|Net AUM|Closing AUM|AUM)\s*[:\-₹]?\s*([0-9,]+\.[0-9]+)', page_text, re.IGNORECASE)
            if m_aum: aum_me = m_aum.group(1)

            aaum = None
            m_aaum = re.search(r'(?:AAUM|Monthly Average|Average AUM)\s*[:\-₹]?\s*([0-9,]+\.[0-9]+)', page_text, re.IGNORECASE)
            if m_aaum: aaum = m_aaum.group(1)

            allot_date = None
            m_allot = re.search(r'(?:Date of Allotment|Allotment Date|Inception Date)\s*[:\-]?\s*([^\n]+)', page_text, re.IGNORECASE)
            if m_allot: allot_date = m_allot.group(1).strip()

            benchmark = None
            m_bench = re.search(r'Benchmark\s*[:\-]?\s*([^\n]+)', page_text, re.IGNORECASE)
            if m_bench: benchmark = m_bench.group(1).strip()

            category = None
            m_cat = re.search(r'(?:Scheme Category|Category of Scheme|Category)\s*[:\-]?\s*([^\n]+)', page_text, re.IGNORECASE)
            if m_cat: category = m_cat.group(1).strip()

            fund_mgr = None
            m_mgr = re.search(r'Fund Manager[s]?\s*[:\-]?\s*([^\n]+)', page_text, re.IGNORECASE)
            if m_mgr: fund_mgr = m_mgr.group(1).strip()

            ter_reg = None
            m_tr = re.search(r'Regular(?:\s+Plan)?\s*[:\-]?\s*([0-9]+\.[0-9]+%?)', page_text, re.IGNORECASE)
            if m_tr: ter_reg = m_tr.group(1)

            ter_dir = None
            m_td = re.search(r'Direct(?:\s+Plan)?\s*[:\-]?\s*([0-9]+\.[0-9]+%?)', page_text, re.IGNORECASE)
            if m_td: ter_dir = m_td.group(1)

            turnover = None
            m_to = re.search(r'(?:Portfolio Turnover|Turnover Ratio)\s*[:\-]?\s*([0-9]+\.[0-9]+(?:\s+times)?%?)', page_text, re.IGNORECASE)
            if m_to: turnover = m_to.group(1)

            volatility = None
            m_vol = re.search(r'(?:Volatility|Standard Deviation)\s*[:\-]?\s*([0-9]+\.[0-9]+%?)', page_text, re.IGNORECASE)
            if m_vol: volatility = m_vol.group(1)

            beta = None
            m_beta = re.search(r'Beta\s*[:\-]?\s*([0-9]+\.[0-9]+)', page_text, re.IGNORECASE)
            if m_beta: beta = m_beta.group(1)

            rsquared = None
            m_r2 = re.search(r'(?:R Squared|R-Squared)\s*[:\-]?\s*([0-9]+\.[0-9]+)', page_text, re.IGNORECASE)
            if m_r2: rsquared = m_r2.group(1)

            sharpe = None
            m_sh = re.search(r'Sharpe Ratio\s*[:\-]?\s*([0-9]+\.[0-9]+)', page_text, re.IGNORECASE)
            if m_sh: sharpe = m_sh.group(1)

            exit_load = None
            m_el = re.search(r'Exit Load\s*[:\-]?\s*([^\n]+)', page_text, re.IGNORECASE)
            if m_el: exit_load = m_el.group(1).strip()

            min_inv = None
            m_mi = re.search(r'Minimum Investment Amount\s*[:\-]?\s*([^\n]+)', page_text, re.IGNORECASE)
            if m_mi: min_inv = m_mi.group(1).strip()

            if any([nav_reg, nav_dir, aum_me, aaum, allot_date, benchmark]):
                data['features'].append({
                    'Company Name': amc_name,
                    'Scheme Name': current_scheme,
                    'Factsheet Date': factsheet_date,
                    'Scheme Category': category or '-',
                    'Fund Manager(s)': fund_mgr or '-',
                    'Date of Allotment': allot_date or '-',
                    'Benchmark': benchmark or '-',
                    'Net AUM (₹ Cr)': aum_me or '-',
                    'Monthly Average AUM / AAUM (₹ Cr)': aaum or '-',
                    'Plan Available': 'Regular Plan and Direct Plan',
                    'Minimum Investment Amount': min_inv or '-',
                    'NAV Direct Growth (₹)': nav_dir or '-',
                    'NAV Regular Growth (₹)': nav_reg or '-',
                    'NAV Direct IDCW (₹)': nav_idcw_dir or '-',
                    'NAV Regular IDCW (₹)': nav_idcw_reg or '-',
                    'Volatility / Std Dev (%)': volatility or '-',
                    'Beta': beta or '-',
                    'R-Squared': rsquared or '-',
                    'Sharpe Ratio': sharpe or '-',
                    'Portfolio Turnover Ratio': turnover or '-',
                    'TER Regular Plan (%)': ter_reg or '-',
                    'TER Direct Plan (%)': ter_dir or '-',
                    'Exit Load': exit_load or '-',
                    'Factsheet Source File': fname,
                    'PDF Page': pno + 1
                })

        # --- Sector Allocation ---
        if "Sector Allocation" in page_text or "Top 10 Sectors" in page_text or "Industry Allocation" in page_text:
            for b in blocks:
                lines = [l.strip() for l in b[4].split('\n') if l.strip()]
                if len(lines) == 2 and RE_PCT.match(lines[1]):
                    s_name = lines[0]
                    if s_name not in ["Sector", "Value", "Industry", "Total", "% to NAV", "Direct Plan", "Regular Plan", "Growth", "IDCW"]:
                        data['sectors'].append({
                            'Company Name': amc_name,
                            'Scheme Name': current_scheme,
                            'Factsheet Date': factsheet_date,
                            'Sector / Industry': s_name,
                            '% Allocation': lines[1]
                        })

        # --- Portfolio Holdings ---
        if "PORTFOLIO" in page_text or "Issuer Name" in page_text or "% of NAV" in page_text or "% to NAV" in page_text or "Stocks (Top 20" in page_text:
            current_category = "Equity & Equity Related"
            has_derivative_col = "derivative" in page_text.lower() or "hedged" in page_text.lower()
            has_benchmark_col = "benchmark %" in page_text.lower() or "+/-" in page_text

            for b in blocks:
                if b[1] < 90 or b[1] > 780:
                    continue
                lines = [l.strip() for l in b[4].split('\n') if l.strip()]
                idx = 0
                while idx < len(lines):
                    line = lines[idx]
                    if line in ["Instrument Type/Issuer Name", "Company/Issuer", "Issuer Name", "Industry",
                                "Rating", "Industry+ /Rating", "% of NAV", "% to NAV",
                                "% exposure of Derivative", "PORTFOLIO", "Portfolio %", "Benchmark %", "+/- (%)"]:
                        idx += 1
                        continue

                    if line in ["EQUITY", "EQUITY & EQUITY RELATED", "Equity & Equity related", "Equity Shares",
                                "DEBT", "MONEY MARKET", "TREPS", "REITs", "EXCHANGE TRADED FUND",
                                "Cash & Other Receivables", "Derivatives Index / Stock Futures", "GRAND TOTAL", "Total"]:
                        current_category = line
                        idx += 1
                        continue

                    pct_idx = -1
                    for j in range(idx, min(idx + 5, len(lines))):
                        if RE_PCT.match(lines[j]):
                            pct_idx = j
                            break

                    if pct_idx != -1 and pct_idx > idx:
                        issuer = lines[idx]
                        ind_parts = lines[idx+1:pct_idx]
                        ind = " ".join(ind_parts) if ind_parts else "-"
                        pct = lines[pct_idx]

                        bench_pct = "-"
                        weight_diff = "-"
                        deriv_pct = "0.00%"

                        curr = pct_idx + 1
                        if has_benchmark_col:
                            if curr < len(lines) and RE_PCT.match(lines[curr]):
                                bench_pct = lines[curr]
                                curr += 1
                            if curr < len(lines) and RE_PCT.match(lines[curr]):
                                weight_diff = lines[curr]
                                curr += 1
                        elif has_derivative_col:
                            if curr < len(lines) and RE_PCT.match(lines[curr]):
                                deriv_pct = lines[curr]
                                curr += 1

                        idx = curr

                        if len(issuer) > 2 and not any(bad in issuer.lower() for bad in
                            ["fund manager", "exit load", "turnover", "as on", "past performance", "pursuant to", "allotment date"]):
                            data['holdings'].append({
                                'Company Name': amc_name,
                                'Scheme Name': current_scheme,
                                'Factsheet Date': factsheet_date,
                                'Asset Class / Category': current_category,
                                'Company / Issuer Name': issuer,
                                'Industry / Rating': ind,
                                'Portfolio % / % of NAV': pct,
                                'Benchmark %': bench_pct,
                                'Over/Underweight (+/- %)': weight_diff,
                                '% Exposure of Derivative': deriv_pct
                            })
                    else:
                        idx += 1

        # --- Lump Sum Performance ---
        if "PERFORMANCE" in page_text or "CAGR(%)" in page_text or "Returns (%)" in page_text or "Performance Report" in page_text:
            for b in blocks:
                lines = [l.strip() for l in b[4].split('\n') if l.strip()]
                if lines and any(k in lines[0] for k in ["1 Year", "3 Years", "5 Years", "10 Years", "Since Inception"]):
                    pcts = [l for l in lines[1:] if RE_PCT.match(l)]
                    if pcts:
                        data['lumpsum'].append({
                            'Company Name': amc_name,
                            'Scheme Name': current_scheme,
                            'Factsheet Date': factsheet_date,
                            'Period / Horizon': lines[0],
                            'Scheme CAGR (%)': pcts[0],
                            'Scheme Benchmark Returns (%)': pcts[1] if len(pcts) > 1 else "-",
                            'Additional Benchmark Returns (%)': pcts[2] if len(pcts) > 2 else "-",
                            'Value of ₹10,000 Invested - Scheme (₹)': '-',
                            'Value of ₹10,000 Invested - Benchmark (₹)': '-',
                            'Value of ₹10,000 Invested - Addl Benchmark (₹)': '-',
                            'NAV as on Date': '-',
                            'Scheme Benchmark Name': 'Scheme Benchmark',
                            'Additional Benchmark Name': 'Additional Benchmark'
                        })

        # --- SIP Performance ---
        if "SIP Performance" in page_text or "SIP RETURNS" in page_text or "SIP - If you were to invest" in page_text or "SIP Returns" in page_text:
            current_sip_scheme = current_scheme
            for b in blocks:
                lines = [l.strip() for l in b[4].split('\n') if l.strip()]
                if not lines:
                    continue
                # If block mentions a specific scheme or plan
                if any(k in lines[0].upper() for k in ["FUND", "SCHEME"]) and any(p in lines[0].upper() for p in ["REGULAR", "DIRECT", "GROWTH"]):
                    current_sip_scheme = lines[0]
                    continue

                if any(lines[0].lower().startswith(h) for h in ["1 year", "3 year", "5 year", "7 year", "10 year", "15 year", "since inception"]):
                    horizon = lines[0]
                    inv_amt = lines[1] if len(lines) > 1 else "-"
                    fund_ret = lines[2] if len(lines) > 2 else "-"
                    bench_ret = lines[3] if len(lines) > 3 else "-"
                    addl_bench = lines[4] if len(lines) > 4 else "-"
                    mkt_val = lines[5] if len(lines) > 5 else "-"
                    data['sip'].append({
                        'Company Name': amc_name,
                        'Scheme Name': current_sip_scheme,
                        'Factsheet Date': factsheet_date,
                        'SIP Horizon / Period': horizon,
                        'Total Amount Invested (₹)': inv_amt,
                        'Market Value as on Date (₹)': mkt_val,
                        'Fund Return % (XIRR)': fund_ret,
                        'Benchmark Return %': bench_ret,
                        'Additional Benchmark Return %': addl_bench
                    })
                elif len(lines) >= 3 and any(k in lines[0] for k in ["Total Amount Invested", "Market Value", "Fund Return", "Benchmark Return", "SIP"]):
                    desc = lines[0]
                    vals = lines[1:]
                    horizons = ["Since Inception", "15 Years", "10 Years", "7 Years", "5 Years", "3 Years", "1 Year"]
                    for h_idx, v in enumerate(vals[:7]):
                        data['sip'].append({
                            'Company Name': amc_name,
                            'Scheme Name': current_sip_scheme,
                            'Factsheet Date': factsheet_date,
                            'SIP Horizon / Period': horizons[h_idx] if h_idx < len(horizons) else f"Period {h_idx+1}",
                            'Total Amount Invested (₹)': v if "Total Amount" in desc else "-",
                            'Market Value as on Date (₹)': v if "Market Value" in desc or "Mkt Value" in desc else "-",
                            'Fund Return % (XIRR)': v if "Fund Return" in desc else "-",
                            'Benchmark Return %': v if "Benchmark Return" in desc else "-",
                            'Additional Benchmark Return %': "-"
                        })

        # --- IDCW History ---
        if "INCOME DISTRIBUTION" in page_text or "IDCW" in page_text or "DIVIDEND" in page_text:
            for b in blocks:
                full_btext = " ".join([l.strip() for l in b[4].split('\n') if l.strip()])
                date_m = re.search(r'([A-Za-z]{3}\s+\d{2},\s*\d{4}|\d{2}-[A-Za-z]{3}-\d{4})', full_btext)
                if date_m:
                    rec_date = date_m.group(1)
                    nums = re.findall(r'(\d+\.\d+)', full_btext)
                    if len(nums) >= 3:
                        data['idcw'].append({
                            'Company Name': amc_name,
                            'Scheme Name': current_scheme,
                            'Option / Plan': 'IDCW Option',
                            'Record Date': rec_date,
                            'Quantum / Distribution (₹/unit)': nums[0],
                            'Face Value (₹/unit)': nums[1] if len(nums) > 3 else "10.00",
                            'NAV Cum IDCW / Prevailing (₹)': nums[2] if len(nums) > 2 else "-",
                            'NAV Ex IDCW (₹)': nums[3] if len(nums) > 3 else "-",
                            'Direct Plan Quantum (₹/unit)': nums[4] if len(nums) > 5 else "-",
                            'Direct Plan NAV (₹)': nums[6] if len(nums) > 6 else "-"
                        })
    doc.close()


# -----------------------------------------------------------------------
# 5. DATA SANITIZATION & EXCEL WRITER
# -----------------------------------------------------------------------
def sanitize_text(v):
    """Remove illegal XML characters that crash openpyxl."""
    if isinstance(v, str):
        return re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1F\uD800-\uDFFF\uFFFE\uFFFF]', '', v)
    return v

def clean_df(df):
    return df.map(sanitize_text) if not df.empty else df

def apply_executive_formatting(wb):
    """Apply corporate-grade formatting to all sheets."""
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    border_thin = Border(
        left=Side(style='thin', color='E0E0E0'),
        right=Side(style='thin', color='E0E0E0'),
        top=Side(style='thin', color='E0E0E0'),
        bottom=Side(style='thin', color='E0E0E0')
    )

    for sheetname in wb.sheetnames:
        ws = wb[sheetname]
        ws.views.sheetView[0].showGridLines = True
        ws.freeze_panes = 'A2'

        for col in range(1, ws.max_column + 1):
            cell = ws.cell(row=1, column=col)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.row_dimensions[1].height = 28

        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                cell.border = border_thin
                val_str = str(cell.value or '')
                max_len = max(max_len, min(len(val_str), 50))
            ws.column_dimensions[col_letter].width = max(12, max_len + 4)

        ws.auto_filter.ref = ws.dimensions


STANDARD_HEADERS = {
    'Scheme Details & Features': [
        'Company Name', 'Scheme Name', 'Factsheet Date', 'Scheme Category',
        'Fund Manager(s)', 'Date of Allotment', 'Benchmark', 'Net AUM (₹ Cr)',
        'Monthly Average AUM / AAUM (₹ Cr)', 'Plan Available',
        'Minimum Investment Amount', 'NAV Direct Growth (₹)', 'NAV Regular Growth (₹)',
        'NAV Direct IDCW (₹)', 'NAV Regular IDCW (₹)', 'Volatility / Std Dev (%)',
        'Beta', 'R-Squared', 'Sharpe Ratio', 'Portfolio Turnover Ratio',
        'TER Regular Plan (%)', 'TER Direct Plan (%)', 'Exit Load',
        'Factsheet Source File', 'PDF Page'
    ],
    'Portfolio Holdings': [
        'Company Name', 'Scheme Name', 'Factsheet Date', 'Asset Class / Category',
        'Company / Issuer Name', 'Industry / Rating', 'Portfolio % / % of NAV',
        'Benchmark %', 'Over/Underweight (+/- %)', '% Exposure of Derivative'
    ],
    'Sector Allocation': [
        'Company Name', 'Scheme Name', 'Factsheet Date', 'Sector / Industry', '% Allocation'
    ],
    'Lump Sum Performance': [
        'Company Name', 'Scheme Name', 'Factsheet Date', 'Period / Horizon',
        'Scheme CAGR (%)', 'Scheme Benchmark Returns (%)', 'Additional Benchmark Returns (%)',
        'Value of ₹10,000 Invested - Scheme (₹)', 'Value of ₹10,000 Invested - Benchmark (₹)',
        'Value of ₹10,000 Invested - Addl Benchmark (₹)', 'NAV as on Date',
        'Scheme Benchmark Name', 'Additional Benchmark Name'
    ],
    'SIP Performance': [
        'Company Name', 'Scheme Name', 'Factsheet Date', 'SIP Horizon / Period',
        'Total Amount Invested (₹)', 'Market Value as on Date (₹)',
        'Fund Return % (XIRR)', 'Benchmark Return %', 'Additional Benchmark Return %'
    ],
    'IDCW History': [
        'Company Name', 'Scheme Name', 'Option / Plan', 'Record Date',
        'Quantum / Distribution (₹/unit)', 'Face Value (₹/unit)',
        'NAV Cum IDCW / Prevailing (₹)', 'NAV Ex IDCW (₹)',
        'Direct Plan Quantum (₹/unit)', 'Direct Plan NAV (₹)'
    ]
}

def build_sheet_dataframe(records, sheet_name):
    std_cols = STANDARD_HEADERS[sheet_name]
    if records:
        df = clean_df(pd.DataFrame(records).drop_duplicates())
        for col in std_cols:
            if col not in df.columns:
                df[col] = '-'
        return df[std_cols]
    else:
        return pd.DataFrame(columns=std_cols)


def write_excel(data, output_path):
    """Write extracted data to a formatted Excel workbook with all 6 standard sheets guaranteed."""
    df_features = build_sheet_dataframe(data['features'], 'Scheme Details & Features')
    df_holdings = build_sheet_dataframe(data['holdings'], 'Portfolio Holdings')
    df_sectors = build_sheet_dataframe(data['sectors'], 'Sector Allocation')
    df_lumpsum = build_sheet_dataframe(data['lumpsum'], 'Lump Sum Performance')
    df_sip = build_sheet_dataframe(data['sip'], 'SIP Performance')
    df_idcw = build_sheet_dataframe(data['idcw'], 'IDCW History')

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df_features.to_excel(writer, sheet_name='Scheme Details & Features', index=False)
        df_holdings.to_excel(writer, sheet_name='Portfolio Holdings', index=False)
        df_sectors.to_excel(writer, sheet_name='Sector Allocation', index=False)
        df_lumpsum.to_excel(writer, sheet_name='Lump Sum Performance', index=False)
        df_sip.to_excel(writer, sheet_name='SIP Performance', index=False)
        df_idcw.to_excel(writer, sheet_name='IDCW History', index=False)

    # Apply formatting
    wb = openpyxl.load_workbook(output_path)
    apply_executive_formatting(wb)
    wb.save(output_path)

    return {
        'features': len(df_features),
        'holdings': len(df_holdings),
        'sectors': len(df_sectors),
        'lumpsum': len(df_lumpsum),
        'sip': len(df_sip),
        'idcw': len(df_idcw),
    }


# -----------------------------------------------------------------------
# 6. PROCESS A SINGLE DATED FOLDER
# -----------------------------------------------------------------------
def process_dated_folder(date_str):
    """Extract all factsheet data from a single dated folder and produce Excel."""
    date_folder = os.path.join(WORKSPACE, date_str)

    if not os.path.isdir(date_folder):
        print(f"[ERROR] Dated folder does not exist: {date_folder}")
        return False

    print("=" * 68)
    print(f"  FACTSHEET DATA EXTRACTION ENGINE")
    print(f"  Source Folder : {date_folder}")
    print(f"  Date          : {date_str}")
    print("=" * 68)

    # Discover factsheets
    amc_files = find_factsheets_in_folder(date_folder)
    total_amcs = len(amc_files)
    total_pdfs = sum(len(v) for v in amc_files.values())
    print(f"\n[Discovery] Found {total_pdfs} PDFs across {total_amcs} AMCs\n")

    if total_pdfs == 0:
        print("[WARNING] No PDFs found in this folder. Skipping.")
        return False

    # Initialize data containers
    data = {
        'features': [],
        'holdings': [],
        'sectors': [],
        'lumpsum': [],
        'sip': [],
        'idcw': [],
    }

    # Process each AMC
    count = 0
    for amc_name, file_entries in sorted(amc_files.items()):
        for entry in file_entries:
            score, is_large, mtime, fname, fpath = entry
            count += 1
            print(f"  [{count}/{total_pdfs}] {amc_name:35} -> {fname}")
            sys.stdout.flush()
            try:
                if "Mirae Asset" in amc_name and "active-factsheet" in fpath.lower():
                    parse_mirae_asset_pdf(amc_name, fpath, data)
                else:
                    parse_pdf_general(amc_name, fpath, data)
            except Exception as e:
                print(f"    [ERROR] Failed to parse: {e}")

    # Write Excel
    output_excel = os.path.join(date_folder, f"Mutual_Fund_Data_{date_str}.xlsx")
    print(f"\n[Excel] Writing to: {output_excel}")
    sys.stdout.flush()

    counts = write_excel(data, output_excel)

    print(f"\n{'=' * 68}")
    print(f"  EXTRACTION SUMMARY FOR: {date_str}")
    print(f"  Scheme Features   : {counts['features']:,} rows")
    print(f"  Portfolio Holdings : {counts['holdings']:,} rows")
    print(f"  Sector Allocations : {counts['sectors']:,} rows")
    print(f"  Lump Sum Perf     : {counts['lumpsum']:,} rows")
    print(f"  SIP Performance   : {counts['sip']:,} rows")
    print(f"  IDCW History      : {counts['idcw']:,} rows")
    print(f"  Output Excel      : {output_excel}")
    print(f"  File Size         : {os.path.getsize(output_excel)/(1024*1024):.2f} MB")
    print(f"{'=' * 68}\n")

    return True


# -----------------------------------------------------------------------
# 7. MAIN ENTRY POINT
# -----------------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser(description="Mutual Fund Factsheet Data Extraction to Excel")
    parser.add_argument("--date", type=str, default=None,
                        help="Target date folder (YYYY-MM-DD). Defaults to today.")
    parser.add_argument("--all", action="store_true",
                        help="Process ALL dated folders in workspace.")
    args = parser.parse_args()

    if args.all:
        # Find all dated folders (YYYY-MM-DD pattern)
        dated_folders = sorted([
            d for d in os.listdir(WORKSPACE)
            if os.path.isdir(os.path.join(WORKSPACE, d)) and re.match(r'^\d{4}-\d{2}-\d{2}$', d)
        ])
        print(f"[All Mode] Found {len(dated_folders)} dated folders: {dated_folders}")
        for df in dated_folders:
            process_dated_folder(df)
    else:
        target_date = args.date
        if not target_date:
            today_str = datetime.datetime.now().strftime("%Y-%m-%d")
            today_folder = os.path.join(WORKSPACE, today_str)
            if os.path.isdir(today_folder):
                target_date = today_str
            else:
                dated_folders = sorted([
                    d for d in os.listdir(WORKSPACE)
                    if os.path.isdir(os.path.join(WORKSPACE, d)) and re.match(r'^\d{4}-\d{2}-\d{2}$', d)
                ])
                if dated_folders:
                    target_date = dated_folders[-1]
                else:
                    target_date = today_str
        process_dated_folder(target_date)


if __name__ == "__main__":
    main()

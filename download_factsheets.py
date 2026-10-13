import os
import re
import sys
import time
import datetime
import argparse
import requests
from bs4 import BeautifulSoup
import urllib.parse
import urllib3
import logging
import concurrent.futures
from download_registry import DownloadRegistry, normalize_amc_name, normalize_filename

sys.stdout.reconfigure(encoding='utf-8')

# Suppress insecure SSL warnings
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
logging.getLogger("urllib3").setLevel(logging.ERROR)

CUSTOM_URLS = {
    "Axis Mutual Fund": "https://www.axismf.com/knowledge-centre/factsheets",
    "SBI Mutual Fund": "https://www.sbimf.com/en-us/factsheets",
    "HDFC Mutual Fund": "https://www.hdfcfund.com/information/factsheet",
    "Kotak Mahindra Mutual Fund": "https://www.kotakmf.com/Information/factsheet",
    "ICICI Prudential Mutual Fund": "https://www.icicipruamc.com/downloads/monthly-factsheet",
    "Nippon India Mutual Fund": "https://mf.nipponindiaim.com/investor-services/downloads/factsheets",
    "DSP Mutual Fund": "https://www.dspim.com/about-us/mandatory-disclosure/factsheet",
    "UTI Mutual Fund": "https://www.utimf.com/about-us/financials-and-disclosures/factsheet",
    "Bandhan Mutual Fund": "https://bandhanmutual.com/downloads/factsheet",
    "ASK MUTUAL FUND": "https://www.askfinancials.com/ask-mutual-fund/downloads/factsheet",
    "Bank of India Mutual Fund": "https://www.boimf.in/downloads/factsheet",
    "Baroda BNP Paribas Mutual Fund": "https://www.barodabnpparibasmf.in/downloads/factsheet",
    "Edelweiss Mutual Fund": "https://www.edelweissmf.com/downloads/factsheet",
    "Franklin Templeton Mutual Fund": "https://www.franklintempletonindia.com/downloads/factsheet",
    "Jio BlackRock Mutual Fund": "https://www.jioblackrockamc.com/statutory-disclosure/fund-documents/factsheet",
    "Tata Mutual Fund": "https://www.tatamutualfund.com/downloads/factsheet",
    "Mirae Asset Mutual Fund": "https://www.miraeassetmf.co.in/downloads/factsheet"
}

def is_genuine_pdf_url(url):
    """Ensure link is actually pointing to a PDF, not Excel, zip, or HTML pages."""
    try:
        parsed = urllib.parse.urlparse(url)
        path = parsed.path.lower()
        query = parsed.query.lower()
        bad_exts = ('.xls', '.xlsx', '.zip', '.html', '.htm', '.aspx', '.php', '.doc', '.docx', '.rar', '.csv', '.png', '.jpg', '.jpeg', '.gif', '.svg')
        if any(path.endswith(ext) for ext in bad_exts):
            if '.pdf' not in query:
                return False
        return ('.pdf' in path) or ('.pdf' in query)
    except Exception:
        return False

import fitz

def is_valid_pdf_file(filepath):
    """Deep structural integrity verification using PyMuPDF."""
    if not os.path.exists(filepath) or os.path.getsize(filepath) < 1024:
        return False
    try:
        with open(filepath, 'rb') as f:
            if f.read(4) != b'%PDF':
                return False
        doc = fitz.open(filepath)
        if doc.is_encrypted or len(doc) == 0:
            doc.close()
            return False
        doc.close()
        return True
    except Exception:
        return False

def clean_filename_from_url(url):
    url = str(url).strip()
    parsed = urllib.parse.urlparse(url)
    path_name = parsed.path.split('/')[-1].strip()
    
    # If path name is generic or missing (e.g. download.aspx, getfile), check query parameters for .pdf
    if (not path_name or any(path_name.lower().endswith(ext) for ext in ['.aspx', '.php', '.ashx', '.do', '.jsp'])) and '.pdf' in parsed.query.lower():
        params = urllib.parse.parse_qs(parsed.query)
        for _, vals in params.items():
            for v in vals:
                if '.pdf' in v.lower():
                    path_name = v.split('/')[-1].strip()
                    break
                    
    filename = urllib.parse.unquote(path_name).strip()
    while filename.lower().endswith('.pdf'):
        filename = filename[:-4].strip()
    filename += '.pdf'
    filename = re.sub(r'[<>:"/\\|?*]', '_', filename).strip()
    return filename

def download_and_validate_file(url, target_folder, filename):
    try:
        if not url.startswith('http') or not is_genuine_pdf_url(url):
            return False, None
            
        safe_url = urllib.parse.quote(url, safe=":/#?=&")
        filepath = os.path.join(target_folder, filename)
        
        # If file already exists and passes deep structural verification
        if is_valid_pdf_file(filepath):
            return True, filepath
        elif os.path.exists(filepath):
            try: os.remove(filepath)
            except Exception: pass
                
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            'Accept': 'application/pdf,application/octet-stream,*/*'
        }
        
        # Connect timeout: 5s, Read timeout: 60s for large factsheets
        response = requests.get(safe_url, headers=headers, stream=True, verify=False, timeout=(5, 60))
        if response.status_code == 200:
            with open(filepath, 'wb') as out_file:
                for chunk in response.iter_content(chunk_size=32768):
                    out_file.write(chunk)
                    
            if not is_valid_pdf_file(filepath):
                if os.path.exists(filepath):
                    try: os.remove(filepath)
                    except Exception: pass
                return False, None
                
            return True, filepath
        return False, None
    except Exception:
        if os.path.exists(filepath):
            try: os.remove(filepath)
            except Exception: pass
        return False, None

def get_mutual_funds_master():
    url = 'https://www.amfiindia.com/online-center/download-factsheets'
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'}
    try:
        response = requests.get(url, headers=headers, verify=False, timeout=(4, 8))
        chunks = re.findall(r'self\.__next_f\.push\(\[1,\"(.*?)\"\]\)', response.text)
        full_string = ""
        for chunk in chunks:
            try:
                full_string += bytes(chunk, "utf-8").decode("unicode_escape")
            except:
                full_string += chunk
        matches = re.findall(r'\"mf_name\":\"(.*?)\".*?\"amc_monthly_mf_factsheets\":\"(.*?)\"', full_string)
        if matches:
            return list(set(matches))
    except Exception:
        pass
        
    fallback_file = os.path.join(r"e:\mutual_fund", "full_string.txt")
    if os.path.exists(fallback_file):
        with open(fallback_file, "r", encoding="utf-8") as f:
            content = f.read()
        matches = re.findall(r'\"mf_name\":\"(.*?)\".*?\"amc_monthly_mf_factsheets\":\"(.*?)\"', content)
        return list(set(matches))
    return []

def is_factsheet_candidate(url, text=""):
    """Accurately identify genuine factsheet documents and filter out KYC, notices, and forms."""
    u_lower = url.lower()
    t_lower = text.lower()
    combined = f"{u_lower} {t_lower}"
    
    # Negative filters: Eliminate waste files
    noise_keywords = [
        'kyc', 'ckyc', 'nomination', 'nominee', 'grievance', 'awareness',
        'addendum', 'notice', 'change-in-kmp', 'change_in_kmp', 'change-in-isc',
        'caution', 'empanelment', 'application-form', 'application_form',
        'transaction-slip', 'transaction_slip', 'unaudited', 'annual-report',
        'annual_report', 'unclaimed', 'non-business-day', 'redressal', 'sai', 'sid'
    ]
    if any(k in combined for k in noise_keywords):
        return False
        
    normalized = re.sub(r'[\-_/.:\s]+', ' ', combined)
    # Positive factsheet keywords
    factsheet_keywords = [
        'factsheet', 'fact sheet', 'fact_sheet', 'fact-sheet',
        'fund facts', 'fund factsheet', 'fund_facts', 'fund performance',
        'performance report', 'fund performance report',
        'monthly portfolio', 'fundbook', 'fund-book', 'fund book',
        'portfolio disclosure', 'monthly facts', 'snapshot', 'fund-note',
        'fund note', 'product-note', 'fund_book'
    ]
    return any(k in combined or k in normalized for k in factsheet_keywords)

def extract_pdf_urls_from_page(page_html, base_url):
    found_urls = []
    soup = BeautifulSoup(page_html, 'html.parser')
    for a in soup.find_all('a', href=True):
        href = a['href']
        text = a.text
        full_url = urllib.parse.urljoin(base_url, href)
        if is_genuine_pdf_url(full_url) and is_factsheet_candidate(full_url, text):
            found_urls.append(full_url)
            
    for u in re.findall(r'https?://[^\s\"\'\<\>\[\]\{\}]+\.pdf', page_html, re.I):
        if is_genuine_pdf_url(u) and is_factsheet_candidate(u):
            found_urls.append(u)
        
    for rel_pdf in re.findall(r'(?:/|\\/)[^\s\"\'\<\>\[\]\{\}]+\.pdf', page_html, re.I):
        clean_rel = rel_pdf.replace('\\/', '/')
        full_url = urllib.parse.urljoin(base_url, clean_rel)
        if is_genuine_pdf_url(full_url) and is_factsheet_candidate(full_url):
            found_urls.append(full_url)
        
    return list(set(found_urls))

def process_amc(amc_name, amfi_link, registry):
    target_link = CUSTOM_URLS.get(amc_name, amfi_link)
    if not target_link:
        return amc_name, 0, 0, 0
        
    fs_target_dir = registry.get_amc_folder(amc_name, category="Factsheets")
    
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'}
    page_html = ""
    try:
        r = requests.get(target_link, headers=headers, verify=False, timeout=(4, 8))
        if r.status_code == 200:
            page_html = r.text
    except Exception:
        pass
        
    pdf_links = extract_pdf_urls_from_page(page_html, target_link)
    
    total_found = len(pdf_links)
    skipped_existing = 0
    newly_downloaded = 0
    
    for url in pdf_links:
        filename = clean_filename_from_url(url)
        
        # --- DEDUPLICATION CHECK ---
        # Checks all previous dated folders (e.g. 2026-10-10) and registry!
        is_downloaded, existing_path = registry.is_already_downloaded(amc_name, filename, url)
        if is_downloaded:
            skipped_existing += 1
            continue
            
        ok, downloaded_path = download_and_validate_file(url, fs_target_dir, filename)
        if ok:
            registry.register_download(amc_name, filename, downloaded_path, url)
            newly_downloaded += 1
            print(f"  [{amc_name}] NEW FACTSHEET DOWNLOADED: {filename} ({os.path.getsize(downloaded_path)/(1024*1024):.2f} MB)", flush=True)
            
    return amc_name, total_found, skipped_existing, newly_downloaded

def run_incremental_download(target_date=None):
    date_str = target_date or datetime.datetime.now().strftime("%Y-%m-%d")
    print(f"==================================================================", flush=True)
    print(f"MUTUAL FUND MONTHLY FACTSHEET INCREMENTAL DOWNLOADER", flush=True)
    print(f"Run Date Folder: {date_str}", flush=True)
    print(f"==================================================================", flush=True)
    
    registry = DownloadRegistry(workspace_dir=r"e:\mutual_fund", current_date_str=date_str)
    funds = get_mutual_funds_master()
    print(f"[Master] Found {len(funds)} AMC companies to scan online.\n", flush=True)
    
    total_scanned = 0
    total_skipped = 0
    total_new = 0
    
    with concurrent.futures.ThreadPoolExecutor(max_workers=8) as executor:
        future_to_amc = {executor.submit(process_amc, name, link, registry): name for name, link in funds}
        done_count = 0
        for future in concurrent.futures.as_completed(future_to_amc):
            done_count += 1
            try:
                amc_name, f_cnt, s_cnt, n_cnt = future.result()
                total_scanned += f_cnt
                total_skipped += s_cnt
                total_new += n_cnt
                status_str = f"{n_cnt} new downloaded, {s_cnt} already present (skipped)" if n_cnt > 0 else f"Up to date ({s_cnt} existing skipped)"
                print(f"[{done_count}/{len(funds)}] {amc_name:35} -> {status_str}", flush=True)
            except Exception as e:
                amc_name = future_to_amc[future]
                print(f"[{done_count}/{len(funds)}] {amc_name:35} -> Error scanning: {e}", flush=True)
                
    registry.save_registry_state()
    
    print("\n==================================================================", flush=True)
    print(f"INCREMENTAL DOWNLOAD SUMMARY FOR: {date_str}", flush=True)
    print(f"Total Online Links Scanned     : {total_scanned}", flush=True)
    print(f"Already Downloaded (Skipped)   : {total_skipped}", flush=True)
    print(f"Brand New Factsheets Downloaded: {total_new}", flush=True)
    print(f"Target Output Folder           : e:\\mutual_fund\\{date_str}\\", flush=True)
    print(f"==================================================================", flush=True)
    return total_new

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Incremental Mutual Fund Factsheet Downloader")
    parser.add_argument("--date", type=str, default=None, help="Target date folder (YYYY-MM-DD)")
    args = parser.parse_args()
    run_incremental_download(target_date=args.date)

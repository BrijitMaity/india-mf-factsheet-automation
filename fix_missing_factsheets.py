import os
import re
import sys
import glob
import datetime
import argparse
import requests
from bs4 import BeautifulSoup
import urllib.parse
import urllib3
from download_registry import DownloadRegistry, normalize_amc_name, normalize_filename

sys.stdout.reconfigure(encoding='utf-8')
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

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

def download_and_validate_file(url, folder, filename):
    try:
        if not url.startswith('http'):
            return False, None
            
        filepath = os.path.join(folder, filename)
        if is_valid_pdf_file(filepath):
            return True, filepath
        elif os.path.exists(filepath):
            try: os.remove(filepath)
            except Exception: pass
            
        headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'}
        response = requests.get(url, headers=headers, stream=True, verify=False, timeout=(5, 45))
        
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

def fix_missing(target_date=None):
    date_str = target_date or datetime.datetime.now().strftime("%Y-%m-%d")
    print(f"==================================================================", flush=True)
    print(f"AGGREGATOR FALLBACK SCAN FOR MISSING FACTSHEETS", flush=True)
    print(f"Target Date Folder: {date_str}", flush=True)
    print(f"==================================================================", flush=True)
    
    registry = DownloadRegistry(workspace_dir=r"e:\mutual_fund", current_date_str=date_str)
    
    # Check all AMCs in the workspace (from previous folders or full_string.txt)
    fallback_file = os.path.join(r"e:\mutual_fund", "full_string.txt")
    all_amcs = set()
    if os.path.exists(fallback_file):
        with open(fallback_file, "r", encoding="utf-8") as f:
            content = f.read()
        for name in re.findall(r'\"mf_name\":\"(.*?)\"', content):
            all_amcs.add(name)
            
    # Also add AMCs from any dated folders in the workspace
    workspace_dir = r"e:\mutual_fund"
    for df in os.listdir(workspace_dir):
        p = os.path.join(workspace_dir, df)
        if os.path.isdir(p) and re.match(r'^\d{4}-\d{2}-\d{2}$', df):
            for d in os.listdir(p):
                if os.path.isdir(os.path.join(p, d)):
                    all_amcs.add(d)
                    
    print(f"[Fallback] Checking {len(all_amcs)} AMCs for missing factsheet coverage...", flush=True)
    
    total_new = 0
    total_skipped = 0
    
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'}
    
    for amc_name in sorted(all_amcs):
        norm_amc = normalize_amc_name(amc_name)
        existing_count = len(registry.cache_by_amc.get(norm_amc, {}))
        
        # Only query aggregator if AMC has very few/zero factsheets
        if existing_count < 2:
            print(f"[{amc_name}] Only {existing_count} factsheets found. Querying Advisorkhoj aggregator...", flush=True)
            search_name = amc_name.replace(" ", "-").replace("&", "").replace("(", "").replace(")", "").strip("-")
            variations = [
                search_name,
                search_name.replace("MUTUAL-FUND", "Mutual-Fund"),
                amc_name.replace(" ", "-")
            ]
            
            dest_dir = registry.get_amc_folder(amc_name, category="Factsheets")
            amc_downloaded = 0
            
            current_year = datetime.datetime.now().year
            years_to_check = [str(current_year), str(current_year - 1), str(current_year - 2)]
            
            for var in set(variations):
                for year in years_to_check:
                    url = f"https://www.advisorkhoj.com/mutual-funds-research/factsheet/{var}/{year}"
                    try:
                        r = requests.get(url, headers=headers, timeout=(4, 8))
                        if r.status_code == 200:
                            soup = BeautifulSoup(r.text, 'html.parser')
                            for a in soup.find_all('a', href=True):
                                href = a['href']
                                if '.pdf' in href.lower() or 'download' in href.lower():
                                    raw_fn = href.split('/')[-1].split('?')[0]
                                    raw_fn = urllib.parse.unquote(raw_fn)
                                    if not raw_fn.lower().endswith('.pdf'):
                                        raw_fn += '.pdf'
                                    clean_fn = re.sub(r'[<>:"/\\|?*]', '_', raw_fn).strip()
                                    
                                    # Deduplication check
                                    is_down, existing_path = registry.is_already_downloaded(amc_name, clean_fn, href)
                                    if is_down:
                                        total_skipped += 1
                                        continue
                                        
                                    ok, fpath = download_and_validate_file(href, dest_dir, clean_fn)
                                    if ok:
                                        registry.register_download(amc_name, clean_fn, fpath, href)
                                        amc_downloaded += 1
                                        total_new += 1
                                        print(f"  [{amc_name}] Downloaded from aggregator: {clean_fn}", flush=True)
                    except Exception:
                        pass
                if amc_downloaded > 0:
                    break
                    
    if total_new > 0:
        registry.save_registry_state()
        
    print(f"Fallback check completed. New: {total_new}, Skipped existing: {total_skipped}", flush=True)
    return total_new

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Aggregator fallback for missing factsheets")
    parser.add_argument("--date", type=str, default=None, help="Target date folder (YYYY-MM-DD)")
    args = parser.parse_args()
    fix_missing(target_date=args.date)

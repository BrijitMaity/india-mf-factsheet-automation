import os
import re
import json
import datetime
import urllib.parse
import hashlib

REGISTRY_FILENAME = "download_registry.json"

def normalize_amc_name(name):
    """Normalize AMC name to alphanumeric lowercase for robust cross-folder matching."""
    return re.sub(r'[^a-zA-Z0-9]', '', str(name)).lower()

def normalize_filename(fname):
    """Normalize filename for comparison (lowercase, trimmed, strip extension)."""
    clean = urllib.parse.unquote(str(fname)).lower().strip()
    while clean.endswith('.pdf'):
        clean = clean[:-4].strip()
    # Replace separators with standard underscore
    clean = re.sub(r'[\s\-_]+', '_', clean)
    return clean

class DownloadRegistry:
    def __init__(self, workspace_dir=r"e:\mutual_fund", current_date_str=None):
        self.workspace_dir = os.path.abspath(workspace_dir)
        if current_date_str:
            self.current_date = current_date_str
        else:
            self.current_date = datetime.datetime.now().strftime("%Y-%m-%d")
            
        self.registry_file = os.path.join(self.workspace_dir, REGISTRY_FILENAME)
        self.cache_by_amc = {} # norm_amc -> set of (normalized_filename, raw_filename, path)
        self.url_cache = set()
        
        self._load_registry()
        self._scan_all_previous_folders()

    def _load_registry(self):
        """Loads cached metadata from disk if available."""
        if os.path.exists(self.registry_file):
            try:
                with open(self.registry_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    for item in data.get("downloads", []):
                        norm_amc = item.get("norm_amc")
                        fn = item.get("filename", "")
                        url = item.get("url", "")
                        if url:
                            self.url_cache.add(url.strip())
                        if norm_amc:
                            if norm_amc not in self.cache_by_amc:
                                self.cache_by_amc[norm_amc] = {}
                            self.cache_by_amc[norm_amc][normalize_filename(fn)] = item.get("filepath", "")
            except Exception as e:
                print(f"[Registry] Warning: Could not read existing registry: {e}")

    def _scan_all_previous_folders(self):
        """Scans all dated directories in the workspace to catalog every existing PDF."""
        entries = os.listdir(self.workspace_dir)
        dated_folders = []
        for e in entries:
            full_p = os.path.join(self.workspace_dir, e)
            if os.path.isdir(full_p) and re.match(r'^\d{4}-\d{2}-\d{2}$', e):
                # We check all folders that are NOT the currently active run folder
                # (and also scan active folder to handle resumable runs)
                dated_folders.append(e)
                
        dated_folders.sort()
        count = 0
        for df in dated_folders:
            df_path = os.path.join(self.workspace_dir, df)
            for amc in os.listdir(df_path):
                amc_path = os.path.join(df_path, amc)
                if not os.path.isdir(amc_path):
                    continue
                norm_amc = normalize_amc_name(amc)
                if norm_amc not in self.cache_by_amc:
                    self.cache_by_amc[norm_amc] = {}
                    
                for sub in ["Factsheets", "Forms_and_Information"]:
                    sub_p = os.path.join(amc_path, sub)
                    if os.path.exists(sub_p):
                        for f in os.listdir(sub_p):
                            if f.lower().endswith('.pdf'):
                                fpath = os.path.join(sub_p, f)
                                if os.path.getsize(fpath) > 0:
                                    norm_fn = normalize_filename(f)
                                    # Save first seen path (or prefer older date)
                                    if norm_fn not in self.cache_by_amc[norm_amc]:
                                        self.cache_by_amc[norm_amc][norm_fn] = fpath
                                        count += 1
                                        
        total_tracked = sum(len(v) for v in self.cache_by_amc.values())
        print(f"[Registry] Cataloged {total_tracked} total unique PDFs ({count} new from disk scan) across {len(self.cache_by_amc)} AMCs from folders: {dated_folders}")

    def is_already_downloaded(self, amc_name, filename, url=None):
        """
        Returns (True, filepath) if this PDF was already downloaded in any previous or current folder.
        Returns (False, None) if this is a new factsheet that should be downloaded.
        """
        # 1. URL check
        if url and url.strip() in self.url_cache:
            return True, "Cached by URL"

        norm_amc = normalize_amc_name(amc_name)
        norm_fn = normalize_filename(filename)

        # 2. Check exact AMC cache
        if norm_amc in self.cache_by_amc:
            if norm_fn in self.cache_by_amc[norm_amc]:
                return True, self.cache_by_amc[norm_amc][norm_fn]

        # 3. Fuzzy AMC match (e.g., 'Axis Mutual Fund' vs 'Axis')
        if len(norm_amc) >= 4:
            for stored_amc, files in self.cache_by_amc.items():
                if len(stored_amc) >= 4 and (norm_amc in stored_amc or stored_amc in norm_amc):
                    if norm_fn in files:
                        return True, files[norm_fn]

        return False, None

    def register_download(self, amc_name, filename, filepath, url=None):
        """Registers a newly downloaded file so future checks immediately recognize it."""
        norm_amc = normalize_amc_name(amc_name)
        norm_fn = normalize_filename(filename)
        
        if norm_amc not in self.cache_by_amc:
            self.cache_by_amc[norm_amc] = {}
        self.cache_by_amc[norm_amc][norm_fn] = filepath
        
        if url:
            self.url_cache.add(url.strip())

    def unregister_file(self, filepath):
        """Removes a file from in-memory cache if purged/deleted."""
        norm_fp = os.path.abspath(filepath).lower()
        for norm_amc, file_map in list(self.cache_by_amc.items()):
            for norm_fn, p in list(file_map.items()):
                if os.path.abspath(p).lower() == norm_fp:
                    del file_map[norm_fn]
                    return True
        return False

    def get_amc_folder(self, amc_name, category="Factsheets"):
        """Returns the destination folder inside the current date root."""
        # Sanitize AMC folder name
        safe_name = "".join([c for c in amc_name if c.isalnum() or c in (' ', '_', '-')]).strip()
        target_dir = os.path.join(self.workspace_dir, self.current_date, safe_name, category)
        if not os.path.exists(target_dir):
            os.makedirs(target_dir, exist_ok=True)
        return target_dir

    def save_registry_state(self):
        """Persists cataloged files to download_registry.json."""
        records = []
        for norm_amc, file_map in self.cache_by_amc.items():
            for norm_fn, fpath in file_map.items():
                records.append({
                    "norm_amc": norm_amc,
                    "filename": os.path.basename(fpath),
                    "filepath": fpath
                })
        try:
            with open(self.registry_file, 'w', encoding='utf-8') as f:
                json.dump({"updated_at": datetime.datetime.now().isoformat(), "total_files": len(records), "downloads": records}, f, indent=2)
            print(f"[Registry] Saved state with {len(records)} entries to {REGISTRY_FILENAME}")
        except Exception as e:
            print(f"[Registry] Warning: Could not save registry: {e}")

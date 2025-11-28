import pandas as pd
import numpy as np
import requests
import re
import os
import io
import warnings
import base64
import sys
from datetime import datetime
from collections import defaultdict
from pathlib import Path

from cryptography.fernet import Fernet, InvalidToken
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC

from config import (
    INTENDED_COLS, COLUMN_ALIASES, ENCRYPTED_GITHUB_TOKEN,
    HIDDEN_PASSWORD, ENCRYPTION_SALT, unprotect_sensitive_data
)

class InvalidFileTypeError(Exception):
    pass

def get_encryption_key(password: str) -> bytes:
    kdf = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=ENCRYPTION_SALT, iterations=480000)
    return base64.urlsafe_b64encode(kdf.derive(password.encode()))

def encrypt_data(data: bytes, key: bytes) -> bytes:
    return Fernet(key).encrypt(data)

def decrypt_data(encrypted_data: bytes, key: bytes) -> bytes:
    try:
        return Fernet(key).decrypt(encrypted_data)
    except InvalidToken:
        # Try legacy decryption with original password
        try:
            from cryptography.hazmat.primitives import hashes
            from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
            kdf = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=ENCRYPTION_SALT, iterations=480000)
            legacy_key = base64.urlsafe_b64encode(kdf.derive("doleadmin".encode()))
            return Fernet(legacy_key).decrypt(encrypted_data)
        except Exception:
            raise InvalidToken("Decryption failed. The cache file may be corrupt or from an old version.")

def normalize_col(col):
    return str(col).strip().replace("_", "").replace(" ", "").lower()

def clean_str(s):
    if not isinstance(s, str): return ""
    return s.strip().lower()

def normalize_name(val):
    if pd.isna(val) or not isinstance(val, str) or not val.strip(): return ""
    cleaned_val = re.sub(r'\s+', ' ', val.replace(',', ' ')).strip()
    try:
        if float(cleaned_val) == int(float(cleaned_val)): cleaned_val = str(int(float(cleaned_val)))
    except (ValueError, TypeError): pass
    return " ".join([p.capitalize() if p.lower() != 'ma.' else 'Ma.' for p in cleaned_val.split()])

def normalize_date(val):
    if pd.isna(val) or str(val).strip() == "": return ""
    if isinstance(val, (datetime, pd.Timestamp)): return val.strftime("%Y-%m-%d")
    v = str(val).strip()
    
    # Handle datetime strings with time components (like "10/02/198812:00:00am")
    # Extract just the date part by removing time components
    if re.search(r'\d{1,2}:\d{2}:\d{2}(am|pm)?', v, re.IGNORECASE):
        # Remove time part (12:00:00am, 12:00:00pm, etc.)
        v = re.sub(r'\d{1,2}:\d{2}:\d{2}(am|pm)?', '', v, flags=re.IGNORECASE).strip()
    
    formats_to_try = [
        '%d/%m/%Y',
        '%m/%d/%Y',
        '%Y-%m-%d',
        '%Y/%m/%d',
        '%m-%d-%Y',
        '%d-%m-%Y'
    ]
    for fmt in formats_to_try:
        try: return datetime.strptime(v, fmt).strftime("%Y-%m-%d")
        except (ValueError, TypeError): continue
    try:
        with warnings.catch_warnings():
            warnings.simplefilter("ignore", UserWarning)
            parsed_date = pd.to_datetime(v, errors="coerce")
            if pd.notna(parsed_date):
                return parsed_date.strftime("%Y-%m-%d")
    except (ValueError, TypeError, AttributeError): 
        pass
    return ""

def normalize_sex(val):
    """Standardizes variations of sex/gender into 'Male', 'Female', or empty."""
    if pd.isna(val) or not isinstance(val, str) or not val.strip():
        return ""
    
    sex_clean = val.strip().lower()
    
    if sex_clean.startswith('m'):
        return "Male"
    elif sex_clean.startswith('f'):
        return "Female"
    
    return ""

def normalize_city(val):
    """Normalizes city names to prevent detection evasion through city variations."""
    if pd.isna(val) or not isinstance(val, str) or not val.strip():
        return ""
    
    city_clean = val.strip().lower()
    
    # Remove common suffixes and prefixes
    city_clean = re.sub(r'\b(city|municipality|municipal|town|mun\.?|barangay|brgy|bgy|lgu)\b', '', city_clean).strip()
    
    # Remove common separators and normalize spacing
    city_clean = re.sub(r'[,\-_\.\s]+', ' ', city_clean).strip()
    
    # Handle common abbreviations and variations
    abbreviations = {
        'st': 'saint',
        'sto': 'santo', 
        'sta': 'santa',
        'san': 'san',
        'mt': 'mount',
        'n': 'north',
        'e': 'east', 
        's': 'south',
        'w': 'west'
    }
    
    words = city_clean.split()
    normalized_words = []
    
    for word in words:
        # Expand abbreviations
        if word in abbreviations:
            normalized_words.append(abbreviations[word])
        else:
            normalized_words.append(word)
    
    # Join and capitalize properly
    result = ' '.join(normalized_words)
    
    # Remove leading connector left from patterns like "City of X" / "Municipality of X"
    result = re.sub(r'^of\s+', '', result).strip()
    
    # Remove parenthetical descriptors e.g., "(Capital)", "(City Proper)"
    result = re.sub(r'\s*\([^)]*\)\s*', ' ', result).strip()
    # Collapse extra spaces after removals
    result = re.sub(r'\s+', ' ', result)
    
    # Handle special cases for common city name variations across MIMAROPA provinces
    special_cases = {
        # Oriental Mindoro
        'calapan': 'calapan',
        'calapancity': 'calapan',
        'calapan city': 'calapan',
        'baco': 'baco',
        'naujan': 'naujan',
        'victoria': 'victoria',
        'socorro': 'socorro',
        'pinamalayan': 'pinamalayan',
        'gloria': 'gloria',
        'pola': 'pola',
        'bulalacao': 'bulalacao',
        'mansalay': 'mansalay',
        'roxas': 'roxas',
        'bansud': 'bansud',
        'bongabong': 'bongabong',
        
        # Occidental Mindoro
        'mamburao': 'mamburao',
        'san jose': 'san jose',
        'sanjose': 'san jose',
        'lubang': 'lubang',
        'looc': 'looc',
        'magsaysay': 'magsaysay',
        'paluan': 'paluan',
        'rizal': 'rizal',
        'sablayan': 'sablayan',
        'santa cruz': 'santa cruz',
        'santacruz': 'santa cruz',
        'calintaan': 'calintaan',
        'abra de ilog': 'abra de ilog',
        
        # Marinduque
        'boac': 'boac',
        'buenavista': 'buenavista',
        'gasan': 'gasan',
        'mogpog': 'mogpog',
        'santa cruz marinduque': 'santa cruz',
        'torrijos': 'torrijos',
        
        # Romblon
        'romblon': 'romblon',
        'alcantara': 'alcantara',
        'banton': 'banton',
        'cajidiocan': 'cajidiocan',
        'calatrava': 'calatrava',
        'concepcion': 'concepcion',
        'corcuera': 'corcuera',
        'ferrol': 'ferrol',
        'looc romblon': 'looc',
        'magdiwang': 'magdiwang',
        'odiongan': 'odiongan',
        'san agustin': 'san agustin',
        'sanagustin': 'san agustin',
        'san andres': 'san andres',
        'sanandres': 'san andres',
        'san fernando': 'san fernando',
        'sanfernando': 'san fernando',
        'san jose romblon': 'san jose',
        'santa fe': 'santa fe',
        'santafe': 'santa fe',
        'santa maria': 'santa maria',
        'santamaria': 'santa maria',
        
        # Palawan
        'puerto princesa': 'puerto princesa',
        'puerto princesa city': 'puerto princesa',
        'aborlan': 'aborlan',
        'agutaya': 'agutaya',
        'araceli': 'araceli',
        'balabac': 'balabac',
        'bataraza': 'bataraza',
        'brookes point': 'brookes point',
        'brookespoint': 'brookes point',
        'busuanga': 'busuanga',
        'cagayancillo': 'cagayancillo',
        'coron': 'coron',
        'culion': 'culion',
        'dumaran': 'dumaran',
        'el nido': 'el nido',
        'elnido': 'el nido',
        'linapacan': 'linapacan',
        'magsaysay palawan': 'magsaysay',
        'narra': 'narra',
        'quezon palawan': 'quezon',
        'rizal palawan': 'rizal',
        'roxas palawan': 'roxas',
        'san vicente': 'san vicente',
        'sanvicente': 'san vicente',
        'sofronio espanola': 'sofronio espanola',
        'taytay': 'taytay',
        
        # Common variations
        'santo tomas': 'santo tomas',
        'santotomas': 'santo tomas',
        'san antonio': 'san antonio',
        'sanantonio': 'san antonio',
        'san miguel': 'san miguel',
        'sanmiguel': 'san miguel',
        'san pedro': 'san pedro',
        'sanpedro': 'san pedro',
        'san pablo': 'san pablo',
        'sanpablo': 'san pablo',
        'san juan': 'san juan',
        'sanjuan': 'san juan',
        'san carlos': 'san carlos',
        'sancarlos': 'san carlos',
        'san rafael': 'san rafael',
        'sanrafael': 'san rafael',
        'san luis': 'san luis',
        'sanluis': 'san luis'
    }
    
    if result in special_cases:
        result = special_cases[result]
    
    return result.title() if result else ""


def normalize_batch_name(val):
    """Preserve the original batch name text from the source.
    Only trims and collapses whitespace/newlines so values like
    'Caloocan 9031' remain exactly as provided (no forced 'Batch <n>').
    """
    if pd.isna(val) or not str(val).strip():
        return ""
    s = str(val)
    # Normalize whitespace: remove tabs/newlines, collapse multiple spaces
    s = re.sub(r'[\t\r\n]+', ' ', s)
    s = re.sub(r'\s+', ' ', s).strip()
    return s

def smart_remap_columns_to_intended(df, is_officials_file=False):
    normalized_cols = {normalize_col(c): c for c in df.columns}
    has_first = any(alias in normalized_cols for alias in COLUMN_ALIASES["First Name"])
    has_last = any(alias in normalized_cols for alias in COLUMN_ALIASES["Last Name"])
    col_map = {}
    base_cols = {"First Name", "Middle Name", "Last Name", "Suffix", "Birthdate", "City", "Sex", "Contact Number"}
    # Include Batch Name as an optional informational column
    base_cols.update({"Batch Name"})
    if is_officials_file: base_cols.update({"Position", "Barangay"})
    aliases_to_map = set(base_cols)
    if not (has_first and has_last): aliases_to_map.add("FullName")
    for group in aliases_to_map:
        target = "First Name" if group == "FullName" else group
        for alias in COLUMN_ALIASES.get(group, []):
            if alias in normalized_cols and normalized_cols[alias] not in col_map:
                col_map[normalized_cols[alias]] = target
                break
    df = df.rename(columns=col_map)
    final_cols = list(INTENDED_COLS)
    # Append Batch Name if present among columns or mapped
    if "Batch Name" in df.columns:
        final_cols.append("Batch Name")
    if is_officials_file: final_cols.extend(["Position", "Barangay"])
    for col in final_cols:
        if col not in df.columns: df[col] = ""
    return df[final_cols].copy()

def parse_full_name_column(df):
    SUFFIXES = {"jr", "sr", "i", "ii", "iii", "iv", "v"}
    mask = (df['First Name'].notna()) & (df['First Name'].str.contains(' |,', na=False)) & (df['Last Name'].fillna('').str.strip() == '')
    for i in df[mask].index:
        name_str = str(df.at[i, "First Name"]).strip()
        if ',' in name_str:
            parts = name_str.split(',')
            last_name = parts[0].strip()
            other_parts = parts[1].strip().split() if len(parts) > 1 else []
            first_name = other_parts.pop(0) if other_parts else ""
            suffix = other_parts.pop(-1).upper().replace(".", "") if other_parts and other_parts[-1].lower().replace(".", "") in SUFFIXES else ""
            middle_name = " ".join(other_parts)
        else:
            parts = name_str.split()
            suffix = parts.pop(-1).upper().replace(".", "") if len(parts) > 1 and parts[-1].lower().replace(".", "") in SUFFIXES else ""
            last_name = parts.pop(-1) if parts else ""
            first_name = parts.pop(0) if parts else ""
            middle_name = " ".join(parts)
        df.loc[i, ["First Name", "Middle Name", "Last Name", "Suffix"]] = [first_name, middle_name, last_name, suffix]
    return df

def get_auth_headers(key):
    try:
        decrypted_token = unprotect_sensitive_data(ENCRYPTED_GITHUB_TOKEN)
        return {'Authorization': f'token {decrypted_token}'}
    except Exception:
        return None

def check_internet(timeout=3):
    try:
        requests.get("https://www.google.com", timeout=timeout)
        return True
    except (requests.ConnectionError, requests.Timeout):
        return False

def smart_download_pat(url, local_path, meta_path, auth_headers, encryption_key, encrypt_locally=True):
    headers = auth_headers.copy() if auth_headers else {}
    if os.path.exists(meta_path):
        try:
            with open(meta_path, "r") as f:
                etag = f.read().strip()
                if etag: headers['If-None-Match'] = etag
        except Exception: pass
    try:
        r = requests.get(url, headers=headers, timeout=10)
        if r.status_code == 304: return 'UP_TO_DATE', "Cache is up-to-date."
        if r.status_code == 200:
            content_to_save = r.content
            if encrypt_locally:
                content_to_save = encrypt_data(content_to_save, encryption_key)

            with open(local_path, "wb") as f: f.write(content_to_save)

            new_etag = r.headers.get('ETag')
            if new_etag:
                with open(meta_path, 'w') as f: f.write(new_etag)
            return 'UPDATED', "Cache updated successfully."
        return 'HTTP_ERROR', f"File not found on server (HTTP {r.status_code}). Check token or repo permissions."
    except requests.exceptions.RequestException as e:
        return 'NETWORK_ERROR', f"Could not check for updates: {e}."

def download_file_with_logging(log_callback, file_type, download_func, *args, **kwargs):
    status, message = download_func(*args, **kwargs)
    prefix = {'UP_TO_DATE': '✅', 'UPDATED': '✅', 'NETWORK_ERROR': '⚠️', 'HTTP_ERROR': '❌'}.get(status, 'ℹ️')
    log_callback(f"{prefix} [{file_type}] {message}")

def update_remote_files(app_data, encryption_key, log_callback, province_urls, global_urls):
    auth_headers = get_auth_headers(encryption_key)
    if not auth_headers:
        log_callback("❌ CRITICAL: Could not decrypt GitHub token. Cannot check for updates.")
        return False

    files_to_check = {
        "MasterDB": (province_urls["master_db"], app_data.master_db_path, app_data.master_db_meta_path),
        "OfficialsDB": (province_urls["officials"], app_data.officials_db_path, app_data.officials_db_meta_path),
        "Nickname": (global_urls["NICKNAME_CSV_URL"], app_data.nickname_path, app_data.nickname_meta_path)
    }
    if check_internet():
        log_callback("🌐 Internet found. Preparing to check remote files...")
        for name, (url, path, meta_path) in files_to_check.items():
            download_file_with_logging(log_callback, name, smart_download_pat, url, path, meta_path, auth_headers, encryption_key, encrypt_locally=True)
        return True
    else:
        log_callback("⚠️ No internet connection detected. Proceeding with local caches.")
        for name, (_, path, _) in files_to_check.items():
            if path.exists(): log_callback(f"✅ [{name}] Using local cache.")
            else: log_callback(f"❌ [{name}] No internet and no local cache found.")
        return False

def load_nickname_map(app_data, encryption_key, log_callback=None):
    try:
        if not os.path.exists(app_data.nickname_path): return {}
        with open(app_data.nickname_path, "rb") as f: content_bytes = decrypt_data(f.read(), encryption_key)
        
        original_stderr = sys.stderr; sys.stderr = open(os.devnull, 'w')
        try: df = pd.read_csv(io.StringIO(content_bytes.decode(errors='ignore')))
        finally: sys.stderr.close(); sys.stderr = original_stderr
        
        if not {'nickname', 'formal_name'}.issubset(df.columns):
            if log_callback: log_callback("❌ [Nickname] File has incorrect format.")
            return {}
        nickname_map = defaultdict(list)
        for _, row in df.iterrows():
            nick, formal = clean_str(row['nickname']), clean_str(row['formal_name'])
            if nick and formal: nickname_map[nick].append(formal)
        return dict(nickname_map)
    except Exception as e:
        if log_callback: log_callback(f"❌ [Nickname] Failed to load map: {e}")
        return {}

# --- FUNCTION RESTORED ---
def load_raw_file(filepath, encryption_key):
    """Decrypts and loads a CSV file into a pandas DataFrame."""
    try:
        with open(filepath, "rb") as f:
            content_bytes = decrypt_data(f.read(), encryption_key)
        
        original_stderr = sys.stderr
        sys.stderr = open(os.devnull, 'w')
        try:
            df = pd.read_csv(io.StringIO(content_bytes.decode(errors='ignore')), dtype=str, engine='python', on_bad_lines='warn')
        finally:
            sys.stderr.close()
            sys.stderr = original_stderr
            
        return df
    except Exception as e:
        raise InvalidFileTypeError(f"Could not read the file. It may be invalid or corrupted: {e}")
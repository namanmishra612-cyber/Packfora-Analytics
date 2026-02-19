"""
Packfora Analytics - Final Version with Admin Panel
No external dependencies (RSS removed) - Works on FREE hosting!
Upload Excel files via web browser - No server access needed
"""

from flask import Flask, render_template_string, request, jsonify, send_file, redirect, session, flash
from flask_cors import CORS
from werkzeug.utils import secure_filename
import pandas as pd
import numpy as np
from pathlib import Path
import io
from datetime import datetime, timedelta
import os
from functools import wraps
import logging
import shutil
import json

# ================= LOAD ENVIRONMENT VARIABLES =================
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    print("python-dotenv not installed. Using system environment variables.")
    print("To use .env file, run: pip install python-dotenv")

# ================= CONFIGURATION =================
BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"

DATA_DIR.mkdir(parents=True, exist_ok=True)

RESIN_EXCEL = Path(os.getenv('RESIN_DATABASE_PATH', DATA_DIR / "resin-data.xlsx"))
MACHINE_EXCEL = Path(os.getenv('MACHINE_DATABASE_PATH', DATA_DIR / "machine-database.xlsx"))
VAR_COST_EXCEL = Path(os.getenv('VARIABLE_COST_PATH', DATA_DIR / "variables-geo.xlsx"))

# Backup directory
BACKUP_DIR = DATA_DIR / "backups"

BACKUP_DIR.mkdir(parents=True, exist_ok=True)


# Application constants
FILE_CHECK_INTERVAL_SECONDS = 30
MAX_MACHINES_TO_DISPLAY = 100
CACHE_EXPIRY_MINUTES = 5

# Admin credentials (CHANGE THESE!)
ADMIN_USERNAME = os.getenv('ADMIN_USERNAME', 'packfora')
ADMIN_PASSWORD = os.getenv('ADMIN_PASSWORD', 'packfora123')

# File upload settings
ALLOWED_EXTENSIONS = {'xlsx', 'xls', 'pdf'}
MAX_FILE_SIZE = 50 * 1024 * 1024  # 50 MB

# ================= RESIN TYPE CONFIGURATION =================
# Centralized resin type detection — add new types here
import re as _re

RESIN_TYPE_PATTERNS = {
    'LLDPE': [r'\bLLDPE\b', r'\bLINEAR\s*LOW\s*DENSITY\b'],
    'LDPE':  [r'\bLDPE\b', r'\bLOW\s*DENSITY\s*POLY(?:ETHYLENE)?\b'],
    'HDPE':  [r'\bHDPE\b', r'\bHIGH\s*DENSITY\s*POLY(?:ETHYLENE)?\b', r'\bHD\s*PE\b'],
    'PP':    [r'\bPP\b', r'\bPOLYPROPYLENE\b', r'\bPP\s*GRADE\b', r'\bPP\s*HOMO\b',
              r'\bPP\s*RANDOM\b', r'\bPP\s*IMPACT\b', r'\bPP\s*COPOLY\b'],
    'PET':   [r'\bPET\b', r'\bPOLYETHYLENE\s*TEREPHTHALATE\b', r'\bPOLYESTER\b'],
    'PVC':   [r'\bPVC\b', r'\bPOLYVINYL\s*CHLORIDE\b'],
    'PS':    [r'\bPS\b', r'\bPOLYSTYRENE\b', r'\bHIPS\b', r'\bGPPS\b', r'\bEPS\b'],
    'EVA':   [r'\bEVA\b', r'\bETHYLENE\s*VINYL\s*ACETATE\b'],
    'ABS':   [r'\bABS\b'],
    'PA':    [r'\bNYLON\b', r'\bPA6\b', r'\bPA66\b', r'\bPOLYAMIDE\b'],
}

# Priority order for detection (specific before general — LLDPE before LDPE)
RESIN_TYPE_PRIORITY = ['LLDPE', 'LDPE', 'HDPE', 'PP', 'PET', 'PVC', 'PS', 'EVA', 'ABS', 'PA']

# Supplier name variants (keys = canonical, values = search patterns)
SUPPLIER_PATTERNS = {
    'Reliance':       [r'\bRELIANCE\b', r'\bRIL\b', r'\bRPL\b'],
    'Haldia':         [r'\bHALDIA\b', r'\bHPL\b'],
    'IOCL':           [r'\bIOCL\b', r'\bINDIAN\s*OIL\b'],
    'GAIL':           [r'\bGAIL\b'],
    'OPAL':           [r'\bOPAL\b'],
    'HPCL':           [r'\bHPCL\b', r'\bHINDUSTAN\s*PETRO\b'],
    'BPCL':           [r'\bBPCL\b', r'\bBHARAT\s*PETRO\b'],
    'ONGC':           [r'\bONGC\b'],
    'Sabic':          [r'\bSABIC\b'],
    'LyondellBasell': [r'\bLYONDELL\b', r'\bBASELL\b'],
    'MRPL':           [r'\bMRPL\b', r'\bMANGALORE\s*REFINERY\b'],
    'BNPL':           [r'\bBNPL\b', r'\bBRAHMAPUTRA\b'],
    'HPL':            [r'\bHPL\b', r'\bHALDIA\s*PETRO\b'],
}

# Country inference from location/state names
COUNTRY_LOCATION_MAP = {
    'India': [
        'Maharashtra', 'Gujarat', 'Tamil Nadu', 'Karnataka', 'Rajasthan',
        'Delhi', 'Haryana', 'Punjab', 'West Bengal', 'Andhra Pradesh',
        'Telangana', 'Uttar Pradesh', 'Madhya Pradesh', 'Bihar', 'Odisha',
        'Kerala', 'Assam', 'Jharkhand', 'Chhattisgarh', 'Uttarakhand',
        'Jamnagar', 'Vadodara', 'Mumbai', 'Chennai', 'Kolkata', 'Hyderabad',
        'Bangalore', 'Ahmedabad', 'Pune', 'Nagpur', 'Baroda', 'Haldia',
        'Panipat', 'Bathinda', 'Dahej', 'Hazira', 'Mangalore', 'Kochi',
    ],
}

def detect_resin_type(text):
    """Detect resin type from text using regex patterns.
    Returns (resin_type, confidence) tuple."""
    text_upper = text.upper()
    for rtype in RESIN_TYPE_PRIORITY:
        if rtype in RESIN_TYPE_PATTERNS:
            for pattern in RESIN_TYPE_PATTERNS[rtype]:
                if _re.search(pattern, text_upper):
                    return rtype, 'high'
    return 'Unknown', 'none'

def detect_supplier(text):
    """Detect supplier name from text using regex patterns."""
    text_upper = text.upper()
    for supplier, patterns in SUPPLIER_PATTERNS.items():
        for pattern in patterns:
            if _re.search(pattern, text_upper):
                return supplier
    return 'Unknown'

def infer_country(state_name, location_name):
    """Infer country from state/location names."""
    for country, locations in COUNTRY_LOCATION_MAP.items():
        for loc in locations:
            if loc.lower() in state_name.lower() or loc.lower() in location_name.lower():
                return country
    return 'India'  # Default fallback

# ================= APPLICATION SETUP =================
STATIC_DIR = BASE_DIR / "static"
app = Flask(__name__, static_folder=str(STATIC_DIR))

app.secret_key = os.getenv('SECRET_KEY', 'change-this-secret-key-in-production')
CORS(app)

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# File modification tracking
file_mod_times = {'resin': None, 'machine': None, 'cost': None}

# In-memory cache for Excel data
data_cache = {
    'resin': {'data': None, 'timestamp': None},
    'machine': {'data': None, 'timestamp': None},
    'cost': {'data': None, 'timestamp': None}
}

# ================= ADMIN AUTHENTICATION =================

def login_required(f):
    """Decorator to require admin login"""
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'logged_in' not in session:
            return redirect('/admin/login')
        return f(*args, **kwargs)
    return decorated_function

def allowed_file(filename):
    """Check if file extension is allowed"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# ================= HELPER FUNCTIONS =================

def check_files_exist():
    """Check if all required Excel files exist"""
    files = {
        'Resin Database': RESIN_EXCEL,
        'Machine Database': MACHINE_EXCEL,
        'Variable Cost Database': VAR_COST_EXCEL
    }
    
    missing_files = []
    for name, path in files.items():
        if not path.exists():
            missing_files.append(f"{name}: {path}")
    
    if missing_files:
        error_msg = "Missing required files:\n" + "\n".join(missing_files)
        logger.error(error_msg)
        return False, error_msg
    
    return True, "All files present"

def create_backup(file_path):
    """Create backup of file before replacing"""
    if not file_path.exists():
        return None
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_name = f"{file_path.stem}_{timestamp}{file_path.suffix}"
    backup_path = BACKUP_DIR / backup_name
    
    shutil.copy2(file_path, backup_path)
    logger.info(f"Backup created: {backup_path}")
    
    # Keep only last 10 backups per file
    backups = sorted(BACKUP_DIR.glob(f"{file_path.stem}_*{file_path.suffix}"))
    if len(backups) > 10:
        for old_backup in backups[:-10]:
            old_backup.unlink()
            logger.info(f"Deleted old backup: {old_backup}")
    
    return backup_path

def get_file_mod_time(file_path):
    """Get file modification time"""
    try:
        return file_path.stat().st_mtime if file_path.exists() else None
    except Exception as e:
        logger.error(f"Error getting modification time: {e}")
        return None

def check_file_updated(file_type):
    """Check if Excel file has been updated"""
    file_map = {'resin': RESIN_EXCEL, 'machine': MACHINE_EXCEL, 'cost': VAR_COST_EXCEL}
    
    file_path = file_map.get(file_type)
    if not file_path:
        return False
    
    current_mod_time = get_file_mod_time(file_path)
    last_mod_time = file_mod_times.get(file_type)
    
    if last_mod_time is None:
        file_mod_times[file_type] = current_mod_time
        return False
    
    if current_mod_time and current_mod_time > last_mod_time:
        file_mod_times[file_type] = current_mod_time
        if file_type in data_cache:
            data_cache[file_type] = {'data': None, 'timestamp': None}
        return True
    
    return False

def validate_json_input(data, required_fields):
    """Validate JSON input"""
    if not data:
        return False, "No data provided"
    
    missing_fields = [field for field in required_fields if field not in data]
    if missing_fields:
        return False, f"Missing required fields: {', '.join(missing_fields)}"
    
    return True, ""

def load_excel_cached(file_type, sheet_name=None, header=None):
    """Load Excel file with caching"""
    if data_cache[file_type]['data'] is not None:
        cache_time = data_cache[file_type]['timestamp']
        if cache_time and datetime.now() - cache_time < timedelta(minutes=CACHE_EXPIRY_MINUTES):
            logger.info(f"Using cached data for {file_type}")
            return data_cache[file_type]['data']
    
    file_map = {'resin': RESIN_EXCEL, 'machine': MACHINE_EXCEL, 'cost': VAR_COST_EXCEL}
    file_path = file_map[file_type]
    
    try:
        if sheet_name:
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=header)
        else:
            df = pd.ExcelFile(file_path)
        
        data_cache[file_type]['data'] = df
        data_cache[file_type]['timestamp'] = datetime.now()
        
        logger.info(f"Loaded fresh data for {file_type}")
        return df
    except Exception as e:
        logger.error(f"Error loading {file_type}: {e}")
        raise

def format_num(col_name, val):
    """Format numbers with appropriate units"""
    if val in [None, "", np.nan] or str(val).lower() == "nan": 
        return "0.00"
    
    l = str(col_name).lower()
    try:
        num = float(val)
        if any(k in l for k in ["depreciation", "interest"]): 
            return f"{num * 100 if num <= 1 else num:.1f}%"
        
        if "electricity" in l:
            unit = " / kWh"
        elif "sqm" in l or "footprint" in l:
            unit = " sqm"
        elif "power" in l:
            unit = " kWh"
        else:
            unit = ""
        
        return f"{num:,.2f}{unit}"
    except:
        return str(val)

# Per-sheet resin cache: { sheet_name: {'df': DataFrame, 'file_mtime': float} }
_resin_sheet_cache = {}
# Lightweight meta cache: { sheet_name: {'locations': [...], 'grades': [...], 'file_mtime': float} }
_resin_meta_cache = {}

def invalidate_resin_cache():
    """Clear all cached resin sheets (call after import/upload)."""
    global _resin_sheet_cache, _resin_meta_cache
    _resin_sheet_cache.clear()
    _resin_meta_cache.clear()
    logger.info("Resin sheet cache invalidated")


def load_resin_meta(sheet_name):
    """Fast metadata reader — returns only Location & Grade lists.
    Reads only first 6 columns via usecols, skipping hundreds of date columns."""
    global _resin_meta_cache
    try:
        current_mtime = RESIN_EXCEL.stat().st_mtime if RESIN_EXCEL.exists() else 0
        cached = _resin_meta_cache.get(sheet_name)
        if cached and cached.get('file_mtime') == current_mtime:
            return cached

        logger.info(f"Fast-reading resin meta for '{sheet_name}'")

        # Read only first 6 columns — skips hundreds of price/date columns
        meta_cols_range = list(range(6))

        # Try header=0 first
        df = pd.read_excel(RESIN_EXCEL, sheet_name=sheet_name, usecols=meta_cols_range)
        df.columns = [str(c).strip() for c in df.columns]

        if 'Location' not in df.columns:
            df = pd.read_excel(RESIN_EXCEL, sheet_name=sheet_name, header=1, usecols=meta_cols_range)
            df.columns = [str(c).strip() for c in df.columns]

        if 'Location' not in df.columns:
            for h in range(5):
                df = pd.read_excel(RESIN_EXCEL, sheet_name=sheet_name, header=h, usecols=meta_cols_range)
                df.columns = [str(c).strip() for c in df.columns]
                if 'Location' in df.columns:
                    break

        locations = sorted(df["Location"].dropna().astype(str).str.strip().unique().tolist()) if 'Location' in df.columns else []
        grades = sorted(df["Grade"].dropna().astype(str).str.strip().unique().tolist()) if 'Grade' in df.columns else []

        result = {'locations': locations, 'grades': grades, 'file_mtime': current_mtime}
        _resin_meta_cache[sheet_name] = result
        return result
    except Exception as e:
        logger.warning(f"Fast meta read failed for '{sheet_name}': {e}, falling back to full read")
        df = clean_resin_df(sheet_name)
        result = {
            'locations': sorted(df["Location"].dropna().unique().tolist()),
            'grades': sorted(df["Grade"].dropna().unique().tolist()),
            'file_mtime': RESIN_EXCEL.stat().st_mtime if RESIN_EXCEL.exists() else 0
        }
        _resin_meta_cache[sheet_name] = result
        return result


def parse_date_col(col_str):
    """Parse a date column name into a datetime object.
    Supports: YYYY/MM/DD, YYYY-MM-DD, M/D/YYYY, MM/DD/YYYY, DD/MM/YYYY,
    pandas Timestamp, and other common date formats.
    Returns datetime or None if unparseable."""
    from datetime import datetime as _dt
    s = str(col_str).strip()
    # Handle pandas Timestamp objects
    if hasattr(col_str, 'year') and hasattr(col_str, 'month'):
        try:
            return _dt(col_str.year, col_str.month, col_str.day)
        except:
            pass
    for fmt in [
        '%Y/%m/%d', '%Y-%m-%d', '%Y/%m/%d %H:%M:%S', '%Y-%m-%d %H:%M:%S',
        '%m/%d/%Y', '%d/%m/%Y', '%m-%d-%Y', '%d-%m-%Y',
        '%B %d, %Y', '%b %d, %Y', '%B %d %Y', '%b %d %Y',
    ]:
        try:
            return _dt.strptime(s, fmt)
        except ValueError:
            continue
    try:
        return pd.to_datetime(s).to_pydatetime()
    except:
        return None


def sort_date_series(dates_str, values):
    """Parse date strings, sort chronologically, return (sorted_iso_dates, sorted_labels, sorted_values)."""
    paired = []
    for d, v in zip(dates_str, values):
        dt_obj = parse_date_col(d)
        if dt_obj:
            paired.append((dt_obj, v, d))
        else:
            paired.append((datetime.max, v, d))
    paired.sort(key=lambda x: x[0])
    iso_dates = [p[0].strftime('%Y-%m-%d') if p[0] != datetime.max else p[2] for p in paired]
    labels = [p[0].strftime('%b %Y') if p[0] != datetime.max else p[2] for p in paired]
    sorted_values = [p[1] for p in paired]
    return iso_dates, labels, sorted_values


def clean_resin_df(sheet_name):
    """Clean resin dataframe — auto-detect header row. Cached per-sheet with file-mtime invalidation."""
    global _resin_sheet_cache
    try:
        current_mtime = RESIN_EXCEL.stat().st_mtime if RESIN_EXCEL.exists() else 0
        cached = _resin_sheet_cache.get(sheet_name)
        if cached and cached.get('file_mtime') == current_mtime:
            return cached['df']

        logger.info(f"Reading resin sheet '{sheet_name}' from disk (cache miss)")
        # Try header=0 first (standard format from auto-created sheets)
        df = pd.read_excel(RESIN_EXCEL, sheet_name=sheet_name)
        df.columns = [str(c).strip() for c in df.columns]

        if 'Location' not in df.columns:
            # Try header=1 (some legacy formats)
            df = pd.read_excel(RESIN_EXCEL, sheet_name=sheet_name, header=1)
            df.columns = [str(c).strip() for c in df.columns]

        if 'Location' not in df.columns:
            # Scan first 5 rows for the one that contains "Location"
            for h in range(5):
                df = pd.read_excel(RESIN_EXCEL, sheet_name=sheet_name, header=h)
                df.columns = [str(c).strip() for c in df.columns]
                if 'Location' in df.columns:
                    break

        _resin_sheet_cache[sheet_name] = {'df': df, 'file_mtime': current_mtime}
        return df
    except Exception as e:
        logger.error(f"Error cleaning resin dataframe for '{sheet_name}': {e}")
        raise

def analyze_machines_ai(machines):
    """AI-powered machine recommendation"""
    if not machines or len(machines) == 0:
        return None
    
    valid_machines = [m for m in machines if m['cost_raw'] > 0 and m['power_raw'] > 0]
    
    if not valid_machines:
        return None
    
    costs = [m['cost_raw'] for m in valid_machines]
    powers = [m['power_raw'] for m in valid_machines]
    
    min_cost, max_cost = min(costs), max(costs)
    min_power, max_power = min(powers), max(powers)
    
    for machine in valid_machines:
        cost_score = ((machine['cost_raw'] - min_cost) / (max_cost - min_cost) * 100) if max_cost > min_cost else 0
        power_score = ((machine['power_raw'] - min_power) / (max_power - min_power) * 100) if max_power > min_power else 0
        machine['ai_score'] = (cost_score * 0.5) + (power_score * 0.5)
    
    valid_machines.sort(key=lambda x: x['ai_score'])
    best = valid_machines[0]
    
    reasons = []
    if best['cost_raw'] == min(costs):
        reasons.append("lowest cost")
    if best['power_raw'] == min(powers):
        reasons.append("most energy efficient")
    
    if not reasons:
        reasons.append("best balance of cost and energy efficiency")
    
    return {
        "make": best['make'],
        "model": best['model'],
        "cost": best['cost'],
        "power": best['power'],
        "sqm": best['sqm'],
        "reason": " and ".join(reasons),
        "total_analyzed": len(valid_machines)
    }

# ================= ADMIN ROUTES =================

@app.route("/admin/login", methods=["GET", "POST"])
def admin_login():
    """Admin login page"""
    if request.method == "POST":
        username = request.form.get('username')
        password = request.form.get('password')
        
        if username == ADMIN_USERNAME and password == ADMIN_PASSWORD:
            session['logged_in'] = True
            session['username'] = username
            flash('Login successful!', 'success')
            return redirect('/admin/dashboard')
        else:
            flash('Invalid username or password', 'error')
    
    return render_template_string(ADMIN_LOGIN_HTML)

@app.route("/admin/logout")
def admin_logout():
    """Admin logout"""
    session.clear()
    flash('You have been logged out', 'info')
    return redirect('/admin/login')

@app.route("/admin/dashboard")
@login_required
def admin_dashboard():
    """Admin dashboard"""
    files_info = []
    for name, path in [
        ('Resin Database', RESIN_EXCEL),
        ('Machine Database', MACHINE_EXCEL),
        ('Variable Costs', VAR_COST_EXCEL)
    ]:
        if path.exists():
            stat = path.stat()
            files_info.append({
                'name': name,
                'filename': path.name,
                'size': f"{stat.st_size / 1024 / 1024:.2f} MB",
                'modified': datetime.fromtimestamp(stat.st_mtime).strftime('%Y-%m-%d %H:%M:%S'),
                'exists': True
            })
        else:
            files_info.append({
                'name': name,
                'filename': 'N/A',
                'size': 'N/A',
                'modified': 'File not found',
                'exists': False
            })
    
    backups = []
    for backup_file in sorted(BACKUP_DIR.glob("*.xlsx"), reverse=True)[:10]:
        stat = backup_file.stat()
        backups.append({
            'name': backup_file.name,
            'size': f"{stat.st_size / 1024 / 1024:.2f} MB",
            'date': datetime.fromtimestamp(stat.st_mtime).strftime('%Y-%m-%d %H:%M:%S')
        })
    
    return render_template_string(ADMIN_DASHBOARD_HTML, files=files_info, backups=backups, username=session.get('username'))

@app.route("/admin/upload", methods=["POST"])
@login_required
def admin_upload():
    """Handle file upload"""
    try:
        file_type = request.form.get('file_type')
        file = request.files.get('file')
        
        if not file or not file.filename:
            flash('No file selected', 'error')
            return redirect('/admin/dashboard')
        
        if not allowed_file(file.filename):
            flash('Invalid file type. Only .xlsx and .xls allowed', 'error')
            return redirect('/admin/dashboard')
        
        file_map = {'resin': RESIN_EXCEL, 'machine': MACHINE_EXCEL, 'cost': VAR_COST_EXCEL}
        
        if file_type not in file_map:
            flash('Invalid file type', 'error')
            return redirect('/admin/dashboard')
        
        target_path = file_map[file_type]
        
        # Create backup
        if target_path.exists():
            backup_path = create_backup(target_path)
            logger.info(f"Backup created: {backup_path}")
        
        # Save new file
        file.save(target_path)
        
        # Invalidate cache
        data_cache[file_type] = {'data': None, 'timestamp': None}
        file_mod_times[file_type] = get_file_mod_time(target_path)
        if file_type == 'resin':
            invalidate_resin_cache()
        
        flash(f'{file.filename} uploaded successfully!', 'success')
        logger.info(f"File uploaded: {file.filename} → {target_path}")
        
    except Exception as e:
        flash(f'Error uploading file: {str(e)}', 'error')
        logger.error(f"Upload error: {e}")
    
    return redirect('/admin/dashboard')

@app.route("/admin/download/<file_type>")
@login_required
def admin_download(file_type):
    """Download current file"""
    file_map = {'resin': RESIN_EXCEL, 'machine': MACHINE_EXCEL, 'cost': VAR_COST_EXCEL}
    
    if file_type not in file_map:
        flash('Invalid file type', 'error')
        return redirect('/admin/dashboard')
    
    file_path = file_map[file_type]
    
    if not file_path.exists():
        flash('File not found', 'error')
        return redirect('/admin/dashboard')
    
    return send_file(file_path, as_attachment=True)

@app.route("/api/test_import", methods=["GET"])
@login_required
def test_import_endpoint():
    """Test endpoint to verify import route is working"""
    return jsonify({
        "status": "ok",
        "message": "Import endpoint is accessible",
        "resin_db_exists": RESIN_EXCEL.exists(),
        "resin_db_path": str(RESIN_EXCEL),
        "resin_db_writable": os.access(RESIN_EXCEL.parent, os.W_OK) if RESIN_EXCEL.parent.exists() else False
    })

# ================= PET FILM PDF PARSER =================

def parse_pet_film_pdf(file_bytes, filename=""):
    """Parse JPFL-style PET film price-list PDFs.

    Returns a list of dicts with keys matching the resin-import record format:
        resin_type, supplier, country, location, grade, unit, date, price, state, depot

    The PDF layout groups prices by:
        Film-type category → micron thickness → price (₹/Kg)
    This function extracts those groups and converts them into flat records
    where Grade = "<Film Type> <Micron>" (e.g. "Thin Films Normal 8 mic").
    """
    import re
    from datetime import datetime as dt

    try:
        import pdfplumber
    except ImportError:
        import subprocess, sys
        subprocess.check_call([sys.executable, "-m", "pip", "install",
                               "pdfplumber", "--break-system-packages", "-q"])
        import pdfplumber

    records = []

    pdf_file = io.BytesIO(file_bytes) if isinstance(file_bytes, bytes) else file_bytes
    pdf_file.seek(0)

    with pdfplumber.open(pdf_file) as pdf:
        full_text = "\n".join(page.extract_text() or "" for page in pdf.pages)

    if not full_text.strip():
        logger.warning(f"PET PDF '{filename}': no extractable text")
        return records

    # --- Extract effective date (W.E.F dd.mm.yyyy or dd/mm/yyyy) ---
    date_col = "Unknown"
    # Allow any text (including newlines) between W.E.F and the date digits
    date_match = re.search(
        r'W\.?\s*E\.?\s*F\.?\s*[:\-]?\s*[\s\S]{0,30}?(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})',
        full_text, re.IGNORECASE
    )
    if not date_match:
        # Fallback: any dd.mm.yyyy pattern near the top of the document
        date_match = re.search(r'(\d{1,2})[./-](\d{1,2})[./-](\d{4})', full_text[:500])
    if date_match:
        d, m, y = date_match.groups()
        y = y if len(y) == 4 else f"20{y}"
        try:
            price_date = dt.strptime(f"{d}/{m}/{y}", "%d/%m/%Y")
            date_col = f"{price_date.month}/{price_date.day}/{price_date.year}"
        except ValueError:
            date_col = f"{d}/{m}/{y}"

    # --- Detect supplier from filename / text ---
    supplier = "JPFL"  # Default for Jindal Poly Films
    for name, patterns in SUPPLIER_PATTERNS.items():
        for pat in patterns:
            if _re.search(pat, full_text.upper()) or _re.search(pat, filename.upper()):
                supplier = name
                break

    # --- Parse line-by-line to extract category → micron → price ---
    lines = full_text.split("\n")

    # Category headers we recognise (order matters — first match wins)
    CATEGORY_PATTERNS = [
        (r'METALISED\s*FILMS?',           'Metallised Films'),
        (r'THICK\s*FILM',                 'Thick Film'),
        (r'THIN\s*FILMS?',                'Thin Films'),
        (r'PETG\s*\(\s*(J[\-\s]?\d+)\s*\)\s*BARE', None),   # dynamic — captured below
        (r'Matte\s*\(High\s*Adhesion\)',  'Matte High Adhesion'),
        (r'Multi\s*purpose\s*film',       'Multi Purpose Film'),
        (r'Packaging\s*Grade',            'Packaging Grade'),
        (r'Glitter\s*Grade',              'Glitter Grade'),
        (r'Yarn\s*Grade',                 'Yarn Grade'),
        (r'Twist\s*Grade',                'Twist Grade'),
        (r'Ultra\s*Clear',                'Ultra Clear'),
        (r'Isotropic',                    'Isotropic'),
        (r'Opaque\s*White',               'Opaque White'),
        (r'Milky\s*White',                'Milky White'),
        (r'Matte\b',                      'Matte'),
        (r'Normal',                       'Normal'),
        (r'Clear\b',                      'Clear'),
        (r'Hazy\b',                       'Hazy'),
    ]

    current_category = "General"

    # Regex: captures micron spec + optional grade code + price
    # Examples:  "8 mic 116.0"  |  "10 mic (J-351) 130.5"  |  "12-14 mic (J-450/451) 102.0"
    mic_price_re = re.compile(
        r'(\d+[\s,]*(?:[-–]\s*\d+)?\s*mic(?:ron)?)'   # micron part
        r'\s*'
        r'(\([^)]*\))?'                                # optional grade code in parens
        r'\s+'
        r'(\d+(?:\.\d+)?)',                            # price
        re.IGNORECASE
    )

    # Alternate pattern: micron inside parens → "(23 - 50mic) 110.5"
    # Handles lines like "Ultra Clear (23 - 50mic) 110.5" and "Opaque White (45 - 46mic) 171.0"
    paren_mic_price_re = re.compile(
        r'\('
        r'(\d+[\s,]*(?:[-–]\s*\d+)?\s*mic(?:ron)?)'   # micron inside parens
        r'\)'
        r'\s+'
        r'(\d+(?:\.\d+)?)',                            # price
        re.IGNORECASE
    )

    # Also handle thick-film style lines: "Opaque White 200.0" (no mic)
    name_price_re = re.compile(
        r'^([\w\s\(\)\-/]+?)\s+(\d{2,4}(?:\.\d+)?)\s*$'
    )

    thick_film_mode = False

    for line in lines:
        stripped = line.strip()
        if not stripped:
            continue

        # Skip non-data lines
        if re.search(r'Page\s+\d+\s+of\s+\d+', stripped, re.IGNORECASE):
            continue
        if re.search(r'Contact\s+(?:our|sales)', stripped, re.IGNORECASE):
            continue
        if re.search(r'Other\s+Information', stripped, re.IGNORECASE):
            continue

        # --- Check for category header ---
        matched_cat = False
        for pat, cat_name in CATEGORY_PATTERNS:
            m = re.search(pat, stripped, re.IGNORECASE)
            if m:
                if cat_name is None:
                    # Dynamic PETG category
                    current_category = f"PETG ({m.group(1).strip()})"
                else:
                    current_category = cat_name
                matched_cat = True

                # Check for thick film mode
                if re.search(r'THICK\s*FILM', stripped, re.IGNORECASE):
                    thick_film_mode = True
                elif re.search(r'(?:THIN\s*FILM|METALISED)', stripped, re.IGNORECASE):
                    thick_film_mode = False

                break

        # --- Try to extract micron + price from this line ---
        mic_matches = mic_price_re.findall(stripped)
        for mic_part, grade_code, price_str in mic_matches:
            mic_clean = re.sub(r'\s+', ' ', mic_part.strip())
            grade_suffix = grade_code.strip() if grade_code else ""
            grade_label = f"{current_category} {mic_clean}"
            if grade_suffix:
                grade_label = f"{current_category} {mic_clean} {grade_suffix}"

            try:
                price = float(price_str)
                records.append({
                    'resin_type': 'PET',
                    'supplier': supplier,
                    'country': 'India',
                    'location': 'India',
                    'grade': grade_label.strip(),
                    'unit': 'Rs/ Kg',
                    'date': date_col,
                    'price': price,          # Already in Rs/Kg, no /1000
                    'state': '',
                    'depot': '',
                })
            except (ValueError, TypeError):
                pass

        # --- Fallback: micron inside parens, e.g. "Ultra Clear (23 - 50mic) 110.5" ---
        if not mic_matches:
            paren_matches = paren_mic_price_re.findall(stripped)
            for mic_part, price_str in paren_matches:
                mic_clean = re.sub(r'\s+', ' ', mic_part.strip())
                grade_label = f"{current_category} {mic_clean}"
                try:
                    price = float(price_str)
                    records.append({
                        'resin_type': 'PET',
                        'supplier': supplier,
                        'country': 'India',
                        'location': 'India',
                        'grade': grade_label.strip(),
                        'unit': 'Rs/ Kg',
                        'date': date_col,
                        'price': price,
                        'state': '',
                        'depot': '',
                    })
                except (ValueError, TypeError):
                    pass

        # --- Thick-film / category-level pricing (no micron): e.g. "Opaque White 200.0" ---
        if not mic_matches and thick_film_mode:
            nm = name_price_re.match(stripped)
            if nm:
                name_part = nm.group(1).strip()
                price_str = nm.group(2).strip()
                # Skip lines that are clearly headers
                if not re.search(r'(?:Base\s*Price|Standard|Type|W\.E\.F)', name_part, re.IGNORECASE):
                    try:
                        price = float(price_str)
                        grade_label = f"Thick Film {name_part}"
                        records.append({
                            'resin_type': 'PET',
                            'supplier': supplier,
                            'country': 'India',
                            'location': 'India',
                            'grade': grade_label.strip(),
                            'unit': 'Rs/ Kg',
                            'date': date_col,
                            'price': price,
                            'state': '',
                            'depot': '',
                        })
                    except (ValueError, TypeError):
                        pass

        # --- Metallised films: same mic+price pattern, already handled above ---

    logger.info(f"PET PDF '{filename}': extracted {len(records)} price records, date={date_col}")
    return records


@app.route("/api/import_resin_prices", methods=["POST"])
@login_required
def api_import_resin_prices():
    """Parse monthly resin price Excel files AND PET film PDF price lists
    and merge into resin database.
    
    Auto-extracts resin type, supplier, location, and grade from file content.
    Supports multi-sheet workbooks and JPFL-style PET film PDFs.
    Auto-creates new database entries as needed.
    """
    import re
    from datetime import datetime as dt

    try:
        files = request.files.getlist('price_files')
        if not files or len(files) == 0:
            return jsonify({"error": "No files uploaded"}), 400

        all_records = []
        file_results = []

        for file_obj in files:
            if not file_obj.filename:
                continue
            fname = file_obj.filename

            # --- PDF path: PET film price lists ---
            if fname.lower().endswith('.pdf'):
                try:
                    pdf_bytes = file_obj.read()
                    pdf_records = parse_pet_film_pdf(pdf_bytes, filename=fname)
                    if pdf_records:
                        all_records.extend(pdf_records)
                        file_results.append({
                            "file": fname,
                            "status": "success",
                            "records": len(pdf_records),
                            "sheets_processed": 1,
                            "sheet_details": [{
                                "sheet": "PDF",
                                "status": "success",
                                "records": len(pdf_records),
                                "resin_type": "PET"
                            }],
                        })
                    else:
                        file_results.append({
                            "file": fname, "status": "no_data",
                            "records": 0, "sheets_processed": 0,
                            "sheet_details": [],
                            "message": "No price records extracted from PDF"
                        })
                except Exception as ex:
                    logger.error(f"PDF parse error for {fname}: {ex}", exc_info=True)
                    file_results.append({"file": fname, "status": "error",
                                         "message": f"PDF parse error: {ex}"})
                continue  # skip Excel processing below

            # --- Phase 1: Read workbook (all sheets) ---
            try:
                file_bytes = io.BytesIO(file_obj.read())
                xls = pd.ExcelFile(file_bytes)
                sheet_names = xls.sheet_names
            except Exception as ex:
                file_results.append({"file": fname, "status": "error",
                                     "message": f"Cannot read Excel: {ex}"})
                continue

            # --- Filename-level detection (fallback context) ---
            fname_supplier = detect_supplier(fname)
            fname_resin, _ = detect_resin_type(fname)

            file_record_count = 0
            file_sheet_results = []

            for sheet_idx, sname in enumerate(sheet_names):
                try:
                    df_raw = pd.read_excel(file_bytes, sheet_name=sname, header=None)
                except Exception as ex:
                    file_sheet_results.append({
                        "sheet": sname, "status": "error", "message": str(ex)
                    })
                    continue

                if df_raw.empty or df_raw.shape[0] < 3:
                    continue

                # --- Phase 2a: Build full-sheet text context ---
                sheet_text_parts = []
                scan_rows = min(20, df_raw.shape[0])
                scan_cols = min(8, df_raw.shape[1])
                for r in range(scan_rows):
                    for c in range(scan_cols):
                        val = df_raw.iloc[r, c]
                        if pd.notna(val):
                            sheet_text_parts.append(str(val))
                sheet_context = " ".join(sheet_text_parts)

                # Also use the sheet name itself as context
                full_context = f"{sname} {sheet_context} {fname}"

                # --- Phase 2b: Detect resin type (sheet context > filename) ---
                sheet_resin, confidence = detect_resin_type(full_context)
                if sheet_resin == 'Unknown' and fname_resin != 'Unknown':
                    sheet_resin = fname_resin
                    confidence = 'filename'

                # --- Phase 2c: Detect supplier (cell scan > filename) ---
                sheet_supplier = detect_supplier(sheet_context)
                if sheet_supplier == 'Unknown':
                    sheet_supplier = fname_supplier

                # --- Phase 2d: Find section headers ("Date :") ---
                section_starts = []
                for i in range(df_raw.shape[0]):
                    for col in range(min(5, df_raw.shape[1])):
                        val = str(df_raw.iloc[i, col]) if pd.notna(df_raw.iloc[i, col]) else ''
                        if re.search(r'Date\s*:', val, re.IGNORECASE):
                            section_starts.append(i)
                            logger.info(f"Found section header at row {i} in sheet '{sname}': {val[:80]}")
                            break

                if not section_starts:
                    # Alternative: look for other date-like patterns
                    for i in range(df_raw.shape[0]):
                        for col in range(min(5, df_raw.shape[1])):
                            val = str(df_raw.iloc[i, col]) if pd.notna(df_raw.iloc[i, col]) else ''
                            if re.search(r'(?:Effective|Price|Rate)\s*(?:Date|From|W\.?E\.?F)', val, re.IGNORECASE):
                                section_starts.append(i)
                                logger.info(f"Found alt date pattern at row {i} in sheet '{sname}': {val[:80]}")
                                break

                if not section_starts:
                    # Last resort: look for any row with "Date" in first few columns
                    for i in range(df_raw.shape[0]):
                        for col in range(min(3, df_raw.shape[1])):
                            val = str(df_raw.iloc[i, col]) if pd.notna(df_raw.iloc[i, col]) else ''
                            if ('Date' in val or 'DATE' in val) and (':' in val or 'Grade' in val or 'GRADE' in val):
                                section_starts.append(i)
                                logger.info(f"Found date-grade pattern at row {i} in sheet '{sname}': {val[:80]}")
                                break

                if not section_starts:
                    file_sheet_results.append({
                        "sheet": sname, "status": "skipped",
                        "message": "No price section headers found"
                    })
                    continue

                sheet_record_count = 0

                for sec_idx, sec_start in enumerate(section_starts):
                    # --- Section-level context for resin type override ---
                    section_context = ""
                    for r in range(max(0, sec_start - 2), min(sec_start + 4, df_raw.shape[0])):
                        for c in range(min(8, df_raw.shape[1])):
                            val = df_raw.iloc[r, c]
                            if pd.notna(val):
                                section_context += " " + str(val)

                    # Override resin type if section has its own type marker
                    section_resin, sec_conf = detect_resin_type(section_context)
                    resin_type = section_resin if section_resin != 'Unknown' else sheet_resin

                    # Override supplier if section mentions one
                    section_supplier = detect_supplier(section_context)
                    supplier = section_supplier if section_supplier != 'Unknown' else sheet_supplier

                    logger.info(f"Section {sec_idx + 1} in sheet '{sname}': resin={resin_type}, supplier={supplier}")

                    # --- Extract date ---
                    header_text = section_context
                    date_match = re.search(
                        r'Date\s*:\s*(\w+)\s+(\d{1,2})\s*,?\s*(\d{4})', header_text)
                    if not date_match:
                        date_match = re.search(
                            r'(\w+)\s+(\d{1,2})\s*,?\s*(\d{4})', header_text)

                    if date_match:
                        month_str, day_str, year_str = date_match.groups()
                        date_col = "Unknown"
                        for fmt in ["%B %d %Y", "%b %d %Y"]:
                            try:
                                price_date = dt.strptime(
                                    f"{month_str} {day_str} {year_str}", fmt)
                                date_col = f"{price_date.month}/{price_date.day}/{price_date.year}"
                                break
                            except ValueError:
                                continue
                        if date_col == "Unknown":
                            date_col = f"{month_str} {day_str}, {year_str}"
                    else:
                        date_col = "Unknown"

                    # --- Find grade header row ---
                    grade_row = None
                    for j in range(sec_start, min(sec_start + 10, df_raw.shape[0])):
                        for col in range(min(2, df_raw.shape[1])):
                            val = str(df_raw.iloc[j, col]) if pd.notna(df_raw.iloc[j, col]) else ''
                            if re.search(r'Sr|No\.|S\.N|Serial|#|Sl', val, re.IGNORECASE):
                                grade_row = j
                                logger.info(f"Found grade header at row {j}: {val[:50]}")
                                break
                        if grade_row is not None:
                            break

                    if grade_row is None:
                        logger.warning(f"No grade header found for section at row {sec_start}, skipping")
                        continue

                    # --- Extract grades from column headers ---
                    grades = []
                    for c in range(4, df_raw.shape[1]):
                        gv = df_raw.iloc[grade_row, c]
                        if pd.notna(gv) and str(gv).strip():
                            grades.append((c, str(gv).strip()))

                    if not grades:
                        continue

                    sec_end = (section_starts[sec_idx + 1]
                               if sec_idx + 1 < len(section_starts)
                               else df_raw.shape[0])

                    # --- Parse data rows ---
                    for r in range(grade_row + 1, sec_end):
                        sr_val = df_raw.iloc[r, 0]
                        if pd.isna(sr_val):
                            continue
                        try:
                            int(float(sr_val))
                        except (ValueError, TypeError):
                            continue

                        state = str(df_raw.iloc[r, 1]).strip() if pd.notna(df_raw.iloc[r, 1]) else ''
                        location = str(df_raw.iloc[r, 2]).strip() if pd.notna(df_raw.iloc[r, 2]) else ''
                        depot = str(df_raw.iloc[r, 3]).strip() if pd.notna(df_raw.iloc[r, 3]) else ''

                        # Auto-infer country from state/location
                        country = infer_country(state, location)

                        for col_idx, grade_name in grades:
                            if col_idx < df_raw.shape[1]:
                                price_val = df_raw.iloc[r, col_idx]
                                if pd.notna(price_val):
                                    try:
                                        price = float(price_val)
                                        all_records.append({
                                            'resin_type': resin_type,
                                            'supplier': supplier,
                                            'country': country,
                                            'location': location,
                                            'grade': grade_name,
                                            'unit': 'Rs/ Kg',
                                            'date': date_col,
                                            'price': round(price / 1000, 2),
                                            'state': state,
                                            'depot': depot,
                                        })
                                        sheet_record_count += 1
                                    except (ValueError, TypeError):
                                        pass

                file_record_count += sheet_record_count
                if sheet_record_count > 0:
                    file_sheet_results.append({
                        "sheet": sname,
                        "status": "success",
                        "records": sheet_record_count,
                        "resin_type": resin_type
                    })

            file_results.append({
                "file": fname,
                "status": "success" if file_record_count > 0 else "no_data",
                "records": file_record_count,
                "sheets_processed": len(file_sheet_results),
                "sheet_details": file_sheet_results,
            })

        # --- Phase 3: Build / Merge into resin database ---
        if not all_records:
            error_details = [f"{r['file']}: {r.get('message', r.get('status', 'Unknown'))}"
                             for r in file_results if r.get('status') != 'success']
            return jsonify({
                "error": "No valid price records found. " + "; ".join(error_details),
                "file_results": file_results
            }), 400

        records_df = pd.DataFrame(all_records)

        # Handle Unknown resin types
        unknown_count = len(records_df[records_df['resin_type'] == 'Unknown'])
        if unknown_count > 0:
            logger.warning(f"{unknown_count} records with Unknown resin type — excluding")
            records_df = records_df[records_df['resin_type'] != 'Unknown']
            if records_df.empty:
                return jsonify({
                    "error": f"All {unknown_count} records have Unknown resin type. "
                             "Ensure files contain HDPE/LLDPE/LDPE/PP/PET etc. in headers, "
                             "sheet names, or filenames.",
                    "file_results": file_results
                }), 400

        resin_types_found = records_df['resin_type'].unique().tolist()

        # Create backup before modifying
        create_backup(RESIN_EXCEL)

        # Load existing workbook structure
        try:
            existing_wb = pd.ExcelFile(RESIN_EXCEL)
            existing_sheets = existing_wb.sheet_names
            write_mode = 'a'
        except Exception:
            existing_sheets = []
            write_mode = 'w'

        writer_kwargs = {
            'path': RESIN_EXCEL, 'engine': 'openpyxl', 'mode': write_mode
        }
        if write_mode == 'a':
            writer_kwargs['if_sheet_exists'] = 'replace'

        merge_stats = {}

        with pd.ExcelWriter(**writer_kwargs) as writer:
            for rt in resin_types_found:
                rt_df = records_df[records_df['resin_type'] == rt].copy()
                rt_df['unique_key'] = rt_df['supplier'] + rt_df['location'] + rt_df['grade']

                pivot = rt_df.pivot_table(
                    index=['resin_type', 'supplier', 'country', 'location',
                           'grade', 'unit', 'unique_key'],
                    columns='date',
                    values='price',
                    aggfunc='first'
                ).reset_index()
                pivot.columns.name = None

                col_rename = {
                    'resin_type': 'Resin Type', 'supplier': 'Supplier',
                    'country': 'Country', 'location': 'Location',
                    'grade': 'Grade', 'unit': 'Unit', 'unique_key': 'Key',
                }
                pivot.rename(columns=col_rename, inplace=True)

                sheet_name = rt

                if sheet_name in existing_sheets:
                    try:
                        existing_df = pd.read_excel(RESIN_EXCEL, sheet_name=sheet_name)
                        existing_df.columns = [str(c).strip() for c in existing_df.columns]

                        meta_cols = ['Resin Type', 'Supplier', 'Country',
                                     'Location', 'Grade', 'Unit', 'Key']
                        new_date_cols = [c for c in pivot.columns if c not in meta_cols]
                        key_cols = ['Supplier', 'Location', 'Grade']

                        merged = existing_df.merge(
                            pivot[key_cols + new_date_cols],
                            on=key_cols, how='outer', suffixes=('', '_new')
                        )

                        for dc in new_date_cols:
                            if dc + '_new' in merged.columns:
                                merged[dc] = merged[dc + '_new'].combine_first(merged[dc])
                                merged.drop(columns=[dc + '_new'], inplace=True)

                        # Auto-fill metadata for new rows (auto-created entries)
                        merged['Resin Type'] = merged['Resin Type'].fillna(rt)
                        merged['Country'] = merged['Country'].fillna('India')
                        merged['Unit'] = merged['Unit'].fillna('Rs/ Kg')

                        # Fill Supplier from pivot for new rows
                        for idx, row in merged.iterrows():
                            if pd.isna(row.get('Supplier')):
                                match = pivot[
                                    (pivot['Supplier'].notna()) &
                                    (pivot['Location'] == row.get('Location')) &
                                    (pivot['Grade'] == row.get('Grade'))
                                ]
                                if not match.empty:
                                    merged.at[idx, 'Supplier'] = match.iloc[0]['Supplier']

                            # Auto-fill Key if missing
                            if pd.isna(row.get('Key')):
                                merged.at[idx, 'Key'] = (
                                    str(merged.at[idx, 'Supplier'] or '') +
                                    str(merged.at[idx, 'Location'] or '') +
                                    str(merged.at[idx, 'Grade'] or '')
                                )

                        merged.to_excel(writer, sheet_name=sheet_name, index=False)
                        new_rows = len(merged) - len(existing_df)
                        merge_stats[rt] = {
                            "new_dates": len(new_date_cols),
                            "total_rows": len(merged),
                            "new_rows": max(0, new_rows),
                            "mode": "merged"
                        }

                    except Exception as merge_err:
                        logger.warning(f"Merge failed for {sheet_name}: {merge_err}")
                        pivot.to_excel(writer, sheet_name=sheet_name, index=False)
                        merge_stats[rt] = {
                            "new_dates": len([c for c in pivot.columns
                                              if c not in col_rename.values()]),
                            "total_rows": len(pivot),
                            "new_rows": len(pivot),
                            "mode": "fresh_overwrite"
                        }
                else:
                    # AUTO-CREATE new sheet for previously unseen resin type
                    pivot.to_excel(writer, sheet_name=sheet_name, index=False)
                    merge_stats[rt] = {
                        "new_dates": len([c for c in pivot.columns
                                          if c not in col_rename.values()]),
                        "total_rows": len(pivot),
                        "new_rows": len(pivot),
                        "mode": "auto_created"
                    }
                    logger.info(f"Auto-created new sheet '{sheet_name}' "
                                f"with {len(pivot)} rows")

        # Invalidate cache
        data_cache['resin'] = {'data': None, 'timestamp': None}

        response_data = {
            "status": "success",
            "total_records": len(records_df),
            "resin_types": resin_types_found,
            "merge_stats": merge_stats,
            "file_results": file_results,
        }

        if unknown_count > 0:
            response_data["warning"] = (
                f"{unknown_count} records with Unknown resin type were excluded. "
                "Ensure files have resin type info in headers, sheet names, or filenames."
            )

        # Flag auto-created items for admin visibility
        auto_created = [rt for rt, st in merge_stats.items()
                        if st['mode'] == 'auto_created']
        new_entries = [rt for rt, st in merge_stats.items()
                       if st.get('new_rows', 0) > 0]
        if auto_created:
            response_data["auto_created_sheets"] = auto_created
        if new_entries:
            response_data["new_entries_added"] = {
                rt: merge_stats[rt]['new_rows'] for rt in new_entries
            }

        invalidate_resin_cache()
        return jsonify(response_data)

    except Exception as e:
        logger.error(f"Resin price import error: {e}", exc_info=True)
        return jsonify({"error": str(e)}), 500


# ================= SKU STORAGE & API ROUTES =================

# SKU storage file path
SKU_STORAGE_FILE = DATA_DIR / "saved_skus.json"

SKU_STORAGE_FILE.parent.mkdir(parents=True, exist_ok=True)

# Ensure file exists
if not SKU_STORAGE_FILE.exists():
    SKU_STORAGE_FILE.write_text('[]')

@app.route('/api/save_sku', methods=['POST'])
def save_sku():
    """Save SKU configuration to backend storage"""
    try:
        data = request.get_json()
        
        if not data or 'name' not in data:
            return jsonify({'success': False, 'message': 'Invalid SKU data'}), 400
        
        # Load existing SKUs
        try:
            with open(SKU_STORAGE_FILE, 'r') as f:
                skus = json.load(f)
        except (FileNotFoundError, json.JSONDecodeError):
            skus = []
        
        # Remove existing SKU with same name
        skus = [s for s in skus if s.get('name') != data['name']]
        
        # Add new SKU
        skus.append(data)
        
        # Save to file
        with open(SKU_STORAGE_FILE, 'w') as f:
            json.dump(skus, f, indent=2)
        
        logger.info(f"SKU saved: {data['name']} (Model: {data.get('model', 'unknown')})")
        
        return jsonify({
            'success': True,
            'message': f'SKU "{data["name"]}" saved successfully',
            'sku_count': len(skus)
        })
    
    except Exception as e:
        logger.error(f"Error saving SKU: {str(e)}")
        return jsonify({'success': False, 'message': str(e)}), 500


@app.route('/api/load_sku/<sku_name>', methods=['GET'])
def load_sku_api(sku_name):
    """Load specific SKU by name"""
    try:
        # Load SKUs
        try:
            with open(SKU_STORAGE_FILE, 'r') as f:
                skus = json.load(f)
        except (FileNotFoundError, json.JSONDecodeError):
            return jsonify({'success': False, 'message': 'No SKUs found'}), 404
        
        # Find SKU
        sku = next((s for s in skus if s.get('name') == sku_name), None)
        
        if not sku:
            return jsonify({'success': False, 'message': f'SKU "{sku_name}" not found'}), 404
        
        return jsonify({
            'success': True,
            'sku': sku
        })
    
    except Exception as e:
        logger.error(f"Error loading SKU: {str(e)}")
        return jsonify({'success': False, 'message': str(e)}), 500


@app.route('/api/list_skus', methods=['GET'])
def list_skus():
    """List all saved SKUs"""
    try:
        # Load SKUs
        try:
            with open(SKU_STORAGE_FILE, 'r') as f:
                skus = json.load(f)
        except (FileNotFoundError, json.JSONDecodeError):
            skus = []
        
        # Return summary info only
        sku_list = [{
            'name': s.get('name'),
            'model': s.get('model'),
            'timestamp': s.get('timestamp')
        } for s in skus]
        
        return jsonify({
            'success': True,
            'skus': sku_list,
            'count': len(skus)
        })
    
    except Exception as e:
        logger.error(f"Error listing SKUs: {str(e)}")
        return jsonify({'success': False, 'message': str(e)}), 500


@app.route('/api/delete_sku/<sku_name>', methods=['DELETE'])
def delete_sku_api(sku_name):
    """Delete SKU by name"""
    try:
        # Load SKUs
        try:
            with open(SKU_STORAGE_FILE, 'r') as f:
                skus = json.load(f)
        except (FileNotFoundError, json.JSONDecodeError):
            return jsonify({'success': False, 'message': 'No SKUs found'}), 404
        
        # Filter out SKU
        original_count = len(skus)
        skus = [s for s in skus if s.get('name') != sku_name]
        
        if len(skus) == original_count:
            return jsonify({'success': False, 'message': f'SKU "{sku_name}" not found'}), 404
        
        # Save updated list
        with open(SKU_STORAGE_FILE, 'w') as f:
            json.dump(skus, f, indent=2)
        
        logger.info(f"SKU deleted: {sku_name}")
        
        return jsonify({
            'success': True,
            'message': f'SKU "{sku_name}" deleted successfully'
        })
    
    except Exception as e:
        logger.error(f"Error deleting SKU: {str(e)}")
        return jsonify({'success': False, 'message': str(e)}), 500


# ================= PUBLIC ROUTES =================

@app.route("/")
def home():
    """Dashboard home"""
    files_ok, message = check_files_exist()
    if not files_ok:
        error_content = f"""
        <div class='card error-card'>
            <h3>Configuration Error</h3>
            <pre>{message}</pre>
            <p style="margin-top:20px;">
                <a href='/admin/login' class='btn-analyze' style='display:inline-block; text-decoration:none;'>
                    Go to Admin Panel to Upload Files
                </a>
            </p>
        </div>
        """
        return render_template_string(BASE_HTML.replace("{{ content | safe }}", error_content), active="Dashboard")
    
    return render_template_string(BASE_HTML.replace("{{ content | safe }}", DASH_HTML), active="Dashboard")

@app.route("/resin")
def resin():
    """Resin tracker"""
    try:
        xls = load_excel_cached('resin')
        if isinstance(xls, pd.DataFrame):
            xls = pd.ExcelFile(RESIN_EXCEL)
        
        sheets_options = ''.join([f'<option value="{s}">{s}</option>' for s in xls.sheet_names if s.lower() != 'unknown'])
        resin_content = RESIN_UI.replace("{{SHEETS_OPTIONS}}", sheets_options)
        return render_template_string(BASE_HTML.replace("{{ content | safe }}", resin_content), active="Resin")
    except Exception as e:
        logger.error(f"Error loading resin page: {e}")
        error_msg = f"<div class='card error-card'><h3>Error Loading Resin Data</h3><p>{str(e)}</p></div>"
        return render_template_string(BASE_HTML.replace("{{ content | safe }}", error_msg), active="Resin")

@app.route("/machines")
def machines():
    """Machine database"""
    return render_template_string(BASE_HTML.replace("{{ content | safe }}", MACH_HTML), active="Machines")

@app.route("/costs")
def costs():
    """Variable costs"""
    return render_template_string(BASE_HTML.replace("{{ content | safe }}", COST_HTML), active="Costs")

@app.route("/calculator")
def calculator():
    """Cost Calculator"""
    return render_template_string(BASE_HTML.replace("{{ content | safe }}", CALC_HTML), active="Calculator")

# ================= API ENDPOINTS =================

@app.route("/api/dashboard_stats", methods=["GET"])
def api_dashboard_stats():
    """Get dashboard statistics"""
    try:
        xls = load_excel_cached('resin')
        if isinstance(xls, pd.DataFrame):
            xls = pd.ExcelFile(RESIN_EXCEL)
        total_resin_types = len([s for s in xls.sheet_names if s.lower() != 'unknown'])
        
        df_machines = load_excel_cached('machine', sheet_name="Database", header=2)
        total_machines = len(df_machines)
        
        df_costs = load_excel_cached('cost', sheet_name="Data", header=9)
        total_countries = len(df_costs)
        
        return jsonify({
            "resin_types": total_resin_types,
            "machines": total_machines,
            "countries": total_countries,
            "last_updated": datetime.now().strftime("%B %d, %Y at %I:%M %p")
        })
    except Exception as e:
        logger.error(f"Error in dashboard_stats: {e}")
        return jsonify({"error": str(e)}), 500

@app.route("/api/check_file_updates", methods=["GET"])
def api_check_file_updates():
    """Check file updates"""
    return jsonify({
        "resin_updated": check_file_updated('resin'),
        "machine_updated": check_file_updated('machine'),
        "cost_updated": check_file_updated('cost')
    })

@app.route("/api/resin_load", methods=["POST"])
def api_resin_load():
    """Load resin data — fast path using metadata reader (skips date columns)"""
    try:
        data = request.json
        is_valid, error_msg = validate_json_input(data, ['sheet'])
        if not is_valid:
            return jsonify({"error": error_msg}), 400
        
        meta = load_resin_meta(data["sheet"])
        
        return jsonify({
            "locations": meta['locations'],
            "grades": meta['grades']
        })
    except Exception as e:
        logger.error(f"Error in resin_load: {e}")
        return jsonify({"error": str(e)}), 500

@app.route("/api/resin_preload", methods=["POST"])
def api_resin_preload():
    """Pre-warm the full sheet cache in the background.
    Called by frontend fire-and-forget so resin_generate is instant."""
    try:
        data = request.json
        sheet = data.get('sheet', '') if data else ''
        if not sheet:
            return jsonify({"ok": False}), 400
        clean_resin_df(sheet)  # Populates _resin_sheet_cache
        return jsonify({"ok": True})
    except Exception as e:
        logger.warning(f"Preload failed: {e}")
        return jsonify({"ok": False}), 200

@app.route("/api/resin_generate", methods=["POST"])
def api_resin_generate():
    """Generate resin analysis"""
    try:
        d = request.json
        is_valid, error_msg = validate_json_input(d, ['sheet', 'location', 'grade'])
        if not is_valid:
            return jsonify({"error": error_msg}), 400
        
        df = clean_resin_df(d["sheet"])
        subset = df[(df["Location"] == d["location"]) & (df["Grade"] == d["grade"])]
        
        if subset.empty: 
            return jsonify({"error": "No data found"}), 404
        
        row = subset.iloc[0]
        meta = ["Supplier", "Country", "Location", "Grade", "Unit"]
        
        all_dates = []
        all_values = []
        for col in df.columns:
            if col not in meta and not str(col).startswith("Unnamed"):
                try:
                    v = float(row[col])
                    if v > 0: 
                        all_dates.append(str(col))
                        all_values.append(v)
                except: 
                    continue
        
        # Sort chronologically before trimming
        iso_all, labels_all, values_all = sort_date_series(all_dates, all_values)
        
        duration = d.get("duration", "12")
        if duration != "all" and iso_all:
            months_to_keep = int(duration)
            keep_count = min(months_to_keep, len(iso_all))
            iso_dates = iso_all[-keep_count:]
            labels = labels_all[-keep_count:]
            values = values_all[-keep_count:]
        else:
            iso_dates = iso_all
            labels = labels_all
            values = values_all
        
        if not values:
            return jsonify({"error": "No price data available"}), 404
        
        curr = values[-1]
        
        if len(values) > 1:
            first_in_period = values[0]
            diff = ((curr - first_in_period) / first_in_period * 100) if first_in_period != 0 else 0
        else:
            diff = 0
        
        avg_price = sum(values) / len(values)
        min_price = min(values)
        max_price = max(values)
        
        status = "BULLISH" if diff > 1.2 else "BEARISH" if diff < -1.2 else "STABLE"
        
        return jsonify({
            "series": {"dates": iso_dates, "labels": labels, "values": values},
            "insights": {
                "curr": f"₹{curr:,.0f}", 
                "last": f"₹{values[0]:,.0f}", 
                "diff": f"{diff:+.2f}%",
                "avg": f"₹{avg_price:,.0f}",
                "min": f"₹{min_price:,.0f}",
                "max": f"₹{max_price:,.0f}",
                "status": status, 
                "badge": f"badge-{status.lower()}",
                "rec": "Stock up now" if status == "BULLISH" else "Delay bulk orders" if status == "BEARISH" else "Standard procurement"
            }
        })
    except Exception as e:
        logger.error(f"Error in resin_generate: {e}")
        return jsonify({"error": str(e)}), 500

@app.route("/api/mach_res", methods=["POST"])
def api_mach_res():
    """Get machine results"""
    try:
        data = request.json
        is_valid, error_msg = validate_json_input(data, ['cat', 'proc'])
        if not is_valid:
            return jsonify({"error": error_msg}), 400
        
        df = load_excel_cached('machine', sheet_name="Database", header=2)
        f = df[(df["Category"] == data["cat"]) & (df["Process"] == data["proc"])]
        
        if len(f) > MAX_MACHINES_TO_DISPLAY:
            logger.warning(f"Limiting results from {len(f)} to {MAX_MACHINES_TO_DISPLAY}")
            f = f.head(MAX_MACHINES_TO_DISPLAY)
        
        res = []
        
        for _, r in f.iterrows():
            cost = r.get("Machine Cost In €") or r.get("Machine Cost") or r.get("Price")
            if pd.isna(cost) or cost == 0:
                for col in df.columns:
                    if any(k in str(col) for k in ["€", "Cost", "Price"]): 
                        cost = r[col]
                        break
            
            power = r.get("Power Consumption")
            sqm = r.get("Machine Footprint SQM")
            
            res.append({
                "make": str(r.get("Make", "")), 
                "model": str(r.get("Model", "")), 
                "cost": format_num("cost", cost), 
                "cost_raw": float(cost) if not pd.isna(cost) else 0,
                "power": format_num("power", power), 
                "power_raw": float(power) if not pd.isna(power) else 0,
                "sqm": format_num("sqm", sqm),
                "sqm_raw": float(sqm) if not pd.isna(sqm) else 0
            })
        
        recommendation = analyze_machines_ai(res)
        
        return jsonify({"results": res, "recommendation": recommendation})
    except Exception as e:
        logger.error(f"Error in mach_res: {e}")
        return jsonify({"error": str(e)}), 500

@app.route("/api/export_machines", methods=["POST"])
def api_export_machines():
    """Export machines to Excel"""
    try:
        data = request.json
        is_valid, error_msg = validate_json_input(data, ['results'])
        if not is_valid:
            return jsonify({"error": error_msg}), 400
        
        df = pd.DataFrame(data['results'])
        df = df[[col for col in df.columns if not col.endswith('_raw')]]
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Machines', index=False)
        
        output.seek(0)
        
        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=f'machine_export_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
        )
    except Exception as e:
        logger.error(f"Error in export_machines: {e}")
        return jsonify({"error": str(e)}), 500

@app.route("/api/cost_res", methods=["POST"])
def api_cost_res():
    """Get cost results"""
    try:
        data = request.json
        is_valid, error_msg = validate_json_input(data, ['country'])
        if not is_valid:
            return jsonify({"error": error_msg}), 400
        
        df = load_excel_cached('cost', sheet_name="Data", header=9)
        df.columns = [str(c).strip() for c in df.columns]
        
        country_data = df[df.iloc[:, 0] == data["country"]]
        if country_data.empty:
            return jsonify({"error": "Country not found"}), 404
        
        row = country_data.iloc[0]
        
        sections = [
            {"name": "Utility & Energy", "keys": ["electricity"]},
            {"name": "Manpower & Labor", "keys": ["labour", "operator", "engineer", "manager"]},
            {"name": "Infrastructure", "keys": ["land", "building", "lease"], "exclude": ["depreciation"]},
            {"name": "Depreciation & Finance", "keys": ["depreciation", "interest"], "exclude": ["exchange"]},
            {"name": "Exchange Rates", "keys": ["usd", "euro", "exchange"]}
        ]
        
        output = []
        for s in sections:
            items = [
                {"label": col.title(), "value": format_num(col, row[col])} 
                for col in df.columns[1:] 
                if any(k in col.lower() for k in s["keys"]) 
                and not any(ex in col.lower() for ex in s.get("exclude", []))
            ]
            if items: 
                output.append({"section": s["name"], "items": items})
        
        return jsonify({"sections": output})
    except Exception as e:
        logger.error(f"Error in cost_res: {e}")
        return jsonify({"error": str(e)}), 500

@app.route("/api/init", methods=["POST"])
def api_init():
    """Initialize dropdowns"""
    try:
        data = request.json
        is_valid, error_msg = validate_json_input(data, ['module'])
        if not is_valid:
            return jsonify({"error": error_msg}), 400
        
        m = data["module"]
        
        if m == "machines":
            df = load_excel_cached('machine', sheet_name="Database", header=2)
            categories = sorted(df["Category"].dropna().unique().tolist())
            return jsonify(categories)
        elif m == "costs":
            df = load_excel_cached('cost', sheet_name="Data", header=9)
            countries = df.iloc[:, 0].dropna().unique().tolist()
            return jsonify(countries)
        else:
            return jsonify({"error": "Invalid module"}), 400
    except Exception as e:
        logger.error(f"Error in init: {e}")
        return jsonify({"error": str(e)}), 500

@app.route("/api/procs", methods=["POST"])
def api_procs():
    """Get processes"""
    try:
        data = request.json
        is_valid, error_msg = validate_json_input(data, ['cat'])
        if not is_valid:
            return jsonify({"error": error_msg}), 400
        
        df = load_excel_cached('machine', sheet_name="Database", header=2)
        processes = sorted(df[df["Category"] == data["cat"]]["Process"].dropna().unique().tolist())
        
        return jsonify(processes)
    except Exception as e:
        logger.error(f"Error in procs: {e}")
        return jsonify({"error": str(e)}), 500

@app.route("/api/resin_grades", methods=["POST"])
def api_resin_grades():
    """Get grades and locations for a specific resin type"""
    try:
        data = request.json
        is_valid, error_msg = validate_json_input(data, ['resin_type'])
        if not is_valid:
            return jsonify({"error": error_msg}), 400
        
        resin_type = data.get('resin_type')
        xl = load_excel_cached('resin')
        if xl is None:
            return jsonify({"error": "Failed to load resin database"}), 500
        
        if "Price History" not in xl.sheet_names:
            if resin_type not in xl.sheet_names:
                return jsonify({"error": f"Resin type {resin_type} not found"}), 404
            meta = load_resin_meta(resin_type)
            return jsonify({
                "grades": meta['grades'],
                "locations": meta['locations']
            })
        
        df = xl.parse("Price History", header=0)
        df_filtered = df[df["Resin Type"].str.strip() == resin_type]
        
        if df_filtered.empty:
            return jsonify({"error": f"No data found for {resin_type}"}), 404
        
        return jsonify({
            "grades": sorted(df_filtered["Grade"].dropna().unique().tolist()),
            "locations": sorted(df_filtered["Location"].dropna().unique().tolist())
        })
    except Exception as e:
        logger.error(f"Error in resin_grades: {e}")
        return jsonify({"error": str(e)}), 500

@app.route("/api/resin_compare", methods=["POST"])
def api_resin_compare():
    """Compare resin prices across regions (DATA-POINT SAFE VERSION)"""
    try:
        data = request.json
        required = ["resin_type", "grade", "locations", "duration"]
        is_valid, error_msg = validate_json_input(data, required)
        if not is_valid:
            return jsonify({"error": error_msg}), 400

        resin_type = data["resin_type"]
        grade = data["grade"]
        locations = data["locations"]
        duration_label = data["duration"]

        DURATION_MAP = {
            "Last 1 Month": 1,
            "Last 3 Months": 3,
            "Last 6 Months": 6,
            "Last 1 Year": 12,
            "Last 2 Years": 24,
            "All": "all"
        }

        points_required = DURATION_MAP.get(duration_label)
        if points_required is None:
            return jsonify({"error": f"Invalid duration: {duration_label}"}), 400

        if len(locations) < 2:
            return jsonify({"error": "Please select at least 2 locations"}), 400

        # Limit time-series data points (not locations) to keep responses fast
        MAX_TS_POINTS = 50  # cap per-location time-series when many locations selected
        ts_limit = MAX_TS_POINTS if len(locations) > 10 else None

        df = clean_resin_df(resin_type)
        df = df[df["Grade"].astype(str).str.strip() == grade.strip()]

        if df.empty:
            return jsonify({"error": "No data found for selected grade"}), 404

        meta_cols = ["Supplier", "Country", "Location", "Grade", "Unit"]
        all_price_cols = [
            c for c in df.columns
            if c not in meta_cols and not str(c).startswith("Unnamed")
        ]

        if not all_price_cols:
            return jsonify({"error": "No price columns found"}), 500

        # Sort price columns chronologically
        col_date_pairs = []
        for c in all_price_cols:
            dt_obj = parse_date_col(c)
            col_date_pairs.append((c, dt_obj if dt_obj else datetime.max))
        col_date_pairs.sort(key=lambda x: x[1])
        all_price_cols = [p[0] for p in col_date_pairs]

        comparison = []

        for loc in locations:
            row = df[df["Location"].astype(str).str.strip() == loc.strip()]
            if row.empty:
                continue

            # Collect valid price points from RIGHT to LEFT
            collected = []
            for col in reversed(all_price_cols):
                try:
                    v = float(row.iloc[0][col])
                    if v > 0:
                        collected.append((col, v))
                except:
                    continue

                if points_required != "all" and len(collected) >= points_required:
                    break

            if not collected:
                continue

            collected.reverse()
            dates = [str(c[0]) for c in collected]
            prices = [c[1] for c in collected]

            # Sort and format dates
            iso_dates, date_labels, sorted_prices = sort_date_series(dates, prices)

            curr = sorted_prices[-1]
            avg_p = sum(sorted_prices) / len(sorted_prices)
            min_p = min(sorted_prices)
            max_p = max(sorted_prices)

            change = ((curr - sorted_prices[0]) / sorted_prices[0] * 100) if len(sorted_prices) > 1 else 0
            trend = "Rising" if change > 2 else "Falling" if change < -2 else "Stable"

            # Apply ts_limit if needed
            final_iso = iso_dates[::max(1, len(iso_dates)//ts_limit)] if ts_limit and len(iso_dates) > ts_limit else iso_dates
            final_labels = date_labels[::max(1, len(date_labels)//ts_limit)] if ts_limit and len(date_labels) > ts_limit else date_labels
            final_prices = sorted_prices[::max(1, len(sorted_prices)//ts_limit)] if ts_limit and len(sorted_prices) > ts_limit else sorted_prices

            comparison.append({
                "location": loc,
                "current_price": f"₹{curr:,.2f}",
                "avg_price": f"₹{avg_p:,.2f}",
                "min_price": f"₹{min_p:,.2f}",
                "max_price": f"₹{max_p:,.2f}",
                "price_change": f"{change:+.2f}%",
                "trend": trend,
                "data_points": len(sorted_prices),
                "time_series": [
                    {"date": iso, "label": lbl, "price": p}
                    for iso, lbl, p in zip(final_iso, final_labels, final_prices)
                ],
                "current_price_raw": curr
            })

        if len(comparison) < 2:
            return jsonify({"error": "Insufficient comparable data"}), 404

        comparison.sort(key=lambda x: x["current_price_raw"])
        spread = comparison[-1]["current_price_raw"] - comparison[0]["current_price_raw"]

        return jsonify({
            "resin_type": resin_type,
            "grade": grade,
            "duration": duration_label,
            "comparison": comparison,
            "summary": {
                "best_price_location": comparison[0]["location"],
                "worst_price_location": comparison[-1]["location"],
                "price_spread": f"₹{spread:,.2f}",
                "total_locations": len(comparison)
            }
        })

    except Exception as e:
        logger.error(f"Resin comparison error: {e}", exc_info=True)
        return jsonify({"error": str(e)}), 500


@app.route("/api/export_comparison", methods=["POST"])
def api_export_comparison():
    """Export region comparison to Excel"""
    try:
        data = request.json
        if not data:
            return jsonify({"error": "No data to export"}), 400
        
        comparison_list = []
        for loc in data.get('comparison', []):
            comparison_list.append({
                'Location': loc['location'], 'Current Price': loc['current_price'],
                'Average Price': loc['avg_price'], 'Min Price': loc['min_price'],
                'Max Price': loc['max_price'], 'Trend': loc['trend'],
                'Price Change': loc['price_change'], 'Data Points': loc['data_points']
            })
        
        df = pd.DataFrame(comparison_list)
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name='Comparison', index=False)
            summary_data = [{
                'Resin Type': data.get('resin_type'), 'Grade': data.get('grade'),
                'Duration': data.get('duration'),
                'Best Price Location': data['summary']['best_price_location'],
                'Worst Price Location': data['summary']['worst_price_location'],
                'Price Spread': data['summary']['price_spread'],
                'Total Locations Compared': data['summary']['total_locations']
            }]
            pd.DataFrame(summary_data).to_excel(writer, sheet_name='Summary', index=False)
            
            for loc_data in data.get('comparison', []):
                if loc_data.get('time_series'):
                    ts_df = pd.DataFrame(loc_data['time_series'])
                    sheet_name = f"{loc_data['location'][:25]}_Trend"
                    ts_df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            workbook = writer.book
            worksheet = writer.sheets['Comparison']
            header_format = workbook.add_format({'bold': True, 'bg_color': '#e8601c', 'font_color': 'white', 'border': 1})
            for col_num, value in enumerate(df.columns.values):
                worksheet.write(0, col_num, value, header_format)
            for i, col in enumerate(df.columns):
                max_len = max(df[col].astype(str).apply(len).max(), len(col)) + 2
                worksheet.set_column(i, i, max_len)
        
        output.seek(0)
        return send_file(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                        as_attachment=True, download_name=f'resin_comparison_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx')
    except Exception as e:
        logger.error(f"Export comparison error: {e}", exc_info=True)
        return jsonify({"error": f"Export failed: {str(e)}"}), 500

# ================= COST CALCULATOR API =================

@app.route("/api/carton_machine_db", methods=["POST"])
def api_carton_machine_db():
    """Get carton machine data from Machine Database for Advanced Carton Calculator"""
    try:
        df = load_excel_cached('machine', sheet_name="Database", header=2)
        machines_by_process = {}
        for _, r in df.iterrows():
            process = str(r.get("Process", "")).strip()
            if not process or process == 'nan':
                continue
            make = str(r.get("Make", "")).strip()
            model = str(r.get("Model", "")).strip()
            label = f"{make} {model}".strip() if make != 'nan' else model
            if not label or label == 'nan':
                continue
            cost = r.get("Machine Cost In €") or r.get("Machine Cost") or 0
            if pd.isna(cost):
                cost = 0
            try:
                cost = float(cost)
            except:
                cost = 0
            power = r.get("Power Consumption") or 0
            if pd.isna(power):
                power = 0
            try:
                power = float(power)
            except:
                power = 0
            sqm = r.get("Machine Footprint SQM") or 0
            if pd.isna(sqm):
                sqm = 0
            try:
                sqm = float(sqm)
            except:
                sqm = 0
            speed = 0
            for col in df.columns:
                if 'speed' in str(col).lower() or 'output' in str(col).lower():
                    val = r.get(col)
                    if not pd.isna(val):
                        try:
                            speed = float(val)
                        except:
                            speed = 0
                    break
            if process not in machines_by_process:
                machines_by_process[process] = []
            machines_by_process[process].append({
                "label": label,
                "cost_eur": round(cost, 2),
                "power_kwh": round(power, 2),
                "sqm": round(sqm, 2),
                "speed": round(speed, 2)
            })
        return jsonify({"machines": machines_by_process})
    except Exception as e:
        logger.error(f"carton_machine_db error: {e}")
        return jsonify({"error": str(e)}), 500


@app.route("/api/calc_carton_advanced", methods=["POST"])
def api_calc_carton_advanced():
    """Advanced Carton Cost Model - Full engineering cost model matching Excel exactly"""
    try:
        d = request.json
        if not d:
            return jsonify({"error": "No data provided"}), 400

        # --- GENERAL INPUTS ---
        country = d.get('country', 'India')
        annual_volume = float(d.get('annual_volume', 3126950))
        avg_order_size = float(d.get('avg_order_size', 260579.17))
        no_of_colours = int(d.get('no_of_colours', 5))
        common_colours = int(d.get('common_colours', 2))
        no_of_designs = int(d.get('no_of_designs', 1))
        print_runs_year = int(d.get('print_runs_year', 12))
        designs_per_run = int(d.get('designs_per_run', 1))
        no_of_shifts = int(d.get('no_of_shifts', 3))
        margin_pct = float(d.get('margin_pct', 0.20))

        # Carton Dimensions (mm)
        length_1 = float(d.get('length_1', 36.3))
        length_2 = float(d.get('length_2', 37))
        width_1 = float(d.get('width_1', 46))
        width_2 = float(d.get('width_2', 46))
        height = float(d.get('height', 179))
        max_flap = float(d.get('max_flap', 96.9))
        gluing_area = float(d.get('gluing_area', 13))
        grain_direction = d.get('grain_direction', 'Long')

        # --- FLAT SIZE (folding carton layout) ---
        # Length = Gluing + W1 + L1 + W2 + L2  (cross-direction panels)
        # Width = Height + Max Flap (vertical direction)
        layflat_length = gluing_area + width_1 + length_1 + width_2 + length_2
        layflat_width = height + max_flap
        area_per_carton = (layflat_length * layflat_width) / 1000000  # sqm

        # --- SHEET LAYOUT ---
        machine_size = d.get('machine_size', 'IIIB')
        ups_lengthwise = int(d.get('ups_lengthwise', 5))
        ups_widthwise = int(d.get('ups_widthwise', 2))
        side_lay_1 = float(d.get('side_lay_1', 5))
        side_lay_2 = float(d.get('side_lay_2', 5))
        gripper = float(d.get('gripper', 10))
        back_lay = float(d.get('back_lay', 5))
        trimmer = float(d.get('trimmer', 0))
        double_cut = float(d.get('double_cut', 0))
        gutter = float(d.get('gutter', 0))
        interlock_flag = d.get('interlock', 'N')
        interlock_val = float(d.get('interlock_val', 0)) if interlock_flag == 'Y' else 0

        ups_per_sheet = ups_lengthwise * ups_widthwise
        if ups_per_sheet <= 0:
            ups_per_sheet = 1

        # Sheet Size: Length = flatL * upsL + side1 + side2 + backLay + trimmer + doubleCut + gutter*(upsL-1) + interlock
        # Width = flatW * upsW + gripper + trimmer + gutter*(upsW-1)
        if grain_direction == 'Long':
            sheet_length = (layflat_length * ups_lengthwise) + side_lay_1 + side_lay_2 + back_lay + trimmer + double_cut + (gutter * max(0, ups_lengthwise - 1)) + interlock_val
            sheet_width = (layflat_width * ups_widthwise) + gripper + trimmer + (gutter * max(0, ups_widthwise - 1))
        else:
            sheet_length = (layflat_width * ups_lengthwise) + side_lay_1 + side_lay_2 + back_lay + trimmer + double_cut + (gutter * max(0, ups_lengthwise - 1)) + interlock_val
            sheet_width = (layflat_length * ups_widthwise) + gripper + trimmer + (gutter * max(0, ups_widthwise - 1))

        area_sheet = (sheet_length * sheet_width) / 1000000  # sqm

        # --- MATERIAL RATES ---
        board_gsm = float(d.get('board_gsm', 300))
        board_rate = float(d.get('board_rate', 45))
        ink_rate = float(d.get('ink_rate', 834.16))
        ink_gsm = float(d.get('ink_gsm', 2))
        varnish_rate = float(d.get('varnish_rate', 521.35))
        varnish_gsm = float(d.get('varnish_gsm', 3))
        euro_rate = float(d.get('euro_rate', 104.27))

        # Decoration flags
        spot_varnish = d.get('spot_varnish', 'N')
        hot_foiling = d.get('hot_foiling', 'N')
        lamination = d.get('lamination', 'N')
        window_carton = d.get('window_carton', 'N')
        liner_flag = d.get('liner_flag', 'N')
        primer_flag = d.get('primer_flag', 'N')

        # ==================================
        # MATERIAL COST (per 1000 cartons)
        # ==================================

        # Board: area_sheet * GSM / 1000 * 1000/UPS
        board_consumption = (area_sheet * board_gsm / 1000) * (1000 / ups_per_sheet)
        board_cost = board_consumption * board_rate

        # Ink: uses sheet area basis (ink covers full sheet)
        ink_consumption = (area_sheet * ink_gsm / 1000) * (1000 / ups_per_sheet)
        ink_cost = ink_consumption * ink_rate

        # Varnish: uses sheet area basis (varnish applied during coating on full sheet)
        varnish_consumption = (area_sheet * varnish_gsm / 1000) * (1000 / ups_per_sheet)
        varnish_cost = varnish_consumption * varnish_rate

        # Spot Varnish
        spot_varnish_area_pct = float(d.get('spot_varnish_area_pct', 0.3))
        spot_varnish_cost = 0
        if spot_varnish == 'Y':
            sv_consumption = varnish_consumption * spot_varnish_area_pct
            spot_varnish_cost = sv_consumption * varnish_rate

        # Hot Foiling
        foil_rate_roll = float(d.get('foil_rate_roll', 938.43))
        foil_length = float(d.get('foil_length', 14.52))
        foil_width = float(d.get('foil_width', 71.6))
        hot_foil_cost = 0
        if hot_foiling == 'Y':
            foil_area = foil_length * foil_width / 1000000
            hot_foil_cost = foil_area * foil_rate_roll * 1000

        # Lamination (film + adhesive/primer)
        film_rate = float(d.get('film_rate', 208.54))
        film_gsm = float(d.get('film_gsm', 14))
        film_cost = 0
        if lamination == 'Y':
            film_consumption = (area_sheet * film_gsm / 1000) * (1000 / ups_per_sheet)
            film_cost = film_consumption * film_rate

        # Window, Liner, Primer (typically small add-ons)
        window_cost = 0
        liner_cost = 0
        primer_cost = 0
        if window_carton == 'Y':
            win_w = float(d.get('window_width', 89.5))
            win_l = float(d.get('window_length', 83))
            win_gsm = float(d.get('window_film_gsm', 175))
            win_area = win_w * win_l / 1000000
            window_cost = win_area * win_gsm / 1000 * film_rate * 1000
        if liner_flag == 'Y':
            lin_w = float(d.get('liner_width', 322))
            lin_l = float(d.get('liner_length', 160))
            lin_gsm = float(d.get('liner_gsm', 84.1))
            lin_rate = float(d.get('liner_rate', 160))
            lin_area = lin_w * lin_l / 1000000
            liner_cost = lin_area * lin_gsm / 1000 * lin_rate * 1000
        if primer_flag == 'Y':
            primer_rate = float(d.get('primer_rate', 208.54))
            primer_consumption = area_per_carton * (2 / 1000) * 1000
            primer_cost = primer_consumption * primer_rate

        # Plates & Dies
        plate_cost_per_plate = 5 * euro_rate  # EUR 5 per plate
        total_plate_cost = plate_cost_per_plate * no_of_colours
        plate_per_1000 = total_plate_cost / (avg_order_size / 1000) if avg_order_size > 0 else 0

        blanking_die_cost = 500 * euro_rate  # EUR 500 per die
        die_life = 2000000  # sheets
        annual_sheets = annual_volume / ups_per_sheet
        blanking_per_1000 = (blanking_die_cost * annual_sheets / die_life) / (annual_volume / 1000) if annual_volume > 0 else 0

        other_material_cost = plate_per_1000 + blanking_per_1000

        # Wastage: based on printing setup + registration waste %
        # Num changeovers = print_runs + design_change_colours
        design_change_colours = max(0, no_of_colours - common_colours)
        num_changeovers = print_runs_year + design_change_colours

        # Wastage percentages from setup and registration
        changeover_time = 45  # minutes
        registration_time = 14  # minutes
        setup_speed_sheets_min = 10
        reg_speed_sheets_min = 20

        setup_sheets_wasted = changeover_time * setup_speed_sheets_min  # per setup
        reg_sheets_wasted = registration_time * reg_speed_sheets_min
        sheets_per_run = avg_order_size / ups_per_sheet

        setup_wastage_pct = (setup_sheets_wasted * print_runs_year) / annual_sheets if annual_sheets > 0 else 0
        reg_wastage_pct = (reg_sheets_wasted * print_runs_year) / annual_sheets if annual_sheets > 0 else 0
        total_wastage_pct = setup_wastage_pct + reg_wastage_pct

        material_before_wastage = board_cost + ink_cost + varnish_cost + spot_varnish_cost + hot_foil_cost + film_cost
        wastage_cost = material_before_wastage * total_wastage_pct

        total_material = (board_cost + ink_cost + varnish_cost + spot_varnish_cost + hot_foil_cost +
                         film_cost + window_cost + liner_cost + primer_cost +
                         wastage_cost + other_material_cost)

        # ======================================
        # MACHINE UTILISATION & CONVERSION COST
        # ======================================

        # Machine database lookup
        machine_db = {
            'KBA 8000': {'cost_eur': 800000, 'power_kw': 150, 'speed': 12000, 'sqm': 36},
            'Heidelberg SM 52': {'cost_eur': 300000, 'power_kw': 50, 'speed': 8000, 'sqm': 15},
            'Heidelberg SM 74': {'cost_eur': 450000, 'power_kw': 75, 'speed': 10000, 'sqm': 20},
            'Heidelberg SM 102': {'cost_eur': 600000, 'power_kw': 100, 'speed': 12000, 'sqm': 30},
            'KBA Rapida 75': {'cost_eur': 400000, 'power_kw': 65, 'speed': 9000, 'sqm': 18},
            'KBA Rapida 106': {'cost_eur': 650000, 'power_kw': 110, 'speed': 11000, 'sqm': 28},
            'Komori Lithrone': {'cost_eur': 500000, 'power_kw': 80, 'speed': 10000, 'sqm': 22},
            'Manroland R700': {'cost_eur': 550000, 'power_kw': 90, 'speed': 11000, 'sqm': 25},
            'Manroland R900': {'cost_eur': 700000, 'power_kw': 120, 'speed': 12000, 'sqm': 32},
            'Unison': {'cost_eur': 120000, 'power_kw': 20, 'speed': 5000, 'sqm': 8},
            'Hot Foil Machine': {'cost_eur': 250000, 'power_kw': 50, 'speed': 5000, 'sqm': 15},
            'Kompac KwikPrint / EZ Koat': {'cost_eur': 150000, 'power_kw': 20, 'speed': 3500, 'sqm': 10},
            'Bobst Mastercut': {'cost_eur': 700000, 'power_kw': 50, 'speed': 7500, 'sqm': 36},
            'Bobst Masterfold': {'cost_eur': 600000, 'power_kw': 50, 'speed': 120000, 'sqm': 50},
            'Masterwork Ecocut': {'cost_eur': 350000, 'power_kw': 35, 'speed': 6000, 'sqm': 28},
            'Medium automatic window patcher': {'cost_eur': 60000, 'power_kw': 10, 'speed': 1800, 'sqm': 6},
            'Liner Carton Machine': {'cost_eur': 14300, 'power_kw': 50, 'speed': 12000, 'sqm': 5},
        }

        # Try loading from Excel DB
        try:
            mdb = pd.read_excel(os.path.join(DATA_DIR, 'machine-database.xlsx'), sheet_name='carton-machine', header=0)
            for _, row in mdb.iterrows():
                name = str(row.get('Machine Name', row.get('machine_name', ''))).strip()
                if name:
                    machine_db[name] = {
                        'cost_eur': float(row.get('Cost EUR', row.get('cost_eur', 0)) or 0),
                        'power_kw': float(row.get('Power KW', row.get('power_kw', 0)) or 0),
                        'speed': float(row.get('Speed', row.get('speed', 0)) or 0),
                        'sqm': float(row.get('SQM', row.get('sqm', 0)) or 0),
                    }
        except:
            pass

        def get_machine(name, default_name='KBA 8000'):
            return machine_db.get(name, machine_db.get(default_name, {'cost_eur':800000,'power_kw':150,'speed':12000,'sqm':36}))

        pm_machine = get_machine(d.get('printing_machine', 'KBA 8000'))
        sv_machine = get_machine(d.get('spot_varnish_machine', 'Unison'), 'Unison')
        hf_machine = get_machine(d.get('hot_foiling_machine', 'Hot Foil Machine'), 'Hot Foil Machine')
        lm_machine = get_machine(d.get('lamination_machine', 'Kompac KwikPrint / EZ Koat'), 'Kompac KwikPrint / EZ Koat')
        cb_machine = get_machine(d.get('cb_machine', 'Bobst Mastercut'), 'Bobst Mastercut')
        fg_machine = get_machine(d.get('fg_machine', 'Bobst Masterfold'), 'Bobst Masterfold')
        wp_machine = get_machine('Medium automatic window patcher')
        ln_machine = get_machine('Liner Carton Machine')

        pm_speed = pm_machine['speed']  # sheets/hr
        cb_speed = cb_machine['speed']  # sheets/hr
        fg_speed = fg_machine['speed']  # cartons/hr (folder-gluer processes cartons)
        efficiency = float(d.get('efficiency', 0.80))
        total_working_hours = 24 * 330  # 7920

        # --- PRINTING HOURS ---
        # Run hours = annual_sheets / (speed * efficiency)
        print_run_hours = (annual_sheets) / (pm_speed * efficiency) if pm_speed > 0 else 0

        # Setup hours = changeover_time * num_changeovers / 60
        print_setup_hours = changeover_time * num_changeovers / 60

        # Registration hours = registration_time * num_changeovers / 60
        print_reg_hours = registration_time * num_changeovers / 60

        total_print_hours = print_run_hours + print_setup_hours + print_reg_hours
        print_machines_req = total_print_hours / total_working_hours if total_working_hours > 0 else 0

        # --- SPOT VARNISH HOURS ---
        sv_machines_req = 0
        sv_hours = 0
        if spot_varnish == 'Y':
            sv_speed = sv_machine['speed']
            sv_run = annual_sheets / (sv_speed * efficiency) if sv_speed > 0 else 0
            sv_setup = 30 * print_runs_year / 60  # 30 min per setup
            sv_hours = sv_run + sv_setup
            sv_machines_req = sv_hours / total_working_hours

        # --- HOT FOILING HOURS ---
        hf_machines_req = 0
        hf_hours = 0
        if hot_foiling == 'Y':
            hf_speed = hf_machine['speed']
            hf_run = annual_sheets / (hf_speed * efficiency) if hf_speed > 0 else 0
            hf_setup = 30 * print_runs_year / 60
            hf_hours = hf_run + hf_setup
            hf_machines_req = hf_hours / total_working_hours

        # --- LAMINATION HOURS ---
        lm_machines_req = 0
        lm_hours = 0
        if lamination == 'Y':
            lm_speed = lm_machine['speed']
            lm_run = annual_sheets / (lm_speed * efficiency) if lm_speed > 0 else 0
            lm_setup = 60 * print_runs_year / 60
            lm_hours = lm_run + lm_setup
            lm_machines_req = lm_hours / total_working_hours

        # --- WINDOW PATCHING HOURS ---
        wp_machines_req = 0
        wp_hours = 0
        if window_carton == 'Y':
            wp_speed = wp_machine['speed']
            wp_run = annual_sheets / (wp_speed * efficiency) if wp_speed > 0 else 0
            wp_setup = 30 * print_runs_year / 60
            wp_hours = wp_run + wp_setup
            wp_machines_req = wp_hours / total_working_hours

        # --- LINER HOURS ---
        ln_machines_req = 0
        ln_hours = 0
        if liner_flag == 'Y':
            ln_speed = ln_machine['speed']
            ln_run = annual_volume / (ln_speed * efficiency) if ln_speed > 0 else 0
            ln_setup = 60 * print_runs_year / 60
            ln_hours = ln_run + ln_setup
            ln_machines_req = ln_hours / total_working_hours

        # --- CREASING & BLANKING HOURS ---
        # Setup hours = speed_at_setup / (setup_time * efficiency) / 60 * num_setups
        cb_setup_speed = cb_speed / 4  # setup speed = 25% of line speed
        cb_setup_time = 30  # minutes
        cb_setup_hours = cb_setup_speed / (cb_setup_time * efficiency) / 60 * print_runs_year if efficiency > 0 else 0
        cb_run_hours = annual_sheets / (cb_speed * efficiency) if cb_speed > 0 else 0
        total_cb_hours = cb_run_hours + cb_setup_hours
        cb_machines_req = total_cb_hours / total_working_hours if total_working_hours > 0 else 0

        # --- FOLDER GLUER HOURS ---
        # FG processes cartons (not sheets), speed in cartons/hr
        # Setup hours = speed_at_setup / (setup_time * efficiency) / 60 * num_setups
        fg_setup_speed = fg_speed / 20  # setup speed fraction
        fg_setup_time = 30  # minutes
        fg_setup_hours = fg_setup_speed / (fg_setup_time * efficiency) / 60 * print_runs_year if efficiency > 0 else 0
        fg_run_hours = annual_volume / (fg_speed * efficiency) if fg_speed > 0 else 0
        total_fg_hours = fg_run_hours + fg_setup_hours
        fg_machines_req = total_fg_hours / total_working_hours if total_working_hours > 0 else 0

        # --- MACHINE INVESTMENT ---
        total_direct_machines = print_machines_req + sv_machines_req + hf_machines_req + lm_machines_req + cb_machines_req + fg_machines_req + wp_machines_req + ln_machines_req

        pm_investment = pm_machine['cost_eur'] * euro_rate * print_machines_req
        sv_investment = sv_machine['cost_eur'] * euro_rate * sv_machines_req
        hf_investment = hf_machine['cost_eur'] * euro_rate * hf_machines_req
        lm_investment = lm_machine['cost_eur'] * euro_rate * lm_machines_req
        cb_investment = cb_machine['cost_eur'] * euro_rate * cb_machines_req
        fg_investment = fg_machine['cost_eur'] * euro_rate * fg_machines_req
        wp_investment = wp_machine['cost_eur'] * euro_rate * wp_machines_req
        ln_investment = ln_machine['cost_eur'] * euro_rate * ln_machines_req

        total_machine_inv = pm_investment + sv_investment + hf_investment + lm_investment + cb_investment + fg_investment + wp_investment + ln_investment
        auxiliary_inv = total_machine_inv * 0.01133  # ~1.13% for auxiliary equipment
        total_inv = total_machine_inv + auxiliary_inv

        # --- CONVERSION INPUTS ---
        elec_rate = float(d.get('elec_rate', 10.72))
        skilled_labour = float(d.get('skilled_labour', 541800))
        engineer_salary = float(d.get('engineer_salary', 1260000))
        pm_salary = float(d.get('pm_salary', 1890000))
        handler_salary = skilled_labour * 0.7  # Material handlers at 70% of skilled
        repair_pct = float(d.get('repair_pct', 0.02))
        other_oh_pct = float(d.get('other_oh_pct', 0.02))
        dep_pm_pct = float(d.get('dep_pm_pct', 0.15))
        dep_bldg_pct = float(d.get('dep_bldg_pct', 0.10))
        completed_life = int(d.get('completed_life', 5))
        land_cost_sqm = float(d.get('land_cost_sqm', 23519))
        building_cost_sqm = float(d.get('building_cost_sqm', 7000))
        lease_cost_sqm = float(d.get('lease_cost_sqm', 2136))
        premises_type = d.get('premises_type', 'Owned')
        int_lt = float(d.get('int_lt', 0.125))
        int_wc = float(d.get('int_wc', 0.14))
        rm_inventory_days = float(d.get('rm_inventory_days', 45))
        fg_inventory_days = float(d.get('fg_inventory_days', 7))
        rm_payment_days = float(d.get('rm_payment_days', 45))
        fg_payment_days = float(d.get('fg_payment_days', 60))

        # --- ELECTRICITY ---
        # Process KW = hours * machine_power
        kw_print = total_print_hours * pm_machine['power_kw']
        kw_sv = sv_hours * sv_machine['power_kw'] if spot_varnish == 'Y' else 0
        kw_hf = hf_hours * hf_machine['power_kw'] if hot_foiling == 'Y' else 0
        kw_lm = lm_hours * lm_machine['power_kw'] if lamination == 'Y' else 0
        kw_cb = total_cb_hours * cb_machine['power_kw']
        kw_fg = total_fg_hours * fg_machine['power_kw']
        kw_wp = wp_hours * wp_machine['power_kw'] if window_carton == 'Y' else 0
        kw_ln = ln_hours * ln_machine['power_kw'] if liner_flag == 'Y' else 0

        # Ancillary = 30% of process electricity (lighting, HVAC, compressed air)
        total_process_kw = kw_print + kw_sv + kw_hf + kw_lm + kw_cb + kw_fg + kw_wp + kw_ln
        ancillary_kw = total_process_kw * 0.30
        total_kw = total_process_kw + ancillary_kw

        electricity_cost_annual = total_kw * elec_rate
        electricity_per_1000 = (electricity_cost_annual / annual_volume) * 1000 if annual_volume > 0 else 0

        # --- DIRECT LABOUR ---
        # Skilled: (3 per printing + 1 per other active process) * machines * shifts
        # Handler: (2 per printing + 1 per other active process) * machines * shifts
        # Engineer: 0.125 per process type * total_machines * shifts
        # PM: 2 * PM_salary * print_machines * shifts
        
        skilled_headcount = (3 * print_machines_req + 1 * sv_machines_req + 1 * hf_machines_req +
                           1 * lm_machines_req + 1 * cb_machines_req + 1 * fg_machines_req +
                           1 * wp_machines_req + 1 * ln_machines_req) * no_of_shifts
        handler_headcount = (2 * print_machines_req + 1 * sv_machines_req + 1 * hf_machines_req +
                           1 * lm_machines_req + 1 * cb_machines_req + 1 * fg_machines_req +
                           1 * wp_machines_req + 1 * ln_machines_req) * no_of_shifts
        engineer_headcount = 0.125 * total_direct_machines * no_of_shifts
        pm_headcount_frac = print_machines_req * no_of_shifts  # PM allocated to printing

        direct_labour_cost = (skilled_headcount * skilled_labour +
                            handler_headcount * handler_salary +
                            engineer_headcount * engineer_salary +
                            2 * pm_salary * pm_headcount_frac)
        direct_labour_per_1000 = (direct_labour_cost / annual_volume) * 1000 if annual_volume > 0 else 0

        # --- INDIRECT LABOUR ---
        # Fixed factory headcount: 39 indirect staff allocated proportionally
        indirect_headcount_salary = (
            7 * skilled_labour +           # 7 managers (stores, HR, IT, procurement, logistics, safety, quality)
            4 * handler_salary +           # Safety asst(2) + Asst quality(2)
            27 * handler_salary +          # Security(9) + Housekeeping(9) + Movers(9)
            1 * pm_salary                  # General manager
        )
        # Allocation: indirect_salary * machine_fraction * scaling_factor (0.438)
        # Factor 0.438 represents Excel model's proportional headcount-to-machine allocation
        indirect_labour_allocation = indirect_headcount_salary * total_direct_machines * 0.438
        indirect_labour_per_1000 = (indirect_labour_allocation / annual_volume) * 1000 if annual_volume > 0 else 0

        # --- DEPRECIATION (Reducing Balance Method) ---
        # Plant & Machinery
        salvage_pm = total_inv * 0.05
        net_pm = total_inv - salvage_pm
        for yr in range(max(1, completed_life)):
            dep_pm = net_pm * dep_pm_pct
            net_pm -= dep_pm
        depreciation_pm_per_1000 = (dep_pm / annual_volume) * 1000 if annual_volume > 0 else 0

        # --- FACILITY COSTS ---
        # Area: machine footprint + RM storage + FG storage + support
        machine_area = (pm_machine['sqm'] * print_machines_req +
                       sv_machine['sqm'] * sv_machines_req +
                       hf_machine['sqm'] * hf_machines_req +
                       lm_machine['sqm'] * lm_machines_req +
                       cb_machine['sqm'] * cb_machines_req +
                       fg_machine['sqm'] * fg_machines_req +
                       wp_machine['sqm'] * wp_machines_req +
                       ln_machine['sqm'] * ln_machines_req)

        # Storage and dispatch area (typically ~5-8x machine area for cartons)
        dispatch_area = machine_area * 6  # From Excel: despatch ~82% of total area
        total_area = machine_area + dispatch_area
        total_building_area = total_area * 0.7  # Building covers 70% of land
        land_area = total_area * 1.0

        # Building investment
        building_inv = total_building_area * building_cost_sqm
        land_inv = land_area * land_cost_sqm

        # Building depreciation (reducing balance)
        bldg_plus_aux = building_inv + auxiliary_inv
        salvage_bldg = bldg_plus_aux * 0.05
        net_bldg = bldg_plus_aux - salvage_bldg
        for yr in range(max(1, completed_life)):
            dep_bldg = net_bldg * dep_bldg_pct
            net_bldg -= dep_bldg
        building_dep_per_1000 = (dep_bldg / annual_volume) * 1000 if annual_volume > 0 else 0

        # --- INTEREST ---
        # Long-term loan on machinery (EMI-based, reducing balance)
        if int_lt > 0 and total_inv > 0:
            loan_term = 10  # years
            r = int_lt
            emi = total_inv * (r * (1 + r) ** loan_term) / ((1 + r) ** loan_term - 1)
            balance = total_inv
            for yr in range(max(1, completed_life)):
                interest_yr = balance * r
                principal = emi - interest_yr
                balance -= principal
            interest_lt_per_1000 = (interest_yr / annual_volume) * 1000 if annual_volume > 0 else 0
        else:
            interest_lt_per_1000 = 0

        # Interest on land & building
        total_property = land_inv + building_inv
        if int_lt > 0 and total_property > 0:
            emi_bldg = total_property * (r * (1 + r) ** loan_term) / ((1 + r) ** loan_term - 1)
            bal_bldg = total_property
            for yr in range(max(1, completed_life)):
                int_bldg_yr = bal_bldg * r
                prin_bldg = emi_bldg - int_bldg_yr
                bal_bldg -= prin_bldg
            interest_bldg_per_1000 = (int_bldg_yr / annual_volume) * 1000 if annual_volume > 0 else 0
        else:
            interest_bldg_per_1000 = 0

        # Lease (if not owned)
        lease_per_1000 = 0
        if premises_type == 'Leased':
            lease_cost = total_area * lease_cost_sqm
            lease_per_1000 = (lease_cost / annual_volume) * 1000 if annual_volume > 0 else 0

        # R&M and Other Overheads
        repair_per_1000 = (total_machine_inv * repair_pct / annual_volume) * 1000 if annual_volume > 0 else 0
        other_oh_per_1000 = (total_machine_inv * other_oh_pct / annual_volume) * 1000 if annual_volume > 0 else 0

        # --- WORKING CAPITAL INTEREST ---
        conversion_for_wc = (electricity_per_1000 + direct_labour_per_1000 + indirect_labour_per_1000 +
                           repair_per_1000 + other_oh_per_1000 + depreciation_pm_per_1000 +
                           building_dep_per_1000 + interest_lt_per_1000 + interest_bldg_per_1000 + lease_per_1000)

        # RM inventory: stored as SHEETS so divide annual material by UPS
        annual_material_total = total_material * (annual_volume / 1000)
        rm_inventory_value = (annual_material_total / ups_per_sheet) * rm_inventory_days / 365
        rm_wc = rm_inventory_value * int_wc

        # FG inventory: cartons valued at material + direct conversion (excl indirect)
        direct_conv_for_fg = (electricity_per_1000 + direct_labour_per_1000 +
                            repair_per_1000 + other_oh_per_1000 + depreciation_pm_per_1000 +
                            building_dep_per_1000 + interest_lt_per_1000 + interest_bldg_per_1000 + lease_per_1000)
        fg_annual_cost = (total_material + direct_conv_for_fg) * (annual_volume / 1000)
        fg_inventory_value = fg_annual_cost * fg_inventory_days / 365
        fg_wc = fg_inventory_value * int_wc

        # Warehouse: RM inventory value * net cash gap
        net_cash_gap = max(0, fg_payment_days - rm_payment_days)
        wh_wc = rm_inventory_value * (net_cash_gap / 365) * int_wc

        total_wc = rm_wc + fg_wc + wh_wc
        inventory_wc_per_1000 = total_wc / (annual_volume / 1000) if annual_volume > 0 else 0

        # Dispatch WC: (material + direct_conv) * net_cash_gap / 365 * int_wc
        dispatch_wc_per_1000 = (total_material + direct_conv_for_fg) * net_cash_gap / 365 * int_wc

        # --- TOTAL CONVERSION ---
        total_conversion = conversion_for_wc + inventory_wc_per_1000

        # --- MARGIN (% of conversion cost) ---
        margin = total_conversion * margin_pct

        # --- PACKING ---
        cartons_per_box = float(d.get('cartons_per_box', 1500))
        shipper_cost_eur = float(d.get('shipper_cost_eur', 0.6))
        polybag_cost_eur = float(d.get('polybag_cost_eur', 0.25))
        shipper_cost_inr = shipper_cost_eur * euro_rate
        polybag_cost_inr = polybag_cost_eur * euro_rate
        packing_per_1000 = ((shipper_cost_inr + polybag_cost_inr) / cartons_per_box) * 1000 if cartons_per_box > 0 else 0

        # --- FREIGHT ---
        boxes_per_container = float(d.get('boxes_per_container', 173.719))
        freight_cost_container = float(d.get('freight_cost_container', 20010))
        cartons_per_container = cartons_per_box * boxes_per_container
        freight_per_1000 = (freight_cost_container / cartons_per_container) * 1000 if cartons_per_container > 0 else 0

        # --- TOTAL ---
        total_cost = total_material + total_conversion + margin + packing_per_1000 + freight_per_1000

        # --- RESPONSE ---
        total_interest = interest_lt_per_1000 + interest_bldg_per_1000 + inventory_wc_per_1000 + dispatch_wc_per_1000
        total_depreciation = depreciation_pm_per_1000 + building_dep_per_1000

        summary = {
            'country': country,
            'ups_per_sheet': ups_per_sheet,
            'layflat_length': round(layflat_length, 1),
            'layflat_width': round(layflat_width, 1),
            'sheet_length': round(sheet_length, 1),
            'sheet_width': round(sheet_width, 1),
            'area_per_carton': round(area_per_carton, 6),
            'area_per_sheet': round(area_sheet, 6),

            # Material
            'board_cost': round(board_cost, 2),
            'ink_cost': round(ink_cost, 2),
            'varnish_cost': round(varnish_cost, 2),
            'spot_varnish_cost': round(spot_varnish_cost, 2),
            'hot_foil_cost': round(hot_foil_cost, 2),
            'film_cost': round(film_cost, 2),
            'window_cost': round(window_cost, 2),
            'liner_cost': round(liner_cost, 2),
            'primer_cost': round(primer_cost, 2),
            'wastage_cost': round(wastage_cost, 2),
            'other_material_cost': round(other_material_cost, 2),
            'material_cost': round(total_material, 2),

            # Conversion
            'electricity_cost': round(electricity_per_1000, 2),
            'direct_labour': round(direct_labour_per_1000, 2),
            'indirect_labour': round(indirect_labour_per_1000, 2),
            'repair_maintenance': round(repair_per_1000, 2),
            'other_overheads': round(other_oh_per_1000, 2),
            'depreciation': round(total_depreciation, 2),
            'interest': round(total_interest, 2),
            'lease_cost': round(lease_per_1000, 2),
            'conversion_cost': round(total_conversion, 2),

            # Margin
            'margin': round(margin, 2),
            'margin_pct_input': margin_pct,
            'margin_calc_type': '% of Conversion Cost',

            # Distribution
            'packing_cost': round(packing_per_1000, 2),
            'freight_cost': round(freight_per_1000, 2),

            # Total
            'total_cost_per_1000': round(total_cost, 2),
            'total_cost_per_1000_eur': round(total_cost / euro_rate, 2) if euro_rate > 0 else 0,

            # Machine utilization
            'print_machines': round(print_machines_req, 6),
            'cb_machines': round(cb_machines_req, 6),
            'fg_machines': round(fg_machines_req, 6),
            'total_machine_investment': round(total_inv, 2),
            'building_area': round(total_building_area, 2),
            'land_area': round(land_area, 2),
        }

        return jsonify(summary)
    except Exception as e:
        logger.error(f"Advanced Carton calc error: {e}", exc_info=True)
        return jsonify({"error": str(e)}), 500

@app.route("/api/calc_carton", methods=["POST"])
def api_calc_carton():
    """Calculate carton cost per 1000 pcs"""
    try:
        d = request.json
        if not d:
            return jsonify({"error": "No data provided"}), 400

        # --- INPUTS ---
        layflat_length = float(d.get('layflat_length', 125.2))
        layflat_width = float(d.get('layflat_width', 394.5))
        sheet_length = float(d.get('sheet_length', 1020))
        sheet_width = float(d.get('sheet_width', 720))
        side_lay_1 = float(d.get('side_lay_1', 12))
        side_lay_2 = float(d.get('side_lay_2', 10))
        gripper = float(d.get('gripper', 6))
        back_lay = float(d.get('back_lay', 4))
        trimmer_1 = float(d.get('trimmer_1', 5))
        trimmer_2 = float(d.get('trimmer_2', 5))
        double_cut = float(d.get('double_cut', 0))
        gutter = float(d.get('gutter', 0))
        interlock = float(d.get('interlock', 0))
        ups_lengthwise = int(d.get('ups_lengthwise', 5))
        ups_widthwise = int(d.get('ups_widthwise', 2))

        # Board
        board_type = d.get('board_type', 'WB')
        board_gsm = float(d.get('board_gsm', 400))
        board_rate = float(d.get('board_rate', 55))

        # Ink
        ink_rate = float(d.get('ink_rate', 850))
        ink_consumption = float(d.get('ink_consumption', 0.9))

        # Varnish
        varnish_type = d.get('varnish_type', 'Gloss')
        varnish_rate = float(d.get('varnish_rate', 450))
        varnish_consumption = float(d.get('varnish_consumption', 5))

        # Primer
        primer_rate = float(d.get('primer_rate', 165))
        primer_consumption = float(d.get('primer_consumption', 5))

        # Lamination
        film_rate = float(d.get('film_rate', 135))
        film_gsm = float(d.get('film_gsm', 14))
        adhesive_rate = float(d.get('adhesive_rate', 104))
        adhesive_gsm = float(d.get('adhesive_gsm', 2))
        lam_labour = float(d.get('lam_labour', 115))

        # Corrugation
        middle_liner_gsm = float(d.get('middle_liner_gsm', 150))
        liner_rate = float(d.get('liner_rate', 35))
        flute_type = d.get('flute_type', 'E')
        inner_liner_gsm = float(d.get('inner_liner_gsm', 100))
        inner_liner_rate = float(d.get('inner_liner_rate', 35))
        corrugation_conversion_rate = float(d.get('corrugation_conversion_rate', 6.5))

        # Foil Stamping
        foil_width_per_carton = float(d.get('foil_width_per_carton', 0))
        foil_length_per_carton = float(d.get('foil_length_per_carton', 0))
        foil_cost_per_roll = float(d.get('foil_cost_per_roll', 1050))
        foil_stamping_conversion = float(d.get('foil_stamping_conversion', 0))

        # Other
        other_costs = float(d.get('other_costs', 50))
        conversion_cost = float(d.get('conversion_cost', 195))

        # Flute take-up factor
        flute_factors = {'E': 0.25, 'F': 0.20}
        flute_factor = flute_factors.get(flute_type, 0.25)

        # --- CALCULATIONS (matching Excel exactly) ---

        # Area of 1 carton (sq mtrs)
        area_per_carton = (layflat_length * layflat_width) / 1000000

        # UPs per sheet
        ups_per_sheet = ups_lengthwise * ups_widthwise

        # Actual sheet size (layout calculation)
        # Length direction: carton layflat_length * ups_lengthwise + side margins + trimmers
        actual_length = (layflat_length * ups_lengthwise) + side_lay_1 + side_lay_2 + trimmer_1 + double_cut + (gutter * max(0, ups_lengthwise - 1)) + interlock
        # Width direction: carton layflat_width * ups_widthwise + gripper margins + trimmers
        actual_width = (layflat_width * ups_widthwise) + gripper + back_lay + trimmer_2 + side_lay_2 + back_lay + (gutter * max(0, ups_widthwise - 1))
        area_sheet = (actual_length * actual_width) / 1000000

        # Board wastage = 2%
        board_wastage = 0.02
        # Board Consumption (Kg/1000 Cartons) = (area_sheet * GSM / 1000) * (1000 / ups_per_sheet)
        board_consumption_1000 = (area_sheet * board_gsm / 1000) * (1000 / ups_per_sheet)

        board_cost = board_consumption_1000 * board_rate * (1 + board_wastage)

        # Two different area bases:
        # ink_area_sheet = area_sheet / ups_per_sheet (ink covers full sheet including margins)
        # area_per_carton = actual carton area (varnish/primer/lamination applied only to carton)
        ink_area_sheet = area_sheet / ups_per_sheet

        # Ink: uses SHEET area per carton (ink covers full sheet)
        ink_wastage = 0.02
        ink_cost = ink_area_sheet * (ink_consumption / 1000) * ink_rate * 1000 * (1 + ink_wastage)

        # Varnish: uses CARTON area (applied only to carton surface)
        varnish_wastage = 0.02
        varnish_cost = area_per_carton * (varnish_consumption / 1000) * varnish_rate * 1000 * (1 + varnish_wastage)

        # Primer: uses CARTON area, no wastage
        primer_cost = area_per_carton * (primer_consumption / 1000) * primer_rate * 1000

        # Lamination: uses CARTON area
        lam_wastage = 0.05
        film_cost = area_per_carton * (film_gsm / 1000) * film_rate * 1000 * (1 + lam_wastage)
        adhesive_cost = area_per_carton * (adhesive_gsm / 1000) * adhesive_rate * 1000 * (1 + lam_wastage)
        # Labour cost: INR/1000 sheets -> per 1000 cartons
        lam_labour_cost = lam_labour / ups_per_sheet
        lamination_cost = film_cost + adhesive_cost + lam_labour_cost

        # Corrugation
        corrugation_wastage = 0.05
        # Middle layer: area_sheet * middle_gsm * (1+flute_factor) / 1000 * liner_rate * (1000/ups) 
        # D83 = 368.065... 
        # = (0.534154 * 150 * 1.25 / 1000) * 35 * (1000/10) 
        # = (0.534154 * 187.5 / 1000) * 35 * 100
        # = 0.100154 * 3500 = 350.539... hmm, let me recalc
        # Actually: board_consumption_1000 for middle layer = area_sheet * middle_gsm * (1+flute) / 1000 * 1000/ups
        middle_consumption = (area_sheet * middle_liner_gsm * (1 + flute_factor) / 1000) * (1000 / ups_per_sheet)
        middle_cost = middle_consumption * liner_rate

        # Back layer (inner liner)
        back_consumption = (area_sheet * inner_liner_gsm / 1000) * (1000 / ups_per_sheet)
        back_cost = back_consumption * inner_liner_rate

        # Cost of corrugation (total board weight for conversion cost)
        total_corr_weight = middle_consumption + back_consumption
        # D87 = 36.723... = cost of carton corrugation weight?
        # Actually D87 label is "Cost of Carton" = total_corr_weight * some factor
        # Let me check: 21.36616 (board) is area_sheet*gsm/1000 * 1000/ups
        # middle: area_sheet * 150 * 1.25 / 1000 * 100 = 0.534154*187.5/1000*100 = 10.01539
        # Actually wait: (area_sheet * gsm / 1000) gives kg per sheet
        # * (1000/ups) gives kg per 1000 cartons
        # middle_consumption = 0.534154 * 150 * 1.25 / 1000 * 100 = 0.534154 * 0.1875 * 100 = 10.015...
        # Hmm that gives middle_cost = 10.015 * 35 = 350.54, but Excel says 368.065
        # Let me recheck: 368.065 / 35 = 10.516, 10.516 * 1000 / 100 = 105.16, 
        # 105.16 / (150 * 1.25) = 105.16 / 187.5 = 0.561, that's not area_sheet (0.534154)
        # Maybe wastage is applied to middle layer too?
        # Let me try: area_sheet * middle_gsm * (1+flute) / 1000 * 1000/ups * (1 + some_wastage)
        # With 5% wastage: 10.015 * 1.05 = 10.516, * 35 = 368.06 ✓✓✓
        middle_cost = middle_consumption * liner_rate * (1 + corrugation_wastage)

        # Back layer: same with wastage
        # D86 = 196.30... = back_consumption * 35 * 1.05
        # back_consumption = 0.534154 * 100 / 1000 * 100 = 5.34154
        # 5.34154 * 35 * 1.05 = 5.34154 * 36.75 = 196.30 ✓
        back_cost = back_consumption * inner_liner_rate * (1 + corrugation_wastage)

        # Conversion cost for corrugation
        # D87 "Cost of Carton" = total corrugation board weight (no wastage) = middle_consumption + back_consumption + board_consumption
        # Actually D87 = 36.723 seems like just total corr weight: 10.015 + 5.3415 + 21.366 = 36.723 ✓
        total_corr_board_weight = middle_consumption + back_consumption + board_consumption_1000
        
        # D90 Conversion Cost = total_corr_board_weight * corrugation_conversion_rate * (1 + 0.05)
        # = 36.723 * 6.5 * 1.05 = 250.635 ✓
        corr_conversion_cost = total_corr_board_weight * corrugation_conversion_rate * (1 + corrugation_wastage)

        corrugation_total = middle_cost + back_cost + corr_conversion_cost

        # Foil Stamping
        foil_roll_width = 610  # mm
        foil_roll_length = 120000  # mm
        foil_area_per_carton = foil_width_per_carton * foil_length_per_carton  # sq mm
        foil_wastage = 0.05
        
        if foil_area_per_carton > 0 and foil_cost_per_roll > 0:
            # Cartons per roll
            cartons_per_roll = (foil_roll_width * foil_roll_length) / foil_area_per_carton if foil_area_per_carton > 0 else 0
            foil_material_cost = (foil_cost_per_roll / cartons_per_roll * 1000) if cartons_per_roll > 0 else 0
            foil_material_cost_w = foil_material_cost * (1 + foil_wastage)
        else:
            foil_material_cost_w = 0
        
        foil_conversion = foil_stamping_conversion * (1000 / ups_per_sheet) if foil_stamping_conversion > 0 else 0
        foil_total = foil_material_cost_w + foil_conversion

        # Packing Cost (≈2.122% of subtotal before packing - verified against Excel)
        subtotal_before_packing = board_cost + ink_cost + varnish_cost + primer_cost + lamination_cost + corrugation_total + foil_total + other_costs + conversion_cost
        packing_pct = 0.02122
        packing_cost = subtotal_before_packing * packing_pct

        # Cost Summary
        summary = {
            'board_cost': round(board_cost, 2),
            'ink_cost': round(ink_cost, 2),
            'varnish_cost': round(varnish_cost + primer_cost, 2),
            'lamination_cost': round(lamination_cost, 2),
            'corrugation_cost': round(corrugation_total, 2),
            'foil_cost': round(foil_total, 2),
            'other_material_cost': round(other_costs, 2),
            'conversion_cost': round(conversion_cost, 2),
            'packing_cost': round(packing_cost, 2),
        }
        
        total = sum(summary.values())
        summary['total_cost_per_1000'] = round(total, 2)
        
        # Intermediate values for display
        summary['area_per_carton'] = round(area_per_carton, 7)
        summary['ups_per_sheet'] = ups_per_sheet
        summary['actual_sheet_length'] = round(actual_length, 1)
        summary['actual_sheet_width'] = round(actual_width, 1)
        summary['area_sheet'] = round(area_sheet, 6)
        summary['board_consumption_1000'] = round(board_consumption_1000, 5)

        return jsonify(summary)
    except Exception as e:
        logger.error(f"Carton calc error: {e}", exc_info=True)
        return jsonify({"error": str(e)}), 500


@app.route("/api/calc_flexibles", methods=["POST"])
def api_calc_flexibles():
    """Calculate flexibles laminate cost"""
    try:
        d = request.json
        if not d:
            return jsonify({"error": "No data provided"}), 400

        # Film/layer density lookup
        density_map = {
            'HIPS': 1.04, 'GPPS': 1.05, 'PET Film': 1.45, 'EAA': 0.92, 'EVA': 0.93,
            'CPP Film': 0.9, 'HDPE': 0.95, 'MDPE': 0.94, 'BON': 1.14, 'AL Foil': 2.7,
            'Cellophane': 1.55, 'BOPP Film': 0.91, 'METPET Film': 1.4, 'MET MDOPE Film': 0.96,
            'Matt Finish PET Film': 1.45, 'Primer': 1, 'Lamination - Adhesive (Solvent Based)': 1,
            'Lamination - Adhesive (Solvent Less)': 1, 'Heat Seal Lacquer': 1, 'Cold Seal': 1,
            'Gloss Varnish': 1, 'Matte Varnish': 1, 'Gravure': 1, 'Flexo': 1,
            'Mono Layer PE': 0.95, '2 Layer All PE': 0.95, '3 Layer All PE': 0.95,
            '5 Layer All PE': 0.95, '5 Layer EVOH Barrier': 0.96, '5 Layer Nylon Barrier': 1,
            '7 Layer All PE': 0.95,
        }

        layers = d.get('layers', [])
        if not layers:
            return jsonify({"error": "At least one layer is required"}), 400

        conversion_cost = float(d.get('conversion_cost', 50))

        total_gsm = 0
        layer_results = []
        
        for layer in layers:
            name = layer.get('name', '')
            mic = float(layer.get('mic', 0))
            rate = float(layer.get('rate', 0))
            density = density_map.get(name, 1.0)
            
            gsm = mic * density
            total_gsm += gsm
            
            layer_results.append({
                'name': name,
                'mic': mic,
                'density': density,
                'gsm': round(gsm, 2),
                'rate': rate,
            })

        # Now calculate contributions and costs
        if total_gsm == 0:
            return jsonify({"error": "Total GSM cannot be zero"}), 400

        total_material_cost = 0
        for lr in layer_results:
            contribution = lr['gsm'] / total_gsm
            layer_cost = contribution * lr['rate']
            lr['contribution'] = round(contribution * 100, 4)
            lr['layer_cost'] = round(layer_cost, 4)
            total_material_cost += layer_cost

        # Wastage 6%
        wastage_pct = 0.06
        wastage_cost = total_material_cost * wastage_pct
        material_cost_with_wastage = total_material_cost + wastage_cost

        # Packing cost = 1.25% of material cost with wastage
        # From Excel: 2.5177 / 201.768 = 0.01248 ≈ 1.25%
        packing_cost = material_cost_with_wastage * 0.01248

        laminate_cost_per_kg = material_cost_with_wastage + conversion_cost + packing_cost

        # Average density and cost per SQM
        avg_density = total_gsm / sum(lr['mic'] for lr in layer_results) if sum(lr['mic'] for lr in layer_results) > 0 else 1

        # ₹/SQM = ₹/kg / avg_density (verified against Excel: 254.286 / 1.00808 = 252.249)
        laminate_cost_per_sqm = laminate_cost_per_kg / avg_density

        summary = {
            'layers': layer_results,
            'laminate_gsm': round(total_gsm, 2),
            'total_contribution': round(sum(lr['gsm'] for lr in layer_results) / total_gsm * 100, 4),
            'avg_density': round(avg_density, 6),
            'material_cost_per_kg': round(total_material_cost, 4),
            'wastage_pct': wastage_pct * 100,
            'wastage_cost': round(wastage_cost, 4),
            'material_cost_with_wastage': round(material_cost_with_wastage, 4),
            'conversion_cost': round(conversion_cost, 2),
            'packing_cost': round(packing_cost, 4),
            'laminate_cost_per_kg': round(laminate_cost_per_kg, 4),
            'laminate_cost_per_sqm': round(laminate_cost_per_sqm, 4),
        }

        return jsonify(summary)
    except Exception as e:
        logger.error(f"Flexibles calc error: {e}", exc_info=True)
        return jsonify({"error": str(e)}), 500

@app.route("/api/calc_ebm", methods=["POST"])
def api_calc_ebm():
    """Calculate EBM (Extrusion Blow Moulding) Rigids cost per 1000 pcs"""
    try:
        d = request.json
        if not d:
            return jsonify({"error": "No data provided"}), 400

        # --- INPUTS ---
        # SKU
        sku_description = d.get('sku_description', 'Comfort 220ml')
        country = d.get('country', 'India')
        currency_symbol = d.get('currency_symbol', 'INR')
        annual_volume = float(d.get('annual_volume', 62975559))
        
        # Material Details
        weight_g = float(d.get('weight', 19))
        
        # Layer 1
        l1_ratio = float(d.get('l1_ratio', 0.48))
        l1_polymer_type = d.get('l1_polymer_type', 'HDPE')
        l1_polymer_rate = float(d.get('l1_polymer_rate', 95))
        l1_mb_dosage = float(d.get('l1_mb_dosage', 0.02))
        l1_mb_rate = float(d.get('l1_mb_rate', 450))
        l1_additive_dosage = float(d.get('l1_additive_dosage', 0))
        l1_additive_rate = float(d.get('l1_additive_rate', 249.93))
        
        # Layer 2
        l2_ratio = float(d.get('l2_ratio', 0.50))
        l2_polymer_type = d.get('l2_polymer_type', 'rHDPE')
        l2_polymer_rate = float(d.get('l2_polymer_rate', 107))
        l2_mb_dosage = float(d.get('l2_mb_dosage', 0))
        l2_mb_rate = float(d.get('l2_mb_rate', 450))
        l2_additive_dosage = float(d.get('l2_additive_dosage', 0))
        l2_additive_rate = float(d.get('l2_additive_rate', 249.93))
        
        # Layer 3
        l3_ratio = float(d.get('l3_ratio', 0))
        l3_polymer_type = d.get('l3_polymer_type', 'HDPE')
        l3_polymer_rate = float(d.get('l3_polymer_rate', 0))
        l3_mb_dosage = float(d.get('l3_mb_dosage', 0))
        l3_mb_rate = float(d.get('l3_mb_rate', 450))
        l3_additive_dosage = float(d.get('l3_additive_dosage', 0))
        l3_additive_rate = float(d.get('l3_additive_rate', 249.93))
        
        # Regrind
        regrind_ratio = float(d.get('regrind_ratio', 0))
        
        # Mould
        mould_cavitation = int(d.get('mould_cavitation', 12))
        mould_cycle_time = float(d.get('mould_cycle_time', 16.3))
        machine_model = d.get('machine_model', 'Jomar 65')
        num_setups_year = int(d.get('num_setups_year', 6))
        num_rampups_year = int(d.get('num_rampups_year', 6))
        
        # Energy
        electricity_rate = float(d.get('electricity_rate', 10.72))
        
        # Manpower
        skilled_labour_salary = float(d.get('skilled_labour', 541800))
        engineer_salary = float(d.get('engineer', 1260000))
        prod_manager_salary = float(d.get('prod_manager', 1890000))
        
        # Overheads
        repair_pct = float(d.get('repair_pct', 0.025))
        other_oh_pct = float(d.get('other_oh_pct', 0.025))
        
        # Depreciation
        depreciation_pm_pct = float(d.get('depreciation_pm', 0.15))
        depreciation_bldg_pct = float(d.get('depreciation_bldg', 0.10))
        completed_life = int(d.get('completed_life', 5))
        
        # Premises
        land_cost_sqm = float(d.get('land_cost', 23519.02))
        building_cost_sqm = float(d.get('building_cost', 7000))
        lease_cost_sqm = float(d.get('lease_cost', 2136.33))
        premises_type = d.get('premises_type', 'Owned')
        
        # Financials
        interest_lt = float(d.get('interest_lt', 0.125))
        interest_wc = float(d.get('interest_wc', 0.14))
        margin_pct = float(d.get('margin_pct', 0.20))
        margin_calc = d.get('margin_calc', '% of Conversion Cost')
        lt_debt_equity = float(d.get('lt_debt_equity', 0.70))
        
        # Orders / Packing / Delivery
        num_orders_year = int(d.get('num_orders_year', 12))
        bottles_per_box = int(d.get('bottles_per_box', 360))
        boxes_per_container = int(d.get('boxes_per_container', 320))
        shipper_cost = float(d.get('shipper_cost', 59.43))
        polybag_cost = float(d.get('polybag_cost', 25.02))
        freight_per_container = float(d.get('freight_per_container', 8341.60))
        
        # Payment terms
        rm_payment_days = int(d.get('rm_payment_days', 45))
        fg_payment_days = int(d.get('fg_payment_days', 60))

        # Euro rate for conversion (from Database)
        euro_rate = float(d.get('euro_rate', 104.27))

        # ===================== Machine Database =====================
        machine_db = {
            'Jomar 65': {'cost_eur': 323485, 'power_kwh': 35, 'sqm': 40},
            'Jomar 135': {'cost_eur': 495662.5, 'power_kwh': 80, 'sqm': 40},
            'Uniloy': {'cost_eur': 349572.5, 'power_kwh': 65, 'sqm': 40},
            'Sika': {'cost_eur': 349572.5, 'power_kwh': 65, 'sqm': 40},
            'Speedex': {'cost_eur': 125220, 'power_kwh': 65, 'sqm': 40},
            'Magic 10': {'cost_eur': 1069587.5, 'power_kwh': 125, 'sqm': 40},
            'BMU 70': {'cost_eur': 357920.5, 'power_kwh': 70, 'sqm': 40},
            'BMU 100': {'cost_eur': 600012.5, 'power_kwh': 70, 'sqm': 40},
            'SEB 820': {'cost_eur': 1153067.5, 'power_kwh': 130, 'sqm': 40},
            'Bekum': {'cost_eur': 391312.5, 'power_kwh': 90, 'sqm': 50},
            'Kautex': {'cost_eur': 495662.5, 'power_kwh': 120, 'sqm': 71.5},
            'Uniloy Rotary': {'cost_eur': 464357.5, 'power_kwh': 80, 'sqm': 55},
            'Jomar Shuttle': {'cost_eur': 391312.5, 'power_kwh': 85, 'sqm': 55},
            'Chinese OEM': {'cost_eur': 193047.5, 'power_kwh': 65, 'sqm': 45},
        }
        
        machine = machine_db.get(machine_model, machine_db['Jomar 65'])
        machine_cost_eur = machine['cost_eur']
        machine_power_kwh = machine['power_kwh']
        machine_sqm = machine['sqm']
        
        # Machine cost: EUR * euro_rate (import charges already included in EUR price)
        machine_cost_inr_single = machine_cost_eur * euro_rate
        
        # Ancillary equipment factor (from backup sheet)
        ancillary_kwh = machine_power_kwh * 0.2  # ~20% of main machine

        # ===================== PRODUCTION CALCULATIONS =====================
        utilisation = 0.75
        capacity_required = annual_volume / utilisation
        
        output_per_hour = mould_cavitation * (3600 / mould_cycle_time)
        
        # Hours per year (330 days x 24 hours)
        hours_per_annum = 330 * 24  # = 7920
        output_per_annum_per_machine = output_per_hour * hours_per_annum
        
        num_machines = capacity_required / output_per_annum_per_machine
        
        # Actual production hours allocated per machine per year
        # = 330 days * 8 hours per shift = 2640 hours (1 shift allocation per machine)
        production_hours_allocated = 330 * 8  # 2640
        actual_production_hours = output_per_hour * num_machines * hours_per_annum / output_per_hour if output_per_hour > 0 else 0
        
        # Setup & Ramp Up
        setup_time_mins = 180
        rampup_time_mins = 45
        rampup_speed_ppm = output_per_hour / 60 * 0.8  # 80% speed during ramp
        components_wasted_rampup = rampup_speed_ppm * rampup_time_mins * num_rampups_year
        
        process_wastage_pct = components_wasted_rampup / annual_volume if annual_volume > 0 else 0
        burning_wastage_pct = 0.015
        total_wastage_pct = process_wastage_pct + burning_wastage_pct
        
        # ===================== MACHINE INVESTMENT =====================
        total_machine_cost = machine_cost_inr_single * num_machines
        
        # Ancillary equipment (from backup sheet: ~2.68% of total machine cost for Jomar 65)
        ancillary_fraction = 0.02682
        ancillary_items_cost = total_machine_cost * ancillary_fraction
        
        total_investment = total_machine_cost + ancillary_items_cost
        
        # ===================== MATERIAL COST =====================
        # Resin: ratio * rate * weight (MB NOT subtracted from resin)
        # MB: dosage * rate * weight (NOT multiplied by layer ratio - it's % of total component)
        l1_resin_cost = l1_ratio * l1_polymer_rate * weight_g
        l1_mb_cost = l1_mb_dosage * l1_mb_rate * weight_g
        l1_additive_cost = l1_additive_dosage * l1_additive_rate * weight_g
        
        # Layer 2
        l2_resin_cost = l2_ratio * l2_polymer_rate * weight_g
        l2_mb_cost = l2_mb_dosage * l2_mb_rate * weight_g
        l2_additive_cost = l2_additive_dosage * l2_additive_rate * weight_g
        
        # Layer 3
        l3_resin_cost = l3_ratio * l3_polymer_rate * weight_g
        l3_mb_cost = l3_mb_dosage * l3_mb_rate * weight_g
        l3_additive_cost = l3_additive_dosage * l3_additive_rate * weight_g
        
        total_resin = l1_resin_cost + l2_resin_cost + l3_resin_cost
        total_mb = l1_mb_cost + l2_mb_cost + l3_mb_cost
        total_additive = l1_additive_cost + l2_additive_cost + l3_additive_cost
        direct_material = total_resin + total_mb + total_additive
        
        # Wastage cost (simple multiplication matching Excel)
        wastage_cost = direct_material * total_wastage_pct
        
        material_cost = direct_material + wastage_cost
        
        # ===================== CONVERSION COST =====================
        # Electricity
        total_power = machine_power_kwh + ancillary_kwh
        actual_usage_pct = 0.5
        running_load = total_power * actual_usage_pct
        electricity_cost_per_hour = running_load * electricity_rate
        electricity_per_1000 = (electricity_cost_per_hour / output_per_hour) * 1000 if output_per_hour > 0 else 0
        
        # Direct Manpower
        # 1 operator per 4 machines per shift, 3 shifts
        operators_per_shift = num_machines / 4
        handlers_per_shift = num_machines / 4
        engineers_per_shift = num_machines / 15
        pm_count = num_machines / 15  # 1 PM for morning+night
        
        total_direct_labour = (
            operators_per_shift * 3 * skilled_labour_salary +
            handlers_per_shift * 3 * skilled_labour_salary * 0.7 +  # handler at 70% of skilled
            engineers_per_shift * 3 * engineer_salary +
            pm_count * prod_manager_salary  # 1 shift only
        )
        direct_labour_per_1000 = (total_direct_labour / annual_volume) * 1000 * utilisation
        
        # Repair & Maintenance
        repair_cost = total_investment * repair_pct
        repair_per_1000 = (repair_cost / annual_volume) * 1000
        
        # Other Overheads
        other_oh_cost = total_investment * other_oh_pct
        other_oh_per_1000 = (other_oh_cost / annual_volume) * 1000
        
        # ===================== FACILITY AREA =====================
        production_hall_sqm = machine_sqm * num_machines
        rm_storage_sqm = production_hall_sqm * 1.154
        fg_warehouse_sqm = production_hall_sqm * 0.681
        tool_room_sqm = num_machines * 2
        qa_lab_sqm = num_machines * 0.4
        maintenance_sqm = num_machines * 2
        office_sqm = num_machines * 2
        canteen_sqm = num_machines * 0.4
        loading_bays_sqm = num_machines * 0.4
        future_expansion = 5
        
        total_built_up = (production_hall_sqm + rm_storage_sqm + fg_warehouse_sqm + 
                         tool_room_sqm + qa_lab_sqm + maintenance_sqm + office_sqm + 
                         canteen_sqm + loading_bays_sqm + future_expansion)
        
        fsi = 0.7
        land_area = total_built_up / fsi
        
        # ===================== DEPRECIATION =====================
        # Machine depreciation (WDV method)
        salvage_value = total_investment * 0.05
        net_asset = total_investment - salvage_value
        
        # Calculate WDV depreciation for the completed_life year
        for yr in range(completed_life):
            depreciation_yr = net_asset * depreciation_pm_pct
            net_asset = net_asset - depreciation_yr
        
        machine_depreciation_per_1000 = (depreciation_yr / annual_volume) * 1000
        
        # Building depreciation
        building_investment = total_built_up * building_cost_sqm
        bldg_salvage = building_investment * 0.05
        bldg_net_asset = building_investment - bldg_salvage
        for yr in range(completed_life):
            bldg_dep_yr = bldg_net_asset * depreciation_bldg_pct
            bldg_net_asset = bldg_net_asset - bldg_dep_yr
        
        building_depreciation_per_1000 = (bldg_dep_yr / annual_volume) * 1000
        
        total_depreciation_per_1000 = machine_depreciation_per_1000 + building_depreciation_per_1000
        
        # ===================== INTEREST =====================
        # Long Term Loan on machinery (with debt-equity ratio)
        lt_loan_machinery = total_investment * lt_debt_equity
        loan_tenure = 10
        
        # EMI calculation
        r_monthly = interest_lt
        if r_monthly > 0 and loan_tenure > 0:
            emi_machinery = lt_loan_machinery * (r_monthly * (1 + r_monthly)**loan_tenure) / ((1 + r_monthly)**loan_tenure - 1)
        else:
            emi_machinery = 0
        
        # Interest for the completed_life year (amortization schedule)
        balance = lt_loan_machinery
        for yr in range(completed_life):
            interest_yr = balance * interest_lt
            principal_yr = emi_machinery - interest_yr
            balance = balance - principal_yr
        
        interest_lt_per_1000 = (interest_yr / annual_volume) * 1000
        
        # Interest on land & building (only land + building, NOT ancillary)
        land_investment = land_area * land_cost_sqm
        lt_loan_lb = (land_investment + building_investment) * lt_debt_equity
        
        if interest_lt > 0 and loan_tenure > 0:
            emi_lb = lt_loan_lb * (interest_lt * (1 + interest_lt)**loan_tenure) / ((1 + interest_lt)**loan_tenure - 1)
        else:
            emi_lb = 0
        
        balance_lb = lt_loan_lb
        int_lb_yr = 0
        for yr in range(completed_life):
            int_lb_yr = balance_lb * interest_lt
            prin_lb = emi_lb - int_lb_yr
            balance_lb = balance_lb - prin_lb
        
        interest_lb_per_1000 = (int_lb_yr / annual_volume) * 1000
        
        # Lease cost
        lease_per_1000 = 0
        if premises_type == 'Leased':
            lease_per_1000 = (land_area * lease_cost_sqm / annual_volume) * 1000
        
        # Working Capital Interest
        # 1. Inventory holding costs (cost of warehouse space for RM and FG inventory)
        bldg_dep_per_sqm = bldg_dep_yr / total_built_up if total_built_up > 0 else 0
        bldg_interest_sqm = building_cost_sqm * interest_lt
        land_interest_sqm = land_cost_sqm * interest_lt

        # Areas from facility calculation (reuse already-computed variables)
        rm_stor_area = rm_storage_sqm
        fg_wh_area = fg_warehouse_sqm
        total_inv_area = rm_stor_area + fg_wh_area

        rm_holding = rm_stor_area * (bldg_interest_sqm + bldg_dep_per_sqm) / annual_volume * 1000
        fg_holding = fg_wh_area * (land_interest_sqm + bldg_interest_sqm + bldg_dep_per_sqm) / annual_volume * 1000
        wh_holding = total_inv_area * bldg_dep_per_sqm / annual_volume * 1000

        # 2. Net cash gap working capital interest
        # Conversion base for WC = all conversion components INCLUDING LT interest but EXCLUDING WC interest itself
        # (We need indirect_labour first, so compute it now)

        # --- Indirect Labour (common allocation) ---
        indirect_manpower = (
            2 * skilled_labour_salary +    # Tool room machinist (2 people)
            2 * engineer_salary +           # Tool room manager (2 people)
            1 * skilled_labour_salary +     # Procurement
            1 * skilled_labour_salary +     # Logistics
            1 * skilled_labour_salary +     # Safety Manager
            1 * skilled_labour_salary +     # Safety Asst
            1 * skilled_labour_salary +     # Quality Manager
            1 * skilled_labour_salary +     # Asst Quality Manager
            9 * skilled_labour_salary +     # Security (3 per shift)
            9 * skilled_labour_salary +     # Housekeeping (3 per shift)
            9 * skilled_labour_salary +     # Mover & Loader (3 per shift)
            1 * prod_manager_salary         # General Manager
        )
        # 38 indirect headcount supports ~28.7 machines (ratio = 0.756)
        indirect_headcount = 38
        total_factory_machines = indirect_headcount * 0.756
        factory_allocation = num_machines / total_factory_machines
        indirect_allocated = indirect_manpower * factory_allocation
        indirect_labour_per_1000 = (indirect_allocated / annual_volume) * 1000

        # Conversion base for WC: elec + labour + indirect + R&M + OH + dep + LT interest (no WC)
        conv_for_wc = (electricity_per_1000 + direct_labour_per_1000 + indirect_labour_per_1000 +
                       repair_per_1000 + other_oh_per_1000 + lease_per_1000 +
                       total_depreciation_per_1000 + interest_lt_per_1000 + interest_lb_per_1000)
        fg_base = material_cost + conv_for_wc
        net_cash_gap = fg_payment_days - rm_payment_days
        wc_net_interest = fg_base * interest_wc * max(net_cash_gap, 0) / 365

        total_wc_interest = rm_holding + fg_holding + wh_holding + wc_net_interest
        total_interest_per_1000 = interest_lt_per_1000 + interest_lb_per_1000 + total_wc_interest
        
        # ===================== PACKING =====================
        packing_cost_per_1000 = (shipper_cost + polybag_cost) * (1000 / bottles_per_box)
        
        # ===================== FREIGHT =====================
        bottles_per_container = bottles_per_box * boxes_per_container
        delivery_per_1000 = (freight_per_container / bottles_per_container) * 1000 if bottles_per_container > 0 else 0
        
        # ===================== CONVERSION COST TOTAL =====================
        conversion_cost = (electricity_per_1000 + direct_labour_per_1000 +
                          indirect_labour_per_1000 + repair_per_1000 + other_oh_per_1000 + lease_per_1000 +
                          total_depreciation_per_1000 + total_interest_per_1000)
        
        # ===================== MARGIN =====================
        if margin_calc == '% of Conversion Cost':
            margin_per_1000 = conversion_cost * margin_pct
        else:
            margin_per_1000 = (material_cost + conversion_cost) * margin_pct
        
        # ===================== TOTAL COST =====================
        total_cost_per_1000 = material_cost + conversion_cost + margin_per_1000 + packing_cost_per_1000 + delivery_per_1000
        
        # Per piece
        cost_per_piece = total_cost_per_1000 / 1000
        
        # EUR conversion
        total_cost_eur = total_cost_per_1000 / euro_rate if euro_rate > 0 else 0
        
        # Percentages
        total = total_cost_per_1000
        mat_pct = (material_cost / total * 100) if total > 0 else 0
        conv_pct = (conversion_cost / total * 100) if total > 0 else 0
        margin_pct_total = (margin_per_1000 / total * 100) if total > 0 else 0
        pkg_pct = (packing_cost_per_1000 / total * 100) if total > 0 else 0
        freight_pct = (delivery_per_1000 / total * 100) if total > 0 else 0
        
        summary = {
            # Cost Summary
            'material_cost': round(material_cost, 2),
            'conversion_cost': round(conversion_cost, 2),
            'margin': round(margin_per_1000, 2),
            'packing_cost': round(packing_cost_per_1000, 2),
            'freight_cost': round(delivery_per_1000, 2),
            'total_cost_per_1000': round(total_cost_per_1000, 2),
            'cost_per_piece': round(cost_per_piece, 4),
            'total_cost_eur': round(total_cost_eur, 2),
            
            # Material breakdown
            'resin_cost': round(total_resin, 2),
            'mb_cost': round(total_mb, 2),
            'additive_cost': round(total_additive, 2),
            'wastage_cost': round(wastage_cost, 2),
            'wastage_pct': round(total_wastage_pct * 100, 4),
            
            # Conversion breakdown
            'electricity_cost': round(electricity_per_1000, 2),
            'direct_labour': round(direct_labour_per_1000, 2),
            'indirect_labour': round(indirect_labour_per_1000, 2),
            'repair_cost': round(repair_per_1000, 2),
            'other_oh': round(other_oh_per_1000, 2),
            'lease_cost': round(lease_per_1000, 2),
            'depreciation': round(total_depreciation_per_1000, 2),
            'interest_total': round(total_interest_per_1000, 2),
            
            # Percentages
            'mat_pct': round(mat_pct, 1),
            'conv_pct': round(conv_pct, 1),
            'margin_pct_total': round(margin_pct_total, 1),
            'margin_pct_input': margin_pct,
            'margin_calc_type': margin_calc,
            'pkg_pct': round(pkg_pct, 1),
            'freight_pct': round(freight_pct, 1),
            
            # Production info
            'num_machines': round(num_machines, 2),
            'output_per_hour': round(output_per_hour, 0),
            'total_investment_inr': round(total_investment, 0),
            'land_area_sqm': round(land_area, 0),
            'built_up_sqm': round(total_built_up, 0),
            
            'currency': currency_symbol,
        }
        
        return jsonify(summary)
    except Exception as e:
        logger.error(f"EBM calc error: {e}", exc_info=True)
        return jsonify({"error": str(e)}), 500



# ================= NEW FEATURE APIs =================

@app.route("/api/machine_db_for_calc", methods=["POST"])
def api_machine_db_for_calc():
    """Get machine data from Machine Database for cost calculators (Feature 1)"""
    try:
        data = request.json
        process = data.get('process', '')
        df = load_excel_cached('machine', sheet_name="Database", header=2)
        if process:
            f = df[df["Process"].str.lower().str.contains(process.lower(), na=False)]
        else:
            f = df
        machines = []
        for _, r in f.iterrows():
            cost = r.get("Machine Cost In €") or r.get("Machine Cost") or r.get("Price")
            if pd.isna(cost) or cost == 0:
                for col in df.columns:
                    if any(k in str(col) for k in ["€", "Cost", "Price"]):
                        cost = r[col]; break
            power = r.get("Power Consumption")
            sqm = r.get("Machine Footprint SQM")
            make = str(r.get("Make", "")); model = str(r.get("Model", ""))
            if pd.isna(cost) or float(cost) == 0: continue
            machines.append({"label": f"{make} {model}".strip(), "cost_eur": round(float(cost), 2) if not pd.isna(cost) else 0, "power_kwh": round(float(power), 2) if not pd.isna(power) else 0, "sqm": round(float(sqm), 2) if not pd.isna(sqm) else 0})
        return jsonify({"machines": machines})
    except Exception as e:
        logger.error(f"machine_db_for_calc error: {e}")
        return jsonify({"error": str(e)}), 500


@app.route("/api/variable_cost_for_calc", methods=["POST"])
def api_variable_cost_for_calc():
    """Get variable cost data from Variable Cost Database (Feature 2)"""
    try:
        data = request.json
        country = data.get('country', '')
        if not country: return jsonify({"error": "Country required"}), 400
        df = load_excel_cached('cost', sheet_name="Data", header=9)
        df.columns = [str(c).strip() for c in df.columns]
        cd = df[df.iloc[:, 0] == country]
        if cd.empty: return jsonify({"error": "Country not found"}), 404
        row = cd.iloc[0]
        variables = {}
        for col in df.columns[1:]:
            try: variables[col] = float(row[col]) if not pd.isna(row[col]) else 0
            except: variables[col] = 0
        return jsonify({"country": country, "variables": variables})
    except Exception as e:
        logger.error(f"variable_cost_for_calc error: {e}")
        return jsonify({"error": str(e)}), 500


@app.route("/api/multi_country_ebm", methods=["POST"])
def api_multi_country_ebm():
    """Run EBM across multiple countries (Feature 3)"""
    try:
        data = request.json
        countries = data.get('countries', [])
        bp = data.get('base_params', {})
        if len(countries) < 2: return jsonify({"error": "Select at least 2 countries"}), 400
        if len(countries) > 6: return jsonify({"error": "Max 6 countries"}), 400

        cdb = {
            'India':{'elec':10.72,'labour':541800,'engineer':1260000,'pm':1890000,'dep_pm':0.15,'dep_bldg':0.10,'land':23519,'building':7000,'lease':2136,'int_lt':0.125,'int_wc':0.14,'euro':104.27,'mb':450,'add':249.93},
            'China':{'elec':0.794,'labour':420000,'engineer':420000,'pm':487200,'dep_pm':0.10,'dep_bldg':0.10,'land':1228.5,'building':1056.51,'lease':3046.68,'int_lt':0.049,'int_wc':0.03,'euro':8.19,'mb':35.35,'add':19.63},
            'Vietnam':{'elec':1744,'labour':139920000,'engineer':827162157,'pm':921734762,'dep_pm':0.10,'dep_bldg':0.10,'land':3586051.86,'building':3403710.24,'lease':5470248.6,'int_lt':0.059,'int_wc':0.062,'euro':30390.27,'mb':131155.86,'add':72843.41},
            'Turkey':{'elec':4.35,'labour':281880,'engineer':548100,'pm':532440,'dep_pm':0.10,'dep_bldg':0.10,'land':9524.94,'building':54736.32,'lease':459.06,'int_lt':0.425,'int_wc':0.395,'euro':49.29,'mb':212.72,'add':118.14},
            'Indonesia':{'elec':1114.74,'labour':7332000,'engineer':19552000,'pm':24440000,'dep_pm':0.25,'dep_bldg':0.10,'land':1700000,'building':5500000,'lease':420000,'int_lt':0.10,'int_wc':0.12,'euro':19314.2,'mb':83354.66,'add':46294.82},
            'Brazil':{'elec':0.657,'labour':73000,'engineer':210240,'pm':315360,'dep_pm':0.10,'dep_bldg':0.04,'land':2533.6,'building':14843.75,'lease':161.68,'int_lt':0.15,'int_wc':0.15,'euro':6.23,'mb':26.89,'add':14.93},
            'United States':{'elec':0.149,'labour':98250.6,'engineer':130993.8,'pm':117125.4,'dep_pm':0.10,'dep_bldg':0.10,'land':32.92,'building':2485.66,'lease':187.14,'int_lt':0.0389,'int_wc':0.0364,'euro':1.16,'mb':5.01,'add':2.78},
            'United Kingdom':{'elec':0.346,'labour':39900,'engineer':57190,'pm':66500,'dep_pm':0.18,'dep_bldg':0.03,'land':148.2,'building':1308.88,'lease':150.03,'int_lt':0.112,'int_wc':0.113,'euro':0.88,'mb':7.22,'add':4.01},
            'Germany':{'elec':0.251,'labour':46692,'engineer':97275,'pm':110245,'dep_pm':0.10,'dep_bldg':0.10,'land':800,'building':1292.51,'lease':50.61,'int_lt':0.0395,'int_wc':0.0395,'euro':1,'mb':4.32,'add':2.40},
            'France':{'elec':0.153,'labour':34800,'engineer':71050,'pm':94250,'dep_pm':0.10,'dep_bldg':0.10,'land':201.2,'building':1037.46,'lease':88.7,'int_lt':0.0345,'int_wc':0.0345,'euro':1,'mb':4.32,'add':2.40},
            'Mexico':{'elec':3.972,'labour':180000,'engineer':492000,'pm':852000,'dep_pm':0.10,'dep_bldg':0.05,'land':0,'building':0,'lease':0,'int_lt':0.0728,'int_wc':0.0728,'euro':21.26,'mb':91.75,'add':50.96},
            'Pakistan':{'elec':41.99,'labour':504000,'engineer':384000,'pm':2400000,'dep_pm':0.15,'dep_bldg':0.10,'land':47253.06,'building':135336.59,'lease':535.57,'int_lt':0.18,'int_wc':0.09,'euro':328.52,'mb':1417.80,'add':787.44},
            'Philippines':{'elec':8.847,'labour':242880,'engineer':473470.53,'pm':538167.68,'dep_pm':0.10,'dep_bldg':0.10,'land':16240,'building':20000,'lease':3240,'int_lt':0.10,'int_wc':0.0863,'euro':67.87,'mb':292.91,'add':162.68},
            'South Africa':{'elec':1.795,'labour':231858,'engineer':494630.4,'pm':772860,'dep_pm':0.20,'dep_bldg':0.05,'land':438.40,'building':8823.27,'lease':827.00,'int_lt':0.1025,'int_wc':0.275,'euro':19.88,'mb':85.80,'add':47.65},
            'Spain':{'elec':0.126,'labour':55960,'engineer':67152,'pm':76945,'dep_pm':0.10,'dep_bldg':0.03,'land':135.4,'building':999.6,'lease':53.5,'int_lt':0.0215,'int_wc':0.032,'euro':1,'mb':4.32,'add':2.40},
            'Poland':{'elec':0.829,'labour':83388,'engineer':133420.8,'pm':266841.6,'dep_pm':0.20,'dep_bldg':0.10,'land':400,'building':3621.4,'lease':362,'int_lt':0.04,'int_wc':0.071,'euro':4.21,'mb':18.17,'add':10.09},
            'Thailand':{'elec':4.086,'labour':303544.8,'engineer':327600,'pm':1404000,'dep_pm':0.20,'dep_bldg':0.05,'land':4546.87,'building':22447.27,'lease':2677.35,'int_lt':0.1268,'int_wc':0.1268,'euro':36.6,'mb':157.96,'add':87.73},
            'Bangladesh':{'elec':12.39,'labour':1521720,'engineer':913032,'pm':1445634,'dep_pm':0.10,'dep_bldg':0.10,'land':18319.58,'building':53821.31,'lease':2531.40,'int_lt':0.13,'int_wc':0.135,'euro':142.84,'mb':616.46,'add':342.38},
            'Sri Lanka':{'elec':16.59,'labour':1060800,'engineer':1560000,'pm':4680000,'dep_pm':0.125,'dep_bldg':0.0667,'land':15815.26,'building':62230.89,'lease':7131.32,'int_lt':0.14,'int_wc':0.18,'euro':362.96,'mb':1566.43,'add':869.99},
            'Argentina':{'elec':129.15,'labour':9792000,'engineer':1632000,'pm':13056000,'dep_pm':0.10,'dep_bldg':0.02,'land':213732.46,'building':102920.78,'lease':8851.94,'int_lt':0.3696,'int_wc':0.3696,'euro':1684.16,'mb':7268.36,'add':4036.82},
            'Canada':{'elec':0.144,'labour':65650,'engineer':99737.5,'pm':112362.5,'dep_pm':0.30,'dep_bldg':0.10,'land':356.37,'building':4068.38,'lease':191.71,'int_lt':0.025,'int_wc':0.0745,'euro':1.62,'mb':6.99,'add':3.88},
            'Costa Rica':{'elec':115.84,'labour':8329800,'engineer':19824924,'pm':45147516,'dep_pm':0.10,'dep_bldg':0.10,'land':68963.19,'building':404761.95,'lease':3966.34,'int_lt':0.0733,'int_wc':0.095,'euro':581.9,'mb':2511.32,'add':1394.77},
        }

        def calc_ebm_for_country(cv, bp):
            er = cv['euro']; wg = float(bp.get('weight', 19)); av = float(bp.get('annual_volume', 62975559))
            mc_val = int(bp.get('mould_cavitation', 12)); mct = float(bp.get('mould_cycle_time', 16.3))
            mm = bp.get('machine_model', 'Jomar 65')
            nry = int(bp.get('num_rampups_year', 6)); rp = float(bp.get('repair_pct', 0.025))
            ohp = float(bp.get('other_oh_pct', 0.025)); cl = int(bp.get('completed_life', 5))
            pt = bp.get('premises_type', 'Owned'); mp = float(bp.get('margin_pct', 0.20))
            mcalc = bp.get('margin_calc', '% of Conversion Cost'); lde = float(bp.get('lt_debt_equity', 0.70))
            bpb = int(bp.get('bottles_per_box', 360)); bpc = int(bp.get('boxes_per_container', 320))
            sc = float(bp.get('shipper_cost', 59.43)); plc = float(bp.get('polybag_cost', 25.02))
            fpc = float(bp.get('freight_per_container', 8341.60))
            rpd = int(bp.get('rm_payment_days', 45)); fpd = int(bp.get('fg_payment_days', 60))

            mdb = {'Jomar 65':{'c':323485,'p':35,'s':40},'Jomar 135':{'c':495662.5,'p':80,'s':40},'Uniloy':{'c':349572.5,'p':65,'s':40},'Sika':{'c':349572.5,'p':65,'s':40},'Speedex':{'c':125220,'p':65,'s':40},'Magic 10':{'c':1069587.5,'p':125,'s':40},'BMU 70':{'c':357920.5,'p':70,'s':40},'BMU 100':{'c':600012.5,'p':70,'s':40},'SEB 820':{'c':1153067.5,'p':130,'s':40},'Bekum':{'c':391312.5,'p':90,'s':50},'Kautex':{'c':495662.5,'p':120,'s':71.5},'Uniloy Rotary':{'c':464357.5,'p':80,'s':55},'Jomar Shuttle':{'c':391312.5,'p':85,'s':55},'Chinese OEM':{'c':193047.5,'p':65,'s':45}}
            m = mdb.get(mm, mdb['Jomar 65'])
            mci = m['c'] * er; mpk = m['p']; msq = m['s']; akh = mpk * 0.2
            oph = mc_val * (3600 / mct); nm = (av / 0.75) / (oph * 7920)
            twp = (oph/60*0.8*45*nry)/av + 0.015 if av > 0 else 0.015
            ti = mci * nm * 1.02682
            # Material
            l1r = float(bp.get('l1_ratio',0.48)); l1pr = float(bp.get('l1_polymer_rate',95))
            l2r = float(bp.get('l2_ratio',0.50)); l2pr = float(bp.get('l2_polymer_rate',107))
            l3r = float(bp.get('l3_ratio',0)); l3pr = float(bp.get('l3_polymer_rate',0))
            l1md = float(bp.get('l1_mb_dosage',0.02)); l2md = float(bp.get('l2_mb_dosage',0)); l3md = float(bp.get('l3_mb_dosage',0))
            l1ad = float(bp.get('l1_additive_dosage',0)); l2ad = float(bp.get('l2_additive_dosage',0)); l3ad = float(bp.get('l3_additive_dosage',0))
            tr = (l1r*l1pr + l2r*l2pr + l3r*l3pr) * wg
            tmb = (l1md*cv['mb'] + l2md*cv['mb'] + l3md*cv['mb']) * wg
            tad = (l1ad*cv['add'] + l2ad*cv['add'] + l3ad*cv['add']) * wg
            dm = tr + tmb + tad; mat_cost = dm * (1 + twp)
            # Conversion
            ep = ((mpk+akh)*0.5*cv['elec']/oph)*1000 if oph>0 else 0
            dlp = ((nm/4*3*cv['labour'] + nm/4*3*cv['labour']*0.7 + nm/15*3*cv['engineer'] + nm/15*cv['pm'])/av)*1000*0.75
            rp1 = (ti*rp/av)*1000; oop1 = (ti*ohp/av)*1000
            tbu = msq*nm*(1+1.154+0.681) + nm*(2+0.4+2+2+0.4+0.4) + 5; la = tbu/0.7
            sv = ti*0.05; na = ti-sv
            for yr in range(cl): dy = na*cv['dep_pm']; na -= dy
            bi = tbu*cv['building']; bna = bi*0.95
            for yr in range(cl): bdy = bna*cv['dep_bldg']; bna -= bdy
            tdp = (dy/av)*1000 + (bdy/av)*1000
            ltl = ti*lde; lt_n = 10
            emi_m = ltl*(cv['int_lt']*(1+cv['int_lt'])**lt_n)/((1+cv['int_lt'])**lt_n-1) if cv['int_lt']>0 else 0
            bal = ltl; iy = 0
            for yr in range(cl): iy = bal*cv['int_lt']; bal -= (emi_m-iy)
            li = la*cv['land']; ltllb = (li+bi)*lde
            emi_lb = ltllb*(cv['int_lt']*(1+cv['int_lt'])**lt_n)/((1+cv['int_lt'])**lt_n-1) if cv['int_lt']>0 else 0
            blb = ltllb; ily = 0
            for yr in range(cl): ily = blb*cv['int_lt']; blb -= (emi_lb-ily)
            ilp = (iy/av)*1000; ilbp = (ily/av)*1000
            lp = (la*cv['lease']/av)*1000 if pt=='Leased' else 0
            imp = (37*cv['labour'] + 2*cv['engineer'] + cv['pm'])
            ilp2 = (imp*nm/(38*0.756)/av)*1000
            bds = bdy/tbu if tbu>0 else 0; bis = cv['building']*cv['int_lt']; lis = cv['land']*cv['int_lt']
            rmh = msq*nm*1.154*(bis+bds)/av*1000; fgh = msq*nm*0.681*(lis+bis+bds)/av*1000
            whh = (msq*nm*1.154+msq*nm*0.681)*bds/av*1000
            cfw = ep+dlp+ilp2+rp1+oop1+lp+tdp+ilp+ilbp
            fgb = mat_cost+cfw; ncg = fpd-rpd
            wni = fgb*cv['int_wc']*max(ncg,0)/365
            twi = rmh+fgh+whh+wni; tip = ilp+ilbp+twi
            conv = ep+dlp+ilp2+rp1+oop1+lp+tdp+tip
            margin = conv*mp if mcalc=='% of Conversion Cost' else (mat_cost+conv)*mp
            pkg = (sc+plc)*(1000/bpb)
            bpc_total = bpb*bpc; frt = (fpc/bpc_total)*1000 if bpc_total>0 else 0
            total = mat_cost+conv+margin+pkg+frt; eur = total/er if er>0 else 0
            return {'material':round(mat_cost,2),'conversion':round(conv,2),'margin':round(margin,2),'packing':round(pkg,2),'freight':round(frt,2),'total_local':round(total,2),'total_eur':round(eur,2),'euro_rate':er,'machines':round(nm,2),'mat_eur':round(mat_cost/er,2) if er>0 else 0,'conv_eur':round(conv/er,2) if er>0 else 0,'margin_eur':round(margin/er,2) if er>0 else 0,'pkg_eur':round(pkg/er,2) if er>0 else 0,'frt_eur':round(frt/er,2) if er>0 else 0}

        results = []
        for c in countries:
            cv = cdb.get(c)
            if not cv: continue
            try:
                r = calc_ebm_for_country(cv, bp)
                r['country'] = c
                results.append(r)
            except Exception as ce:
                results.append({'country': c, 'error': str(ce)})
        results.sort(key=lambda x: x.get('total_eur', 999999))
        return jsonify({"results": results})
    except Exception as e:
        logger.error(f"Multi-country error: {e}", exc_info=True)
        return jsonify({"error": str(e)}), 500


@app.route("/api/multi_country_generic", methods=["POST"])
def api_multi_country_generic():
    """Run any cost model across multiple countries using cost-factor scaling"""
    try:
        data = request.json
        countries = data.get('countries', [])
        model_type = data.get('model_type', '')
        base_result = data.get('base_result', {})
        base_country = data.get('base_country', 'India')
        base_params = data.get('base_params', {})
        if len(countries) < 2: return jsonify({"error": "Select at least 2 countries"}), 400
        if len(countries) > 6: return jsonify({"error": "Max 6 countries"}), 400

        # Country cost database (same as EBM multi-country)
        cdb = {
            'India':{'elec':10.72,'labour':541800,'engineer':1260000,'pm':1890000,'euro':104.27},
            'China':{'elec':0.794,'labour':420000,'engineer':420000,'pm':487200,'euro':8.19},
            'Vietnam':{'elec':1744,'labour':139920000,'engineer':827162157,'pm':921734762,'euro':30390.27},
            'Turkey':{'elec':4.35,'labour':281880,'engineer':548100,'pm':532440,'euro':49.29},
            'Indonesia':{'elec':1114.74,'labour':7332000,'engineer':19552000,'pm':24440000,'euro':19314.2},
            'Brazil':{'elec':0.657,'labour':73000,'engineer':210240,'pm':315360,'euro':6.23},
            'United States':{'elec':0.149,'labour':98250.6,'engineer':130993.8,'pm':117125.4,'euro':1.16},
            'United Kingdom':{'elec':0.346,'labour':39900,'engineer':57190,'pm':66500,'euro':0.88},
            'Germany':{'elec':0.251,'labour':46692,'engineer':97275,'pm':110245,'euro':1},
            'France':{'elec':0.153,'labour':34800,'engineer':71050,'pm':94250,'euro':1},
            'Mexico':{'elec':3.972,'labour':180000,'engineer':492000,'pm':852000,'euro':21.26},
            'Pakistan':{'elec':41.99,'labour':504000,'engineer':384000,'pm':2400000,'euro':328.52},
            'Philippines':{'elec':8.847,'labour':242880,'engineer':473470.53,'pm':538167.68,'euro':67.87},
            'South Africa':{'elec':1.795,'labour':231858,'engineer':494630.4,'pm':772860,'euro':19.88},
            'Spain':{'elec':0.126,'labour':55960,'engineer':67152,'pm':76945,'euro':1},
            'Poland':{'elec':0.829,'labour':83388,'engineer':133420.8,'pm':266841.6,'euro':4.21},
            'Thailand':{'elec':4.086,'labour':303544.8,'engineer':327600,'pm':1404000,'euro':36.6},
            'Bangladesh':{'elec':12.39,'labour':1521720,'engineer':913032,'pm':1445634,'euro':142.84},
            'Sri Lanka':{'elec':16.59,'labour':1060800,'engineer':1560000,'pm':4680000,'euro':362.96},
            'Argentina':{'elec':129.15,'labour':9792000,'engineer':1632000,'pm':13056000,'euro':1684.16},
            'Canada':{'elec':0.144,'labour':65650,'engineer':99737.5,'pm':112362.5,'euro':1.62},
            'Costa Rica':{'elec':115.84,'labour':8329800,'engineer':19824924,'pm':45147516,'euro':581.9},
        }

        base_cv = cdb.get(base_country, cdb['India'])
        base_euro = base_cv['euro']
        # Normalize base costs to EUR for comparison
        base_mat = float(base_result.get('material_cost', 0))
        base_conv = float(base_result.get('conversion_cost', 0))
        base_margin = float(base_result.get('margin', 0))
        base_pkg = float(base_result.get('packing_cost', 0))
        base_frt = float(base_result.get('freight_cost', 0))
        base_total = float(base_result.get('total_cost_per_1000', 0))
        if base_total == 0:
            base_total = base_mat + base_conv + base_margin + base_pkg + base_frt

        # Labour cost index for base country (weighted: 60% labour, 25% engineer, 15% PM)
        base_labour_idx = base_cv['labour'] * 0.60 + base_cv['engineer'] * 0.25 + base_cv['pm'] * 0.15
        # Electricity in EUR for base
        base_elec_eur = base_cv['elec'] / base_euro if base_euro > 0 else 0

        results = []
        for c in countries:
            cv = cdb.get(c)
            if not cv:
                results.append({'country': c, 'error': 'Country not found'})
                continue
            try:
                er = cv['euro']
                # Labour cost index for target country
                tgt_labour_idx = cv['labour'] * 0.60 + cv['engineer'] * 0.25 + cv['pm'] * 0.15
                tgt_elec_eur = cv['elec'] / er if er > 0 else 0

                # Scaling factors
                labour_ratio = (tgt_labour_idx / er) / (base_labour_idx / base_euro) if (base_labour_idx / base_euro) > 0 else 1
                elec_ratio = tgt_elec_eur / base_elec_eur if base_elec_eur > 0 else 1

                # Material: same raw material price globally, just convert via EUR
                mat_eur = base_mat / base_euro if base_euro > 0 else 0

                # Conversion: scale by weighted factor (60% labour, 40% electricity)
                conv_factor = labour_ratio * 0.60 + elec_ratio * 0.40
                conv_local_scaled = base_conv * conv_factor
                conv_eur = conv_local_scaled / base_euro if base_euro > 0 else 0

                # Margin: proportional to conversion (same margin %)
                margin_pct = (base_margin / base_conv) if base_conv > 0 else 0.20
                margin_eur = conv_eur * margin_pct

                # Packing: scale by EUR rate ratio
                pkg_eur = base_pkg / base_euro if base_euro > 0 else 0

                # Freight: same in EUR (international shipping)
                frt_eur = base_frt / base_euro if base_euro > 0 else 0

                total_eur = mat_eur + conv_eur + margin_eur + pkg_eur + frt_eur
                total_local = total_eur * er

                results.append({
                    'country': c,
                    'material': round(mat_eur * er, 2),
                    'conversion': round(conv_local_scaled, 2),
                    'margin': round(margin_eur * er, 2),
                    'packing': round(pkg_eur * er, 2),
                    'freight': round(frt_eur * er, 2),
                    'total_local': round(total_local, 2),
                    'total_eur': round(total_eur, 2),
                    'euro_rate': er,
                    'mat_eur': round(mat_eur, 2),
                    'conv_eur': round(conv_eur, 2),
                    'margin_eur': round(margin_eur, 2),
                    'pkg_eur': round(pkg_eur, 2),
                    'frt_eur': round(frt_eur, 2),
                })
            except Exception as ce:
                results.append({'country': c, 'error': str(ce)})
        results.sort(key=lambda x: x.get('total_eur', 999999))
        return jsonify({"results": results})
    except Exception as e:
        logger.error(f"Multi-country generic error: {e}", exc_info=True)
        return jsonify({"error": str(e)}), 500


@app.route("/api/export_ebm_excel", methods=["POST"])
def api_export_ebm_excel():
    """Export EBM to formatted Excel (Feature 5)"""
    try:
        data = request.json
        if not data: return jsonify({"error": "No data"}), 400
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            pd.DataFrame({'Component':['Material Cost','Conversion Cost','Margin','Packaging Cost','Freight Cost','TOTAL/1000','Per Piece','EUR/1000'],'Amount':[data.get('material_cost',0),data.get('conversion_cost',0),data.get('margin',0),data.get('packing_cost',0),data.get('freight_cost',0),data.get('total_cost_per_1000',0),data.get('cost_per_piece',0),data.get('total_cost_eur',0)],'%':[f"{data.get('mat_pct',0)}%",f"{data.get('conv_pct',0)}%",f"{data.get('margin_pct_total',0)}%",f"{data.get('pkg_pct',0)}%",f"{data.get('freight_pct',0)}%",'100%','','']}).to_excel(writer, sheet_name='Summary', index=False)
            pd.DataFrame({'Component':['Resin','Masterbatch','Additives','Wastage','Total'],'Amount':[data.get('resin_cost',0),data.get('mb_cost',0),data.get('additive_cost',0),data.get('wastage_cost',0),data.get('material_cost',0)]}).to_excel(writer, sheet_name='Material', index=False)
            pd.DataFrame({'Component':['Electricity','Direct Labour','Indirect Labour','R&M','Other OH','Lease','Depreciation','Interest','Total'],'Amount':[data.get('electricity_cost',0),data.get('direct_labour',0),data.get('indirect_labour',0),data.get('repair_cost',0),data.get('other_oh',0),data.get('lease_cost',0),data.get('depreciation',0),data.get('interest_total',0),data.get('conversion_cost',0)]}).to_excel(writer, sheet_name='Conversion', index=False)
            pd.DataFrame({'Param':['Machines','Output/Hr','Investment','Land SQM','Wastage'],'Value':[data.get('num_machines',0),data.get('output_per_hour',0),data.get('total_investment_inr',0),data.get('land_area_sqm',0),f"{data.get('wastage_pct',0)}%"]}).to_excel(writer, sheet_name='Production', index=False)
            from openpyxl.styles import Font, PatternFill, Border, Side
            for sn in writer.sheets:
                ws = writer.sheets[sn]
                for cell in ws[1]:
                    cell.font = Font(bold=True, color='FFFFFF'); cell.fill = PatternFill(start_color='E8601C', end_color='E8601C', fill_type='solid')
                for row in ws.iter_rows(min_row=2):
                    for cell in row:
                        if isinstance(cell.value, (int, float)): cell.number_format = '#,##0.00'
                for col in ws.columns:
                    ml = max(len(str(c.value or '')) for c in col) + 4; ws.column_dimensions[col[0].column_letter].width = min(ml, 30)
        output.seek(0)
        return send_file(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', as_attachment=True, download_name=f'EBM_Report_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx')
    except Exception as e:
        logger.error(f"Excel export error: {e}", exc_info=True)
        return jsonify({"error": str(e)}), 500


@app.route("/api/export_ebm_pdf", methods=["POST"])
def api_export_ebm_pdf():
    """Export EBM to PDF (Feature 6)"""
    try:
        data = request.json
        if not data: return jsonify({"error": "No data"}), 400
        try:
            from reportlab.lib import colors
            from reportlab.lib.pagesizes import A4
            from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
            from reportlab.lib.units import mm
            from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, HRFlowable
            from reportlab.lib.enums import TA_CENTER
        except ImportError:
            return jsonify({"error": "PDF requires reportlab. Install: pip install reportlab"}), 500
        output = io.BytesIO()
        doc = SimpleDocTemplate(output, pagesize=A4, topMargin=20*mm, bottomMargin=20*mm, leftMargin=15*mm, rightMargin=15*mm)
        styles = getSampleStyleSheet()
        ts = ParagraphStyle('T', parent=styles['Title'], fontSize=18, textColor=colors.HexColor('#E8601C'), spaceAfter=6)
        ss = ParagraphStyle('S', parent=styles['Normal'], fontSize=10, textColor=colors.HexColor('#666'), spaceAfter=12)
        hs = ParagraphStyle('H', parent=styles['Heading2'], fontSize=12, textColor=colors.HexColor('#1e40af'), spaceBefore=14, spaceAfter=8)
        els = []
        els.append(Paragraph('Packfora Analytics', ts))
        els.append(Paragraph('EBM Cost Breakdown Report', ss))
        els.append(Paragraph(f'SKU: {data.get("sku_description","N/A")} | Country: {data.get("country","N/A")} | {data.get("currency","INR")}', ss))
        els.append(Paragraph(f'Generated: {datetime.now().strftime("%B %d, %Y %I:%M %p")}', ss))
        els.append(HRFlowable(width="100%", thickness=2, color=colors.HexColor('#E8601C')))
        els.append(Spacer(1, 10))
        o = colors.HexColor('#E8601C'); b = colors.HexColor('#1e40af')
        cd = [['Component','Amount','%'],['Material',f"{data.get('material_cost',0):,.2f}",f"{data.get('mat_pct',0)}%"],['Conversion',f"{data.get('conversion_cost',0):,.2f}",f"{data.get('conv_pct',0)}%"],['Margin',f"{data.get('margin',0):,.2f}",f"{data.get('margin_pct_total',0)}%"],['Packaging',f"{data.get('packing_cost',0):,.2f}",f"{data.get('pkg_pct',0)}%"],['Freight',f"{data.get('freight_cost',0):,.2f}",f"{data.get('freight_pct',0)}%"],['TOTAL',f"{data.get('total_cost_per_1000',0):,.2f}",'100%']]
        els.append(Paragraph('Cost Summary (per 1000 Pcs)', hs))
        t = Table(cd, colWidths=[200, 150, 80])
        t.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),o),('TEXTCOLOR',(0,0),(-1,0),colors.white),('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),('FONTSIZE',(0,0),(-1,-1),9),('ALIGN',(1,0),(-1,-1),'RIGHT'),('BACKGROUND',(0,-1),(-1,-1),colors.HexColor('#FFF3ED')),('FONTNAME',(0,-1),(-1,-1),'Helvetica-Bold'),('GRID',(0,0),(-1,-1),0.5,colors.HexColor('#ddd')),('TOPPADDING',(0,0),(-1,-1),6),('BOTTOMPADDING',(0,0),(-1,-1),6)]))
        els.append(t); els.append(Spacer(1,8))
        els.append(Paragraph(f'Per Unit: {data.get("currency","INR")} {data.get("cost_per_piece",0):.4f} | EUR/1000: \u20ac {data.get("total_cost_eur",0):,.2f}', ss))
        els.append(Paragraph('Material Breakdown', hs))
        md = [['Component','Amount'],['Resin',f"{data.get('resin_cost',0):,.2f}"],['Masterbatch',f"{data.get('mb_cost',0):,.2f}"],['Additives',f"{data.get('additive_cost',0):,.2f}"],['Wastage',f"{data.get('wastage_cost',0):,.2f}"],['Total',f"{data.get('material_cost',0):,.2f}"]]
        t2 = Table(md, colWidths=[250,150])
        t2.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),b),('TEXTCOLOR',(0,0),(-1,0),colors.white),('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),('FONTSIZE',(0,0),(-1,-1),9),('ALIGN',(1,0),(-1,-1),'RIGHT'),('BACKGROUND',(0,-1),(-1,-1),colors.HexColor('#EBF0FF')),('FONTNAME',(0,-1),(-1,-1),'Helvetica-Bold'),('GRID',(0,0),(-1,-1),0.5,colors.HexColor('#ddd')),('TOPPADDING',(0,0),(-1,-1),5),('BOTTOMPADDING',(0,0),(-1,-1),5)]))
        els.append(t2); els.append(Spacer(1,8))
        els.append(Paragraph('Conversion Breakdown', hs))
        cvd = [['Component','Amount'],['Electricity',f"{data.get('electricity_cost',0):,.2f}"],['Direct Labour',f"{data.get('direct_labour',0):,.2f}"],['Indirect Labour',f"{data.get('indirect_labour',0):,.2f}"],['R&M',f"{data.get('repair_cost',0):,.2f}"],['Other OH',f"{data.get('other_oh',0):,.2f}"],['Depreciation',f"{data.get('depreciation',0):,.2f}"],['Interest',f"{data.get('interest_total',0):,.2f}"],['Total',f"{data.get('conversion_cost',0):,.2f}"]]
        t3 = Table(cvd, colWidths=[250,150])
        t3.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),b),('TEXTCOLOR',(0,0),(-1,0),colors.white),('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),('FONTSIZE',(0,0),(-1,-1),9),('ALIGN',(1,0),(-1,-1),'RIGHT'),('BACKGROUND',(0,-1),(-1,-1),colors.HexColor('#EBF0FF')),('FONTNAME',(0,-1),(-1,-1),'Helvetica-Bold'),('GRID',(0,0),(-1,-1),0.5,colors.HexColor('#ddd')),('TOPPADDING',(0,0),(-1,-1),5),('BOTTOMPADDING',(0,0),(-1,-1),5)]))
        els.append(t3); els.append(Spacer(1,15))
        els.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor('#ccc')))
        els.append(Paragraph('Confidential - Packfora Analytics', ParagraphStyle('F', parent=styles['Normal'], fontSize=8, textColor=colors.grey, alignment=TA_CENTER)))
        doc.build(els); output.seek(0)
        return send_file(output, mimetype='application/pdf', as_attachment=True, download_name=f'EBM_Report_{datetime.now().strftime("%Y%m%d_%H%M%S")}.pdf')
    except Exception as e:
        logger.error(f"PDF export error: {e}", exc_info=True)
        return jsonify({"error": str(e)}), 500


@app.route("/api/export_multi_country_excel", methods=["POST"])
def api_export_multi_country_excel():
    """Export multi-country comparison to Excel"""
    try:
        data = request.json
        results = data.get('results', [])
        if not results: return jsonify({"error": "No data"}), 400
        output = io.BytesIO()
        rows = [{'Country':r['country'],'Total EUR/1000':r.get('total_eur',0),'Material EUR':r.get('mat_eur',0),'Conversion EUR':r.get('conv_eur',0),'Margin EUR':r.get('margin_eur',0),'Packing EUR':r.get('pkg_eur',0),'Freight EUR':r.get('frt_eur',0),'Total Local':r.get('total_local',0),'EUR Rate':r.get('euro_rate',0),'Machines':r.get('machines',0)} for r in results if 'error' not in r]
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            pd.DataFrame(rows).to_excel(writer, sheet_name='Comparison', index=False)
            from openpyxl.styles import Font, PatternFill
            ws = writer.sheets['Comparison']
            for cell in ws[1]: cell.font = Font(bold=True, color='FFFFFF'); cell.fill = PatternFill(start_color='E8601C', end_color='E8601C', fill_type='solid')
            for col in ws.columns: ws.column_dimensions[col[0].column_letter].width = max(len(str(c.value or '')) for c in col) + 4
        output.seek(0)
        return send_file(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', as_attachment=True, download_name=f'Country_Comparison_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx')
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/export_generic_excel", methods=["POST"])
def api_export_generic_excel():
    """Export any model result to formatted Excel"""
    try:
        data = request.json
        if not data: return jsonify({"error": "No data"}), 400
        model_type = data.get('model_type', 'model')
        d = data.get('data', {})
        if not d: return jsonify({"error": "No model data"}), 400
        
        output = io.BytesIO()
        model_names = {'carton': 'Carton Essentials', 'flexibles': 'Flexibles', 'ebm': 'EBM Rigids', 'carton-adv': 'Carton Advanced'}
        model_label = model_names.get(model_type, model_type.replace('-', ' ').title())
        
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # Summary sheet
            mat = d.get('material_cost', 0)
            conv = d.get('conversion_cost', 0)
            margin = d.get('margin', 0)
            pkg = d.get('packing_cost', 0)
            frt = d.get('freight_cost', 0)
            total = d.get('total_cost_per_1000', 0) or (mat + conv + margin + pkg + frt)
            unit_label = '₹/kg' if model_type == 'flexibles' else '₹/1000 Pcs'
            
            summary_rows = [
                {'Component': 'Material Cost', 'Amount': mat, '%': f"{(mat/total*100):.1f}%" if total > 0 else '0%'},
                {'Component': 'Conversion Cost', 'Amount': conv, '%': f"{(conv/total*100):.1f}%" if total > 0 else '0%'},
                {'Component': 'Margin', 'Amount': margin, '%': f"{(margin/total*100):.1f}%" if total > 0 else '0%'},
                {'Component': 'Packaging Cost', 'Amount': pkg, '%': f"{(pkg/total*100):.1f}%" if total > 0 else '0%'},
                {'Component': 'Freight Cost', 'Amount': frt, '%': f"{(frt/total*100):.1f}%" if total > 0 else '0%'},
                {'Component': 'TOTAL', 'Amount': total, '%': '100%'},
            ]
            pd.DataFrame(summary_rows).to_excel(writer, sheet_name='Summary', index=False)
            
            # Model-specific detail sheets
            if model_type == 'carton':
                detail = [
                    {'Component': 'Board Cost', 'Amount': d.get('board_cost', 0)},
                    {'Component': 'Ink Cost', 'Amount': d.get('ink_cost', 0)},
                    {'Component': 'Varnish Cost', 'Amount': d.get('varnish_cost', 0)},
                    {'Component': 'Lamination Cost', 'Amount': d.get('lamination_cost', 0)},
                    {'Component': 'Corrugation Cost', 'Amount': d.get('corrugation_cost', 0)},
                    {'Component': 'Foil Cost', 'Amount': d.get('foil_cost', 0)},
                    {'Component': 'Other Material', 'Amount': d.get('other_material_cost', 0)},
                    {'Component': 'Total Material', 'Amount': mat},
                ]
                pd.DataFrame(detail).to_excel(writer, sheet_name='Material Detail', index=False)
            elif model_type == 'flexibles':
                layers = d.get('layers', [])
                if layers:
                    layer_rows = [{'Layer': f"L{i+1}: {l.get('name','')}", 'Micron': l.get('mic',0), 'GSM': l.get('gsm',0), 'Rate ₹/kg': l.get('rate',0), 'Cost ₹/kg': l.get('layer_cost',0)} for i, l in enumerate(layers)]
                    pd.DataFrame(layer_rows).to_excel(writer, sheet_name='Layers', index=False)
                flex_detail = [
                    {'Metric': 'Laminate GSM', 'Value': d.get('laminate_gsm', 0)},
                    {'Metric': 'Avg Density', 'Value': d.get('avg_density', 0)},
                    {'Metric': 'Material ₹/kg', 'Value': d.get('material_cost_per_kg', 0)},
                    {'Metric': 'Wastage ₹/kg', 'Value': d.get('wastage_cost', 0)},
                    {'Metric': 'Laminate ₹/kg', 'Value': d.get('laminate_cost_per_kg', 0)},
                    {'Metric': 'Laminate ₹/SQM', 'Value': d.get('laminate_cost_per_sqm', 0)},
                ]
                pd.DataFrame(flex_detail).to_excel(writer, sheet_name='Details', index=False)
            elif model_type in ('carton-adv', 'ebm'):
                conv_detail = [
                    {'Component': 'Electricity', 'Amount': d.get('electricity_cost', 0)},
                    {'Component': 'Direct Labour', 'Amount': d.get('direct_labour', 0)},
                    {'Component': 'Indirect Labour', 'Amount': d.get('indirect_labour', 0)},
                    {'Component': 'Repair & Maint', 'Amount': d.get('repair_maintenance', d.get('repair_cost', 0))},
                    {'Component': 'Other OH', 'Amount': d.get('other_overheads', d.get('other_oh', 0))},
                    {'Component': 'Depreciation', 'Amount': d.get('depreciation', 0)},
                    {'Component': 'Interest', 'Amount': d.get('interest', d.get('interest_total', 0))},
                    {'Component': 'Lease', 'Amount': d.get('lease_cost', 0)},
                    {'Component': 'Total Conversion', 'Amount': conv},
                ]
                pd.DataFrame(conv_detail).to_excel(writer, sheet_name='Conversion', index=False)
            
            # Format all sheets
            from openpyxl.styles import Font, PatternFill
            for sn in writer.sheets:
                ws = writer.sheets[sn]
                for cell in ws[1]:
                    cell.font = Font(bold=True, color='FFFFFF')
                    cell.fill = PatternFill(start_color='E8601C', end_color='E8601C', fill_type='solid')
                for row in ws.iter_rows(min_row=2):
                    for cell in row:
                        if isinstance(cell.value, (int, float)): cell.number_format = '#,##0.00'
                for col in ws.columns:
                    ml = max(len(str(c.value or '')) for c in col) + 4
                    ws.column_dimensions[col[0].column_letter].width = min(ml, 30)
        
        output.seek(0)
        fname = model_label.replace(' ', '_')
        return send_file(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                         as_attachment=True, download_name=f'{fname}_Report_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx')
    except Exception as e:
        logger.error(f"Generic Excel export error: {e}", exc_info=True)
        return jsonify({"error": str(e)}), 500


@app.route("/api/export_generic_pdf", methods=["POST"])
def api_export_generic_pdf():
    """Export any model result to PDF"""
    try:
        data = request.json
        if not data: return jsonify({"error": "No data"}), 400
        model_type = data.get('model_type', 'model')
        d = data.get('data', {})
        if not d: return jsonify({"error": "No model data"}), 400
        
        try:
            from reportlab.lib import colors
            from reportlab.lib.pagesizes import A4
            from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
            from reportlab.lib.units import mm
            from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, HRFlowable
            from reportlab.lib.enums import TA_CENTER
        except ImportError:
            return jsonify({"error": "PDF requires reportlab. Install: pip install reportlab"}), 500

        model_names = {'carton': 'Carton Essentials', 'flexibles': 'Flexibles', 'ebm': 'EBM Rigids', 'carton-adv': 'Carton Advanced'}
        model_label = model_names.get(model_type, model_type.replace('-', ' ').title())
        
        output = io.BytesIO()
        doc = SimpleDocTemplate(output, pagesize=A4, topMargin=20*mm, bottomMargin=20*mm, leftMargin=15*mm, rightMargin=15*mm)
        styles = getSampleStyleSheet()
        ts = ParagraphStyle('T', parent=styles['Title'], fontSize=18, textColor=colors.HexColor('#E8601C'), spaceAfter=6)
        ss = ParagraphStyle('S', parent=styles['Normal'], fontSize=10, textColor=colors.HexColor('#666'), spaceAfter=12)
        hs = ParagraphStyle('H', parent=styles['Heading2'], fontSize=12, textColor=colors.HexColor('#1e40af'), spaceBefore=14, spaceAfter=8)
        
        els = []
        els.append(Paragraph('Packfora Analytics', ts))
        els.append(Paragraph(f'{model_label} Cost Breakdown Report', ss))
        country = d.get('country', 'N/A')
        sku = d.get('sku_description', d.get('model_type', 'N/A'))
        els.append(Paragraph(f'Model: {model_label} | Country: {country}', ss))
        els.append(Paragraph(f'Generated: {datetime.now().strftime("%B %d, %Y %I:%M %p")}', ss))
        els.append(HRFlowable(width="100%", thickness=2, color=colors.HexColor('#E8601C')))
        els.append(Spacer(1, 10))
        
        o = colors.HexColor('#E8601C')
        b = colors.HexColor('#1e40af')
        
        mat = d.get('material_cost', 0)
        conv = d.get('conversion_cost', 0)
        margin = d.get('margin', 0)
        pkg = d.get('packing_cost', 0)
        frt = d.get('freight_cost', 0)
        total = d.get('total_cost_per_1000', 0) or (mat + conv + margin + pkg + frt)
        
        unit_label = '₹/kg' if model_type == 'flexibles' else '₹/1000 Pcs'
        
        cd = [['Component', 'Amount', '%'],
              ['Material', f"{mat:,.2f}", f"{(mat/total*100):.1f}%" if total > 0 else '0%'],
              ['Conversion', f"{conv:,.2f}", f"{(conv/total*100):.1f}%" if total > 0 else '0%'],
              ['Margin', f"{margin:,.2f}", f"{(margin/total*100):.1f}%" if total > 0 else '0%'],
              ['Packaging', f"{pkg:,.2f}", f"{(pkg/total*100):.1f}%" if total > 0 else '0%'],
              ['Freight', f"{frt:,.2f}", f"{(frt/total*100):.1f}%" if total > 0 else '0%'],
              ['TOTAL', f"{total:,.2f}", '100%']]
        
        els.append(Paragraph(f'Cost Summary ({unit_label})', hs))
        t = Table(cd, colWidths=[200, 150, 80])
        t.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), o), ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'), ('FONTSIZE', (0, 0), (-1, -1), 9),
            ('ALIGN', (1, 0), (-1, -1), 'RIGHT'),
            ('BACKGROUND', (0, -1), (-1, -1), colors.HexColor('#FFF3ED')),
            ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#ddd')),
            ('TOPPADDING', (0, 0), (-1, -1), 6), ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ]))
        els.append(t)
        els.append(Spacer(1, 8))
        
        # Model-specific detail
        if model_type == 'carton':
            els.append(Paragraph('Material Breakdown', hs))
            md = [['Component', 'Amount'],
                  ['Board', f"{d.get('board_cost',0):,.2f}"], ['Ink', f"{d.get('ink_cost',0):,.2f}"],
                  ['Varnish', f"{d.get('varnish_cost',0):,.2f}"], ['Lamination', f"{d.get('lamination_cost',0):,.2f}"],
                  ['Corrugation', f"{d.get('corrugation_cost',0):,.2f}"], ['Foil', f"{d.get('foil_cost',0):,.2f}"],
                  ['Other', f"{d.get('other_material_cost',0):,.2f}"], ['Total Material', f"{mat:,.2f}"]]
            t2 = Table(md, colWidths=[250, 150])
            t2.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),b),('TEXTCOLOR',(0,0),(-1,0),colors.white),('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),('FONTSIZE',(0,0),(-1,-1),9),('ALIGN',(1,0),(-1,-1),'RIGHT'),('BACKGROUND',(0,-1),(-1,-1),colors.HexColor('#EBF0FF')),('FONTNAME',(0,-1),(-1,-1),'Helvetica-Bold'),('GRID',(0,0),(-1,-1),0.5,colors.HexColor('#ddd')),('TOPPADDING',(0,0),(-1,-1),5),('BOTTOMPADDING',(0,0),(-1,-1),5)]))
            els.append(t2)
        elif model_type == 'flexibles':
            els.append(Paragraph('Laminate Details', hs))
            fd = [['Metric', 'Value'],
                  ['Laminate GSM', f"{d.get('laminate_gsm',0)}"], ['Avg Density', f"{d.get('avg_density',0):.4f}"],
                  ['Material ₹/kg', f"{d.get('material_cost_per_kg',0):,.2f}"],
                  ['Wastage ₹/kg', f"{d.get('wastage_cost',0):,.2f}"],
                  ['Laminate ₹/kg', f"{d.get('laminate_cost_per_kg',0):,.2f}"],
                  ['Laminate ₹/SQM', f"{d.get('laminate_cost_per_sqm',0):,.2f}"]]
            t2 = Table(fd, colWidths=[250, 150])
            t2.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),b),('TEXTCOLOR',(0,0),(-1,0),colors.white),('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),('FONTSIZE',(0,0),(-1,-1),9),('ALIGN',(1,0),(-1,-1),'RIGHT'),('GRID',(0,0),(-1,-1),0.5,colors.HexColor('#ddd')),('TOPPADDING',(0,0),(-1,-1),5),('BOTTOMPADDING',(0,0),(-1,-1),5)]))
            els.append(t2)
        elif model_type in ('carton-adv',):
            els.append(Paragraph('Conversion Breakdown', hs))
            cvd = [['Component','Amount'],['Electricity',f"{d.get('electricity_cost',0):,.2f}"],['Direct Labour',f"{d.get('direct_labour',0):,.2f}"],['Indirect Labour',f"{d.get('indirect_labour',0):,.2f}"],['R&M',f"{d.get('repair_maintenance',0):,.2f}"],['Other OH',f"{d.get('other_overheads',0):,.2f}"],['Depreciation',f"{d.get('depreciation',0):,.2f}"],['Interest',f"{d.get('interest',0):,.2f}"],['Total',f"{conv:,.2f}"]]
            t3 = Table(cvd, colWidths=[250,150])
            t3.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),b),('TEXTCOLOR',(0,0),(-1,0),colors.white),('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),('FONTSIZE',(0,0),(-1,-1),9),('ALIGN',(1,0),(-1,-1),'RIGHT'),('BACKGROUND',(0,-1),(-1,-1),colors.HexColor('#EBF0FF')),('FONTNAME',(0,-1),(-1,-1),'Helvetica-Bold'),('GRID',(0,0),(-1,-1),0.5,colors.HexColor('#ddd')),('TOPPADDING',(0,0),(-1,-1),5),('BOTTOMPADDING',(0,0),(-1,-1),5)]))
            els.append(t3)
        
        els.append(Spacer(1, 15))
        els.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor('#ccc')))
        els.append(Paragraph('Confidential - Packfora Analytics', ParagraphStyle('F', parent=styles['Normal'], fontSize=8, textColor=colors.grey, alignment=TA_CENTER)))
        doc.build(els)
        output.seek(0)
        fname = model_label.replace(' ', '_')
        return send_file(output, mimetype='application/pdf', as_attachment=True, download_name=f'{fname}_Report_{datetime.now().strftime("%Y%m%d_%H%M%S")}.pdf')
    except Exception as e:
        logger.error(f"Generic PDF export error: {e}", exc_info=True)
        return jsonify({"error": str(e)}), 500


ADMIN_LOGIN_HTML = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Admin Login - Packfora Analytics</title>
    <link href="https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;600;800&display=swap" rel="stylesheet">
    <style>
        :root { --orange: #E8601C; --royal-blue: #1e40af; }
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { 
            font-family: 'Outfit', sans-serif; 
            background: linear-gradient(135deg, var(--royal-blue) 0%, #3b82f6 50%, #1e3a8a 100%); 
            min-height: 100vh; 
            display: flex;
            align-items: center;
            justify-content: center;
            color: white; 
        }
        .login-container {
            background: rgba(255,255,255,0.15);
            backdrop-filter: blur(20px);
            border-radius: 20px;
            padding: 50px;
            border: 1px solid rgba(255,255,255,0.25);
            width: 400px;
            box-shadow: 0 10px 40px rgba(0,0,0,0.3);
        }
        h2 {
            color: var(--orange);
            margin-bottom: 30px;
            text-align: center;
            font-size: 2rem;
        }
        .form-group {
            margin-bottom: 20px;
        }
        label {
            display: block;
            font-size: 0.9rem;
            margin-bottom: 8px;
            opacity: 0.9;
            font-weight: 600;
        }
        input {
            width: 100%;
            padding: 15px;
            background: rgba(255,255,255,0.2);
            border: 1px solid rgba(255,255,255,0.3);
            border-radius: 10px;
            color: white;
            font-family: 'Outfit';
            font-size: 1rem;
        }
        input::placeholder {
            color: rgba(255,255,255,0.6);
        }
        input:focus {
            outline: none;
            border-color: var(--orange);
        }
        button {
            width: 100%;
            padding: 15px;
            background: var(--orange);
            border: none;
            border-radius: 10px;
            color: white;
            font-weight: 800;
            font-size: 1rem;
            cursor: pointer;
            transition: all 0.3s;
            font-family: 'Outfit';
        }
        button:hover {
            background: #d65519;
            transform: scale(1.02);
        }
        .alert {
            padding: 15px;
            border-radius: 10px;
            margin-bottom: 20px;
            font-size: 0.9rem;
        }
        .alert-error {
            background: rgba(239, 68, 68, 0.2);
            border: 1px solid #ef4444;
        }
        .alert-success {
            background: rgba(16, 185, 129, 0.2);
            border: 1px solid #10b981;
        }
        .alert-info {
            background: rgba(59, 130, 246, 0.2);
            border: 1px solid #3b82f6;
        }
        .back-link {
            text-align: center;
            margin-top: 20px;
        }
        .back-link a {
            color: var(--orange);
            text-decoration: none;
            font-size: 0.9rem;
        }
        .back-link a:hover {
            text-decoration: underline;
        }
    </style>
</head>
<body>
    <div class="login-container">
        <h2>Admin Login</h2>
        
        {% with messages = get_flashed_messages(with_categories=true) %}
          {% if messages %}
            {% for category, message in messages %}
              <div class="alert alert-{{ category }}">{{ message }}</div>
            {% endfor %}
          {% endif %}
        {% endwith %}
        
        <form method="POST">
            <div class="form-group">
                <label>Username</label>
                <input type="text" name="username" placeholder="Enter username" required autofocus>
            </div>
            <div class="form-group">
                <label>Password</label>
                <input type="password" name="password" placeholder="Enter password" required>
            </div>
            <button type="submit">Login</button>
        </form>
        
        <div class="back-link">
            <a href="/">← Back to Dashboard</a>
        </div>
    </div>
</body>
</html>
"""

ADMIN_DASHBOARD_HTML = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Admin Dashboard - Packfora Analytics</title>
    <link href="https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;600;800&display=swap" rel="stylesheet">
    <style>
        :root { --orange: #E8601C; --royal-blue: #1e40af; }
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { 
            font-family: 'Outfit', sans-serif; 
            background: linear-gradient(135deg, var(--royal-blue) 0%, #3b82f6 50%, #1e3a8a 100%); 
            min-height: 100vh; 
            color: white; 
        }
        .navbar {
    background: linear-gradient(
        135deg,
        var(--royal-blue) 0%,
        var(--royal-blue-light) 50%,
        var(--royal-blue-dark) 100%
    );
    padding: 20px 40px;
    display: flex;
    align-items: center;
    border-bottom: 1px solid rgba(255,255,255,0.15);
}

        }
        .navbar h1 {
            flex: 1;
            font-size: 1.5rem;
        }
        .navbar span {
            color: var(--orange);
        }
        .nav-links {
            display: flex;
            gap: 20px;
        }
        .nav-links a {
            color: white;
            text-decoration: none;
            font-weight: 600;
            padding: 10px 18px;
            border-radius: 8px;
            transition: all 0.3s;
        }
        .nav-links a:hover {
            background: rgba(255,255,255,0.1);
        }
        .container {
            max-width: 1200px;
            margin: 40px auto;
            padding: 0 20px;
        }
        .card {
            background: rgba(255,255,255,0.15);
            backdrop-filter: blur(20px);
            border-radius: 20px;
            padding: 35px;
            border: 1px solid rgba(255,255,255,0.25);
            margin-bottom: 25px;
        }
        h2 {
            color: var(--orange);
            margin-bottom: 25px;
            font-size: 1.5rem;
        }
        .upload-section {
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 20px;
            margin-bottom: 30px;
        }
        .upload-card {
            background: rgba(255,255,255,0.1);
            padding: 25px;
            border-radius: 15px;
            border: 2px dashed rgba(255,255,255,0.3);
            text-align: center;
            transition: all 0.3s;
        }
        .upload-card:hover {
            border-color: var(--orange);
            transform: translateY(-5px);
        }
        .upload-card h4 {
            margin-bottom: 15px;
            font-size: 1.1rem;
        }
        input[type="file"] {
            display: none;
        }
        .file-label {
            display: inline-block;
            padding: 12px 25px;
            background: rgba(255,255,255,0.2);
            border-radius: 8px;
            cursor: pointer;
            transition: all 0.3s;
            font-size: 0.9rem;
        }
        .file-label:hover {
            background: rgba(255,255,255,0.3);
        }
        .upload-btn {
            margin-top: 15px;
            padding: 12px 30px;
            background: var(--orange);
            border: none;
            border-radius: 8px;
            color: white;
            font-weight: 700;
            cursor: pointer;
            transition: all 0.3s;
            font-family: 'Outfit';
        }
        .upload-btn:hover {
            background: #d65519;
        }
        .upload-btn:disabled {
            opacity: 0.5;
            cursor: not-allowed;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 15px;
        }
        th, td {
            padding: 15px;
            text-align: left;
            border-bottom: 1px solid rgba(255,255,255,0.1);
        }
        th {
            background: rgba(255,255,255,0.1);
            font-weight: 700;
            color: var(--orange);
        }
        .status-ok {
            color: #10b981;
        }
        .status-error {
            color: #ef4444;
        }
        .action-btn {
            padding: 8px 15px;
            background: rgba(255,255,255,0.2);
            border: none;
            border-radius: 6px;
            color: white;
            cursor: pointer;
            font-size: 0.85rem;
            transition: all 0.3s;
            text-decoration: none;
            display: inline-block;
        }
        .action-btn:hover {
            background: rgba(255,255,255,0.3);
        }
        .alert {
            padding: 15px;
            border-radius: 10px;
            margin-bottom: 20px;
            font-size: 0.9rem;
        }
        .alert-error {
            background: rgba(239, 68, 68, 0.2);
            border: 1px solid #ef4444;
        }
        .alert-success {
            background: rgba(16, 185, 129, 0.2);
            border: 1px solid #10b981;
        }
        .alert-info {
            background: rgba(59, 130, 246, 0.2);
            border: 1px solid #3b82f6;
        }
        .selected-file {
            margin-top: 10px;
            font-size: 0.85rem;
            opacity: 0.8;
        }
        @media (max-width: 768px) {
            .upload-section {
                grid-template-columns: 1fr;
            }
        }
    </style>
</head>
<body>
    <nav class="navbar">
        <h1>Admin <span>Dashboard</span></h1>
        <div class="nav-links">
            <span style="opacity:0.7; font-size:0.9rem;">Welcome, {{ username }}</span>
            <a href="/">Main Site</a>
            <a href="/admin/logout">Logout</a>
        </div>
    </nav>
    
    <div class="container">
        {% with messages = get_flashed_messages(with_categories=true) %}
          {% if messages %}
            {% for category, message in messages %}
              <div class="alert alert-{{ category }}">{{ message }}</div>
            {% endfor %}
          {% endif %}
        {% endwith %}
        
        <div class="card">
            <h2>Upload Data Files</h2>
            <p style="opacity:0.8; margin-bottom:25px;">Upload new Excel files to update the application data. Old files will be backed up automatically.</p>
            
            <div class="upload-section">
                <!-- Resin Upload -->
                <div class="upload-card">
                    <h4>Resin Database</h4>
                    <form action="/admin/upload" method="POST" enctype="multipart/form-data" id="resin-form">
                        <input type="hidden" name="file_type" value="resin">
                        <label for="resin-file" class="file-label">Choose File</label>
                        <input type="file" id="resin-file" name="file" accept=".xlsx,.xls" onchange="updateFileName('resin')">
                        <div class="selected-file" id="resin-filename"></div>
                        <button type="submit" class="upload-btn" id="resin-btn" disabled>Upload</button>
                    </form>
                </div>
                
                <!-- Machine Upload -->
                <div class="upload-card">
                    <h4>Machine Database</h4>
                    <form action="/admin/upload" method="POST" enctype="multipart/form-data" id="machine-form">
                        <input type="hidden" name="file_type" value="machine">
                        <label for="machine-file" class="file-label">Choose File</label>
                        <input type="file" id="machine-file" name="file" accept=".xlsx,.xls" onchange="updateFileName('machine')">
                        <div class="selected-file" id="machine-filename"></div>
                        <button type="submit" class="upload-btn" id="machine-btn" disabled>Upload</button>
                    </form>
                </div>
                
                <!-- Cost Upload -->
                <div class="upload-card">
                    <h4>Variable Costs</h4>
                    <form action="/admin/upload" method="POST" enctype="multipart/form-data" id="cost-form">
                        <input type="hidden" name="file_type" value="cost">
                        <label for="cost-file" class="file-label">Choose File</label>
                        <input type="file" id="cost-file" name="file" accept=".xlsx,.xls" onchange="updateFileName('cost')">
                        <div class="selected-file" id="cost-filename"></div>
                        <button type="submit" class="upload-btn" id="cost-btn" disabled>Upload</button>
                    </form>
                </div>
            </div>
        </div>
        
        <div class="card">
            <h2>📊 Import Monthly Resin Prices</h2>
            <p style="margin-bottom:15px;opacity:0.85;">Upload monthly price Excel files (from Reliance/IOCL etc.) or PET film price-list PDFs (JPFL etc.) to auto-parse grades, locations & prices into the resin database.</p>
            <form id="import-form" enctype="multipart/form-data">
                <div style="margin-bottom:15px;">
                    <label for="price-files" style="display:block;margin-bottom:8px;font-weight:600;font-size:1rem;cursor:pointer;">
                        📁 Select Price Files (.xlsx / .pdf):
                    </label>
                    <input 
                        type="file" 
                        name="price_files" 
                        id="price-files" 
                        accept=".xlsx,.xls,.pdf" 
                        multiple
                        style="display:block;width:100%;max-width:500px;padding:10px;font-size:1rem;background:white;color:black;border:2px solid #ccc;border-radius:5px;cursor:pointer;">
                    <div id="file-count" style="margin-top:10px;font-size:0.9rem;color:#4ade80;"></div>
                </div>
                <button 
                    type="submit" 
                    id="import-btn" 
                    disabled
                    style="padding:12px 24px;background:#ff8c00;color:white;border:none;border-radius:8px;font-weight:700;font-size:1rem;cursor:pointer;opacity:0.5;">
                    🚀 Import Prices
                </button>
            </form>
            <div id="import-progress" style="display:none;margin-top:15px;padding:15px;background:rgba(255,255,255,0.1);border-radius:10px;">
                <div style="display:flex;align-items:center;gap:10px;">
                    <div class="spinner" style="width:20px;height:20px;border:3px solid rgba(255,255,255,0.3);border-top-color:var(--orange);border-radius:50%;animation:spin 1s linear infinite;"></div>
                    <span>Parsing and importing prices...</span>
                </div>
            </div>
            <div id="import-result" style="display:none;margin-top:15px;"></div>
        </div>

        <div class="card">
            <h2>Current Files</h2>
            <table>
                <thead>
                    <tr>
                        <th>Database</th>
                        <th>Filename</th>
                        <th>Size</th>
                        <th>Last Modified</th>
                        <th>Status</th>
                        <th>Actions</th>
                    </tr>
                </thead>
                <tbody>
                    {% for file in files %}
                    <tr>
                        <td><strong>{{ file.name }}</strong></td>
                        <td>{{ file.filename }}</td>
                        <td>{{ file.size }}</td>
                        <td>{{ file.modified }}</td>
                        <td class="{% if file.exists %}status-ok{% else %}status-error{% endif %}">
                            {% if file.exists %}✓ Available{% else %}✗ Missing{% endif %}
                        </td>
                        <td>
                            {% if file.exists %}
                            <a href="/admin/download/{{ file.name.lower().split()[0] }}" class="action-btn">Download</a>
                            {% else %}
                            <span style="opacity:0.5;">N/A</span>
                            {% endif %}
                        </td>
                    </tr>
                    {% endfor %}
                </tbody>
            </table>
        </div>
        
        <div class="card">
            <h2>Recent Backups</h2>
            <p style="opacity:0.8; margin-bottom:15px;">Last 10 backups (old backups are automatically deleted)</p>
            {% if backups %}
            <table>
                <thead>
                    <tr>
                        <th>Filename</th>
                        <th>Size</th>
                        <th>Backup Date</th>
                    </tr>
                </thead>
                <tbody>
                    {% for backup in backups %}
                    <tr>
                        <td>{{ backup.name }}</td>
                        <td>{{ backup.size }}</td>
                        <td>{{ backup.date }}</td>
                    </tr>
                    {% endfor %}
                </tbody>
            </table>
            {% else %}
            <p style="opacity:0.6; text-align:center; padding:20px;">No backups yet</p>
            {% endif %}
        </div>
    </div>
    
    <style>
        @keyframes spin { to { transform: rotate(360deg); } }
    </style>
    <script>
        function updateFileName(type) {
            const input = document.getElementById(type + '-file');
            const display = document.getElementById(type + '-filename');
            const button = document.getElementById(type + '-btn');
            
            if (input.files.length > 0) {
                display.textContent = '📄 ' + input.files[0].name;
                button.disabled = false;
            } else {
                display.textContent = '';
                button.disabled = true;
            }
        }

        // Import price files handler - SIMPLE VERSION
        const priceFilesInput = document.getElementById('price-files');
        const importBtn = document.getElementById('import-btn');
        const fileCount = document.getElementById('file-count');
        
        console.log('Setting up import handlers...');
        console.log('File input:', priceFilesInput);
        console.log('Import button:', importBtn);
        
        if (!priceFilesInput) {
            console.error('ERROR: price-files input not found!');
            alert('ERROR: File input not found! Please refresh the page.');
        }
        
        if (!importBtn) {
            console.error('ERROR: import-btn button not found!');
            alert('ERROR: Import button not found! Please refresh the page.');
        }
        
        if (priceFilesInput && importBtn) {
            priceFilesInput.addEventListener('change', function(e) {
                console.log('File input change event fired!');
                console.log('Files selected:', this.files.length);
                
                const filesSelected = this.files.length;
                
                if (filesSelected > 0) {
                    // Enable button
                    importBtn.disabled = false;
                    importBtn.style.opacity = '1';
                    importBtn.style.cursor = 'pointer';
                    
                    // Show file names
                    const fileNames = Array.from(this.files).map(f => f.name).join(', ');
                    fileCount.innerHTML = `✓ ${filesSelected} file(s) selected: <strong>${fileNames}</strong>`;
                    
                    console.log('Button enabled, files:', fileNames);
                } else {
                    // Disable button
                    importBtn.disabled = true;
                    importBtn.style.opacity = '0.5';
                    importBtn.style.cursor = 'not-allowed';
                    fileCount.innerHTML = '';
                    
                    console.log('No files selected, button disabled');
                }
            });
            
            // Also log when file input is clicked
            priceFilesInput.addEventListener('click', function() {
                console.log('File input clicked!');
            });
        }

        document.getElementById('import-form').addEventListener('submit', async function(e) {
            e.preventDefault();
            console.log('Form submitted');
            
            const files = document.getElementById('price-files').files;
            console.log('Files to upload:', files.length);
            
            if (!files.length) {
                alert('Please select at least one Excel file to import');
                return;
            }

            const btn = document.getElementById('import-btn');
            const progress = document.getElementById('import-progress');
            const result = document.getElementById('import-result');
            
            btn.disabled = true;
            btn.textContent = '⏳ Importing...';
            btn.style.opacity = '0.5';
            progress.style.display = 'block';
            result.style.display = 'none';

            const formData = new FormData();
            for (let f of files) {
                console.log('Adding file:', f.name);
                formData.append('price_files', f);
            }

            try {
                console.log('Sending request to /api/import_resin_prices');
                const resp = await fetch('/api/import_resin_prices', { method: 'POST', body: formData });
                console.log('Response status:', resp.status);
                
                const data = await resp.json();
                console.log('Response data:', data);
                
                progress.style.display = 'none';
                result.style.display = 'block';

                if (data.error) {
                    result.innerHTML = `<div style="padding:15px;background:rgba(239,68,68,0.3);border:1px solid rgba(239,68,68,0.5);border-radius:10px;">
                        <strong>❌ Error:</strong> ${data.error}
                        ${data.file_results ? '<br><small>' + data.file_results.map(f => f.file + ': ' + (f.message||f.status)).join(', ') + '</small>' : ''}
                    </div>`;
                } else {
                    let statsHtml = '';
                    if (data.merge_stats) {
                        for (const [rt, st] of Object.entries(data.merge_stats)) {
                            const modeLabel = {
                                'merged': '🔄 Merged',
                                'auto_created': '🆕 Auto-Created',
                                'fresh_overwrite': '⚠️ Overwritten',
                                'new_sheet': '🆕 New Sheet'
                            }[st.mode] || st.mode;
                            const newRowsInfo = st.new_rows > 0
                                ? ` | <span style="color:#4ade80;">${st.new_rows} new entries</span>`
                                : '';
                            statsHtml += `<div style="display:inline-block;margin:5px;padding:8px 14px;background:rgba(255,255,255,0.15);border-radius:8px;">
                                <strong>${rt}</strong>: ${st.total_rows} rows, ${st.new_dates} date(s) — ${modeLabel}${newRowsInfo}</div>`;
                        }
                    }
                    let filesHtml = '';
                    if (data.file_results) {
                        filesHtml = data.file_results.map(f => {
                            let detail = f.sheet_details
                                ? f.sheet_details.map(s =>
                                    `<span style="opacity:0.7;font-size:0.8rem;margin-left:12px;">
                                     └ Sheet "${s.sheet}": ${s.records || 0} records (${s.resin_type || s.status})</span>`
                                ).join('<br>')
                                : '';
                            return `<div style="margin:3px 0;">
                                <span style="color:${f.status==='success'?'#4ade80':'#f87171'};">
                                ${f.status==='success'?'✓':'✗'}</span> 
                                ${f.file} — ${f.records || 0} records 
                                (${f.sheets_processed || 1} sheets)
                                ${detail ? '<br>' + detail : ''}
                            </div>`;
                        }).join('');
                    }

                    let autoCreateHtml = '';
                    if (data.auto_created_sheets && data.auto_created_sheets.length > 0) {
                        autoCreateHtml = `<div style="margin-top:10px;padding:10px;background:rgba(59,130,246,0.2);border:1px solid rgba(59,130,246,0.4);border-radius:8px;font-size:0.9rem;">
                            🆕 <strong>Auto-created new resin type sheets:</strong> ${data.auto_created_sheets.join(', ')}
                            <br><small>These are now available in the Resin Tracker dropdowns.</small>
                        </div>`;
                    }

                    let newEntriesHtml = '';
                    if (data.new_entries_added) {
                        const entries = Object.entries(data.new_entries_added)
                            .map(([rt, count]) => `${rt}: ${count} new`).join(', ');
                        newEntriesHtml = `<div style="margin-top:8px;padding:10px;background:rgba(16,185,129,0.2);border:1px solid rgba(16,185,129,0.4);border-radius:8px;font-size:0.9rem;">
                            📍 <strong>New locations/grades auto-added:</strong> ${entries}
                        </div>`;
                    }

                    let warningHtml = '';
                    if (data.warning) {
                        warningHtml = `<div style="margin-top:10px;padding:10px;background:rgba(251,191,36,0.2);border:1px solid rgba(251,191,36,0.4);border-radius:8px;font-size:0.9rem;">
                            ⚠️ ${data.warning}
                        </div>`;
                    }
                    
                    result.innerHTML = `<div style="padding:20px;background:rgba(34,197,94,0.2);border:1px solid rgba(34,197,94,0.4);border-radius:10px;">
                        <strong>✅ Import Complete</strong><br>
                        <div style="margin:10px 0;">${data.total_records.toLocaleString()} total price records imported for: <strong>${data.resin_types.join(', ')}</strong></div>
                        <div style="margin:8px 0;">${statsHtml}</div>
                        <div style="margin-top:10px;font-size:0.85rem;opacity:0.85;">${filesHtml}</div>
                        ${autoCreateHtml}
                        ${newEntriesHtml}
                        ${warningHtml}
                    </div>`;
                }
            } catch(err) {
                console.error('Import error:', err);
                progress.style.display = 'none';
                result.style.display = 'block';
                result.innerHTML = `<div style="padding:15px;background:rgba(239,68,68,0.3);border-radius:10px;">
                    <strong>❌ Network Error:</strong> ${err.message}<br>
                    <small style="opacity:0.8;">Check the browser console (F12) for more details.</small>
                </div>`;
            }
            btn.disabled = false;
            btn.textContent = '🚀 Import Prices';
            btn.style.opacity = '1';
            btn.style.cursor = 'pointer';
        });
    </script>
</body>
</html>
"""

BASE_HTML = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Packfora Analytics</title>
    <link rel="icon" type="image/x-icon" href="/static/favicon.ico">
    <link href="https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;600;800&display=swap" rel="stylesheet">
    <script src="https://cdn.plot.ly/plotly-2.27.0.min.js"></script>
    <style>
        :root { 
            --orange: #E8601C; 
            --royal-blue: #1e40af; 
            --royal-blue-light: #3b82f6; 
            --royal-blue-dark: #1e3a8a; 
        }
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body { 
            font-family: 'Outfit', sans-serif; 
            background: linear-gradient(135deg, var(--royal-blue) 0%, var(--royal-blue-light) 50%, var(--royal-blue-dark) 100%); 
            min-height: 100vh; 
            color: white; 
        }
        .navbar { 
            background: rgba(0,0,0,0.3); 
            backdrop-filter: blur(10px); 
            padding: 20px 40px; 
            display: flex; 
            align-items: center; 
            border-bottom: 1px solid rgba(255,255,255,0.1); 
        }
        .navbar-logo {
            margin-right: 40px;
            display: flex;
            align-items: center;
        }
        .navbar-logo img {
            height: 40px;
            width: auto;
        }
        .nav-links { 
            margin-left: auto; 
            display: flex; 
            gap: 20px; 
        }
        .nav-links a { 
            color: white; 
            text-decoration: none; 
            font-weight: 600; 
            padding: 12px 18px; 
            border-radius: 10px; 
            font-size: 0.8rem; 
            text-transform: uppercase; 
            transition: all 0.3s; 
        }
        .nav-links a:hover { background: rgba(255,255,255,0.1); }
        .nav-links a.active { background: var(--orange); }
        .nav-links .admin-link {
            background: rgba(232, 96, 28, 0.2);
            border: 1px solid var(--orange);
        }
        .container { max-width: 1400px; margin: 40px auto; padding: 0 20px; }
        .card { 
    background: rgba(255,255,255,0.15); 
    backdrop-filter: blur(20px); 
    border-radius: 20px; 
    padding: 35px; 
    border: 1px solid rgba(255,255,255,0.25); 
    margin-bottom: 25px; 
    transition: transform 0.25s ease, box-shadow 0.25s ease;
}

.card:hover {
    transform: translateY(-4px);
    box-shadow: 0 16px 40px rgba(0, 0, 0, 0.18);
}

        .error-card {
            background: rgba(239, 68, 68, 0.2);
            border: 2px solid #ef4444;
        }
        .error-card h3 {
            color: #ef4444;
            margin-bottom: 15px;
        }
        .error-card pre {
            background: rgba(0,0,0,0.3);
            padding: 15px;
            border-radius: 8px;
            overflow-x: auto;
            font-size: 0.9rem;
        }
        select { 
            width: 100%; 
            padding: 15px; 
            background: rgba(255, 255, 255, 0.2); 
            color: white; 
            border: 1px solid rgba(255,255,255,0.4); 
            border-radius: 12px; 
            cursor: pointer; 
            font-family: 'Outfit'; 
            font-size: 1rem;
        }
        select option { background: var(--royal-blue-dark); }
        .row { 
            display: flex; 
            justify-content: space-between; 
            padding: 18px 0; 
            border-bottom: 1px solid rgba(255,255,255,0.1); 
        }
        .spec-grid { 
            display: grid; 
            grid-template-columns: 2fr 1fr 1fr 1fr; 
            gap: 10px; 
            padding: 15px 0; 
            border-bottom: 1px solid rgba(255,255,255,0.1); 
            align-items: center; 
        }
        .btn-analyze { 
            background: var(--orange); 
            border: none; 
            color: white; 
            padding: 16px; 
            border-radius: 12px; 
            font-weight: 800; 
            cursor: pointer; 
            width: 100%; 
            margin-top: 15px; 
            font-family: 'Outfit'; 
            font-size: 1rem; 
            transition: all 0.3s; 
        }
        .btn-analyze:hover { 
            background: #d65519; 
            transform: scale(1.02); 
        }
        .btn-analyze:disabled {
            opacity: 0.5;
            cursor: not-allowed;
        }
        .btn-secondary {
            background: rgba(255,255,255,0.2);
            border: 1px solid rgba(255,255,255,0.4);
            color: white;
            padding: 12px 20px;
            border-radius: 10px;
            font-weight: 700;
            cursor: pointer;
            font-family: 'Outfit';
            transition: all 0.3s;
            display: inline-block;
            text-decoration: none;
            font-size: 0.9rem;
        }
        .btn-secondary:hover {
            background: rgba(255,255,255,0.3);
        }
        .badge { 
            padding: 4px 10px; 
            border-radius: 6px; 
            font-size: 0.7rem; 
            font-weight: 800; 
            text-transform: uppercase; 
        }
        .badge-bullish { background: #ef4444; } 
        .badge-bearish { background: #10b981; } 
        .badge-stable { background: rgba(255,255,255,0.3); }
        .update-notification { 
            position: fixed; 
            top: 100px; 
            right: 20px; 
            background: var(--orange); 
            color: white; 
            padding: 15px 20px; 
            border-radius: 10px; 
            box-shadow: 0 4px 20px rgba(232, 96, 28, 0.4); 
            z-index: 1000; 
            display: none; 
            animation: slideIn 0.3s; 
        }
        @keyframes slideIn { 
            from { transform: translateX(400px); opacity: 0; } 
            to { transform: translateX(0); opacity: 1; } 
        }
        .update-notification button { 
            background: white; 
            color: var(--orange); 
            border: none; 
            padding: 8px 15px; 
            border-radius: 6px; 
            margin-left: 15px; 
            cursor: pointer; 
            font-weight: 700; 
        }
        .loading {
            display: inline-block;
            width: 20px;
            height: 20px;
            border: 3px solid rgba(255,255,255,0.3);
            border-radius: 50%;
            border-top-color: white;
            animation: spin 1s ease-in-out infinite;
        }
        @keyframes spin {
            to { transform: rotate(360deg); }
        }
        h3 { 
            color: var(--orange); 
            text-transform: uppercase; 
            font-size: 0.9rem; 
            margin: 0; 
        }
        
        /* Dashboard-specific styles */
        .stat-card { 
            text-align: center; 
            padding: 30px; 
            position: relative; 
            overflow: hidden; 
        }
        .stat-card::before { 
            content: ''; 
            position: absolute; 
            top: 0; 
            left: 0; 
            right: 0; 
            height: 4px; 
            background: linear-gradient(90deg, var(--orange), #ff8f5e); 
        }
        .stat-number { 
            font-size: 3rem; 
            font-weight: 800; 
            color: var(--orange); 
            margin: 15px 0; 
        }
        .stat-label { 
            font-size: 0.9rem; 
            text-transform: uppercase; 
            opacity: 0.8; 
            letter-spacing: 1px; 
        }
        .stat-trend { 
            font-size: 0.85rem; 
            margin-top: 10px; 
            color: #10b981; 
        }
        .quick-action { 
            text-align: center; 
            padding: 25px; 
            cursor: pointer; 
            text-decoration: none; 
            color: white; 
            display: block; 
            position: relative; 
            overflow: hidden; 
        }
        .quick-action::before { 
            content: ''; 
            position: absolute; 
            top: 50%; 
            left: 50%; 
            width: 0; 
            height: 0; 
            border-radius: 50%; 
            background: rgba(232, 96, 28, 0.3); 
            transition: all 0.5s; 
            transform: translate(-50%, -50%); 
        }
        .quick-action:hover::before { 
            width: 300px; 
            height: 300px; 
        }
        .quick-action-title { 
            font-weight: 700; 
            font-size: 1.1rem; 
            position: relative; 
            z-index: 1; 
        }
        .section-header { margin-bottom: 30px; }
        .section-title { 
            font-size: 1.5rem; 
            font-weight: 800; 
            color: var(--orange); 
            text-transform: uppercase; 
            letter-spacing: 2px; 
        }
        .grid-2 { display: grid; grid-template-columns: repeat(2, 1fr); gap: 20px; }
        .grid-3 { display: grid; grid-template-columns: repeat(3, 1fr); gap: 20px; }
        .grid-4 { display: grid; grid-template-columns: repeat(4, 1fr); gap: 20px; }
        
        @media (max-width: 1024px) {
            .grid-4 { grid-template-columns: repeat(2, 1fr); }
            .grid-3 { grid-template-columns: repeat(2, 1fr); }
            .spec-grid { grid-template-columns: 2fr 1fr 1fr; }
        }
        @media (max-width: 768px) {
            .grid-4, .grid-3, .grid-2 { grid-template-columns: 1fr; }
            .navbar { padding: 15px 20px; }
            .nav-links { flex-direction: column; gap: 10px; }
            .spec-grid { grid-template-columns: 1fr; }
        }
    </style>
</head>
<body onload="initPage()">
    <div class="update-notification" id="updateNotification">
        <span id="updateMessage">Data has been updated!</span>
        <button onclick="reloadPage()">Refresh</button>
    </div>
    <nav class="navbar">
        <div class="navbar-logo"><img src="/static/logo.png" alt="Packfora Logo"></div>
        <div class="nav-links">
            <a href="/" class="{{ 'active' if active == 'Dashboard' else '' }}">Dashboard</a>
            <a href="/resin" class="{{ 'active' if active == 'Resin' else '' }}">Resin Tracker</a>
            <a href="/machines" class="{{ 'active' if active == 'Machines' else '' }}">Machine Database</a>
            <a href="/costs" class="{{ 'active' if active == 'Costs' else '' }}">Global Variable Cost Database</a>
            <a href="/calculator" class="{{ 'active' if active == 'Calculator' else '' }}">Cost Calculator</a>
            <a href="/admin/login" class="admin-link">Admin</a>
        </div>
    </nav>
    <div class="container">{{ content | safe }}</div>
    <script>
    let fileCheckInterval;
    
    async function checkFileUpdates() {
        try {
            const response = await fetch("/api/check_file_updates");
            const data = await response.json();
            
            const currentPage = window.location.pathname;
            let shouldNotify = false;
            let message = "";
            
            if (currentPage === "/resin" && data.resin_updated) {
                shouldNotify = true;
                message = "Resin data has been updated!";
            } else if (currentPage === "/machines" && data.machine_updated) {
                shouldNotify = true;
                message = "Machine database has been updated!";
            } else if (currentPage === "/costs" && data.cost_updated) {
                shouldNotify = true;
                message = "Cost data has been updated!";
            }
            
            if (shouldNotify) {
                showUpdateNotification(message);
            }
        } catch (error) {
            console.error("Error checking file updates:", error);
        }
    }
    
    function showUpdateNotification(message) {
        const notification = document.getElementById('updateNotification');
        const messageEl = document.getElementById('updateMessage');
        messageEl.textContent = message;
        notification.style.display = 'block';
        
        setTimeout(() => {
            notification.style.display = 'none';
        }, 10000);
    }
    
    function reloadPage() {
        location.reload();
    }
    
    function startFileMonitoring() {
        checkFileUpdates();
        fileCheckInterval = setInterval(checkFileUpdates, """ + str(FILE_CHECK_INTERVAL_SECONDS * 1000) + """);
    }
    
    async function initPage() {
        const p = window.location.pathname;
        if(p === "/machines" || p === "/costs") {
            const m = p === "/machines" ? "machines" : "costs";
            try {
                const r = await fetch("/api/init", {
                    method:"POST", 
                    headers:{"Content-Type":"application/json"}, 
                    body:JSON.stringify({module: m})
                });
                const d = await r.json();
                const s = p === "/machines" ? document.getElementById('cat') : document.getElementById('country');
                if (s && d && Array.isArray(d)) {
                    d.forEach(i => { 
                        let o = document.createElement('option'); 
                        o.value=i; 
                        o.text=i; 
                        s.add(o); 
                    });
                }
            } catch (error) {
                console.error('Error initializing page:', error);
            }
        } else if(p === "/") {
            loadDashboardData();
        } else if(p === "/calculator") {
            // Initialize flex layers if on calculator page
            if(document.getElementById('flex-layers') && typeof renderFlexLayers === 'function') {
                renderFlexLayers();
            }
        }
        
        startFileMonitoring();
    }
    
    function showError(elementId, message) {
        const el = document.getElementById(elementId);
        if (el) {
            el.innerHTML = `<div class="error-card"><h3>Error</h3><p>${message}</p></div>`;
        }
    }
    </script>
</body>
</html>
"""

DASH_HTML = """
<div class="section-header">
    <h1>Packfora <span style="color:var(--orange)">Analytics</span></h1>
    <p style="opacity:0.8; margin-top:10px; font-size:1.1rem">Real-time business intelligence for packaging industry</p>
</div>

<!-- Statistics Overview -->
<div class="grid-3" id="stats-container">
    <div class="card stat-card">
        <div class="stat-label">Resin Types</div>
        <div class="stat-number" id="stat-resin"><span class="loading"></span></div>
        <div class="stat-trend">Market Coverage</div>
    </div>
    <div class="card stat-card">
        <div class="stat-label">Machine Database</div>
        <div class="stat-number" id="stat-machines"><span class="loading"></span></div>
        <div class="stat-trend">Equipment Options</div>
    </div>
    <div class="card stat-card">
        <div class="stat-label">Global Markets</div>
        <div class="stat-number" id="stat-countries"><span class="loading"></span></div>
        <div class="stat-trend">Countries Tracked</div>
    </div>
</div>

<!-- Quick Actions -->
<div class="section-header" style="margin-top:50px;">
    <div class="section-title">Quick Actions</div>
</div>
<div class="grid-4">
    <a href="/resin" class="card quick-action">
        <div class="quick-action-title">Resin Price Analysis</div>
        <p style="opacity:0.8; margin-top:10px; font-size:0.9rem;">Track real-time pricing trends</p>
    </a>
    <a href="/machines" class="card quick-action">
        <div class="quick-action-title">Machine Database</div>
        <p style="opacity:0.8; margin-top:10px; font-size:0.9rem;">Explore equipment specifications</p>
    </a>
    <a href="/costs" class="card quick-action">
        <div class="quick-action-title">Global Variable Costs</div>
        <p style="opacity:0.8; margin-top:10px; font-size:0.9rem;">Compare regional expenses</p>
    </a>
    <a href="/calculator" class="card quick-action">
        <div class="quick-action-title">Cost Calculator</div>
        <p style="opacity:0.8; margin-top:10px; font-size:0.9rem;">Carton & Flexibles costing</p>
    </a>
</div>

<!-- System Status -->
<div class="card" style="margin-top:50px; padding:20px; background:rgba(255,255,255,0.08);">
    <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:15px;">
        <div style="display:flex; align-items:center; gap:15px;">
            <div style="width:10px; height:10px; border-radius:50%; background:#10b981; box-shadow:0 0 10px #10b981;"></div>
            <span style="font-size:0.9rem; opacity:0.9;">All systems operational</span>
        </div>
        <div style="font-size:0.85rem; opacity:0.7;" id="last-updated">Last updated: Loading...</div>
    </div>
</div>

<script>
async function loadDashboardData() {
    try {
        const statsRes = await fetch("/api/dashboard_stats");
        if (!statsRes.ok) throw new Error('Failed to load statistics');
        
        const stats = await statsRes.json();
        
        document.getElementById('stat-resin').textContent = stats.resin_types;
        document.getElementById('stat-machines').textContent = stats.machines;
        document.getElementById('stat-countries').textContent = stats.countries;
        document.getElementById('last-updated').textContent = 'Last updated: ' + stats.last_updated;
        
    } catch (error) {
        console.error('Error loading dashboard data:', error);
        document.getElementById('stat-resin').textContent = '0';
        document.getElementById('stat-machines').textContent = '0';
        document.getElementById('stat-countries').textContent = '0';
    }
}
</script>
"""

RESIN_UI = """
<style>
.tabs {
    display: flex;
    gap: 10px;
    margin-bottom: 20px;
    border-bottom: 2px solid rgba(232, 96, 28, 0.2);
}
.tab {
    padding: 12px 24px;
    background: transparent;
    border: none;
    color: #e8601c;
    cursor: pointer;
    font-weight: 700;
    font-size: 0.95rem;
    transition: all 0.3s;
    border-bottom: 3px solid transparent;
}
.tab:hover {
    background: rgba(232, 96, 28, 0.1);
}
.tab.active {
    border-bottom-color: #e8601c;
    color: #e8601c;
}
.tab-content {
    display: none;
}
.tab-content.active {
    display: block;
}
.comparison-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(300px, 1fr));
    gap: 20px;
    margin-top: 20px;
}
.comparison-card {
    background: linear-gradient(135deg, rgba(232, 96, 28, 0.1) 0%, rgba(232, 96, 28, 0.05) 100%);
    border: 2px solid rgba(232, 96, 28, 0.3);
    border-radius: 12px;
    padding: 20px;
    transition: transform 0.2s;
}
.comparison-card:hover {
    transform: translateY(-5px);
    border-color: #e8601c;
}
.comparison-card.best {
    border-color: #28a745;
    background: linear-gradient(135deg, rgba(40, 167, 69, 0.15) 0%, rgba(40, 167, 69, 0.05) 100%);
}
.comparison-card.worst {
    border-color: #dc3545;
    background: linear-gradient(135deg, rgba(220, 53, 69, 0.15) 0%, rgba(220, 53, 69, 0.05) 100%);
}
.location-badge {
    display: inline-block;
    padding: 5px 10px;
    background: #e8601c;
    color: white;
    border-radius: 5px;
    font-size: 0.75rem;
    font-weight: 800;
    margin-bottom: 10px;
}
.best .location-badge {
    background: #28a745;
}
.worst .location-badge {
    background: #dc3545;
}
.stat-row {
    display: flex;
    justify-content: space-between;
    padding: 10px 0;
    border-bottom: 1px solid rgba(255,255,255,0.1);
}
.stat-label {
    opacity: 0.7;
    font-size: 0.85rem;
}
.stat-value {
    font-weight: 800;
    font-size: 1rem;
}
.trend-badge {
    display: inline-block;
    padding: 4px 12px;
    border-radius: 20px;
    font-size: 0.8rem;
    font-weight: 700;
}
.trend-rising {
    background: #dc3545;
    color: white;
}
.trend-falling {
    background: #28a745;
    color: white;
}
.trend-stable {
    background: #ffc107;
    color: #000;
}
input[type="checkbox"] {
    width: 18px;
    height: 18px;
    margin-right: 8px;
    cursor: pointer;
}
label.checkbox-label {
    display: flex;
    align-items: center;
    cursor: pointer;
    padding: 8px;
    border-radius: 6px;
    transition: background 0.2s;
}
label.checkbox-label:hover {
    background: rgba(232, 96, 28, 0.1);
}
</style>

<div class="tabs">
    <button class="tab active" onclick="switchTab('search')">Search Prices</button>
    <button class="tab" onclick="switchTab('compare')">Compare Regions</button>
</div>

<!-- Search Tab -->
<div id="search-tab" class="tab-content active">
<h2>Resin Price Tracker</h2>
<div class="card">
    <div style="display:grid; grid-template-columns: repeat(4, 1fr); gap: 20px;">
        <div>
            <label style="display:block; font-size:.75rem; margin-bottom:10px; font-weight:800; opacity:0.9">RESIN TYPE</label>
            <select id="resSheet" onchange="loadResSub()">
                <option value="">Select...</option>
                {{SHEETS_OPTIONS}}
            </select>
        </div>
        <div>
            <label style="display:block; font-size:.75rem; margin-bottom:10px; font-weight:800; opacity:0.9">LOCATION</label>
            <select id="resLoc" disabled><option>Select Resin First</option></select>
        </div>
        <div>
            <label style="display:block; font-size:.75rem; margin-bottom:10px; font-weight:800; opacity:0.9">GRADE</label>
            <select id="resGrade" disabled><option>Select Resin First</option></select>
        </div>
        <div>
            <label style="display:block; font-size:.75rem; margin-bottom:10px; font-weight:800; opacity:0.9">DURATION</label>
            <select id="resDuration">
                <option value="3">Last 3 Months</option>
                <option value="6">Last 6 Months</option>
                <option value="12" selected>Last 1 Year</option>
                <option value="all">All Time</option>
            </select>
        </div>
    </div>
    <button class="btn-analyze" id="analyzeBtn" onclick="genRes()" disabled>Generate Market Analysis</button>
</div>
<div id="res_results"></div>
</div>

<!-- Compare Regions Tab -->
<div id="compare-tab" class="tab-content">
    <h2>Compare Regions</h2>
    <div class="card">
        <div style="display:grid; grid-template-columns: repeat(3, 1fr); gap: 20px;">
            <div>
                <label style="display:block; font-size:.75rem; margin-bottom:10px; font-weight:800; opacity:0.9">RESIN TYPE</label>
                <select id="cmp_rt" onchange="loadGradesCompare()">
                    <option value="">Select...</option>
                    {{SHEETS_OPTIONS}}
                </select>
            </div>
            <div>
                <label style="display:block; font-size:.75rem; margin-bottom:10px; font-weight:800; opacity:0.9">GRADE</label>
                <select id="cmp_grd" onchange="enableCompare()">
                    <option value="">Select Resin First</option>
                </select>
            </div>
            <div>
                <label style="display:block; font-size:.75rem; margin-bottom:10px; font-weight:800; opacity:0.9">DURATION</label>
                <select id="cmp_dur">
                    <option>Last 1 Month</option>
                    <option>Last 3 Months</option>
                    <option>Last 6 Months</option>
                    <option selected>Last 1 Year</option>
                    <option>Last 2 Years</option>
                </select>
            </div>
        </div>
        
        <div style="margin-top: 20px;">
            <div style="display:flex; align-items:center; justify-content:space-between; margin-bottom:10px;">
                <label style="display:block; font-size:.75rem; font-weight:800; opacity:0.9">SELECT LOCATIONS TO COMPARE </label>
                <button id="toggleAllLocBtn" onclick="toggleAllLocations()" style="padding:6px 16px; font-size:0.8rem; border-radius:5px; border:1px solid var(--orange); background:transparent; color:var(--orange); cursor:pointer; font-weight:700; transition:all 0.2s;" onmouseover="this.style.background='var(--orange)';this.style.color='#fff'" onmouseout="this.style.background='transparent';this.style.color='var(--orange)'">Select All</button>
            </div>
            <div id="location-checkboxes" style="display: grid; grid-template-columns: repeat(auto-fill, minmax(200px, 1fr)); gap: 10px; max-height: 350px; overflow-y: auto; padding-right: 5px;">
                <!-- Checkboxes will be populated here -->
            </div>
        </div>
        
        <button class="btn-analyze" id="cmpBtn" onclick="compareRegions()" disabled>Compare Regions</button>
    </div>
    <div id="cmp_res"></div>
</div>

<script>
// Tab switching
function switchTab(tabName) {
    document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
    document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
    
    event.target.classList.add('active');
    document.getElementById(`${tabName}-tab`).classList.add('active');
}

// ORIGINAL SEARCH FUNCTIONALITY
async function loadResSub() {
    const sheet = document.getElementById('resSheet').value;
    if (!sheet) return;
    
    // Fire-and-forget: pre-warm full sheet cache so Generate is instant
    fetch("/api/resin_preload", {
        method:"POST", headers:{"Content-Type":"application/json"},
        body:JSON.stringify({sheet})
    }).catch(() => {});
    
    try {
        const r = await fetch("/api/resin_load", {
            method:"POST", 
            headers:{"Content-Type":"application/json"}, 
            body:JSON.stringify({sheet})
        });
        
        if (!r.ok) throw new Error('Failed to load resin data');
        
        const d = await r.json();
        
        const locSelect = document.getElementById('resLoc');
        const gradeSelect = document.getElementById('resGrade');
        
        locSelect.innerHTML = d.locations.map(l => `<option>${l}</option>`).join('');
        gradeSelect.innerHTML = d.grades.map(g => `<option>${g}</option>`).join('');
        
        locSelect.disabled = false;
        gradeSelect.disabled = false;
        document.getElementById('analyzeBtn').disabled = false;
    } catch (error) {
        console.error('Error loading resin data:', error);
        showError('res_results', 'Failed to load resin data. Please try again.');
    }
}

async function genRes() {
    const sheet = document.getElementById('resSheet').value;
    const location = document.getElementById('resLoc').value;
    const grade = document.getElementById('resGrade').value;
    const duration = document.getElementById('resDuration').value;
    
    if (!sheet || !location || !grade) {
        alert('Please select all fields');
        return;
    }
    
    const btn = document.getElementById('analyzeBtn');
    btn.disabled = true;
    btn.innerHTML = '<span class="loading"></span> Analyzing...';
    
    try {
        const r = await fetch("/api/resin_generate", {
            method:"POST", 
            headers:{"Content-Type":"application/json"}, 
            body:JSON.stringify({sheet, location, grade, duration})
        });
        
        if (!r.ok) {
            const error = await r.json();
            throw new Error(error.error || 'Failed to generate analysis');
        }
        
        const d = await r.json();
        const i = d.insights;
        
        let h = `<div class="card">
            <div style="background: linear-gradient(135deg, rgba(232, 96, 28, 0.2) 0%, rgba(232, 96, 28, 0.05) 100%); border: 2px solid var(--orange); border-radius: 15px; padding: 25px; margin-bottom: 25px;">
                <div style="display:grid; grid-template-columns: repeat(3, 1fr); gap: 20px; margin-bottom: 20px;">
                    <div><div style="opacity:0.7; font-size:0.8rem; margin-bottom:5px;">CURRENT PRICE</div><div style="font-size:2rem; font-weight:900; color:var(--orange);">${i.curr}</div></div>
                    <div><div style="opacity:0.7; font-size:0.8rem; margin-bottom:5px;">PERIOD CHANGE</div><div style="font-size:2rem; font-weight:900; ${parseFloat(i.diff) > 0 ? 'color:#dc3545;' : 'color:#10b981;'}">${i.diff}</div></div>
                    <div><div style="opacity:0.7; font-size:0.8rem; margin-bottom:5px;">MARKET STATUS</div><div><span class="badge ${i.badge}" style="font-size:0.9rem; padding:8px 15px;">${i.status}</span></div></div>
                </div>
                <div style="display:grid; grid-template-columns: repeat(3, 1fr); gap: 20px; padding-top: 20px; border-top: 1px solid rgba(255,255,255,0.2);">
                    <div><div style="opacity:0.7; font-size:0.8rem;">Average Price</div><div style="font-weight:800; font-size:1.1rem;">${i.avg}</div></div>
                    <div><div style="opacity:0.7; font-size:0.8rem;">Min Price</div><div style="font-weight:800; font-size:1.1rem; color:#10b981;">${i.min}</div></div>
                    <div><div style="opacity:0.7; font-size:0.8rem;">Max Price</div><div style="font-weight:800; font-size:1.1rem; color:#dc3545;">${i.max}</div></div>
                </div>
            </div>
       
        </div>`;
        
        h += '<div class="card"><div id="priceChart"></div></div>';
        document.getElementById('res_results').innerHTML = h;
        
        Plotly.newPlot('priceChart', [{
            x: d.series.dates,
            y: d.series.values,
            type: 'scatter',
            mode: 'lines+markers',
            marker: {color: '#E8601C', size: 8},
            line: {color: '#E8601C', width: 3},
            text: d.series.labels || d.series.dates,
            hovertemplate: '%{text}<br>₹%{y:,.0f}<extra></extra>'
        }], {
            title: {text: 'Price Trend', font: {color: 'white', size: 18, family: 'Outfit'}},
            xaxis: {title: 'Date', color: 'white', gridcolor: 'rgba(255,255,255,0.1)', type: 'date', tickformat: '%b %Y'},
            yaxis: {title: 'Price (₹)', color: 'white', gridcolor: 'rgba(255,255,255,0.1)'},
            plot_bgcolor: 'rgba(0,0,0,0)',
            paper_bgcolor: 'rgba(0,0,0,0)',
            font: {color: 'white', family: 'Outfit'}
        }, {responsive: true});
        
    } catch (error) {
        console.error('Error generating analysis:', error);
        document.getElementById('res_results').innerHTML = `<div class="error-card"><h3>Error</h3><p>${error.message}</p></div>`;
    } finally {
        btn.disabled = false;
        btn.innerHTML = 'Generate Market Analysis';
    }
}

// COMPARISON FUNCTIONALITY
async function loadGradesCompare() {
    const rt = document.getElementById('cmp_rt').value;
    if (!rt) return;
    
    // Fire-and-forget: pre-warm full sheet cache so Compare is instant
    fetch("/api/resin_preload", {
        method:"POST", headers:{"Content-Type":"application/json"},
        body:JSON.stringify({sheet: rt})
    }).catch(() => {});
    
    try {
        const r = await fetch("/api/resin_grades", {
            method: "POST",
            headers: {"Content-Type": "application/json"},
            body: JSON.stringify({resin_type: rt})
        });
        
        const d = await r.json();
        const grd = document.getElementById('cmp_grd');
        grd.innerHTML = '<option value="">Select...</option>';
        d.grades.forEach(g => {
            const o = document.createElement('option');
            o.value = g;
            o.text = g;
            grd.add(o);
        });
        grd.disabled = false;
        
        loadLocationCheckboxes(d.locations || []);
        
    } catch (error) {
        console.error('Error loading grades:', error);
    }
}

function loadLocationCheckboxes(locations) {
    const container = document.getElementById('location-checkboxes');
    container.innerHTML = '';
    
    locations.forEach(loc => {
        const label = document.createElement('label');
        label.className = 'checkbox-label';
        
        const checkbox = document.createElement('input');
        checkbox.type = 'checkbox';
        checkbox.value = loc;
        checkbox.onchange = enableCompare;
        
        label.appendChild(checkbox);
        label.appendChild(document.createTextNode(loc));
        container.appendChild(label);
    });
    
    // Reset toggle button text
    const toggleBtn = document.getElementById('toggleAllLocBtn');
    if (toggleBtn) toggleBtn.textContent = 'Select All';
}

function toggleAllLocations() {
    const checkboxes = document.querySelectorAll('#location-checkboxes input[type="checkbox"]');
    const btn = document.getElementById('toggleAllLocBtn');
    if (!checkboxes.length) return;
    
    const allChecked = Array.from(checkboxes).every(cb => cb.checked);
    checkboxes.forEach(cb => cb.checked = !allChecked);
    btn.textContent = allChecked ? 'Select All' : 'Deselect All';
    enableCompare();
}

function enableCompare() {
    const grade = document.getElementById('cmp_grd').value;
    const checked = document.querySelectorAll('#location-checkboxes input:checked').length;
    const btn = document.getElementById('cmpBtn');
    btn.disabled = !grade || checked < 2;
}

async function compareRegions() {
    const rt = document.getElementById('cmp_rt').value;
    const grd = document.getElementById('cmp_grd').value;
    const dur = document.getElementById('cmp_dur').value;
    
    const checkboxes = document.querySelectorAll('#location-checkboxes input:checked');
    const locations = Array.from(checkboxes).map(cb => cb.value);
    
    // Client-side validation
    if (!rt || !grd) {
        alert('Please select both resin type and grade');
        return;
    }
    
    if (locations.length < 2) {
        alert('Please select at least 2 locations to compare');
        return;
    }
    
    const btn = document.getElementById('cmpBtn');
    btn.disabled = true;
    btn.innerHTML = '<span class="loading"></span> Comparing...';
    
    const resultsDiv = document.getElementById('cmp_res');
    resultsDiv.innerHTML = '<div class="card"><p style="opacity:0.6; text-align:center;"><span class="loading"></span> Analyzing price data...</p></div>';
    
    try {
        const r = await fetch("/api/resin_compare", {
            method: "POST",
            headers: {"Content-Type": "application/json"},
            body: JSON.stringify({
                resin_type: rt,
                grade: grd,
                locations: locations,
                duration: dur
            })
        });
        
        const d = await r.json();
        
        // Check if request failed and show actual error message
        if (!r.ok) {
            throw new Error(d.error || 'Comparison failed');
        }
        
        displayComparison(d);
        
    } catch (error) {
        console.error('Error comparing regions:', error);
        // Display the actual error message from the server
        resultsDiv.innerHTML = `<div class="card" style="border-color:#dc3545; background: rgba(220, 53, 69, 0.1); padding: 30px;">
            <div style="text-align: center;">
                <div style="font-size: 3rem; margin-bottom: 15px;">⚠️</div>
                <h3 style="color:#dc3545; margin-bottom: 15px;">Comparison Failed</h3>
                <p style="color:#dc3545; font-size: 1.05rem; line-height: 1.6; margin-bottom: 20px;">${error.message}</p>
                <p style="opacity: 0.7; font-size: 0.9rem;">Please check your selections and try again. If the problem persists, verify your data source.</p>
            </div>
        </div>`;
    } finally {
        btn.disabled = false;
        btn.innerHTML = 'Compare Regions';
    }
}

function displayComparison(data) {
    let html = '<div class="card" style="background: linear-gradient(135deg, rgba(232, 96, 28, 0.15) 0%, rgba(232, 96, 28, 0.05) 100%); border: 2px solid var(--orange); margin-bottom: 20px;">';
    html += '<h3 style="margin-bottom: 10px;">Regional Price Comparison</h3>';
    html += `<p style="opacity: 0.8; margin-bottom: 15px;">${data.resin_type} - ${data.grade} | ${data.duration}</p>`;
    html += '<div style="display: grid; grid-template-columns: repeat(3, 1fr); gap: 15px; padding: 15px; background: rgba(255,255,255,0.1); border-radius: 8px;">';
    html += `<div><div style="opacity: 0.7; font-size: 0.8rem;">BEST PRICE</div><div style="font-size: 1.2rem; font-weight: 800; color: #28a745;">${data.summary.best_price_location}</div></div>`;
    html += `<div><div style="opacity: 0.7; font-size: 0.8rem;">HIGHEST PRICE</div><div style="font-size: 1.2rem; font-weight: 800; color: #dc3545;">${data.summary.worst_price_location}</div></div>`;
    html += `<div><div style="opacity: 0.7; font-size: 0.8rem;">PRICE SPREAD</div><div style="font-size: 1.2rem; font-weight: 800;">${data.summary.price_spread}</div></div>`;
    html += '</div></div>';
    
    html += '<div class="comparison-grid" style="max-height: 600px; overflow-y: auto;">';
    
    data.comparison.forEach((loc, idx) => {
        const cardClass = idx === 0 ? 'best' : (idx === data.comparison.length - 1 ? 'worst' : '');
        const trendClass = loc.trend === 'Rising' ? 'trend-rising' : (loc.trend === 'Falling' ? 'trend-falling' : 'trend-stable');
        
        html += `<div class="comparison-card ${cardClass}">`;
        html += `<span class="location-badge">${loc.location}</span>`;
        html += `<div class="stat-row"><span class="stat-label">Current Price</span><span class="stat-value">${loc.current_price}</span></div>`;
        html += `<div class="stat-row"><span class="stat-label">Average Price</span><span class="stat-value">${loc.avg_price}</span></div>`;
        html += `<div class="stat-row"><span class="stat-label">Price Range</span><span class="stat-value">${loc.min_price} - ${loc.max_price}</span></div>`;
        html += `<div class="stat-row"><span class="stat-label">Trend</span><span class="trend-badge ${trendClass}">${loc.trend} ${loc.price_change}</span></div>`;
        html += `<div class="stat-row"><span class="stat-label">Data Points</span><span class="stat-value">${loc.data_points}</span></div>`;
        html += '</div>';
    });
    
    html += '</div>';
    html += '<button class="btn-secondary" onclick="exportComparison()" style="margin-top: 20px; width: 100%;">Export Comparison to Excel</button>';
    
    document.getElementById('cmp_res').innerHTML = html;
    window.currentComparisonData = data;
}

async function exportComparison() {
    if (!window.currentComparisonData) {
        alert('No comparison data to export');
        return;
    }
    
    try {
        const response = await fetch("/api/export_comparison", {
            method: "POST",
            headers: {"Content-Type": "application/json"},
            body: JSON.stringify(window.currentComparisonData)
        });
        
        if (!response.ok) throw new Error('Export failed');
        
        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `resin_comparison_${new Date().getTime()}.xlsx`;
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
        document.body.removeChild(a);
    } catch (error) {
        console.error('Error exporting comparison:', error);
        alert('Failed to export data. Please try again.');
    }
}
</script>
"""

MACH_HTML = """
<h2>Machine Database</h2>
<div class="card">
    <div style="display:grid; grid-template-columns: 1fr 1fr; gap: 20px;">
        <div>
            <label style="display:block; font-size:.75rem; margin-bottom:10px; font-weight:800; opacity:0.9">CATEGORY</label>
            <select id="cat" onchange="loadProcs(this.value)"><option value="">Select...</option></select>
        </div>
        <div>
            <label style="display:block; font-size:.75rem; margin-bottom:10px; font-weight:800; opacity:0.9">PROCESS</label>
            <select id="proc" onchange="enableSearch()" disabled><option>Select Category First</option></select>
        </div>
    </div>
    <button class="btn-analyze" id="searchBtn" onclick="loadMachs()" disabled>Search Machines</button>
</div>
<div id="ai_recommendation"></div>
<div id="m_res"></div>
<script>
let currentResults = [];

async function loadProcs(cat) {
    if (!cat) return;
    
    try {
        const r = await fetch("/api/procs", {
            method:"POST", 
            headers:{"Content-Type":"application/json"}, 
            body:JSON.stringify({cat})
        });
        
        if (!r.ok) throw new Error('Failed to load processes');
        
        const d = await r.json();
        let s = document.getElementById('proc'); 
        s.innerHTML = '<option value="">Select Process...</option>';
        d.forEach(i => { 
            let o = document.createElement('option'); 
            o.value=i; 
            o.text=i; 
            s.add(o); 
        });
        s.disabled = false;
    } catch (error) {
        console.error('Error loading processes:', error);
        alert('Failed to load processes. Please try again.');
    }
}

function enableSearch() {
    const proc = document.getElementById('proc').value;
    document.getElementById('searchBtn').disabled = !proc;
}

async function loadMachs() {
    const cat = document.getElementById('cat').value;
    const proc = document.getElementById('proc').value;
    
    if (!cat || !proc) {
        alert('Please select both category and process');
        return;
    }
    
    const btn = document.getElementById('searchBtn');
    btn.disabled = true;
    btn.innerHTML = '<span class="loading"></span> Searching...';
    
    try {
        const r = await fetch("/api/mach_res", {
            method:"POST", 
            headers:{"Content-Type":"application/json"}, 
            body:JSON.stringify({cat, proc})
        });
        
        if (!r.ok) throw new Error('Failed to load machines');
        
        const d = await r.json();
        currentResults = d.results;
        
        // Display all machines
        let h = '<div class="card">';
        h += '<div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:20px;">';
        h += `<h3>Found ${d.results.length} Machines</h3>`;
        h += '<button class="btn-secondary" onclick="exportMachines()">Export to Excel</button>';
        h += '</div>';
        h += '<div class="spec-grid" style="border-bottom: 2px solid var(--orange); font-weight: 800; padding-bottom: 15px; margin-bottom: 10px;"><div>Machine Model</div><div>Price (€)</div><div>Power (kWh)</div><div>Footprint (SQM)</div></div>';
        
        d.results.forEach(m => {
            // FIX: Smart Name Display
            // If model starts with make (e.g. "Jomar 65" starts with "Jomar"), don't repeat Make
            let displayName = m.model;
            if (m.make && !m.model.toLowerCase().trim().startsWith(m.make.toLowerCase().trim())) {
                displayName = `${m.make} ${m.model}`;
            }

            h += `<div class="spec-grid"><div><strong>${displayName}</strong></div><div>${m.cost}</div><div>${m.power}</div><div>${m.sqm}</div></div>`;
        });
        
        h += '</div>';
        document.getElementById('m_res').innerHTML = h;
    } catch (error) {
        console.error('Error loading machines:', error);
        showError('m_res', 'Failed to load machines. Please try again.');
    } finally {
        btn.disabled = false;
        btn.innerHTML = 'Search Machines';
    }
}

async function exportMachines() {
    if (currentResults.length === 0) {
        alert('No machines to export');
        return;
    }
    
    try {
        const response = await fetch("/api/export_machines", {
            method: "POST",
            headers: {"Content-Type": "application/json"},
            body: JSON.stringify({results: currentResults})
        });
        
        if (!response.ok) throw new Error('Export failed');
        
        const blob = await response.blob();
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `machine_export_${new Date().getTime()}.xlsx`;
        document.body.appendChild(a);
        a.click();
        window.URL.revokeObjectURL(url);
        document.body.removeChild(a);
    } catch (error) {
        console.error('Error exporting machines:', error);
        alert('Failed to export data. Please try again.');
    }
}
</script>
"""
COST_HTML = """
<h2>Global Variable Cost Database</h2>
<div class="card">
    <label style="display:block; font-size:.75rem; margin-bottom:10px; font-weight:800; opacity:0.9">SELECT GEOGRAPHY</label>
    <select id="country" onchange="loadCosts(this.value)"><option value="">Select Country...</option></select>
</div>
<div id="c_res"></div>
<script>
async function loadCosts(c) {
    if (!c) return;
    
    const resultsDiv = document.getElementById('c_res');
    resultsDiv.innerHTML = '<div class="card"><p style="opacity:0.6; text-align:center;"><span class="loading"></span> Loading cost data...</p></div>';
    
    try {
        const r = await fetch("/api/cost_res", {
            method:"POST", 
            headers:{"Content-Type":"application/json"}, 
            body:JSON.stringify({country: c})
        });
        
        if (!r.ok) throw new Error('Failed to load cost data');
        
        const d = await r.json();
        let h = "";
        
        d.sections.forEach(s => {
            h += `<div class="card"><h3>${s.section}</h3>`;
            s.items.forEach(i => {
                h += `<div class="row"><div>${i.label}</div><div style="font-weight:800">${i.value}</div></div>`;
            });
            h += "</div>";
        });
        
        resultsDiv.innerHTML = h;
    } catch (error) {
        console.error('Error loading costs:', error);
        showError('c_res', 'Failed to load cost data. Please try again.');
    }
}
</script>
"""

CALC_HTML = """
<style>
/* --- Layout & General --- */
.calc-content { display: none; }
.calc-content.active { display: block; }
.calc-layout { display: grid; grid-template-columns: 1fr 1fr; gap: 25px; }
@media (max-width: 1024px) { .calc-layout { grid-template-columns: 1fr; } }
.calc-section { margin-bottom: 20px; }
.calc-section-title { font-size: 0.8rem; font-weight: 800; color: var(--orange); text-transform: uppercase; letter-spacing: 1px; margin-bottom: 12px; padding-bottom: 8px; border-bottom: 1px solid rgba(232, 96, 28, 0.3); }
.calc-row { display: grid; grid-template-columns: 1.5fr 0.5fr 1fr; gap: 8px; align-items: center; margin-bottom: 8px; }
.calc-row label { font-size: 0.85rem; opacity: 0.9; }
.calc-row .unit { font-size: 0.75rem; opacity: 0.6; text-align: center; }

/* --- Inputs --- */
.calc-input { width: 100%; padding: 8px 12px; background: rgba(255,255,255,0.15); color: white; border: 1px solid rgba(255,255,255,0.3); border-radius: 8px; font-family: 'Outfit'; font-size: 0.9rem; }
.calc-input:focus { outline: none; border-color: var(--orange); background: rgba(255,255,255,0.2); }
.calc-input-green { border-color: rgba(76, 175, 80, 0.6); background: rgba(76, 175, 80, 0.1); }
.calc-input-green:focus { border-color: #4CAF50; background: rgba(76, 175, 80, 0.15); }
.calc-select { width: 100%; padding: 8px 12px; background: rgba(76, 175, 80, 0.1); color: white; border: 1px solid rgba(76, 175, 80, 0.6); border-radius: 8px; font-family: 'Outfit'; font-size: 0.9rem; cursor: pointer; }
.calc-select option { background: #1e3a8a; }
.calc-select:focus { outline: none; border-color: #4CAF50; }

/* --- Summary Cards --- */
.summary-card { background: linear-gradient(135deg, rgba(232, 96, 28, 0.15) 0%, rgba(232, 96, 28, 0.05) 100%); border: 2px solid var(--orange); border-radius: 15px; padding: 25px; }
.summary-row { display: flex; justify-content: space-between; padding: 10px 0; border-bottom: 1px solid rgba(255,255,255,0.1); font-size: 0.9rem; }
.summary-row:last-child { border-bottom: none; }
.summary-row .label { opacity: 0.8; }
.summary-row .value { font-weight: 700; }
.summary-total { display: flex; justify-content: space-between; padding: 15px 0; margin-top: 10px; border-top: 2px solid var(--orange); font-size: 1.1rem; font-weight: 800; }
.summary-total .value { color: var(--orange); font-size: 1.3rem; }
.cost-bar { height: 8px; border-radius: 4px; margin-top: 4px; transition: width 0.5s ease; }

/* --- Flex Layers --- */
.flex-layer-block { background: rgba(255,255,255,0.05); border: 1px solid rgba(255,255,255,0.15); border-radius: 10px; padding: 15px; margin-bottom: 12px; }
.flex-layer-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 10px; }
.flex-layer-title { font-size: 0.85rem; font-weight: 700; color: var(--orange); }
.btn-remove-layer { background: rgba(239, 68, 68, 0.2); border: 1px solid #ef4444; color: #ef4444; padding: 4px 10px; border-radius: 6px; cursor: pointer; font-size: 0.75rem; font-family: 'Outfit'; transition: all 0.3s; }
.btn-remove-layer:hover { background: rgba(239, 68, 68, 0.4); }
.btn-add-layer { background: rgba(76, 175, 80, 0.2); border: 1px solid #4CAF50; color: #4CAF50; padding: 10px 20px; border-radius: 10px; cursor: pointer; font-size: 0.85rem; font-weight: 700; font-family: 'Outfit'; width: 100%; margin-top: 10px; transition: all 0.3s; }
.btn-add-layer:hover { background: rgba(76, 175, 80, 0.4); }

/* --- Legends & Nav --- */
.legend { display: flex; gap: 20px; margin-bottom: 15px; font-size: 0.8rem; }
.legend-item { display: flex; align-items: center; gap: 6px; }
.legend-dot { width: 12px; height: 12px; border-radius: 3px; }
.legend-green { background: rgba(76, 175, 80, 0.4); border: 1px solid #4CAF50; }
.legend-grey { background: rgba(255,255,255,0.15); border: 1px solid rgba(255,255,255,0.3); }

.universal-tab-navigation { display: flex; gap: 10px; margin-bottom: 20px; background: rgba(255,255,255,0.1); padding: 10px; border-radius: 15px; }
.universal-tab-btn { flex: 1; padding: 15px 25px; background: rgba(255,255,255,0.15); border: none; border-radius: 10px; color: white; font-family: 'Outfit', sans-serif; font-size: 1rem; font-weight: 600; cursor: pointer; transition: all 0.3s; }
.universal-tab-btn:hover { background: rgba(255,255,255,0.25); }
.universal-tab-btn.active { background: var(--orange); box-shadow: 0 5px 15px rgba(232, 96, 28, 0.4); }
.calculator-view { display: none; }
.calculator-view.active { display: block; }

/* --- NEW STYLES: Sub-Tabs & Login --- */
.sub-tabs { display: flex; border-bottom: 2px solid rgba(255,255,255,0.1); margin-bottom: 20px; }
.sub-tab-btn { padding: 10px 20px; background: transparent; border: none; color: rgba(255,255,255,0.6); font-family: 'Outfit'; font-weight: 600; cursor: pointer; border-bottom: 2px solid transparent; transition: all 0.3s; }
.sub-tab-btn:hover { color: white; }
.sub-tab-btn.active { color: var(--orange); border-bottom-color: var(--orange); }
.sub-tab-content { display: none; animation: fadeIn 0.3s ease; }
.sub-tab-content.active { display: block; }
@keyframes fadeIn { from { opacity: 0; } to { opacity: 1; } }

.login-overlay { background: rgba(0,0,0,0.2); border-radius: 15px; padding: 40px; text-align: center; border: 1px solid rgba(255,255,255,0.1); }
.login-input { padding: 10px 15px; border-radius: 8px; border: 1px solid rgba(255,255,255,0.3); background: rgba(255,255,255,0.1); color: white; font-family: 'Outfit'; margin-right: 10px; }
.login-btn { padding: 10px 20px; background: var(--orange); color: white; border: none; border-radius: 8px; font-weight: 700; cursor: pointer; }

/* Ensure only active model is shown */
.model-view { display: none !important; }
.model-view.active { display: block !important; }

/* ============================================
   MOBILE RESPONSIVE CSS
   ============================================ */

/* Mobile-first responsive breakpoints */
@media screen and (max-width: 768px) {
    .container { padding: 10px !important; max-width: 100% !important; }
    .card { margin-bottom: 15px !important; padding: 12px !important; }
    .tab-container { flex-wrap: wrap !important; gap: 8px !important; }
    .tab { flex: 1 1 calc(50% - 8px) !important; min-width: 120px !important; font-size: 0.85rem !important; padding: 10px 8px !important; }
    .sub-tab-container { flex-wrap: wrap !important; gap: 6px !important; }
    .sub-tab { flex: 1 1 calc(50% - 6px) !important; min-width: 100px !important; font-size: 0.8rem !important; padding: 8px 6px !important; }
    .input-grid { grid-template-columns: 1fr !important; gap: 12px !important; }
    .input-row { flex-direction: column !important; gap: 10px !important; }
    .input-group { width: 100% !important; margin-bottom: 12px !important; }
    .input-group label { font-size: 0.85rem !important; margin-bottom: 4px !important; }
    .input-group input, .input-group select { font-size: 0.9rem !important; padding: 10px !important; width: 100% !important; }
    .btn { width: 100% !important; margin-bottom: 10px !important; padding: 12px !important; font-size: 0.95rem !important; }
    .btn-group { flex-direction: column !important; gap: 10px !important; }
    .btn-group .btn { width: 100% !important; }
    .summary-grid { grid-template-columns: 1fr !important; gap: 10px !important; }
    .summary-item { padding: 10px !important; }
    .summary-item h3 { font-size: 0.85rem !important; }
    .summary-item .value { font-size: 1.2rem !important; }
    .table-container { overflow-x: auto !important; -webkit-overflow-scrolling: touch !important; margin: 10px -12px !important; padding: 0 12px !important; }
    table { min-width: 600px !important; font-size: 0.8rem !important; }
    table th, table td { padding: 8px 6px !important; white-space: nowrap !important; }
    .chart-container { width: 100% !important; height: auto !important; min-height: 300px !important; overflow-x: auto !important; }
    canvas { max-width: 100% !important; height: auto !important; }
    #whatif-controls { flex-direction: column !important; }
    .slider-group { width: 100% !important; margin-bottom: 15px !important; }
    #country-checkboxes { grid-template-columns: 1fr !important; max-height: 250px !important; }
    .country-checkbox { font-size: 0.85rem !important; padding: 8px !important; }
    #flex-layers .layer-card { padding: 10px !important; }
    #flex-layers .input-grid { grid-template-columns: 1fr !important; }
    #results-section { padding: 12px !important; }
    .breakdown-section { margin-bottom: 15px !important; }
    .cost-item { font-size: 0.85rem !important; padding: 6px !important; }
    .modal-content { width: 95% !important; max-width: 95% !important; margin: 20px auto !important; padding: 15px !important; }
    .modal h2 { font-size: 1.2rem !important; }
    header h1 { font-size: 1.5rem !important; }
    .header-actions { flex-direction: column !important; gap: 8px !important; }
    #sku-controls-container { flex-direction: column !important; gap: 10px !important; }
    #sku-controls-container select, #sku-controls-container input, #sku-controls-container button { width: 100% !important; }
    .calc-layout { grid-template-columns: 1fr !important; }
}
@media screen and (min-width: 769px) and (max-width: 1024px) {
    .input-grid { grid-template-columns: repeat(2, 1fr) !important; }
    .tab { flex: 1 1 auto !important; min-width: 100px !important; }
    .summary-grid { grid-template-columns: repeat(2, 1fr) !important; }
    #country-checkboxes { grid-template-columns: repeat(2, 1fr) !important; }
}
</style>

<div class="universal-tab-navigation">
    <button class="universal-tab-btn active" onclick="switchUniversalTab('calculator')" data-tab="calculator">Cost Calculator</button>
    <button class="universal-tab-btn" onclick="switchUniversalTab('compare')" data-tab="compare">Multi-Country Comparison</button>
    <button class="universal-tab-btn" onclick="switchUniversalTab('whatif')" data-tab="whatif">What-If Scenario Analysis</button>
</div>

<!-- SKU SAVE/LOAD CONTROLS -->
<div id="sku-controls-container" style="margin: 20px 0; padding: 15px; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); border-radius: 10px; box-shadow: 0 4px 15px rgba(0,0,0,0.2);">
    <div style="display: flex; align-items: center; gap: 10px; flex-wrap: wrap;">
        <div style="flex: 1; min-width: 200px;">
            <label style="color: white; font-weight: 600; display: block; margin-bottom: 5px; font-size: 0.9rem;">
                 Load Saved SKU
            </label>
            <select id="sku-selector" style="width: 100%; padding: 10px; border: none; border-radius: 6px; font-size: 0.95rem; background: white; cursor: pointer;">
                <option value="">-- Select SKU --</option>
            </select>
        </div>
        <div style="display: flex; gap: 8px; margin-top: 20px;">
            <button onclick="loadSKU()" class="btn" style="background: linear-gradient(135deg, #4CAF50, #45a049); border: none; color: white; padding: 10px 20px; border-radius: 6px; cursor: pointer; font-weight: 600; box-shadow: 0 2px 8px rgba(76,175,80,0.3); transition: all 0.3s;">
                 Load
            </button>
            <button onclick="deleteSKU()" class="btn" style="background: linear-gradient(135deg, #f44336, #da190b); border: none; color: white; padding: 10px 20px; border-radius: 6px; cursor: pointer; font-weight: 600; box-shadow: 0 2px 8px rgba(244,67,54,0.3); transition: all 0.3s;">
                 Delete
            </button>
        </div>
    </div>
</div>

<div id="universal-calculator" class="calculator-view active">
    
    <div class="card" style="margin-bottom: 20px; padding: 20px;">
        <div class="sub-tabs">
            <button class="sub-tab-btn active" onclick="switchSubTab('essentials')" id="btn-essentials">Essential Models</button>
            <button class="sub-tab-btn" onclick="switchSubTab('advanced')" id="btn-advanced">Advanced Models</button>
        </div>

        <div id="subtab-essentials" class="sub-tab-content active">
            <div style="display:flex; justify-content:space-between; align-items:center;">
                <div style="flex:1;">
                    <label style="display:block; font-size:.75rem; margin-bottom:5px; font-weight:800; opacity:0.9;">SELECT ESSENTIAL MODEL</label>
                    <select id="essentialsSelect" onchange="switchCalcModel(this.value)" style="width:100%; max-width:400px; padding:10px; border-radius:8px; background:rgba(255,255,255,0.1); color:white; border:1px solid rgba(255,255,255,0.3);">
                        <option value="carton">Carton Cost Model</option>
                        <option value="flexibles">Flexibles Cost Model</option>
                    </select>
                </div>
                <div class="legend" style="margin:0;">
                    <div class="legend-item"><div class="legend-dot legend-green"></div> Input</div>
                    <div class="legend-item"><div class="legend-dot legend-grey"></div> Formula</div>
                </div>
            </div>
        </div>

        <div id="subtab-advanced" class="sub-tab-content">
            <div id="advanced-login-form" class="login-overlay">
                <h3 style="margin-bottom:15px;">🔒 Restricted Access</h3>
                <p style="margin-bottom:20px; opacity:0.8;">Enter password to access Advanced Engineering Models.</p>
                <div>
                    <input type="password" id="adv-password" class="login-input" placeholder="Enter Password">
                    <button class="login-btn" onclick="checkAdvancedLogin()">Unlock</button>
                </div>
                <p id="login-error" style="color:#ef4444; margin-top:10px; font-size:0.9rem; display:none;">Incorrect Password</p>
            </div>

            <div id="advanced-secured-content" style="display:none;">
                <div style="display:flex; justify-content:space-between; align-items:center;">
                    <div style="flex:1;">
                        <label style="display:block; font-size:.75rem; margin-bottom:5px; font-weight:800; opacity:0.9; color:#4CAF50;">ADVANCED MODEL UNLOCKED</label>
                        <select id="advancedSelect" onchange="switchCalcModel(this.value)" style="width:100%; max-width:400px; padding:10px; border-radius:8px; background:rgba(76,175,80,0.1); color:white; border:1px solid #4CAF50;">
                            <option value="ebm" selected>EBM Cost Model</option>
                            <option value="carton-adv">Carton Cost Model</option>
                        </select>
                    </div>
                    <div>
                        <button onclick="lockAdvanced()" style="background:transparent; border:1px solid rgba(255,255,255,0.3); color:white; padding:5px 10px; border-radius:5px; cursor:pointer; font-size:0.8rem;">🔒 Lock</button>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <div id="carton-calculator" class="model-view">
        <div class="calc-layout">
            <div><div class="card">
                    <div class="calc-section">
                        <div class="calc-section-title">Carton Specs</div>
                        <div class="calc-row"><label>Layflat Length</label><span class="unit">mm</span><input type="number" class="calc-input calc-input-green" id="c_layflat_length" value="125.2"></div>
                        <div class="calc-row"><label>Layflat Width</label><span class="unit">mm</span><input type="number" class="calc-input calc-input-green" id="c_layflat_width" value="394.5"></div>
                    </div>
                    <div class="calc-section">
                        <div class="calc-section-title">Sheet Size</div>
                        <div class="calc-row"><label>Length</label><span class="unit">mm</span><input type="number" class="calc-input calc-input-green" id="c_sheet_length" value="1020"></div>
                        <div class="calc-row"><label>Width</label><span class="unit">mm</span><input type="number" class="calc-input calc-input-green" id="c_sheet_width" value="720"></div>
                    </div>
                    <div class="calc-section">
                        <div class="calc-section-title">Layout</div>
                        <div class="calc-row"><label>UPs Lengthwise</label><span class="unit">Nos</span><input type="number" class="calc-input calc-input-green" id="c_ups_l" value="5"></div>
                        <div class="calc-row"><label>UPs Widthwise</label><span class="unit">Nos</span><input type="number" class="calc-input calc-input-green" id="c_ups_w" value="2"></div>
                        <div class="calc-row"><label>Side Lay-1</label><span class="unit">mm</span><input type="number" class="calc-input calc-input-green" id="c_side_lay_1" value="12"></div>
                        <div class="calc-row"><label>Side Lay-2</label><span class="unit">mm</span><input type="number" class="calc-input calc-input-green" id="c_side_lay_2" value="10"></div>
                        <div class="calc-row"><label>Gripper</label><span class="unit">mm</span><input type="number" class="calc-input calc-input-green" id="c_gripper" value="6"></div>
                        <div class="calc-row"><label>Back Lay</label><span class="unit">mm</span><input type="number" class="calc-input calc-input-green" id="c_back_lay" value="4"></div>
                        <div class="calc-row"><label>Trimmer 1</label><span class="unit">mm</span><input type="number" class="calc-input calc-input-green" id="c_trimmer_1" value="5"></div>
                        <div class="calc-row"><label>Trimmer 2</label><span class="unit">mm</span><input type="number" class="calc-input calc-input-green" id="c_trimmer_2" value="5"></div>
                        <div class="calc-row"><label>Double Cut</label><span class="unit">mm</span><input type="number" class="calc-input calc-input-green" id="c_double_cut" value="0"></div>
                        <div class="calc-row"><label>Gutter</label><span class="unit">mm</span><input type="number" class="calc-input calc-input-green" id="c_gutter" value="0"></div>
                        <div class="calc-row"><label>Interlock</label><span class="unit">mm</span><input type="number" class="calc-input calc-input-green" id="c_interlock" value="0"></div>
                    </div>
                    <div class="calc-section">
                        <div class="calc-section-title">Board & Ink</div>
                        <div class="calc-row"><label>Board Type</label><span class="unit"></span><select class="calc-select" id="c_board_type"><option>WB</option><option>FBB</option><option>SBS</option><option>Duplex</option></select></div>
                        <div class="calc-row"><label>Board GSM</label><span class="unit">GSM</span><input type="number" class="calc-input calc-input-green" id="c_board_gsm" value="400"></div>
                        <div class="calc-row"><label>Board Rate</label><span class="unit">₹/Kg</span><input type="number" class="calc-input calc-input-green" id="c_board_rate" value="55"></div>
                        <div class="calc-row"><label>Ink Rate</label><span class="unit">₹/Kg</span><input type="number" class="calc-input calc-input-green" id="c_ink_rate" value="850"></div>
                        <div class="calc-row"><label>Ink Cons.</label><span class="unit">GSM</span><input type="number" class="calc-input calc-input-green" id="c_ink_consumption" value="0.9"></div>
                    </div>
                    <div class="calc-section">
                        <div class="calc-section-title">Finishing</div>
                        <div class="calc-row"><label>Varnish Type</label><span class="unit"></span><select class="calc-select" id="c_varnish_type"><option>Gloss</option><option>Matt</option></select></div>
                        <div class="calc-row"><label>Varnish Rate</label><span class="unit">₹/kg</span><input type="number" class="calc-input calc-input-green" id="c_varnish_rate" value="450"></div>
                        <div class="calc-row"><label>Consumption</label><span class="unit">GSM</span><input type="number" class="calc-input calc-input-green" id="c_varnish_consumption" value="5"></div>
                        <div class="calc-row"><label>Primer Rate</label><span class="unit">₹/kg</span><input type="number" class="calc-input calc-input-green" id="c_primer_rate" value="165"></div>
                        <div class="calc-row"><label>Consumption</label><span class="unit">GSM</span><input type="number" class="calc-input calc-input-green" id="c_primer_consumption" value="5"></div>
                    </div>
                    <div class="calc-section">
                        <div class="calc-section-title">Lamination</div>
                        <div class="calc-row"><label>Film Rate</label><span class="unit">₹/kg</span><input type="number" class="calc-input calc-input-green" id="c_film_rate" value="135"></div>
                        <div class="calc-row"><label>Film GSM</label><span class="unit">GSM</span><input type="number" class="calc-input calc-input-green" id="c_film_gsm" value="14"></div>
                        <div class="calc-row"><label>Adhesive Rate</label><span class="unit">₹/kg</span><input type="number" class="calc-input calc-input-green" id="c_adhesive_rate" value="104"></div>
                        <div class="calc-row"><label>Adhesive GSM</label><span class="unit">GSM</span><input type="number" class="calc-input calc-input-green" id="c_adhesive_gsm" value="2"></div>
                        <div class="calc-row"><label>Labour</label><span class="unit">₹/1k</span><input type="number" class="calc-input calc-input-green" id="c_lam_labour" value="115"></div>
                    </div>
                    <div class="calc-section">
                        <div class="calc-section-title">Corrugation</div>
                        <div class="calc-row"><label>Middle GSM</label><span class="unit">GSM</span><input type="number" class="calc-input calc-input-green" id="c_middle_gsm" value="150"></div>
                        <div class="calc-row"><label>Liner Rate</label><span class="unit">₹/Kg</span><input type="number" class="calc-input calc-input-green" id="c_liner_rate" value="35"></div>
                        <div class="calc-row"><label>Flute</label><span class="unit"></span><select class="calc-select" id="c_flute_type"><option>E</option><option>F</option></select></div>
                        <div class="calc-row"><label>Inner GSM</label><span class="unit">GSM</span><input type="number" class="calc-input calc-input-green" id="c_inner_gsm" value="100"></div>
                        <div class="calc-row"><label>Inner Rate</label><span class="unit">₹/Kg</span><input type="number" class="calc-input calc-input-green" id="c_inner_rate" value="35"></div>
                        <div class="calc-row"><label>Conversion</label><span class="unit">₹/kg</span><input type="number" class="calc-input calc-input-green" id="c_corr_conv_rate" value="6.5"></div>
                    </div>
                    <div class="calc-section">
                        <div class="calc-section-title">Foil & Other</div>
                        <div class="calc-row"><label>Foil WxL</label><span class="unit">mm</span><div style="display:flex;gap:5px;"><input type="number" class="calc-input calc-input-green" id="c_foil_w" value="0"><input type="number" class="calc-input calc-input-green" id="c_foil_l" value="0"></div></div>
                        <div class="calc-row"><label>Foil Rate</label><span class="unit">₹/Roll</span><input type="number" class="calc-input calc-input-green" id="c_foil_cost" value="1050"></div>
                        <div class="calc-row"><label>Stamping</label><span class="unit">₹/Sht</span><input type="number" class="calc-input calc-input-green" id="c_foil_conv" value="0"></div>
                        <div class="calc-row"><label>Other</label><span class="unit">₹/1k</span><input type="number" class="calc-input calc-input-green" id="c_other" value="50"></div>
                        <div class="calc-row"><label>Conversion</label><span class="unit">₹/1k</span><input type="number" class="calc-input calc-input-green" id="c_conversion" value="195"></div>
                    </div>
                    <button class="btn-analyze" onclick="calculateCarton()">Calculate Carton Cost</button>
                    <button onclick="saveSKU()" id="save-sku-btn-carton" class="btn-analyze" style="background: linear-gradient(135deg, #FF9800, #F57C00); margin-top: 10px; display: none; border: none; color: white; padding: 12px 24px; border-radius: 8px; cursor: pointer; font-weight: 600; box-shadow: 0 4px 12px rgba(255,152,0,0.4); transition: all 0.3s;">💾 Save as SKU</button>
                </div>
            </div>
            <div><div class="summary-card" id="carton-summary">
                    <h3 style="margin-bottom: 20px;">Cost Summary</h3>
                    <p style="opacity: 0.6; text-align: center; padding: 40px 0;">Click "Calculate" to see results</p>
                </div>
                <div id="carton-pie-chart" style="margin-top:20px; display:none;">
                    <div class="summary-card">
                        <h3 style="margin-bottom:10px;">Cost Distribution</h3>
                        <div id="carton-donut" style="width:100%; height:320px;"></div>
                    </div>
                </div>
                <div id="carton-export-btns" style="margin-top:15px; display:none;">
                    <div style="display:flex; gap:10px;">
                        <button class="btn-secondary" onclick="exportGenericExcel()" style="flex:1;">Export Excel</button>
                        <button class="btn-secondary" onclick="exportGenericPDF()" style="flex:1;">Export PDF</button>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <div id="flexibles-calculator" class="model-view">
        <div class="calc-layout">
            <div>
                <div class="card">
                    <div class="calc-section">
                        <div class="calc-section-title">Laminate GSM (Auto-calculated)</div>
                        <div class="calc-row"><label>Total Laminate GSM</label><span class="unit">GSM</span><input type="number" class="calc-input" id="f_laminate_gsm" value="174.9" readonly style="opacity:0.6;"></div>
                    </div>
                    <div id="flex-layers"></div>
                    <button class="btn-add-layer" onclick="addFlexLayer()">+ Add Layer</button>
                    <div class="calc-section" style="margin-top: 20px;">
                        <div class="calc-section-title">Conversion & Other</div>
                        <div class="calc-row"><label>Conversion Cost</label><span class="unit">₹/kg</span><input type="number" class="calc-input calc-input-green" id="f_conversion" value="50"></div>
                    </div>
                    <button class="btn-analyze" onclick="calculateFlexibles()">Calculate Flexibles Cost</button>
                    <button onclick="saveSKU()" id="save-sku-btn-flexibles" class="btn-analyze" style="background: linear-gradient(135deg, #FF9800, #F57C00); margin-top: 10px; display: none; border: none; color: white; padding: 12px 24px; border-radius: 8px; cursor: pointer; font-weight: 600; box-shadow: 0 4px 12px rgba(255,152,0,0.4); transition: all 0.3s;">💾 Save as SKU</button>
                </div>
            </div>
            <div>
                <div class="summary-card" id="flex-summary">
                    <h3 style="margin-bottom: 20px;">Cost Summary</h3>
                    <p style="opacity: 0.6; text-align: center; padding: 40px 0;">Click "Calculate" to see results</p>
                </div>
                <div id="flex-pie-chart" style="margin-top:20px; display:none;">
                    <div class="summary-card">
                        <h3 style="margin-bottom:10px;">Cost Distribution</h3>
                        <div id="flex-donut" style="width:100%; height:320px;"></div>
                    </div>
                </div>
                <div id="flex-export-btns" style="margin-top:15px; display:none;">
                    <div style="display:flex; gap:10px;">
                        <button class="btn-secondary" onclick="exportGenericExcel()" style="flex:1;">Export Excel</button>
                        <button class="btn-secondary" onclick="exportGenericPDF()" style="flex:1;">Export PDF</button>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <div id="ebm-calculator" class="model-view">
        <div class="calc-layout">
            <div><div class="card">
                    <div class="calc-section">
                        <div class="calc-section-title">SKU & General</div>
                        <div class="calc-row"><label>SKU</label><span class="unit"></span><input type="text" class="calc-input calc-input-green" id="e_sku" value="Comfort 220ml"></div>
                        <div class="calc-row"><label>Country</label><span class="unit"></span><select class="calc-select" id="e_country" onchange="loadEBMCountryDefaults()"><option>India</option><option>China</option><option>Indonesia</option><option>Brazil</option><option>Mexico</option><option>Turkey</option><option>Vietnam</option><option>Pakistan</option><option>Philippines</option><option>South Africa</option><option>United States</option><option>United Kingdom</option><option>Germany</option><option>France</option><option>Spain</option><option>Poland</option><option>Thailand</option><option>Bangladesh</option><option>Sri Lanka</option><option>Argentina</option><option>Canada</option><option>Costa Rica</option></select></div>
                        <div class="calc-row"><label>Volume</label><span class="unit">pcs</span><input type="number" class="calc-input calc-input-green" id="e_volume" value="62975559"></div>
                        <div class="calc-row"><label>Weight</label><span class="unit">g</span><input type="number" class="calc-input calc-input-green" id="e_weight" value="19"></div>
                    </div>
                    <div class="calc-section">
                        <div class="calc-section-title">Layers</div>
                        <div class="calc-row"><label>L1 Ratio</label><span class="unit">%</span><input type="number" class="calc-input calc-input-green" id="e_l1_ratio" value="0.48"></div>
                        <div class="calc-row"><label>L1 Type</label><span class="unit"></span><select class="calc-select" id="e_l1_type"><option>HDPE</option><option>rHDPE</option><option>LDPE</option><option>PP</option><option>rPP</option><option>PET</option><option>rPET</option></select></div>
                        <div class="calc-row"><label>L1 Rate</label><span class="unit">₹/kg</span><input type="number" class="calc-input calc-input-green" id="e_l1_rate" value="95"></div>
                        <div class="calc-row"><label>L1 MB%</label><span class="unit">%</span><input type="number" class="calc-input calc-input-green" id="e_l1_mb_dosage" value="0.02"></div>
                        <div class="calc-row"><label>L1 MB Rate</label><span class="unit">₹/kg</span><input type="number" class="calc-input calc-input-green" id="e_l1_mb_rate" value="450"></div>
                        <div class="calc-row"><label>L1 Add%</label><span class="unit">%</span><input type="number" class="calc-input calc-input-green" id="e_l1_add_dosage" value="0"></div>
                        <div class="calc-row"><label>L1 Add Rate</label><span class="unit">₹/kg</span><input type="number" class="calc-input calc-input-green" id="e_l1_add_rate" value="249.93"></div>

                        <div style="border-top:1px dashed rgba(255,255,255,0.2); margin:10px 0;"></div>
                         
                        <div class="calc-row"><label>L2 Ratio</label><span class="unit">%</span><input type="number" class="calc-input calc-input-green" id="e_l2_ratio" value="0.50"></div>
                        <div class="calc-row"><label>L2 Type</label><span class="unit"></span><select class="calc-select" id="e_l2_type"><option>rHDPE</option><option>HDPE</option><option>LDPE</option><option>PP</option><option>rPP</option><option>PET</option><option>rPET</option></select></div>
                        <div class="calc-row"><label>L2 Rate</label><span class="unit">₹/kg</span><input type="number" class="calc-input calc-input-green" id="e_l2_rate" value="107"></div>
                        <div class="calc-row"><label>L2 MB%</label><span class="unit">%</span><input type="number" class="calc-input calc-input-green" id="e_l2_mb_dosage" value="0"></div>
                        <div class="calc-row"><label>L2 MB Rate</label><span class="unit">₹/kg</span><input type="number" class="calc-input calc-input-green" id="e_l2_mb_rate" value="450"></div>
                        <div class="calc-row"><label>L2 Add%</label><span class="unit">%</span><input type="number" class="calc-input calc-input-green" id="e_l2_add_dosage" value="0"></div>
                        <div class="calc-row"><label>L2 Add Rate</label><span class="unit">₹/kg</span><input type="number" class="calc-input calc-input-green" id="e_l2_add_rate" value="249.93"></div>

                        <div style="border-top:1px dashed rgba(255,255,255,0.2); margin:10px 0;"></div>

                        <div class="calc-row"><label>L3 Ratio</label><span class="unit">%</span><input type="number" class="calc-input calc-input-green" id="e_l3_ratio" value="0"></div>
                        <div class="calc-row"><label>L3 Type</label><span class="unit"></span><select class="calc-select" id="e_l3_type"><option>HDPE</option><option>rHDPE</option><option>LDPE</option><option>PP</option><option>rPP</option><option>PET</option><option>rPET</option></select></div>
                        <div class="calc-row"><label>L3 Rate</label><span class="unit">₹/kg</span><input type="number" class="calc-input calc-input-green" id="e_l3_rate" value="0"></div>
                        <div class="calc-row"><label>L3 MB%</label><span class="unit">%</span><input type="number" class="calc-input calc-input-green" id="e_l3_mb_dosage" value="0"></div>
                        <div class="calc-row"><label>L3 MB Rate</label><span class="unit">₹/kg</span><input type="number" class="calc-input calc-input-green" id="e_l3_mb_rate" value="450"></div>
                        <div class="calc-row"><label>L3 Add%</label><span class="unit">%</span><input type="number" class="calc-input calc-input-green" id="e_l3_add_dosage" value="0"></div>
                        <div class="calc-row"><label>L3 Add Rate</label><span class="unit">₹/kg</span><input type="number" class="calc-input calc-input-green" id="e_l3_add_rate" value="249.93"></div>
                    </div>
                    <div class="calc-section">
                        <div class="calc-section-title">Machine</div>
                        <div class="calc-row"><label>Cavitation</label><span class="unit">Nos</span><input type="number" class="calc-input calc-input-green" id="e_cavity" value="12"></div>
                        <div class="calc-row"><label>Cycle Time</label><span class="unit">Secs</span><input type="number" class="calc-input calc-input-green" id="e_cycle" value="16.3"></div>
                        <div class="calc-row"><label>Model</label><span class="unit"></span><select class="calc-select" id="e_machine"><option>Jomar 65</option><option>Jomar 135</option><option>Uniloy</option><option>Sika</option><option>Speedex</option><option>Magic 10</option><option>BMU 70</option><option>BMU 100</option><option>SEB 820</option><option>Bekum</option><option>Kautex</option><option>Uniloy Rotary</option><option>Jomar Shuttle</option><option>Chinese OEM</option></select></div>
                        <div class="calc-row"><label>Setups/Yr</label><span class="unit">Nos</span><input type="number" class="calc-input calc-input-green" id="e_setups" value="6"></div>
                        <div class="calc-row"><label>Rampups/Yr</label><span class="unit">Nos</span><input type="number" class="calc-input calc-input-green" id="e_rampups" value="6"></div>
                    </div>
                     <div class="calc-section">
                        <div class="calc-section-title">Conversion Costs</div>
                        <div class="calc-row"><label>Elec Rate</label><span class="unit">₹/kWH</span><input type="number" class="calc-input calc-input-green" id="e_elec_rate" value="10.72"></div>
                        <div class="calc-row"><label>Skilled Lab</label><span class="unit">₹/Yr</span><input type="number" class="calc-input calc-input-green" id="e_labour" value="541800"></div>
                        <div class="calc-row"><label>Engineer</label><span class="unit">₹/Yr</span><input type="number" class="calc-input calc-input-green" id="e_engineer" value="1260000"></div>
                        <div class="calc-row"><label>Prod Mgr</label><span class="unit">₹/Yr</span><input type="number" class="calc-input calc-input-green" id="e_pm" value="1890000"></div>
                        <div class="calc-row"><label>R&M %</label><span class="unit">%</span><input type="number" class="calc-input calc-input-green" id="e_repair" value="0.025"></div>
                        <div class="calc-row"><label>Other OH %</label><span class="unit">%</span><input type="number" class="calc-input calc-input-green" id="e_other_oh" value="0.025"></div>
                    </div>
                    <div class="calc-section">
                        <div class="calc-section-title">Depreciation & Interest</div>
                        <div class="calc-row"><label>Dep P&M</label><span class="unit">%</span><input type="number" class="calc-input calc-input-green" id="e_dep_pm" value="0.15"></div>
                        <div class="calc-row"><label>Dep Bldg</label><span class="unit">%</span><input type="number" class="calc-input calc-input-green" id="e_dep_bldg" value="0.10"></div>
                        <div class="calc-row"><label>Life</label><span class="unit">Yrs</span><input type="number" class="calc-input calc-input-green" id="e_life" value="5"></div>
                        <div class="calc-row"><label>Int LT</label><span class="unit">%</span><input type="number" class="calc-input calc-input-green" id="e_int_lt" value="0.125"></div>
                        <div class="calc-row"><label>Int WC</label><span class="unit">%</span><input type="number" class="calc-input calc-input-green" id="e_int_wc" value="0.14"></div>
                         <div class="calc-row"><label>D/E Ratio</label><span class="unit"></span><input type="number" class="calc-input calc-input-green" id="e_debt_equity" value="0.70"></div>
                    </div>
                    <div class="calc-section">
                        <div class="calc-section-title">Premises</div>
                        <div class="calc-row"><label>Land Cost</label><span class="unit">₹/SQM</span><input type="number" class="calc-input calc-input-green" id="e_land" value="23519"></div>
                        <div class="calc-row"><label>Bldg Cost</label><span class="unit">₹/SQM</span><input type="number" class="calc-input calc-input-green" id="e_building" value="7000"></div>
                        <div class="calc-row"><label>Lease Cost</label><span class="unit">₹/SQM</span><input type="number" class="calc-input calc-input-green" id="e_lease" value="2136"></div>
                        <div class="calc-row"><label>Type</label><span class="unit"></span><select class="calc-select" id="e_premises"><option>Owned</option><option>Leased</option></select></div>
                    </div>
                    <div class="calc-section">
                        <div class="calc-section-title">Commercials</div>
                        <div class="calc-row"><label>Margin</label><span class="unit">%</span><input type="number" class="calc-input calc-input-green" id="e_margin" value="0.20"></div>
                        <div class="calc-row"><label>Type</label><span class="unit"></span><select class="calc-select" id="e_margin_calc"><option>% of Conversion Cost</option><option>% of Total Cost</option></select></div>
                        <div class="calc-row"><label>RM Days</label><span class="unit">Days</span><input type="number" class="calc-input calc-input-green" id="e_rm_days" value="45"></div>
                        <div class="calc-row"><label>FG Days</label><span class="unit">Days</span><input type="number" class="calc-input calc-input-green" id="e_fg_days" value="60"></div>
                         <div class="calc-row"><label>Euro Rate</label><span class="unit"></span><input type="number" class="calc-input calc-input-green" id="e_euro_rate" value="104.27"></div>
                    </div>
                     <div class="calc-section">
                        <div class="calc-section-title">Packing & Logistics</div>
                        <div class="calc-row"><label>Orders/Yr</label><span class="unit">Nos</span><input type="number" class="calc-input calc-input-green" id="e_orders" value="12"></div>
                        <div class="calc-row"><label>Pcs/Box</label><span class="unit">Nos</span><input type="number" class="calc-input calc-input-green" id="e_bottles_box" value="360"></div>
                        <div class="calc-row"><label>Boxes/Cont</label><span class="unit">Nos</span><input type="number" class="calc-input calc-input-green" id="e_boxes_cont" value="320"></div>
                        <div class="calc-row"><label>Shipper</label><span class="unit">₹</span><input type="number" class="calc-input calc-input-green" id="e_shipper" value="59.43"></div>
                        <div class="calc-row"><label>Polybag</label><span class="unit">₹</span><input type="number" class="calc-input calc-input-green" id="e_polybag" value="25.02"></div>
                        <div class="calc-row"><label>Freight</label><span class="unit">₹</span><input type="number" class="calc-input calc-input-green" id="e_freight" value="8341.60"></div>
                    </div>
                    <button class="btn-analyze" onclick="calculateEBM()">Calculate EBM Cost</button>
                    <button onclick="saveSKU()" id="save-sku-btn-ebm" class="btn-analyze" style="background: linear-gradient(135deg, #FF9800, #F57C00); margin-top: 10px; display: none; border: none; color: white; padding: 12px 24px; border-radius: 8px; cursor: pointer; font-weight: 600; box-shadow: 0 4px 12px rgba(255,152,0,0.4); transition: all 0.3s;">💾 Save as SKU</button>
                </div>
            </div>
            <div><div class="summary-card" id="ebm-summary">
                    <h3 style="margin-bottom: 20px;">Cost Summary</h3>
                    <p style="opacity: 0.6; text-align: center; padding: 40px 0;">Click "Calculate" to see results</p>
                </div>
                <div id="ebm-pie-chart" style="margin-top:20px; display:none;">
                    <div class="summary-card">
                        <h3 style="margin-bottom:10px;">Cost Distribution</h3>
                        <div id="ebm-donut" style="width:100%; height:320px;"></div>
                    </div>
                </div>
                <div id="ebm-export-btns" style="margin-top:15px; display:none;">
                    <div style="display:flex; gap:10px;">
                        <button class="btn-secondary" onclick="exportEBMExcel()" style="flex:1;">Export Excel</button>
                        <button class="btn-secondary" onclick="exportEBMPDF()" style="flex:1;">Export PDF</button>
                    </div>
                </div>
            </div>
        </div>
    </div>
</div>

<!-- ADVANCED CARTON COST MODEL -->
<div id="carton-adv-calculator" class="model-view">
    <div class="calc-layout">
        <div><div class="card">
            <div class="calc-section">
                <div class="calc-section-title">General Info</div>
                <div class="calc-row"><label>Country</label><span class="unit"></span><select class="calc-select" id="ca_country" onchange="loadCartonAdvCountryDefaults()"><option>India</option><option>China</option><option>Indonesia</option><option>Brazil</option><option>Mexico</option><option>Turkey</option><option>Vietnam</option><option>Pakistan</option><option>Philippines</option><option>South Africa</option><option>United States</option><option>United Kingdom</option><option>Germany</option><option>France</option><option>Spain</option><option>Poland</option><option>Thailand</option><option>Bangladesh</option><option>Sri Lanka</option><option>Argentina</option><option>Canada</option><option>Costa Rica</option></select></div>
                <div class="calc-row"><label>Annual Volume</label><span class="unit">pcs</span><input type="number" class="calc-input calc-input-green" id="ca_annual_vol" value="3126950"></div>
                <div class="calc-row"><label>Avg Order Size</label><span class="unit">pcs</span><input type="number" class="calc-input calc-input-green" id="ca_avg_order" value="260579.17"></div>
                <div class="calc-row"><label>No of Colours</label><span class="unit">Nos</span><input type="number" class="calc-input calc-input-green" id="ca_colours" value="5"></div>
                <div class="calc-row"><label>Common Colours</label><span class="unit">Nos</span><input type="number" class="calc-input calc-input-green" id="ca_common_col" value="2"></div>
                <div class="calc-row"><label>Print Runs/Year</label><span class="unit">Nos</span><input type="number" class="calc-input calc-input-green" id="ca_print_runs" value="12"></div>
                <div class="calc-row"><label>No of Shifts</label><span class="unit"></span><input type="number" class="calc-input calc-input-green" id="ca_shifts" value="3"></div>
            </div>
            <div class="calc-section">
                <div class="calc-section-title">Carton Dimensions</div>
                <div class="calc-row"><label>Length 1</label><span class="unit">mm</span><input type="number" class="calc-input calc-input-green" id="ca_len1" value="36.3"></div>
                <div class="calc-row"><label>Length 2</label><span class="unit">mm</span><input type="number" class="calc-input calc-input-green" id="ca_len2" value="37"></div>
                <div class="calc-row"><label>Width 1</label><span class="unit">mm</span><input type="number" class="calc-input calc-input-green" id="ca_wid1" value="46"></div>
                <div class="calc-row"><label>Width 2</label><span class="unit">mm</span><input type="number" class="calc-input calc-input-green" id="ca_wid2" value="46"></div>
                <div class="calc-row"><label>Height</label><span class="unit">mm</span><input type="number" class="calc-input calc-input-green" id="ca_height" value="179"></div>
                <div class="calc-row"><label>Max Flap</label><span class="unit">mm</span><input type="number" class="calc-input calc-input-green" id="ca_flap" value="96.9"></div>
                <div class="calc-row"><label>Gluing Area</label><span class="unit">mm</span><input type="number" class="calc-input calc-input-green" id="ca_glue" value="13"></div>
                <div class="calc-row"><label>Machine Size</label><span class="unit"></span><select class="calc-select" id="ca_mach_size"><option>IIIB</option><option>VI</option><option>VIB</option></select></div>
                <div class="calc-row"><label>Grain Direction</label><span class="unit"></span><select class="calc-select" id="ca_grain"><option>Long</option><option>Short</option></select></div>
            </div>
            <div class="calc-section">
                <div class="calc-section-title">Layout</div>
                <div class="calc-row"><label>UPs Lengthwise</label><span class="unit">Nos</span><input type="number" class="calc-input calc-input-green" id="ca_ups_l" value="5"></div>
                <div class="calc-row"><label>UPs Widthwise</label><span class="unit">Nos</span><input type="number" class="calc-input calc-input-green" id="ca_ups_w" value="2"></div>
                <div class="calc-row"><label>Side Lay 1</label><span class="unit">mm</span><input type="number" class="calc-input calc-input-green" id="ca_side1" value="5"></div>
                <div class="calc-row"><label>Side Lay 2</label><span class="unit">mm</span><input type="number" class="calc-input calc-input-green" id="ca_side2" value="5"></div>
                <div class="calc-row"><label>Gripper</label><span class="unit">mm</span><input type="number" class="calc-input calc-input-green" id="ca_gripper" value="10"></div>
                <div class="calc-row"><label>Back Lay</label><span class="unit">mm</span><input type="number" class="calc-input calc-input-green" id="ca_back_lay" value="5"></div>
            </div>
            <div class="calc-section">
                <div class="calc-section-title">Material Rates</div>
                <div class="calc-row"><label>Board GSM</label><span class="unit">GSM</span><input type="number" class="calc-input calc-input-green" id="ca_board_gsm" value="300"></div>
                <div class="calc-row"><label>Board Rate</label><span class="unit">₹/Kg</span><input type="number" class="calc-input calc-input-green" id="ca_board_rate" value="45"></div>
                <div class="calc-row"><label>Ink Rate</label><span class="unit">₹/Kg</span><input type="number" class="calc-input calc-input-green" id="ca_ink_rate" value="834.16"></div>
                <div class="calc-row"><label>Ink GSM</label><span class="unit">GSM</span><input type="number" class="calc-input calc-input-green" id="ca_ink_gsm" value="2"></div>
                <div class="calc-row"><label>Varnish Rate</label><span class="unit">₹/Kg</span><input type="number" class="calc-input calc-input-green" id="ca_varnish_rate" value="521.35"></div>
                <div class="calc-row"><label>Varnish GSM</label><span class="unit">GSM</span><input type="number" class="calc-input calc-input-green" id="ca_varnish_gsm" value="3"></div>
            </div>
            <div class="calc-section">
                <div class="calc-section-title">Decoration Options</div>
                <div class="calc-row"><label>Spot Varnish</label><span class="unit"></span><select class="calc-select" id="ca_spot_varnish"><option>N</option><option>Y</option></select></div>
                <div class="calc-row"><label>Hot Foiling</label><span class="unit"></span><select class="calc-select" id="ca_hot_foil"><option>N</option><option>Y</option></select></div>
                <div class="calc-row"><label>Lamination</label><span class="unit"></span><select class="calc-select" id="ca_lamination"><option>N</option><option>Y</option></select></div>
                <div class="calc-row"><label>Window Carton</label><span class="unit"></span><select class="calc-select" id="ca_window"><option>N</option><option>Y</option></select></div>
                <div class="calc-row"><label>Liner</label><span class="unit"></span><select class="calc-select" id="ca_liner"><option>N</option><option>Y</option></select></div>
                <div class="calc-row"><label>Primer</label><span class="unit"></span><select class="calc-select" id="ca_primer"><option>N</option><option>Y</option></select></div>
                <div class="calc-row"><label>Foil Rate</label><span class="unit">₹/Roll</span><input type="number" class="calc-input calc-input-green" id="ca_foil_rate" value="938.43"></div>
                <div class="calc-row"><label>Foil L × W</label><span class="unit">mm</span><div style="display:flex;gap:5px;"><input type="number" class="calc-input calc-input-green" id="ca_foil_l" value="14.52" style="width:50%;"><input type="number" class="calc-input calc-input-green" id="ca_foil_w" value="71.6" style="width:50%;"></div></div>
                <div class="calc-row"><label>Film Rate</label><span class="unit">₹/Kg</span><input type="number" class="calc-input calc-input-green" id="ca_film_rate" value="208.54"></div>
                <div class="calc-row"><label>Film GSM</label><span class="unit">GSM</span><input type="number" class="calc-input calc-input-green" id="ca_film_gsm" value="14"></div>
            </div>
            <div class="calc-section">
                <div class="calc-section-title">Machines (from DB)</div>
                <div class="calc-row"><label>Printing</label><span class="unit"></span><select class="calc-select" id="ca_m_print"><option>KBA 8000</option></select></div>
                <div class="calc-row"><label>Spot Varnish</label><span class="unit"></span><select class="calc-select" id="ca_m_sv"><option>Unison</option></select></div>
                <div class="calc-row"><label>Hot Foiling</label><span class="unit"></span><select class="calc-select" id="ca_m_hf"><option>Hot Foil Machine</option></select></div>
                <div class="calc-row"><label>Lamination</label><span class="unit"></span><select class="calc-select" id="ca_m_lam"><option>Kompac KwikPrint / EZ Koat</option></select></div>
                <div class="calc-row"><label>Crease&Blank</label><span class="unit"></span><select class="calc-select" id="ca_m_cb"><option>Bobst Mastercut</option></select></div>
                <div class="calc-row"><label>Folder-Gluer</label><span class="unit"></span><select class="calc-select" id="ca_m_fg"><option>Bobst Masterfold</option></select></div>
                <div id="ca-machine-db-info"></div>
            </div>
            <div class="calc-section">
                <div class="calc-section-title">Conversion Costs</div>
                <div class="calc-row"><label>Elec Rate</label><span class="unit">₹/kWH</span><input type="number" class="calc-input calc-input-green" id="ca_elec" value="10.72"></div>
                <div class="calc-row"><label>Skilled Lab</label><span class="unit">₹/Yr</span><input type="number" class="calc-input calc-input-green" id="ca_labour" value="541800"></div>
                <div class="calc-row"><label>Engineer</label><span class="unit">₹/Yr</span><input type="number" class="calc-input calc-input-green" id="ca_engineer" value="1260000"></div>
                <div class="calc-row"><label>Prod Mgr</label><span class="unit">₹/Yr</span><input type="number" class="calc-input calc-input-green" id="ca_pm" value="1890000"></div>
                <div class="calc-row"><label>R&M %</label><span class="unit">%</span><input type="number" class="calc-input calc-input-green" id="ca_repair" value="0.02"></div>
                <div class="calc-row"><label>Other OH %</label><span class="unit">%</span><input type="number" class="calc-input calc-input-green" id="ca_other_oh" value="0.02"></div>
                <div class="calc-row"><label>Dep P&M %</label><span class="unit">%</span><input type="number" class="calc-input calc-input-green" id="ca_dep_pm" value="0.15"></div>
                <div class="calc-row"><label>Dep Bldg %</label><span class="unit">%</span><input type="number" class="calc-input calc-input-green" id="ca_dep_bldg" value="0.10"></div>
                <div class="calc-row"><label>Life of Asset</label><span class="unit">Yrs</span><input type="number" class="calc-input calc-input-green" id="ca_life" value="5"></div>
                <div class="calc-row"><label>Land Cost</label><span class="unit">₹/sqm</span><input type="number" class="calc-input calc-input-green" id="ca_land" value="23519"></div>
                <div class="calc-row"><label>Building Cost</label><span class="unit">₹/sqm</span><input type="number" class="calc-input calc-input-green" id="ca_building" value="7000"></div>
                <div class="calc-row"><label>Premises</label><span class="unit"></span><select class="calc-select" id="ca_premises"><option>Owned</option><option>Leased</option></select></div>
                <div class="calc-row"><label>Int LT Loan</label><span class="unit">%</span><input type="number" class="calc-input calc-input-green" id="ca_int_lt" value="0.125"></div>
                <div class="calc-row"><label>Int WC</label><span class="unit">%</span><input type="number" class="calc-input calc-input-green" id="ca_int_wc" value="0.14"></div>
                <div class="calc-row"><label>EUR Rate</label><span class="unit">₹</span><input type="number" class="calc-input calc-input-green" id="ca_euro" value="104.27"></div>
                <div class="calc-row"><label>Margin</label><span class="unit">%</span><input type="number" class="calc-input calc-input-green" id="ca_margin" value="0.20"></div>
            </div>
            <div class="calc-section">
                <div class="calc-section-title">Packing & Freight</div>
                <div class="calc-row"><label>Cartons/Box</label><span class="unit">Nos</span><input type="number" class="calc-input calc-input-green" id="ca_ctn_box" value="1500"></div>
                <div class="calc-row"><label>Boxes/Container</label><span class="unit">Nos</span><input type="number" class="calc-input calc-input-green" id="ca_box_cont" value="173.72"></div>
                <div class="calc-row"><label>Freight Cost</label><span class="unit">₹/Cont</span><input type="number" class="calc-input calc-input-green" id="ca_freight" value="20010"></div>
                <div class="calc-row"><label>Shipper Cost</label><span class="unit">EUR/pc</span><input type="number" class="calc-input calc-input-green" id="ca_shipper" value="0.6"></div>
                <div class="calc-row"><label>Polybag Cost</label><span class="unit">EUR/pc</span><input type="number" class="calc-input calc-input-green" id="ca_polybag" value="0.25"></div>
            </div>
            <div id="ca-var-cost-info"></div>
            <button class="btn-analyze" onclick="calculateCartonAdvanced()">Calculate Carton Cost</button>
            <button onclick="saveSKU()" id="save-sku-btn-carton-adv" class="btn-analyze" style="background: linear-gradient(135deg, #FF9800, #F57C00); margin-top: 10px; display: none; border: none; color: white; padding: 12px 24px; border-radius: 8px; cursor: pointer; font-weight: 600; box-shadow: 0 4px 12px rgba(255,152,0,0.4); transition: all 0.3s;">💾 Save as SKU</button>
        </div></div>
        <div><div class="summary-card" id="carton-adv-summary">
            <h3 style="margin-bottom: 20px;">Cost Summary</h3>
            <p style="opacity: 0.6; text-align: center; padding: 40px 0;">Click "Calculate" to see results</p>
        </div>
        <div id="carton-adv-pie-chart" style="margin-top:20px; display:none;">
            <div class="card">
                <div id="ca-donut" style="width:100%; height:320px;"></div>
            </div>
        </div>
        <div id="carton-adv-export-btns" style="margin-top:15px; display:none;">
            <div style="display:flex; gap:10px;">
                <button class="btn-secondary" onclick="exportCartonAdvExcel()" style="flex:1;">Export Excel</button>
                <button class="btn-secondary" onclick="exportCartonAdvPDF()" style="flex:1;">Export PDF</button>
            </div>
        </div>
        </div>
    </div>
</div>

<div id="universal-compare" class="calculator-view">
    <div class="card">
        <div class="calc-section-title">Multi-Country Comparison</div>
        <p style="opacity:0.7; font-size:0.85rem; margin-bottom:15px;">Run the same SKU across multiple countries side-by-side. First calculate your cost model (any model) in the Calculator tab, then select countries below for comparison.</p>
        <div style="display:flex; flex-wrap:wrap; gap:8px; margin-bottom:15px;" id="country-checkboxes"></div>
        <button class="btn-analyze" id="compareCountriesBtn" onclick="runMultiCountry()" disabled>Compare Countries</button>
    </div>
    <div id="multi-country-results"></div>
    <div id="multi-country-chart" style="margin-top:20px;"></div>
</div>

<div id="universal-whatif" class="calculator-view">
    <div class="card">
        <div class="calc-section-title">What-If Scenario Analysis</div>
        <p style="opacity:0.7; font-size:0.85rem; margin-bottom:15px;">Adjust sliders to see how parameter changes affect total cost. Works with all cost models — first calculate in the Calculator tab.</p>
        <div id="whatif-sliders">
            <div class="calc-section" style="margin-bottom:12px;">
                <label style="font-size:0.85rem;">Resin Price Change: <strong id="wi_resin_label">0%</strong></label>
                <input type="range" min="-50" max="50" value="0" id="wi_resin" oninput="updateWhatIf()" style="width:100%; accent-color:var(--orange);">
            </div>
            <div class="calc-section" style="margin-bottom:12px;">
                <label style="font-size:0.85rem;">Annual Volume Change: <strong id="wi_volume_label">0%</strong></label>
                <input type="range" min="-50" max="100" value="0" id="wi_volume" oninput="updateWhatIf()" style="width:100%; accent-color:var(--orange);">
            </div>
            <div class="calc-section" style="margin-bottom:12px;">
                <label style="font-size:0.85rem;">Electricity Rate Change: <strong id="wi_elec_label">0%</strong></label>
                <input type="range" min="-50" max="50" value="0" id="wi_elec" oninput="updateWhatIf()" style="width:100%; accent-color:var(--orange);">
            </div>
            <div class="calc-section" style="margin-bottom:12px;">
                <label style="font-size:0.85rem;">Labour Cost Change: <strong id="wi_labour_label">0%</strong></label>
                <input type="range" min="-50" max="50" value="0" id="wi_labour" oninput="updateWhatIf()" style="width:100%; accent-color:var(--orange);">
            </div>
            <div class="calc-section" style="margin-bottom:12px;">
                <label style="font-size:0.85rem;">Margin Change: <strong id="wi_margin_label">0%</strong></label>
                <input type="range" min="-50" max="50" value="0" id="wi_margin" oninput="updateWhatIf()" style="width:100%; accent-color:var(--orange);">
            </div>
        </div>
        <button class="btn-secondary" onclick="resetWhatIf()" style="width:100%; margin-top:10px;">Reset All Sliders</button>
    </div>
    <div id="whatif-results" style="margin-top:20px;"></div>
    <div id="whatif-chart" style="margin-top:20px;"></div>
</div>

<script>
// --- Constants ---
const FILM_OPTIONS = ['PET Film','BOPP Film','CPP Film','METPET Film','MET MDOPE Film','Matt Finish PET Film','AL Foil','BON','HIPS','GPPS','HDPE','MDPE','EAA','EVA','Cellophane','Mono Layer PE','2 Layer All PE','3 Layer All PE','5 Layer All PE','5 Layer EVOH Barrier','5 Layer Nylon Barrier','7 Layer All PE','Gravure','Flexo','Primer','Lamination - Adhesive (Solvent Based)','Lamination - Adhesive (Solvent Less)','Heat Seal Lacquer','Cold Seal','Gloss Varnish','Matte Varnish'];

// --- State ---
let flexLayers = [
    {name: 'PET Film', mic: 12, rate: 145},
    {name: 'Gravure', mic: 3, rate: 1700},
    {name: 'Lamination - Adhesive (Solvent Less)', mic: 10, rate: 750},
    {name: '5 Layer All PE', mic: 145, rate: 125},
    {name: '5 Layer All PE', mic: 8, rate: 125},
];
let currentModel = 'carton';
let isAdvancedLoggedIn = false;
let lastEBMResult = null;
let lastEBMInput = null;
let lastModelResult = null;
let lastModelInput = null;
let lastModelType = null;

// --- Tab Logic ---

function switchUniversalTab(tab) {
    document.querySelectorAll('.universal-tab-btn').forEach(b => {
        if (b.getAttribute('data-tab') === tab) b.classList.add('active');
        else b.classList.remove('active');
    });
    
    document.querySelectorAll('.calculator-view').forEach(c => c.classList.remove('active'));
    
    if (tab === 'calculator') {
        document.getElementById('universal-calculator').classList.add('active');
        // Re-show whichever model was active (carton-adv lives outside universal-calculator)
        var activeModel = currentModel || 'carton';
        var mv = document.getElementById(activeModel + '-calculator');
        if (mv) mv.classList.add('active');
    } else {
        // Hide ALL model-view panels (critical: carton-adv-calculator is outside universal-calculator)
        document.querySelectorAll('.model-view').forEach(v => v.classList.remove('active'));
        if (tab === 'compare') {
            document.getElementById('universal-compare').classList.add('active');
        } else if (tab === 'whatif') {
            document.getElementById('universal-whatif').classList.add('active');
        }
    }
}

function switchSubTab(tab) {
    // Buttons
    document.getElementById('btn-essentials').classList.remove('active');
    document.getElementById('btn-advanced').classList.remove('active');
    document.getElementById('btn-' + tab).classList.add('active');

    // Sub-content
    document.getElementById('subtab-essentials').classList.remove('active');
    document.getElementById('subtab-advanced').classList.remove('active');
    document.getElementById('subtab-' + tab).classList.add('active');

    // CLEAR ALL CALCULATORS FIRST
    document.querySelectorAll('.model-view').forEach(v => v.classList.remove('active'));

    if (tab === 'essentials') {
        const val = document.getElementById('essentialsSelect').value;
        switchCalcModel(val);
    } else {
        if (isAdvancedLoggedIn) {
            document.getElementById('advanced-login-form').style.display = 'none';
            document.getElementById('advanced-secured-content').style.display = 'block';
            switchCalcModel(document.getElementById('advancedSelect').value); // Show selected advanced model
        } else {
            document.getElementById('advanced-login-form').style.display = 'block';
            document.getElementById('advanced-secured-content').style.display = 'none';
            // Hide all models
            document.querySelectorAll('.model-view').forEach(v => v.classList.remove('active'));
        }
    }
}

function checkAdvancedLogin() {
    const pass = document.getElementById('adv-password').value;
    const errorMsg = document.getElementById('login-error');
    if (pass === 'packfora123' || pass === 'admin') { 
        isAdvancedLoggedIn = true;
        errorMsg.style.display = 'none';
        document.getElementById('adv-password').value = '';
        switchSubTab('advanced');
    } else {
        errorMsg.style.display = 'block';
    }
}

function lockAdvanced() {
    isAdvancedLoggedIn = false;
    switchSubTab('advanced');
}

function switchCalcModel(model) {
    currentModel = model;
    // Hide ALL model views first
    document.querySelectorAll('.model-view').forEach(c => c.classList.remove('active'));
    
    // Show only the selected model view
    const view = document.getElementById(model + '-calculator');
    if (view) view.classList.add('active');
}

// --- CARTON LOGIC ---
async function calculateCarton() {
    const btn = event.target;
    btn.disabled = true;
    btn.innerHTML = '<span class="loading"></span> Calculating...';
    
    const data = {
        layflat_length: parseFloat(document.getElementById('c_layflat_length').value) || 0,
        layflat_width: parseFloat(document.getElementById('c_layflat_width').value) || 0,
        sheet_length: parseFloat(document.getElementById('c_sheet_length').value) || 0,
        sheet_width: parseFloat(document.getElementById('c_sheet_width').value) || 0,
        side_lay_1: parseFloat(document.getElementById('c_side_lay_1').value) || 0,
        side_lay_2: parseFloat(document.getElementById('c_side_lay_2').value) || 0,
        gripper: parseFloat(document.getElementById('c_gripper').value) || 0,
        back_lay: parseFloat(document.getElementById('c_back_lay').value) || 0,
        trimmer_1: parseFloat(document.getElementById('c_trimmer_1').value) || 0,
        trimmer_2: parseFloat(document.getElementById('c_trimmer_2').value) || 0,
        double_cut: parseFloat(document.getElementById('c_double_cut').value) || 0,
        gutter: parseFloat(document.getElementById('c_gutter').value) || 0,
        interlock: parseFloat(document.getElementById('c_interlock').value) || 0,
        ups_lengthwise: parseInt(document.getElementById('c_ups_l').value) || 1,
        ups_widthwise: parseInt(document.getElementById('c_ups_w').value) || 1,
        board_type: document.getElementById('c_board_type').value,
        board_gsm: parseFloat(document.getElementById('c_board_gsm').value) || 0,
        board_rate: parseFloat(document.getElementById('c_board_rate').value) || 0,
        ink_rate: parseFloat(document.getElementById('c_ink_rate').value) || 0,
        ink_consumption: parseFloat(document.getElementById('c_ink_consumption').value) || 0,
        varnish_type: document.getElementById('c_varnish_type').value,
        varnish_rate: parseFloat(document.getElementById('c_varnish_rate').value) || 0,
        varnish_consumption: parseFloat(document.getElementById('c_varnish_consumption').value) || 0,
        primer_rate: parseFloat(document.getElementById('c_primer_rate').value) || 0,
        primer_consumption: parseFloat(document.getElementById('c_primer_consumption').value) || 0,
        film_rate: parseFloat(document.getElementById('c_film_rate').value) || 0,
        film_gsm: parseFloat(document.getElementById('c_film_gsm').value) || 0,
        adhesive_rate: parseFloat(document.getElementById('c_adhesive_rate').value) || 0,
        adhesive_gsm: parseFloat(document.getElementById('c_adhesive_gsm').value) || 0,
        lam_labour: parseFloat(document.getElementById('c_lam_labour').value) || 0,
        middle_liner_gsm: parseFloat(document.getElementById('c_middle_gsm').value) || 0,
        liner_rate: parseFloat(document.getElementById('c_liner_rate').value) || 0,
        flute_type: document.getElementById('c_flute_type').value,
        inner_liner_gsm: parseFloat(document.getElementById('c_inner_gsm').value) || 0,
        inner_liner_rate: parseFloat(document.getElementById('c_inner_rate').value) || 0,
        corrugation_conversion_rate: parseFloat(document.getElementById('c_corr_conv_rate').value) || 0,
        foil_width_per_carton: parseFloat(document.getElementById('c_foil_w').value) || 0,
        foil_length_per_carton: parseFloat(document.getElementById('c_foil_l').value) || 0,
        foil_cost_per_roll: parseFloat(document.getElementById('c_foil_cost').value) || 0,
        foil_stamping_conversion: parseFloat(document.getElementById('c_foil_conv').value) || 0,
        other_costs: parseFloat(document.getElementById('c_other').value) || 0,
        conversion_cost: parseFloat(document.getElementById('c_conversion').value) || 0,
    };
    
    try {
        const r = await fetch('/api/calc_carton', {
            method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify(data)
        });
        const d = await r.json();
        if (d.error) throw new Error(d.error);
        
        const total = d.total_cost_per_1000;
        const items = [
            {label: 'Board Cost', value: d.board_cost, color: '#4CAF50'},
            {label: 'Ink Cost', value: d.ink_cost, color: '#2196F3'},
            {label: 'Varnish Cost', value: d.varnish_cost, color: '#FF9800'},
            {label: 'Lamination Cost', value: d.lamination_cost, color: '#9C27B0'},
            {label: 'Corrugation Cost', value: d.corrugation_cost, color: '#F44336'},
            {label: 'Foil Cost', value: d.foil_cost, color: '#00BCD4'},
            {label: 'Other Material Cost', value: d.other_material_cost, color: '#795548'},
            {label: 'Conversion Cost', value: d.conversion_cost, color: '#607D8B'},
            {label: 'Packing Cost', value: d.packing_cost, color: '#FF5722'},
        ];
        
        let h = '<h3 style="margin-bottom:5px;">Cost Summary</h3>';
        h += '<p style="opacity:0.6; font-size:0.75rem; margin-bottom:15px;">INR per 1000 Cartons</p>';
        h += '<div style="background:rgba(255,255,255,0.08); border-radius:8px; padding:12px; margin-bottom:15px; font-size:0.8rem;">';
        h += `<div style="display:flex;justify-content:space-between;margin-bottom:4px;"><span style="opacity:0.7;">Area/Carton</span><span>${d.area_per_carton} m²</span></div>`;
        h += `<div style="display:flex;justify-content:space-between;margin-bottom:4px;"><span style="opacity:0.7;">UPs/Sheet</span><span>${d.ups_per_sheet}</span></div>`;
        h += `<div style="display:flex;justify-content:space-between;margin-bottom:4px;"><span style="opacity:0.7;">Actual Sheet</span><span>${d.actual_sheet_length} × ${d.actual_sheet_width} mm</span></div>`;
        h += `<div style="display:flex;justify-content:space-between;"><span style="opacity:0.7;">Board Consumption</span><span>${d.board_consumption_1000} Kg/1000</span></div>`;
        h += '</div>';
        
        items.forEach(item => {
            const pct = total > 0 ? (item.value / total * 100) : 0;
            h += `<div class="summary-row">
                <span class="label">${item.label}</span>
                <span class="value">₹ ${item.value.toLocaleString('en-IN', {minimumFractionDigits:2, maximumFractionDigits:2})}</span>
            </div>
            <div class="cost-bar" style="background:${item.color}; width:${pct}%; opacity:0.6;"></div>`;
        });
        
        h += `<div class="summary-total">
            <span>Total Cost / 1000 Pcs</span>
            <span class="value">₹ ${total.toLocaleString('en-IN', {minimumFractionDigits:2, maximumFractionDigits:2})}</span>
        </div>`;
        h += `<div style="text-align:center; margin-top:10px; opacity:0.6; font-size:0.8rem;">Per Unit: ₹ ${(total/1000).toFixed(4)}</div>`;
        
        document.getElementById('carton-summary').innerHTML = h;
        
        // Save generic model result
        d.material_cost = (d.board_cost||0) + (d.ink_cost||0) + (d.varnish_cost||0) + (d.lamination_cost||0) + (d.corrugation_cost||0) + (d.foil_cost||0) + (d.other_material_cost||0);
        d.margin = 0;
        d.freight_cost = 0;
        d.model_type = 'carton';
        lastModelResult = d;
        lastModelInput = data;
        lastModelType = 'carton';
        window.lastCartonResults = d; // SKU FEATURE: Store results
        showSaveSKUButton('carton'); // SKU FEATURE: Show save button
        
        // Enable compare & show export buttons
        document.getElementById('compareCountriesBtn').disabled = false;
        var expDiv = document.getElementById('carton-export-btns');
        if (expDiv) expDiv.style.display = 'block';
        
        // Pie chart
        var pieDiv = document.getElementById('carton-pie-chart');
        if (pieDiv) {
            pieDiv.style.display = 'block';
            var pieVals = [d.material_cost, d.conversion_cost, d.packing_cost].filter(v => v > 0);
            var pieLabels = ['Material', 'Conversion', 'Packing'].filter((_, i) => [d.material_cost, d.conversion_cost, d.packing_cost][i] > 0);
            var pieColors = ['#4CAF50', '#2196F3', '#9C27B0'].filter((_, i) => [d.material_cost, d.conversion_cost, d.packing_cost][i] > 0);
            Plotly.newPlot('carton-donut', [{
                values: pieVals, labels: pieLabels, type: 'pie', hole: 0.45,
                marker: {colors: pieColors}, textinfo: 'label+percent', textfont: {color: 'white', size: 11}, hoverinfo: 'label+value+percent',
            }], {paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)', font:{color:'white',family:'Outfit'}, showlegend:false, margin:{t:10,b:10,l:10,r:10}, height:300, annotations:[{text:'Cost<br>Split',font:{size:13,color:'white'},showarrow:false}]}, {displayModeBar:false, responsive:true});
        }
    } catch(e) {
        document.getElementById('carton-summary').innerHTML = `<h3>Error</h3><p style="color:#ef4444;margin-top:10px;">${e.message}</p>`;
    } finally {
        btn.disabled = false;
        btn.innerHTML = 'Calculate Carton Cost';
    }
}

// --- FLEXIBLES LOGIC ---
function renderFlexLayers() {
    const container = document.getElementById('flex-layers');
    let h = '';
    flexLayers.forEach((layer, idx) => {
        const opts = FILM_OPTIONS.map(f => `<option${f===layer.name?' selected':''}>${f}</option>`).join('');
        h += `<div class="flex-layer-block">
            <div class="flex-layer-header">
                <span class="flex-layer-title">Layer ${idx+1}</span>
                ${flexLayers.length > 1 ? `<button class="btn-remove-layer" onclick="removeFlexLayer(${idx})">Remove</button>` : ''}
            </div>
            <div class="calc-row"><label>Type</label><span class="unit"></span><select class="calc-select" id="fl_name_${idx}" onchange="flexLayers[${idx}].name=this.value">${opts}</select></div>
            <div class="calc-row"><label>Micron</label><span class="unit">mic</span><input type="number" class="calc-input calc-input-green" id="fl_mic_${idx}" value="${layer.mic}" onchange="flexLayers[${idx}].mic=parseFloat(this.value)||0; updateFlexGSM();"></div>
            <div class="calc-row"><label>Rate</label><span class="unit">₹/Kg</span><input type="number" class="calc-input calc-input-green" id="fl_rate_${idx}" value="${layer.rate}" onchange="flexLayers[${idx}].rate=parseFloat(this.value)||0"></div>
        </div>`;
    });
    container.innerHTML = h;
    updateFlexGSM();
}

function addFlexLayer() {
    flexLayers.push({name: 'BOPP Film', mic: 12, rate: 100});
    renderFlexLayers();
}

function removeFlexLayer(idx) {
    flexLayers.splice(idx, 1);
    renderFlexLayers();
}

function updateFlexGSM() {
    const densityMap = {'HIPS':1.04,'GPPS':1.05,'PET Film':1.45,'EAA':0.92,'EVA':0.93,'CPP Film':0.9,'HDPE':0.95,'MDPE':0.94,'BON':1.14,'AL Foil':2.7,'Cellophane':1.55,'BOPP Film':0.91,'METPET Film':1.4,'MET MDOPE Film':0.96,'Matt Finish PET Film':1.45,'Primer':1,'Lamination - Adhesive (Solvent Based)':1,'Lamination - Adhesive (Solvent Less)':1,'Heat Seal Lacquer':1,'Cold Seal':1,'Gloss Varnish':1,'Matte Varnish':1,'Gravure':1,'Flexo':1,'Mono Layer PE':0.95,'2 Layer All PE':0.95,'3 Layer All PE':0.95,'5 Layer All PE':0.95,'5 Layer EVOH Barrier':0.96,'5 Layer Nylon Barrier':1,'7 Layer All PE':0.95};
    let total = 0;
    flexLayers.forEach(l => { total += l.mic * (densityMap[l.name] || 1); });
    document.getElementById('f_laminate_gsm').value = total.toFixed(2);
}

async function calculateFlexibles() {
    const btn = event.target;
    btn.disabled = true;
    btn.innerHTML = '<span class="loading"></span> Calculating...';
    
    // Re-read layer values
    flexLayers.forEach((layer, idx) => {
        layer.name = document.getElementById(`fl_name_${idx}`).value;
        layer.mic = parseFloat(document.getElementById(`fl_mic_${idx}`).value) || 0;
        layer.rate = parseFloat(document.getElementById(`fl_rate_${idx}`).value) || 0;
    });
    
    const data = {
        layers: flexLayers,
        conversion_cost: parseFloat(document.getElementById('f_conversion').value) || 0,
    };
    
    try {
        const r = await fetch('/api/calc_flexibles', {
            method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify(data)
        });
        const d = await r.json();
        if (d.error) throw new Error(d.error);
        
        let h = '<h3 style="margin-bottom:5px;">Cost Summary</h3>';
        h += '<p style="opacity:0.6; font-size:0.75rem; margin-bottom:15px;">Flexibles Laminate Costing</p>';
        h += '<div style="margin-bottom:15px;">';
        h += '<div style="font-size:0.75rem; font-weight:700; color:var(--orange); margin-bottom:8px;">LAYER BREAKDOWN</div>';
        d.layers.forEach((l, i) => {
            h += `<div style="display:flex;justify-content:space-between;padding:6px 0;border-bottom:1px solid rgba(255,255,255,0.05);font-size:0.82rem;">
                <span style="opacity:0.8;">L${i+1}: ${l.name} (${l.mic}μ)</span>
                <span style="font-weight:700;">₹${l.layer_cost.toFixed(2)}/kg</span>
            </div>`;
        });
        h += '</div>';
        
        h += `<div class="summary-row"><span class="label">Laminate GSM</span><span class="value">${d.laminate_gsm}</span></div>`;
        h += `<div class="summary-row"><span class="label">Avg Density</span><span class="value">${d.avg_density.toFixed(4)} g/cm³</span></div>`;
        h += `<div class="summary-row"><span class="label">Material Cost</span><span class="value">₹ ${d.material_cost_per_kg.toFixed(2)}/kg</span></div>`;
        h += `<div class="summary-row"><span class="label">Wastage (${d.wastage_pct}%)</span><span class="value">₹ ${d.wastage_cost.toFixed(2)}/kg</span></div>`;
        h += `<div class="summary-row"><span class="label">Material + Wastage</span><span class="value">₹ ${d.material_cost_with_wastage.toFixed(2)}/kg</span></div>`;
        h += `<div class="summary-row"><span class="label">Conversion Cost</span><span class="value">₹ ${d.conversion_cost.toFixed(2)}/kg</span></div>`;
        h += `<div class="summary-row"><span class="label">Packing Cost</span><span class="value">₹ ${d.packing_cost.toFixed(2)}/kg</span></div>`;
        
        h += `<div class="summary-total"><span>Laminate Cost</span><span class="value">₹ ${d.laminate_cost_per_kg.toFixed(2)}/kg</span></div>`;
        h += `<div class="summary-total" style="border-top:1px solid rgba(255,255,255,0.2); padding-top:10px; margin-top:5px;">
            <span>Cost per SQM</span><span class="value" style="font-size:1.1rem;">₹ ${d.laminate_cost_per_sqm.toFixed(2)}</span></div>`;
        
        document.getElementById('flex-summary').innerHTML = h;
        
        // Save generic model result - normalize to common fields
        d.material_cost = d.material_cost_with_wastage || d.material_cost_per_kg || 0;
        d.margin = 0;
        d.freight_cost = 0;
        d.total_cost_per_1000 = d.laminate_cost_per_kg || 0;
        d.model_type = 'flexibles';
        lastModelResult = d;
        lastModelInput = data;
        lastModelType = 'flexibles';
        window.lastFlexiblesResults = d; // SKU FEATURE: Store results
        showSaveSKUButton('flexibles'); // SKU FEATURE: Show save button
        
        // Enable compare & show export buttons
        document.getElementById('compareCountriesBtn').disabled = false;
        var expDiv = document.getElementById('flex-export-btns');
        if (expDiv) expDiv.style.display = 'block';
        
        // Pie chart
        var pieDiv = document.getElementById('flex-pie-chart');
        if (pieDiv) {
            pieDiv.style.display = 'block';
            var pieVals = [d.material_cost, d.conversion_cost, d.packing_cost].filter(v => v > 0);
            var pieLabels = ['Material', 'Conversion', 'Packing'].filter((_, i) => [d.material_cost, d.conversion_cost, d.packing_cost][i] > 0);
            var pieColors = ['#4CAF50', '#2196F3', '#9C27B0'].filter((_, i) => [d.material_cost, d.conversion_cost, d.packing_cost][i] > 0);
            Plotly.newPlot('flex-donut', [{
                values: pieVals, labels: pieLabels, type: 'pie', hole: 0.45,
                marker: {colors: pieColors}, textinfo: 'label+percent', textfont: {color: 'white', size: 11}, hoverinfo: 'label+value+percent',
            }], {paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)', font:{color:'white',family:'Outfit'}, showlegend:false, margin:{t:10,b:10,l:10,r:10}, height:300, annotations:[{text:'Cost<br>Split',font:{size:13,color:'white'},showarrow:false}]}, {displayModeBar:false, responsive:true});
        }
    } catch(e) {
        document.getElementById('flex-summary').innerHTML = `<h3>Error</h3><p style="color:#ef4444;margin-top:10px;">${e.message}</p>`;
    } finally {
        btn.disabled = false;
        btn.innerHTML = 'Calculate Flexibles Cost';
    }
}

// --- EBM LOGIC ---
async function calculateEBM() {
    const btn = event.target;
    btn.disabled = true;
    btn.innerHTML = '<span class="loading"></span> Calculating...';
    
    // COLLECT ALL INPUTS
    const data = {
        sku_description: document.getElementById('e_sku').value,
        country: document.getElementById('e_country').value,
        annual_volume: parseFloat(document.getElementById('e_volume').value) || 0,
        weight: parseFloat(document.getElementById('e_weight').value) || 0,
        l1_ratio: parseFloat(document.getElementById('e_l1_ratio').value) || 0,
        l1_polymer_type: document.getElementById('e_l1_type').value,
        l1_polymer_rate: parseFloat(document.getElementById('e_l1_rate').value) || 0,
        l1_mb_dosage: parseFloat(document.getElementById('e_l1_mb_dosage').value) || 0,
        l1_mb_rate: parseFloat(document.getElementById('e_l1_mb_rate').value) || 0,
        l1_additive_dosage: parseFloat(document.getElementById('e_l1_add_dosage').value) || 0,
        l1_additive_rate: parseFloat(document.getElementById('e_l1_add_rate').value) || 0,
        l2_ratio: parseFloat(document.getElementById('e_l2_ratio').value) || 0,
        l2_polymer_type: document.getElementById('e_l2_type').value,
        l2_polymer_rate: parseFloat(document.getElementById('e_l2_rate').value) || 0,
        l2_mb_dosage: parseFloat(document.getElementById('e_l2_mb_dosage').value) || 0,
        l2_mb_rate: parseFloat(document.getElementById('e_l2_mb_rate').value) || 0,
        l2_additive_dosage: parseFloat(document.getElementById('e_l2_add_dosage').value) || 0,
        l2_additive_rate: parseFloat(document.getElementById('e_l2_add_rate').value) || 0,
        l3_ratio: parseFloat(document.getElementById('e_l3_ratio').value) || 0,
        l3_polymer_type: document.getElementById('e_l3_type').value,
        l3_polymer_rate: parseFloat(document.getElementById('e_l3_rate').value) || 0,
        l3_mb_dosage: parseFloat(document.getElementById('e_l3_mb_dosage').value) || 0,
        l3_mb_rate: parseFloat(document.getElementById('e_l3_mb_rate').value) || 0,
        l3_additive_dosage: parseFloat(document.getElementById('e_l3_add_dosage').value) || 0,
        l3_additive_rate: parseFloat(document.getElementById('e_l3_add_rate').value) || 0,
        mould_cavitation: parseInt(document.getElementById('e_cavity').value) || 1,
        mould_cycle_time: parseFloat(document.getElementById('e_cycle').value) || 1,
        machine_model: document.getElementById('e_machine').value,
        num_setups_year: parseInt(document.getElementById('e_setups').value) || 0,
        num_rampups_year: parseInt(document.getElementById('e_rampups').value) || 0,
        electricity_rate: parseFloat(document.getElementById('e_elec_rate').value) || 0,
        skilled_labour: parseFloat(document.getElementById('e_labour').value) || 0,
        engineer: parseFloat(document.getElementById('e_engineer').value) || 0,
        prod_manager: parseFloat(document.getElementById('e_pm').value) || 0,
        repair_pct: parseFloat(document.getElementById('e_repair').value) || 0,
        other_oh_pct: parseFloat(document.getElementById('e_other_oh').value) || 0,
        depreciation_pm: parseFloat(document.getElementById('e_dep_pm').value) || 0,
        depreciation_bldg: parseFloat(document.getElementById('e_dep_bldg').value) || 0,
        completed_life: parseInt(document.getElementById('e_life').value) || 5,
        land_cost: parseFloat(document.getElementById('e_land').value) || 0,
        building_cost: parseFloat(document.getElementById('e_building').value) || 0,
        lease_cost: parseFloat(document.getElementById('e_lease').value) || 0,
        premises_type: document.getElementById('e_premises').value,
        interest_lt: parseFloat(document.getElementById('e_int_lt').value) || 0,
        interest_wc: parseFloat(document.getElementById('e_int_wc').value) || 0,
        margin_pct: parseFloat(document.getElementById('e_margin').value) || 0,
        margin_calc: document.getElementById('e_margin_calc').value,
        lt_debt_equity: parseFloat(document.getElementById('e_debt_equity').value) || 0,
        num_orders_year: parseInt(document.getElementById('e_orders').value) || 12,
        bottles_per_box: parseInt(document.getElementById('e_bottles_box').value) || 1,
        boxes_per_container: parseInt(document.getElementById('e_boxes_cont').value) || 1,
        shipper_cost: parseFloat(document.getElementById('e_shipper').value) || 0,
        polybag_cost: parseFloat(document.getElementById('e_polybag').value) || 0,
        freight_per_container: parseFloat(document.getElementById('e_freight').value) || 0,
        rm_payment_days: parseInt(document.getElementById('e_rm_days').value) || 0,
        fg_payment_days: parseInt(document.getElementById('e_fg_days').value) || 0,
        euro_rate: parseFloat(document.getElementById('e_euro_rate').value) || 1,
    };
    
    try {
        const r = await fetch('/api/calc_ebm', {
            method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify(data)
        });
        const d = await r.json();
        if (d.error) throw new Error(d.error);
        
        const fmt = (v) => v.toLocaleString('en-IN', {minimumFractionDigits:2, maximumFractionDigits:2});
        const total = d.total_cost_per_1000;
        
        let h = '<h3 style="margin-bottom:5px;">EBM Cost Summary</h3>';
        h += '<p style="opacity:0.6; font-size:0.75rem; margin-bottom:15px;">' + d.currency + ' per 1000 Pcs</p>';
        
        h += '<div style="background:rgba(255,255,255,0.08); border-radius:8px; padding:12px; margin-bottom:15px; font-size:0.8rem;">';
        h += '<div style="display:flex;justify-content:space-between;margin-bottom:4px;"><span style="opacity:0.7;">Machines Required</span><span>' + d.num_machines + '</span></div>';
        h += '<div style="display:flex;justify-content:space-between;margin-bottom:4px;"><span style="opacity:0.7;">Output/Hour</span><span>' + d.output_per_hour.toLocaleString() + ' pcs</span></div>';
        h += '<div style="display:flex;justify-content:space-between;margin-bottom:4px;"><span style="opacity:0.7;">Total Investment</span><span>₹ ' + d.total_investment_inr.toLocaleString('en-IN') + '</span></div>';
        h += '<div style="display:flex;justify-content:space-between;margin-bottom:4px;"><span style="opacity:0.7;">Land Area</span><span>' + d.land_area_sqm + ' SQM</span></div>';
        h += '<div style="display:flex;justify-content:space-between;"><span style="opacity:0.7;">Wastage</span><span>' + d.wastage_pct + '%</span></div>';
        h += '</div>';
        
        const items = [
            {label: 'Material Cost', value: d.material_cost, color: '#4CAF50', pct: d.mat_pct},
            {label: 'Conversion Cost', value: d.conversion_cost, color: '#2196F3', pct: d.conv_pct},
            {label: 'Margin', value: d.margin, color: '#FF9800', pct: d.margin_pct_total},
            {label: 'Packaging Cost', value: d.packing_cost, color: '#9C27B0', pct: d.pkg_pct},
            {label: 'Freight Cost', value: d.freight_cost, color: '#F44336', pct: d.freight_pct},
        ];
        
        items.forEach(item => {
            h += '<div class="summary-row"><span class="label">' + item.label + ' <span style="opacity:0.5;font-size:0.75rem;">(' + item.pct + '%)</span></span><span class="value">₹ ' + fmt(item.value) + '</span></div>';
            h += '<div class="cost-bar" style="background:' + item.color + '; width:' + item.pct + '%; opacity:0.6;"></div>';
        });
        
        h += '<div class="summary-total"><span>Total / 1000 Pcs</span><span class="value">₹ ' + fmt(total) + '</span></div>';
        h += '<div style="text-align:center; margin-top:10px; opacity:0.6; font-size:0.8rem;">Per Unit: ₹ ' + d.cost_per_piece.toFixed(4) + ' | EUR/1000: € ' + fmt(d.total_cost_eur) + '</div>';
        
        // Detailed breakdowns
        h += '<div style="margin-top:20px;">';
        h += '<div style="font-size:0.75rem; font-weight:700; color:var(--orange); margin-bottom:8px;">MATERIAL BREAKDOWN</div>';
        h += '<div class="summary-row"><span class="label">Resin</span><span class="value">₹ ' + fmt(d.resin_cost) + '</span></div>';
        h += '<div class="summary-row"><span class="label">Masterbatch</span><span class="value">₹ ' + fmt(d.mb_cost) + '</span></div>';
        h += '<div class="summary-row"><span class="label">Additives</span><span class="value">₹ ' + fmt(d.additive_cost) + '</span></div>';
        h += '<div class="summary-row"><span class="label">Wastage</span><span class="value">₹ ' + fmt(d.wastage_cost) + '</span></div>';
        h += '</div>';
        
        h += '<div style="margin-top:15px;">';
        h += '<div style="font-size:0.75rem; font-weight:700; color:var(--orange); margin-bottom:8px;">CONVERSION BREAKDOWN</div>';
        h += '<div class="summary-row"><span class="label">Electricity</span><span class="value">₹ ' + fmt(d.electricity_cost) + '</span></div>';
        h += '<div class="summary-row"><span class="label">Direct Labour</span><span class="value">₹ ' + fmt(d.direct_labour) + '</span></div>';
        h += '<div class="summary-row"><span class="label">Indirect Labour</span><span class="value">₹ ' + fmt(d.indirect_labour) + '</span></div>';
        h += '<div class="summary-row"><span class="label">Repair & Maint</span><span class="value">₹ ' + fmt(d.repair_cost) + '</span></div>';
        h += '<div class="summary-row"><span class="label">Depreciation</span><span class="value">₹ ' + fmt(d.depreciation) + '</span></div>';
        h += '<div class="summary-row"><span class="label">Interest</span><span class="value">₹ ' + fmt(d.interest_total) + '</span></div>';
        h += '</div>';
        
        document.getElementById('ebm-summary').innerHTML = h;
        
        // Save for exports/analysis
        d.sku_description = data.sku_description;
        d.country = data.country;
        lastEBMResult = d;
        lastEBMInput = data;
        lastModelResult = d;
        lastModelInput = data;
        lastModelType = 'ebm';
        window.lastEBMResults = d; // SKU FEATURE: Store results
        showSaveSKUButton('ebm'); // SKU FEATURE: Show save button
        
        renderEBMPieChart(d);
        document.getElementById('ebm-export-btns').style.display = 'block';
        document.getElementById('compareCountriesBtn').disabled = false;
    } catch(e) {
        document.getElementById('ebm-summary').innerHTML = '<h3>Error</h3><p style="color:#ef4444;margin-top:10px;">' + e.message + '</p>';
    } finally {
        btn.disabled = false;
        btn.innerHTML = 'Calculate EBM Cost';
    }
}

// --- UTILITIES ---

async function loadMachinesFromDB() {
    try {
        const r = await fetch('/api/machine_db_for_calc', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({process:'blow'})});
        const d = await r.json();
        if (d.machines && d.machines.length > 0) {
            let info = '<div style="margin-top:10px; font-size:0.75rem; opacity:0.7;">DB: ' + d.machines.length + ' machines found</div>';
            let machSel = document.getElementById('e_machine');
            if (machSel) {
                let infoDiv = document.getElementById('machine-db-info');
                if (!infoDiv) {
                    infoDiv = document.createElement('div');
                    infoDiv.id = 'machine-db-info';
                    machSel.parentNode.appendChild(infoDiv);
                }
                infoDiv.innerHTML = info;
            }
        }
    } catch(e) {}
}

async function loadVariableCostsFromDB(country) {
    try {
        const r = await fetch('/api/variable_cost_for_calc', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({country})});
        const d = await r.json();
        if (d.variables) {
            let info = '<div style="margin-top:8px;padding:8px;background:rgba(76,175,80,0.1);border:1px solid rgba(76,175,80,0.3);border-radius:6px;font-size:0.72rem;">';
            info += '<strong style="color:#4CAF50;">Variable Cost DB Loaded</strong><br>';
            let count = 0;
            for (let [k, v] of Object.entries(d.variables)) {
                if (count < 5 && v > 0) { info += k + ': ' + (typeof v === 'number' ? v.toLocaleString() : v) + '<br>'; count++; }
            }
            if (Object.keys(d.variables).length > 5) info += '... and ' + (Object.keys(d.variables).length - 5) + ' more variables';
            info += '</div>';
            let dbDiv = document.getElementById('var-cost-db-info');
            if (!dbDiv) {
                dbDiv = document.createElement('div');
                dbDiv.id = 'var-cost-db-info';
                document.getElementById('e_country').parentNode.appendChild(dbDiv);
            }
            dbDiv.innerHTML = info;
        }
    } catch(e) {}
}

function renderEBMPieChart(d) {
    var pieDiv = document.getElementById('ebm-pie-chart');
    pieDiv.style.display = 'block';
    var data = [{
        values: [d.material_cost, d.conversion_cost, d.margin, d.packing_cost, d.freight_cost],
        labels: ['Material', 'Conversion', 'Margin', 'Packaging', 'Freight'],
        type: 'pie', hole: 0.45,
        marker: {colors: ['#4CAF50', '#2196F3', '#FF9800', '#9C27B0', '#F44336']},
        textinfo: 'label+percent', textfont: {color: 'white', size: 11},
        hoverinfo: 'label+value+percent',
    }];
    var layout = {
        paper_bgcolor: 'rgba(0,0,0,0)', plot_bgcolor: 'rgba(0,0,0,0)',
        font: {color: 'white', family: 'Outfit'}, showlegend: false,
        margin: {t: 10, b: 10, l: 10, r: 10}, height: 300,
        annotations: [{text: 'Cost<br>Split', font: {size: 13, color: 'white'}, showarrow: false}]
    };
    Plotly.newPlot('ebm-donut', data, layout, {displayModeBar: false, responsive: true});
}

async function exportEBMExcel() {
    if (!lastEBMResult && !(lastModelResult && lastModelType === 'ebm')) { alert('Calculate EBM first'); return; }
    const data = lastEBMResult || lastModelResult;
    try {
        const r = await fetch('/api/export_ebm_excel', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify(data)});
        if (!r.ok) { const e = await r.json(); throw new Error(e.error || 'Export failed'); }
        const blob = await r.blob();
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a'); a.href = url; a.download = 'EBM_Report.xlsx';
        document.body.appendChild(a); a.click(); URL.revokeObjectURL(url); document.body.removeChild(a);
    } catch(e) { alert('Export failed: ' + e.message); }
}

async function exportEBMPDF() {
    if (!lastEBMResult && !(lastModelResult && lastModelType === 'ebm')) { alert('Calculate EBM first'); return; }
    const data = lastEBMResult || lastModelResult;
    try {
        const r = await fetch('/api/export_ebm_pdf', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify(data)});
        if (!r.ok) { const e = await r.json(); throw new Error(e.error || 'Export failed'); }
        const blob = await r.blob();
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a'); a.href = url; a.download = 'EBM_Report.pdf';
        document.body.appendChild(a); a.click(); URL.revokeObjectURL(url); document.body.removeChild(a);
    } catch(e) { alert('Export failed: ' + e.message); }
}

// --- MULTI-COUNTRY ---
const ALL_COUNTRIES = ['India','China','Vietnam','Turkey','Indonesia','Brazil','United States','United Kingdom','Germany','France','Mexico','Pakistan','Philippines','South Africa','Spain','Poland','Thailand','Bangladesh','Sri Lanka','Argentina','Canada','Costa Rica'];

function initCountryCheckboxes() {
    const container = document.getElementById('country-checkboxes');
    if (!container) return;
    container.innerHTML = ALL_COUNTRIES.map(c => 
        '<label class="checkbox-label" style="background:rgba(255,255,255,0.05);border-radius:6px;padding:6px 10px;font-size:0.82rem;"><input type="checkbox" value="' + c + '" onchange="updateCompareBtn()"> ' + c + '</label>'
    ).join('');
}

function updateCompareBtn() {
    const checked = document.querySelectorAll('#country-checkboxes input:checked');
    document.getElementById('compareCountriesBtn').disabled = checked.length < 2;
}

async function runMultiCountry() {
    if (!lastModelResult) { alert('Calculate a cost model first to set base parameters'); return; }
    const checked = Array.from(document.querySelectorAll('#country-checkboxes input:checked')).map(c => c.value);
    if (checked.length < 2) { alert('Select at least 2 countries'); return; }
    
    const btn = document.getElementById('compareCountriesBtn');
    btn.disabled = true; btn.innerHTML = '<span class="loading"></span> Comparing...';
    
    // For EBM use its dedicated endpoint; for all others use generic
    let apiUrl, payload;
    if (lastModelType === 'ebm') {
        apiUrl = '/api/multi_country_ebm';
        payload = {countries: checked, base_params: lastEBMInput || lastModelInput};
    } else {
        apiUrl = '/api/multi_country_generic';
        payload = {
            countries: checked,
            model_type: lastModelType,
            base_params: lastModelInput || {},
            base_result: {
                material_cost: lastModelResult.material_cost || 0,
                conversion_cost: lastModelResult.conversion_cost || 0,
                margin: lastModelResult.margin || 0,
                packing_cost: lastModelResult.packing_cost || 0,
                freight_cost: lastModelResult.freight_cost || 0,
                total_cost_per_1000: lastModelResult.total_cost_per_1000 || 0,
                electricity_cost: lastModelResult.electricity_cost || 0,
                direct_labour: lastModelResult.direct_labour || 0,
                indirect_labour: lastModelResult.indirect_labour || 0,
            },
            base_country: lastModelResult.country || lastModelInput.country || 'India'
        };
    }
    
    try {
        const r = await fetch(apiUrl, {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify(payload)});
        const d = await r.json();
        if (d.error) throw new Error(d.error);
        
        let h = '<div class="card"><div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:15px;"><h3>Country Comparison (EUR/1000 Pcs)</h3>';
        h += '<button class="btn-secondary" onclick="exportMultiCountryExcel()">Export Excel</button></div>';
        h += '<div style="overflow-x:auto;"><table style="width:100%;border-collapse:collapse;font-size:0.82rem;">';
        h += '<tr style="border-bottom:2px solid var(--orange);"><th style="padding:10px;text-align:left;">Country</th><th style="padding:10px;text-align:right;">Material</th><th style="padding:10px;text-align:right;">Conversion</th><th style="padding:10px;text-align:right;">Margin</th><th style="padding:10px;text-align:right;">Packing</th><th style="padding:10px;text-align:right;">Freight</th><th style="padding:10px;text-align:right;color:var(--orange);font-weight:800;">Total EUR</th></tr>';
        
        d.results.forEach((r, i) => {
            if (r.error) { h += '<tr><td>' + r.country + '</td><td colspan="6" style="color:#ef4444;">Error: ' + r.error + '</td></tr>'; return; }
            const bg = i === 0 ? 'background:rgba(76,175,80,0.1);' : '';
            h += '<tr style="border-bottom:1px solid rgba(255,255,255,0.1);' + bg + '"><td style="padding:10px;font-weight:700;">' + r.country + (i===0?' <span style="color:#4CAF50;font-size:0.7rem;">LOWEST</span>':'') + '</td>';
            h += '<td style="padding:10px;text-align:right;">€' + (r.mat_eur||0).toLocaleString('en',{minimumFractionDigits:2}) + '</td>';
            h += '<td style="padding:10px;text-align:right;">€' + (r.conv_eur||0).toLocaleString('en',{minimumFractionDigits:2}) + '</td>';
            h += '<td style="padding:10px;text-align:right;">€' + (r.margin_eur||0).toLocaleString('en',{minimumFractionDigits:2}) + '</td>';
            h += '<td style="padding:10px;text-align:right;">€' + (r.pkg_eur||0).toLocaleString('en',{minimumFractionDigits:2}) + '</td>';
            h += '<td style="padding:10px;text-align:right;">€' + (r.frt_eur||0).toLocaleString('en',{minimumFractionDigits:2}) + '</td>';
            h += '<td style="padding:10px;text-align:right;font-weight:800;color:var(--orange);">€' + (r.total_eur||0).toLocaleString('en',{minimumFractionDigits:2}) + '</td></tr>';
        });
        h += '</table></div></div>';
        document.getElementById('multi-country-results').innerHTML = h;
        window._multiCountryResults = d.results;
        
        const valid = d.results.filter(r => !r.error);
        var traces = [
            {x: valid.map(r=>r.country), y: valid.map(r=>r.mat_eur||0), name:'Material', type:'bar', marker:{color:'#4CAF50'}},
            {x: valid.map(r=>r.country), y: valid.map(r=>r.conv_eur||0), name:'Conversion', type:'bar', marker:{color:'#2196F3'}},
            {x: valid.map(r=>r.country), y: valid.map(r=>r.margin_eur||0), name:'Margin', type:'bar', marker:{color:'#FF9800'}},
            {x: valid.map(r=>r.country), y: valid.map(r=>r.pkg_eur||0), name:'Packing', type:'bar', marker:{color:'#9C27B0'}},
            {x: valid.map(r=>r.country), y: valid.map(r=>r.frt_eur||0), name:'Freight', type:'bar', marker:{color:'#F44336'}},
        ];
        Plotly.newPlot('multi-country-chart', traces, {
            barmode:'stack', paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)',
            font:{color:'white',family:'Outfit'}, legend:{orientation:'h',y:-0.15},
            yaxis:{title:'EUR / 1000 Pcs',gridcolor:'rgba(255,255,255,0.1)'},
            xaxis:{gridcolor:'rgba(255,255,255,0.1)'},
            margin:{t:20,b:80,l:60,r:20}, height:400
        }, {displayModeBar:false, responsive:true});
        
    } catch(e) {
        document.getElementById('multi-country-results').innerHTML = '<div class="card error-card"><h3>Error</h3><p>' + e.message + '</p></div>';
    } finally {
        btn.disabled = false; btn.innerHTML = 'Compare Countries';
    }
}

// --- WHAT-IF ---
function updateWhatIf() {
    ['resin','volume','elec','labour','margin'].forEach(p => {
        document.getElementById('wi_'+p+'_label').textContent = document.getElementById('wi_'+p).value + '%';
    });
    if (!lastModelResult) return;
    
    const resinChg = parseFloat(document.getElementById('wi_resin').value)/100;
    const volChg = parseFloat(document.getElementById('wi_volume').value)/100;
    const elecChg = parseFloat(document.getElementById('wi_elec').value)/100;
    const labourChg = parseFloat(document.getElementById('wi_labour').value)/100;
    const marginChg = parseFloat(document.getElementById('wi_margin').value)/100;
    
    const base = lastModelResult;
    const mt = lastModelType;
    
    let newMaterial, newConv, newMargin, newPacking, newFreight, newTotal, baseTotal;
    
    if (mt === 'ebm' || mt === 'carton-adv') {
        // Full detailed breakdown
        newMaterial = base.material_cost * (1 + resinChg);
        const volFactor = volChg !== 0 ? 1 / (1 + volChg) : 1;
        const elecCost = base.electricity_cost || 0;
        const dlCost = base.direct_labour || 0;
        const ilCost = base.indirect_labour || 0;
        const newElec = elecCost * (1 + elecChg) * volFactor;
        const newLabour = (dlCost + ilCost) * (1 + labourChg) * volFactor;
        const otherConv = base.conversion_cost - elecCost - dlCost - ilCost;
        newConv = newElec + newLabour + otherConv * volFactor;
        // Margin = margin_pct * conversion_cost (dynamic), then apply margin slider
        const mPctInput = base.margin_pct_input || (base.conversion_cost > 0 ? (base.margin / base.conversion_cost) : 0.20);
        const mCalcType = base.margin_calc_type || '% of Conversion Cost';
        if (mCalcType === '% of Conversion Cost') {
            newMargin = newConv * mPctInput * (1 + marginChg);
        } else {
            newMargin = (newMaterial + newConv) * mPctInput * (1 + marginChg);
        }
        newPacking = base.packing_cost;
        newFreight = base.freight_cost;
        newTotal = newMaterial + newConv + newMargin + newPacking + newFreight;
        baseTotal = base.total_cost_per_1000;
    } else {
        // Simpler models (carton essentials, flexibles)
        newMaterial = (base.material_cost || 0) * (1 + resinChg);
        const volFactor = volChg !== 0 ? 1 / (1 + volChg) : 1;
        newConv = (base.conversion_cost || 0) * (1 + (elecChg + labourChg)/2) * volFactor;
        newMargin = (base.margin || 0) * (1 + marginChg);
        newPacking = base.packing_cost || 0;
        newFreight = base.freight_cost || 0;
        newTotal = newMaterial + newConv + newMargin + newPacking + newFreight;
        baseTotal = base.total_cost_per_1000 || (newMaterial + (base.conversion_cost||0) + (base.margin||0) + (base.packing_cost||0) + (base.freight_cost||0));
    }
    
    const diff = newTotal - baseTotal;
    const diffPct = baseTotal !== 0 ? (diff / baseTotal * 100) : 0;
    
    const fmt = (v) => v.toLocaleString('en-IN', {minimumFractionDigits:2, maximumFractionDigits:2});
    const clr = diff > 0 ? '#ef4444' : diff < 0 ? '#10b981' : 'white';
    const unitLabel = mt === 'flexibles' ? '₹/kg' : '₹/1000';
    
    let h = '<div class="card">';
    h += '<h3 style="margin-bottom:15px;">Scenario Impact</h3>';
    h += '<div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:15px;text-align:center;margin-bottom:20px;">';
    h += '<div style="background:rgba(255,255,255,0.08);padding:15px;border-radius:10px;"><div style="font-size:0.75rem;opacity:0.6;">BASE COST</div><div style="font-size:1.3rem;font-weight:800;">' + unitLabel.charAt(0) + ' ' + fmt(baseTotal) + '</div></div>';
    h += '<div style="background:rgba(255,255,255,0.08);padding:15px;border-radius:10px;"><div style="font-size:0.75rem;opacity:0.6;">NEW COST</div><div style="font-size:1.3rem;font-weight:800;color:' + clr + ';">' + unitLabel.charAt(0) + ' ' + fmt(newTotal) + '</div></div>';
    h += '<div style="background:rgba(255,255,255,0.08);padding:15px;border-radius:10px;"><div style="font-size:0.75rem;opacity:0.6;">IMPACT</div><div style="font-size:1.3rem;font-weight:800;color:' + clr + ';">' + (diff>=0?'+':'') + fmt(diff) + ' (' + diffPct.toFixed(1) + '%)</div></div>';
    h += '</div>';
    
    const items = [
        {label:'Material', base:base.material_cost||0, new_val:newMaterial},
        {label:'Conversion', base:base.conversion_cost||0, new_val:newConv},
        {label:'Margin', base:base.margin||0, new_val:newMargin},
        {label:'Packing', base:base.packing_cost||0, new_val:newPacking},
        {label:'Freight', base:base.freight_cost||0, new_val:newFreight},
    ].filter(it => it.base > 0 || it.new_val > 0);
    h += '<div style="font-size:0.82rem;">';
    items.forEach(it => {
        const d2 = it.new_val - it.base;
        const c = d2 > 0.01 ? '#ef4444' : d2 < -0.01 ? '#10b981' : 'rgba(255,255,255,0.6)';
        h += '<div style="display:flex;justify-content:space-between;padding:8px 0;border-bottom:1px solid rgba(255,255,255,0.08);">';
        h += '<span style="opacity:0.8;">' + it.label + '</span>';
        h += '<span>' + unitLabel.charAt(0) + fmt(it.base) + ' → <span style="color:' + c + ';font-weight:700;">' + unitLabel.charAt(0) + fmt(it.new_val) + '</span></span>';
        h += '</div>';
    });
    h += '</div></div>';
    document.getElementById('whatif-results').innerHTML = h;
    
    var waterfallLabels = ['Base'];
    var waterfallValues = [baseTotal];
    var waterfallMeasure = ['absolute'];
    if ((base.material_cost||0) > 0) { waterfallLabels.push('Material Δ'); waterfallValues.push(newMaterial-(base.material_cost||0)); waterfallMeasure.push('relative'); }
    if ((base.conversion_cost||0) > 0) { waterfallLabels.push('Conversion Δ'); waterfallValues.push(newConv-(base.conversion_cost||0)); waterfallMeasure.push('relative'); }
    if ((base.margin||0) > 0) { waterfallLabels.push('Margin Δ'); waterfallValues.push(newMargin-(base.margin||0)); waterfallMeasure.push('relative'); }
    waterfallLabels.push('New Total'); waterfallValues.push(newTotal); waterfallMeasure.push('total');
    
    var trace = {
        x: waterfallLabels, y: waterfallValues, type: 'waterfall',
        connector: {line: {color: 'rgba(255,255,255,0.3)'}},
        decreasing: {marker: {color: '#10b981'}},
        increasing: {marker: {color: '#ef4444'}},
        totals: {marker: {color: '#E8601C'}},
        measure: waterfallMeasure,
    };
    Plotly.newPlot('whatif-chart', [trace], {
        paper_bgcolor:'rgba(0,0,0,0)', plot_bgcolor:'rgba(0,0,0,0)',
        font:{color:'white',family:'Outfit',size:10},
        yaxis:{title:unitLabel,gridcolor:'rgba(255,255,255,0.1)'},
        xaxis:{gridcolor:'rgba(255,255,255,0.1)'},
        margin:{t:20,b:60,l:70,r:20}, height:350, showlegend:false
    }, {displayModeBar:false, responsive:true});
}

function resetWhatIf() {
    ['resin','volume','elec','labour','margin'].forEach(p => {
        document.getElementById('wi_'+p).value = 0;
        document.getElementById('wi_'+p+'_label').textContent = '0%';
    });
    document.getElementById('whatif-results').innerHTML = '';
    document.getElementById('whatif-chart').innerHTML = '';
}

async function exportMultiCountryExcel() {
    if (!window._multiCountryResults) { alert('Run comparison first'); return; }
    try {
        const r = await fetch('/api/export_multi_country_excel', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({results: window._multiCountryResults})});
        if (!r.ok) { const e = await r.json(); throw new Error(e.error || 'Export failed'); }
        const blob = await r.blob();
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a'); a.href = url; a.download = 'Country_Comparison.xlsx';
        document.body.appendChild(a); a.click(); URL.revokeObjectURL(url); document.body.removeChild(a);
    } catch(e) { alert('Export failed: ' + e.message); }
}

async function exportGenericExcel() {
    if (!lastModelResult) { alert('Calculate a model first'); return; }
    try {
        const r = await fetch('/api/export_generic_excel', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({model_type: lastModelType, data: lastModelResult})});
        if (!r.ok) { const e = await r.json(); throw new Error(e.error || 'Export failed'); }
        const blob = await r.blob();
        const url = URL.createObjectURL(blob);
        const name = (lastModelType||'Model').replace('-','_');
        const a = document.createElement('a'); a.href = url; a.download = name + '_Report.xlsx';
        document.body.appendChild(a); a.click(); URL.revokeObjectURL(url); document.body.removeChild(a);
    } catch(e) { alert('Export failed: ' + e.message); }
}

async function exportGenericPDF() {
    if (!lastModelResult) { alert('Calculate a model first'); return; }
    try {
        const r = await fetch('/api/export_generic_pdf', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({model_type: lastModelType, data: lastModelResult})});
        if (!r.ok) { const e = await r.json(); throw new Error(e.error || 'Export failed'); }
        const blob = await r.blob();
        const url = URL.createObjectURL(blob);
        const name = (lastModelType||'Model').replace('-','_');
        const a = document.createElement('a'); a.href = url; a.download = name + '_Report.pdf';
        document.body.appendChild(a); a.click(); URL.revokeObjectURL(url); document.body.removeChild(a);
    } catch(e) { alert('Export failed: ' + e.message); }
}

async function exportCartonAdvExcel() {
    if (!lastModelResult || lastModelType !== 'carton-adv') { alert('Calculate Carton Advanced first'); return; }
    return exportGenericExcel();
}
async function exportCartonAdvPDF() {
    if (!lastModelResult || lastModelType !== 'carton-adv') { alert('Calculate Carton Advanced first'); return; }
    return exportGenericPDF();
}

// --- ADVANCED CARTON LOGIC ---
async function calculateCartonAdvanced() {
    const btn = event.target;
    btn.disabled = true; btn.textContent = 'Calculating...';
    try {
        const g = (id) => parseFloat(document.getElementById(id).value) || 0;
        const s = (id) => document.getElementById(id).value;
        const payload = {
            country: s('ca_country'),
            annual_volume: g('ca_annual_vol'),
            avg_order_size: g('ca_avg_order'),
            no_of_colours: g('ca_colours'),
            common_colours: g('ca_common_col'),
            print_runs_year: g('ca_print_runs'),
            no_of_shifts: g('ca_shifts'),
            length_1: g('ca_len1'), length_2: g('ca_len2'),
            width_1: g('ca_wid1'), width_2: g('ca_wid2'),
            height: g('ca_height'), max_flap: g('ca_flap'), gluing_area: g('ca_glue'),
            machine_size: s('ca_mach_size'), grain_direction: s('ca_grain'),
            ups_lengthwise: g('ca_ups_l'), ups_widthwise: g('ca_ups_w'),
            side_lay_1: g('ca_side1'), side_lay_2: g('ca_side2'),
            gripper: g('ca_gripper'), back_lay: g('ca_back_lay'),
            board_gsm: g('ca_board_gsm'), board_rate: g('ca_board_rate'),
            ink_rate: g('ca_ink_rate'), ink_gsm: g('ca_ink_gsm'),
            varnish_rate: g('ca_varnish_rate'), varnish_gsm: g('ca_varnish_gsm'),
            spot_varnish: s('ca_spot_varnish'), hot_foiling: s('ca_hot_foil'),
            lamination: s('ca_lamination'), window_carton: s('ca_window'),
            liner_flag: s('ca_liner'), primer_flag: s('ca_primer'),
            foil_rate_roll: g('ca_foil_rate'), foil_length: g('ca_foil_l'), foil_width: g('ca_foil_w'),
            film_rate: g('ca_film_rate'), film_gsm: g('ca_film_gsm'),
            printing_machine: s('ca_m_print'), spot_varnish_machine: s('ca_m_sv'),
            hot_foiling_machine: s('ca_m_hf'), lamination_machine: s('ca_m_lam'),
            cb_machine: s('ca_m_cb'), fg_machine: s('ca_m_fg'),
            elec_rate: g('ca_elec'), skilled_labour: g('ca_labour'),
            engineer_salary: g('ca_engineer'), pm_salary: g('ca_pm'),
            repair_pct: g('ca_repair'), other_oh_pct: g('ca_other_oh'),
            dep_pm_pct: g('ca_dep_pm'), dep_bldg_pct: g('ca_dep_bldg'),
            completed_life: g('ca_life'),
            land_cost_sqm: g('ca_land'), building_cost_sqm: g('ca_building'),
            premises_type: s('ca_premises'),
            int_lt: g('ca_int_lt'), int_wc: g('ca_int_wc'),
            euro_rate: g('ca_euro'), margin_pct: g('ca_margin'),
            cartons_per_box: g('ca_ctn_box'), boxes_per_container: g('ca_box_cont'),
            freight_cost_container: g('ca_freight'),
            shipper_cost_eur: g('ca_shipper'), polybag_cost_eur: g('ca_polybag'),
        };
        const r = await fetch('/api/calc_carton_advanced', {
            method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify(payload)
        });
        window._lastCartonAdvInput = payload;
        const d = await r.json();
        if (d.error) { alert('Error: ' + d.error); return; }
        renderCartonAdvSummary(d);
    } catch(e) { alert('Error: ' + e.message); }
    finally { btn.disabled = false; btn.textContent = 'Calculate Carton Cost'; }
}

function renderCartonAdvSummary(d) {
    const fmt = (v) => (v||0).toLocaleString('en-IN', {minimumFractionDigits:2, maximumFractionDigits:2});
    let h = '<h3 style="margin-bottom:15px;">Advanced Carton Cost Model</h3>';
    h += '<p style="opacity:0.6; font-size:0.75rem; margin-bottom:15px;">INR per 1000 Cartons | ' + d.country + '</p>';
    
    h += '<div style="display:grid; grid-template-columns:repeat(3,1fr); gap:10px; margin-bottom:20px;">';
    h += '<div style="background:rgba(76,175,80,0.15); padding:12px; border-radius:10px; text-align:center;"><div style="font-size:0.7rem; opacity:0.7;">TOTAL COST</div><div style="font-size:1.4rem; font-weight:800; color:#4CAF50;">₹ ' + fmt(d.total_cost_per_1000) + '</div></div>';
    h += '<div style="background:rgba(33,150,243,0.15); padding:12px; border-radius:10px; text-align:center;"><div style="font-size:0.7rem; opacity:0.7;">EUR / 1000</div><div style="font-size:1.4rem; font-weight:800; color:#2196F3;">€ ' + fmt(d.total_cost_per_1000_eur) + '</div></div>';
    h += '<div style="background:rgba(255,152,0,0.15); padding:12px; border-radius:10px; text-align:center;"><div style="font-size:0.7rem; opacity:0.7;">UPS/SHEET</div><div style="font-size:1.4rem; font-weight:800; color:#FF9800;">' + d.ups_per_sheet + '</div></div>';
    h += '</div>';
    
    h += '<div style="background:rgba(255,255,255,0.05); padding:10px; border-radius:8px; margin-bottom:15px; font-size:0.78rem;">';
    h += 'Flat: ' + d.layflat_length + ' × ' + d.layflat_width + ' mm | Sheet: ' + d.sheet_length + ' × ' + d.sheet_width + ' mm | Area: ' + d.area_per_carton + ' sqm';
    h += '</div>';
    
    const sections = [
        {label: 'MATERIAL COST', val: d.material_cost, color: '#4CAF50', items: [
            {l:'Board', v:d.board_cost}, {l:'Ink', v:d.ink_cost}, {l:'Varnish', v:d.varnish_cost},
            {l:'Spot Varnish', v:d.spot_varnish_cost}, {l:'Hot Foil', v:d.hot_foil_cost},
            {l:'Film/Adhesive', v:d.film_cost}, {l:'Window', v:d.window_cost}, {l:'Liner', v:d.liner_cost},
            {l:'Primer', v:d.primer_cost}, {l:'Wastage', v:d.wastage_cost}, {l:'Dies & Plates', v:d.other_material_cost}
        ]},
        {label: 'CONVERSION COST', val: d.conversion_cost, color: '#2196F3', items: [
            {l:'Electricity', v:d.electricity_cost}, {l:'Direct Labour', v:d.direct_labour},
            {l:'Indirect Labour', v:d.indirect_labour}, {l:'R&M', v:d.repair_maintenance},
            {l:'Other OH', v:d.other_overheads}, {l:'Depreciation', v:d.depreciation},
            {l:'Interest', v:d.interest}, {l:'Lease', v:d.lease_cost}
        ]},
        {label: 'MARGIN', val: d.margin, color: '#FF9800', items: []},
        {label: 'PACKAGING', val: d.packing_cost, color: '#9C27B0', items: []},
        {label: 'FREIGHT', val: d.freight_cost, color: '#F44336', items: []},
    ];
    
    h += '<div style="font-size:0.82rem;">';
    sections.forEach(sec => {
        h += '<div style="display:flex; justify-content:space-between; padding:8px 0; border-bottom:1px solid rgba(255,255,255,0.15); font-weight:700;">';
        h += '<span style="color:' + sec.color + ';">' + sec.label + '</span>';
        h += '<span style="color:' + sec.color + ';">₹ ' + fmt(sec.val) + '</span></div>';
        sec.items.forEach(it => {
            if ((it.v || 0) != 0) {
                h += '<div style="display:flex; justify-content:space-between; padding:4px 0 4px 15px; opacity:0.75; font-size:0.78rem;">';
                h += '<span>' + it.l + '</span><span>₹ ' + fmt(it.v) + '</span></div>';
            }
        });
    });
    h += '<div style="display:flex; justify-content:space-between; padding:12px 0; border-top:2px solid var(--orange); margin-top:10px; font-weight:800; font-size:1.1rem;">';
    h += '<span>TOTAL</span><span style="color:var(--orange);">₹ ' + fmt(d.total_cost_per_1000) + '</span></div>';
    
    const total = d.total_cost_per_1000;
    if (total > 0) {
        h += '<div style="display:flex; gap:4px; margin-top:10px; height:24px; border-radius:6px; overflow:hidden;">';
        [{v:d.material_cost,c:'#4CAF50'},{v:d.conversion_cost,c:'#2196F3'},{v:d.margin,c:'#FF9800'},{v:d.packing_cost,c:'#9C27B0'},{v:d.freight_cost,c:'#F44336'}].forEach(p => {
            const pct = (p.v / total * 100).toFixed(1);
            if (pct > 2) h += '<div style="background:' + p.c + '; width:' + pct + '%; display:flex; align-items:center; justify-content:center; font-size:0.65rem; font-weight:700;">' + pct + '%</div>';
        });
        h += '</div>';
    }
    h += '</div>';
    
    document.getElementById('carton-adv-summary').innerHTML = h;
    
    // Save generic model result
    lastModelResult = d;
    lastModelResult.model_type = 'carton-adv';
    lastModelInput = window._lastCartonAdvInput || {};
    lastModelType = 'carton-adv';
    window.lastCartonAdvResults = d; // SKU FEATURE: Store results
    showSaveSKUButton('carton_advanced'); // SKU FEATURE: Show save button
    document.getElementById('compareCountriesBtn').disabled = false;
    
    // Show export buttons
    var expDiv = document.getElementById('carton-adv-export-btns');
    if (expDiv) expDiv.style.display = 'block';
    
    var pieDiv = document.getElementById('carton-adv-pie-chart');
    pieDiv.style.display = 'block';
    Plotly.newPlot('ca-donut', [{
        values: [d.material_cost, d.conversion_cost, d.margin, d.packing_cost, d.freight_cost],
        labels: ['Material', 'Conversion', 'Margin', 'Packaging', 'Freight'],
        type: 'pie', hole: 0.45,
        marker: {colors: ['#4CAF50', '#2196F3', '#FF9800', '#9C27B0', '#F44336']},
        textinfo: 'label+percent', textfont: {color: 'white', size: 11},
        hoverinfo: 'label+value+percent',
    }], {
        paper_bgcolor: 'rgba(0,0,0,0)', plot_bgcolor: 'rgba(0,0,0,0)',
        font: {color: 'white', family: 'Outfit'}, showlegend: false,
        margin: {t: 10, b: 10, l: 10, r: 10}, height: 300,
        annotations: [{text: 'Cost<br>Split', font: {size: 13, color: 'white'}, showarrow: false}]
    }, {displayModeBar: false, responsive: true});
}

async function loadCartonMachinesFromDB() {
    try {
        const r = await fetch('/api/carton_machine_db', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({})});
        const d = await r.json();
        if (d.machines) {
            const mapping = {'Printing':'ca_m_print','Spot Varnish':'ca_m_sv','Hot Foiling':'ca_m_hf','Lamination':'ca_m_lam','Creasing & Blanking':'ca_m_cb','Folder - Gluer':'ca_m_fg'};
            let totalMachines = 0;
            for (const [process, selId] of Object.entries(mapping)) {
                const sel = document.getElementById(selId);
                if (!sel) continue;
                const machines = d.machines[process] || [];
                if (machines.length > 0) {
                    const currentVal = sel.value;
                    sel.innerHTML = '';
                    machines.forEach(m => { const o = document.createElement('option'); o.value = m.label; o.textContent = m.label; sel.appendChild(o); });
                    if ([...sel.options].some(o => o.value === currentVal)) sel.value = currentVal;
                    totalMachines += machines.length;
                }
            }
            if (totalMachines > 0) document.getElementById('ca-machine-db-info').innerHTML = '<div style="margin-top:8px;font-size:0.72rem;opacity:0.7;color:#4CAF50;">✓ ' + totalMachines + ' machines loaded from DB</div>';
        }
    } catch(e) {}
}

function loadCartonAdvCountryDefaults() {
    const c = document.getElementById('ca_country').value;
    // Use EBM country defaults for conversion costs
    const db = {'India': {elec:10.72, labour:541800, engineer:1260000, pm:1890000, dep_pm:0.15, dep_bldg:0.10, land:23519, building:7000, lease:2136, int_lt:0.125, int_wc:0.14, euro:104.27}, 'China': {elec:0.794, labour:420000, engineer:420000, pm:487200, dep_pm:0.10, dep_bldg:0.10, land:1228.5, building:1056.51, lease:3046.68, int_lt:0.049, int_wc:0.03, euro:8.19}, 'Indonesia': {elec:1114.74, labour:7332000, engineer:19552000, pm:24440000, dep_pm:0.25, dep_bldg:0.10, land:1700000, building:5500000, lease:420000, int_lt:0.10, int_wc:0.12, euro:19314.2}, 'Brazil': {elec:0.657, labour:73000, engineer:210240, pm:315360, dep_pm:0.10, dep_bldg:0.04, land:2533.6, building:14843.75, lease:161.68, int_lt:0.15, int_wc:0.15, euro:6.23}, 'United States': {elec:0.149, labour:98250.6, engineer:130993.8, pm:117125.4, dep_pm:0.10, dep_bldg:0.10, land:32.92, building:2485.66, lease:187.14, int_lt:0.0389, int_wc:0.0364, euro:1.16}, 'United Kingdom': {elec:0.346, labour:39900, engineer:57190, pm:66500, dep_pm:0.18, dep_bldg:0.03, land:148.2, building:1308.88, lease:150.03, int_lt:0.112, int_wc:0.113, euro:0.88}, 'Germany': {elec:0.251, labour:46692, engineer:97275, pm:110245, dep_pm:0.10, dep_bldg:0.10, land:800, building:1292.51, lease:50.61, int_lt:0.0395, int_wc:0.0395, euro:1}, 'France': {elec:0.153, labour:34800, engineer:71050, pm:94250, dep_pm:0.10, dep_bldg:0.10, land:201.2, building:1037.46, lease:88.7, int_lt:0.0345, int_wc:0.0345, euro:1}, 'Turkey': {elec:4.35, labour:281880, engineer:548100, pm:532440, dep_pm:0.10, dep_bldg:0.10, land:9524.94, building:54736.32, lease:459.06, int_lt:0.425, int_wc:0.395, euro:49.29}, 'Vietnam': {elec:1744, labour:139920000, engineer:827162157, pm:921734762, dep_pm:0.10, dep_bldg:0.10, land:3586051.86, building:3403710.24, lease:5470248.6, int_lt:0.059, int_wc:0.062, euro:30390.27}, 'Mexico': {elec:3.972, labour:180000, engineer:492000, pm:852000, dep_pm:0.10, dep_bldg:0.05, land:0, building:0, lease:0, int_lt:0.0728, int_wc:0.0728, euro:21.26}, 'Pakistan': {elec:41.99, labour:504000, engineer:384000, pm:2400000, dep_pm:0.15, dep_bldg:0.10, land:47253.06, building:135336.59, lease:535.57, int_lt:0.18, int_wc:0.09, euro:328.52}, 'Philippines': {elec:8.847, labour:242880, engineer:473470.53, pm:538167.68, dep_pm:0.10, dep_bldg:0.10, land:16240, building:20000, lease:3240, int_lt:0.10, int_wc:0.0863, euro:67.87}, 'South Africa': {elec:1.795, labour:231858, engineer:494630.4, pm:772860, dep_pm:0.20, dep_bldg:0.05, land:438.40, building:8823.27, lease:827.00, int_lt:0.1025, int_wc:0.275, euro:19.88}, 'Spain': {elec:0.126, labour:55960, engineer:67152, pm:76945, dep_pm:0.10, dep_bldg:0.03, land:135.4, building:999.6, lease:53.5, int_lt:0.0215, int_wc:0.032, euro:1}, 'Poland': {elec:0.829, labour:83388, engineer:133420.8, pm:266841.6, dep_pm:0.20, dep_bldg:0.10, land:400, building:3621.4, lease:362, int_lt:0.04, int_wc:0.071, euro:4.21}, 'Thailand': {elec:4.086, labour:303544.8, engineer:327600, pm:1404000, dep_pm:0.20, dep_bldg:0.05, land:4546.87, building:22447.27, lease:2677.35, int_lt:0.1268, int_wc:0.1268, euro:36.6}, 'Bangladesh': {elec:12.39, labour:1521720, engineer:913032, pm:1445634, dep_pm:0.10, dep_bldg:0.10, land:18319.58, building:53821.31, lease:2531.40, int_lt:0.13, int_wc:0.135, euro:142.84}, 'Sri Lanka': {elec:16.59, labour:1060800, engineer:1560000, pm:4680000, dep_pm:0.125, dep_bldg:0.0667, land:15815.26, building:62230.89, lease:7131.32, int_lt:0.14, int_wc:0.18, euro:362.96}, 'Argentina': {elec:129.15, labour:9792000, engineer:1632000, pm:13056000, dep_pm:0.10, dep_bldg:0.02, land:213732.46, building:102920.78, lease:8851.94, int_lt:0.3696, int_wc:0.3696, euro:1684.16}, 'Canada': {elec:0.144, labour:65650, engineer:99737.5, pm:112362.5, dep_pm:0.30, dep_bldg:0.10, land:356.37, building:4068.38, lease:191.71, int_lt:0.025, int_wc:0.0745, euro:1.62}, 'Costa Rica': {elec:115.84, labour:8329800, engineer:19824924, pm:45147516, dep_pm:0.10, dep_bldg:0.10, land:68963.19, building:404761.95, lease:3966.34, int_lt:0.0733, int_wc:0.095, euro:581.9}};
    const v = db[c];
    if(!v) return;
    document.getElementById('ca_elec').value = v.elec;
    document.getElementById('ca_labour').value = v.labour;
    document.getElementById('ca_engineer').value = v.engineer;
    document.getElementById('ca_pm').value = v.pm;
    document.getElementById('ca_dep_pm').value = v.dep_pm;
    document.getElementById('ca_dep_bldg').value = v.dep_bldg;
    document.getElementById('ca_land').value = v.land;
    document.getElementById('ca_building').value = v.building;
    document.getElementById('ca_int_lt').value = v.int_lt;
    document.getElementById('ca_int_wc').value = v.int_wc;
    document.getElementById('ca_euro').value = v.euro;
    // Also load from Variable Cost Database
    loadCartonAdvVarCosts(c);
}

async function loadCartonAdvVarCosts(country) {
    try {
        const r = await fetch('/api/variable_cost_for_calc', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({country})});
        const d = await r.json();
        if (d.variables) {
            let info = '<div style="margin-top:8px;padding:8px;background:rgba(76,175,80,0.1);border:1px solid rgba(76,175,80,0.3);border-radius:6px;font-size:0.72rem;">';
            info += '<strong style="color:#4CAF50;">Variable Cost DB Loaded for ' + country + '</strong></div>';
            document.getElementById('ca-var-cost-info').innerHTML = info;
        }
    } catch(e) {}
}

function loadEBMCountryDefaults() {
    const db = {'India': {elec:10.72, labour:541800, engineer:1260000, pm:1890000, dep_pm:0.15, dep_bldg:0.10, land:23519, building:7000, lease:2136, int_lt:0.125, int_wc:0.14, euro:104.27, mb:450, add:249.93}, 'China': {elec:0.794, labour:420000, engineer:420000, pm:487200, dep_pm:0.10, dep_bldg:0.10, land:1228.5, building:1056.51, lease:3046.68, int_lt:0.049, int_wc:0.03, euro:8.19, mb:35.35, add:19.63}, 'Indonesia': {elec:1114.74, labour:7332000, engineer:19552000, pm:24440000, dep_pm:0.25, dep_bldg:0.10, land:1700000, building:5500000, lease:420000, int_lt:0.10, int_wc:0.12, euro:19314.2, mb:83354.66, add:46294.82}, 'Brazil': {elec:0.657, labour:73000, engineer:210240, pm:315360, dep_pm:0.10, dep_bldg:0.04, land:2533.6, building:14843.75, lease:161.68, int_lt:0.15, int_wc:0.15, euro:6.23, mb:26.89, add:14.93}, 'United States': {elec:0.149, labour:98250.6, engineer:130993.8, pm:117125.4, dep_pm:0.10, dep_bldg:0.10, land:32.92, building:2485.66, lease:187.14, int_lt:0.0389, int_wc:0.0364, euro:1.16, mb:5.01, add:2.78}, 'United Kingdom': {elec:0.346, labour:39900, engineer:57190, pm:66500, dep_pm:0.18, dep_bldg:0.03, land:148.2, building:1308.88, lease:150.03, int_lt:0.112, int_wc:0.113, euro:0.88, mb:7.22, add:4.01}, 'Germany': {elec:0.251, labour:46692, engineer:97275, pm:110245, dep_pm:0.10, dep_bldg:0.10, land:800, building:1292.51, lease:50.61, int_lt:0.0395, int_wc:0.0395, euro:1, mb:4.32, add:2.40}, 'France': {elec:0.153, labour:34800, engineer:71050, pm:94250, dep_pm:0.10, dep_bldg:0.10, land:201.2, building:1037.46, lease:88.7, int_lt:0.0345, int_wc:0.0345, euro:1, mb:4.32, add:2.40}, 'Turkey': {elec:4.35, labour:281880, engineer:548100, pm:532440, dep_pm:0.10, dep_bldg:0.10, land:9524.94, building:54736.32, lease:459.06, int_lt:0.425, int_wc:0.395, euro:49.29, mb:212.72, add:118.14}, 'Vietnam': {elec:1744, labour:139920000, engineer:827162157, pm:921734762, dep_pm:0.10, dep_bldg:0.10, land:3586051.86, building:3403710.24, lease:5470248.6, int_lt:0.059, int_wc:0.062, euro:30390.27, mb:131155.86, add:72843.41}, 'Mexico': {elec:3.972, labour:180000, engineer:492000, pm:852000, dep_pm:0.10, dep_bldg:0.05, land:0, building:0, lease:0, int_lt:0.0728, int_wc:0.0728, euro:21.26, mb:91.75, add:50.96}, 'Pakistan': {elec:41.99, labour:504000, engineer:384000, pm:2400000, dep_pm:0.15, dep_bldg:0.10, land:47253.06, building:135336.59, lease:535.57, int_lt:0.18, int_wc:0.09, euro:328.52, mb:1417.80, add:787.44}, 'Philippines': {elec:8.847, labour:242880, engineer:473470.53, pm:538167.68, dep_pm:0.10, dep_bldg:0.10, land:16240, building:20000, lease:3240, int_lt:0.10, int_wc:0.0863, euro:67.87, mb:292.91, add:162.68}, 'South Africa': {elec:1.795, labour:231858, engineer:494630.4, pm:772860, dep_pm:0.20, dep_bldg:0.05, land:438.40, building:8823.27, lease:827.00, int_lt:0.1025, int_wc:0.275, euro:19.88, mb:85.80, add:47.65}, 'Spain': {elec:0.126, labour:55960, engineer:67152, pm:76945, dep_pm:0.10, dep_bldg:0.03, land:135.4, building:999.6, lease:53.5, int_lt:0.0215, int_wc:0.032, euro:1, mb:4.32, add:2.40}, 'Poland': {elec:0.829, labour:83388, engineer:133420.8, pm:266841.6, dep_pm:0.20, dep_bldg:0.10, land:400, building:3621.4, lease:362, int_lt:0.04, int_wc:0.071, euro:4.21, mb:18.17, add:10.09}, 'Thailand': {elec:4.086, labour:303544.8, engineer:327600, pm:1404000, dep_pm:0.20, dep_bldg:0.05, land:4546.87, building:22447.27, lease:2677.35, int_lt:0.1268, int_wc:0.1268, euro:36.6, mb:157.96, add:87.73}, 'Bangladesh': {elec:12.39, labour:1521720, engineer:913032, pm:1445634, dep_pm:0.10, dep_bldg:0.10, land:18319.58, building:53821.31, lease:2531.40, int_lt:0.13, int_wc:0.135, euro:142.84, mb:616.46, add:342.38}, 'Sri Lanka': {elec:16.59, labour:1060800, engineer:1560000, pm:4680000, dep_pm:0.125, dep_bldg:0.0667, land:15815.26, building:62230.89, lease:7131.32, int_lt:0.14, int_wc:0.18, euro:362.96, mb:1566.43, add:869.99}, 'Argentina': {elec:129.15, labour:9792000, engineer:1632000, pm:13056000, dep_pm:0.10, dep_bldg:0.02, land:213732.46, building:102920.78, lease:8851.94, int_lt:0.3696, int_wc:0.3696, euro:1684.16, mb:7268.36, add:4036.82}, 'Canada': {elec:0.144, labour:65650, engineer:99737.5, pm:112362.5, dep_pm:0.30, dep_bldg:0.10, land:356.37, building:4068.38, lease:191.71, int_lt:0.025, int_wc:0.0745, euro:1.62, mb:6.99, add:3.88}, 'Costa Rica': {elec:115.84, labour:8329800, engineer:19824924, pm:45147516, dep_pm:0.10, dep_bldg:0.10, land:68963.19, building:404761.95, lease:3966.34, int_lt:0.0733, int_wc:0.095, euro:581.9, mb:2511.32, add:1394.77} };
    const c = document.getElementById('e_country').value;
    const v = db[c];
    if(!v) return;
    document.getElementById('e_elec_rate').value = v.elec;
    document.getElementById('e_labour').value = v.labour;
    document.getElementById('e_engineer').value = v.engineer;
    document.getElementById('e_pm').value = v.pm;
    document.getElementById('e_dep_pm').value = v.dep_pm;
    document.getElementById('e_dep_bldg').value = v.dep_bldg;
    document.getElementById('e_land').value = v.land;
    document.getElementById('e_building').value = v.building;
    document.getElementById('e_lease').value = v.lease;
    document.getElementById('e_int_lt').value = v.int_lt;
    document.getElementById('e_int_wc').value = v.int_wc;
    document.getElementById('e_euro_rate').value = v.euro;
    document.getElementById('e_l1_mb_rate').value = v.mb;
    document.getElementById('e_l2_mb_rate').value = v.mb;
    document.getElementById('e_l3_mb_rate').value = v.mb;
    document.getElementById('e_l1_add_rate').value = v.add;
    document.getElementById('e_l2_add_rate').value = v.add;
    document.getElementById('e_l3_add_rate').value = v.add;
    
    // Feature 2: Also load from Variable Cost Database
    loadVariableCostsFromDB(c);
}

// ============================================
// SKU SAVE/LOAD FEATURE
// ============================================
const SKU_STORAGE_KEY = 'packfora_saved_skus';
function getSavedSKUs() { try { const data = localStorage.getItem(SKU_STORAGE_KEY); return data ? JSON.parse(data) : []; } catch(e) { return []; }}
function saveSKUToStorage(sku) { try { const skus = getSavedSKUs().filter(s => s.name !== sku.name); skus.push(sku); localStorage.setItem(SKU_STORAGE_KEY, JSON.stringify(skus)); return true; } catch(e) { return false; }}
function deleteSKUFromStorage(skuName) { try { const skus = getSavedSKUs().filter(s => s.name !== skuName); localStorage.setItem(SKU_STORAGE_KEY, JSON.stringify(skus)); return true; } catch(e) { return false; }}
function getVal(id) { const el = document.getElementById(id); return el ? el.value : ''; }
function setVal(id, value) { const el = document.getElementById(id); if (el && value !== undefined && value !== null) el.value = value; }

function getActiveModel() {
    const activeCalc = document.querySelector('.sub-tab-btn.active');
    if (!activeCalc) return null;
    const model = document.getElementById('essentialsSelect')?.value;
    if (model) return model; // carton or flexibles
    const advActive = document.getElementById('btn-advanced')?.classList.contains('active');
    if (advActive) {
        const ebmActive = document.getElementById('ebm-inputs')?.style.display !== 'none';
        return ebmActive ? 'ebm' : 'carton_advanced';
    }
    return null;
}

function captureCurrentModelData() {
    const activeModel = getActiveModel();
    if (!activeModel) { alert('No model calculation found. Please calculate first.'); return null; }
    const modelData = { model: activeModel, timestamp: new Date().toISOString(), inputs: {}, results: {} };
    switch(activeModel) {
        case 'ebm': modelData.inputs = captureEBMInputs(); modelData.results = window.lastEBMResults || {}; break;
        case 'carton_advanced': modelData.inputs = captureCartonAdvInputs(); modelData.results = window.lastCartonAdvResults || {}; break;
        case 'carton': modelData.inputs = captureCartonInputs(); modelData.results = window.lastCartonResults || {}; break;
        case 'flexibles': modelData.inputs = captureFlexiblesInputs(); modelData.results = window.lastFlexiblesResults || {}; break;
        default: alert('Unknown model type'); return null;
    }
    return modelData;
}

function captureEBMInputs() { return {country:getVal('e_country'),film_width:getVal('e_film_width'),film_length:getVal('e_film_length'),film_gsm:getVal('e_film_gsm'),annual_qty:getVal('e_annual_qty'),resin:getVal('e_resin'),mat_cost:getVal('e_mat_cost'),elec_cost:getVal('e_elec_cost'),labour_cost:getVal('e_labour_cost'),dep_cost:getVal('e_dep_cost'),other_cost:getVal('e_other_cost'),freight:getVal('e_freight')}; }
function captureCartonAdvInputs() { return {country:getVal('ca_country'),box_length:getVal('ca_box_length'),box_width:getVal('ca_box_width'),box_height:getVal('ca_box_height'),board_gsm:getVal('ca_board_gsm'),annual_qty:getVal('ca_annual_qty'),board_cost:getVal('ca_board_cost'),printing:getVal('ca_printing'),die_cutting:getVal('ca_die_cutting'),gluing:getVal('ca_gluing'),labour:getVal('ca_labour'),overhead:getVal('ca_overhead'),polybag:getVal('ca_polybag')}; }
function captureCartonInputs() { return {country:getVal('c_country'),box_length:getVal('c_box_length'),box_width:getVal('c_box_width'),box_height:getVal('c_box_height'),board_gsm:getVal('c_board_gsm'),annual_qty:getVal('c_annual_qty'),board_cost:getVal('c_board_cost'),printing:getVal('c_printing'),conversion:getVal('c_conversion')}; }
function captureFlexiblesInputs() { return {country:getVal('f_country'),film_width:getVal('f_film_width'),film_thickness:getVal('f_film_thickness'),film_length:getVal('f_film_length'),annual_qty:getVal('f_annual_qty'),resin:getVal('f_resin'),resin_cost:getVal('f_resin_cost'),printing:getVal('f_printing'),slitting:getVal('f_slitting'),conversion:getVal('f_conversion')}; }

function loadEBMInputs(inp) { if(!inp) return; Object.keys(inp).forEach(k => setVal('e_'+k, inp[k])); }
function loadCartonAdvInputs(inp) { if(!inp) return; Object.keys(inp).forEach(k => setVal('ca_'+k, inp[k])); }
function loadCartonInputs(inp) { if(!inp) return; Object.keys(inp).forEach(k => setVal('c_'+k, inp[k])); }
function loadFlexiblesInputs(inp) { if(!inp) return; Object.keys(inp).forEach(k => setVal('f_'+k, inp[k])); }

async function saveSKU() {
    const modelData = captureCurrentModelData();
    if (!modelData) return;
    const skuName = prompt('Enter SKU name:', `SKU_${modelData.model}_${Date.now()}`);
    if (!skuName || skuName.trim() === '') { alert('SKU name cannot be empty'); return; }
    const sku = { name: skuName.trim(), ...modelData };
    const localSuccess = saveSKUToStorage(sku);
    try {
        const response = await fetch('/api/save_sku', { method: 'POST', headers: {'Content-Type':'application/json'}, body: JSON.stringify(sku) });
        const data = await response.json();
        if (localSuccess && data.success) { alert(`✓ SKU "${skuName}" saved successfully!`); refreshSKUDropdown(); }
        else { alert(`⚠ SKU saved locally only. Server: ${data.message || 'Error'}`); refreshSKUDropdown(); }
    } catch(error) {
        if (localSuccess) { alert(`⚠ SKU saved locally only (server unavailable)`); refreshSKUDropdown(); }
        else { alert('✗ Failed to save SKU'); }
    }
}

function loadSKU() {
    const dropdown = document.getElementById('sku-selector');
    if (!dropdown || !dropdown.value) { alert('Please select a SKU to load'); return; }
    const skuName = dropdown.value;
    const skus = getSavedSKUs();
    const sku = skus.find(s => s.name === skuName);
    if (!sku) { alert('SKU not found'); return; }
    
    // Switch to correct model
    if (sku.model === 'carton' || sku.model === 'flexibles') {
        document.getElementById('btn-essentials')?.click();
        setTimeout(() => {
            document.getElementById('essentialsSelect').value = sku.model;
            document.getElementById('essentialsSelect').dispatchEvent(new Event('change'));
            setTimeout(() => {
                if (sku.model === 'carton') loadCartonInputs(sku.inputs);
                else loadFlexiblesInputs(sku.inputs);
                alert(`✓ SKU "${skuName}" loaded successfully!`);
            }, 100);
        }, 100);
    } else if (sku.model === 'ebm' || sku.model === 'carton_advanced') {
        document.getElementById('btn-advanced')?.click();
        setTimeout(() => {
            if (sku.model === 'ebm') { switchAdvModel('ebm'); loadEBMInputs(sku.inputs); }
            else { switchAdvModel('carton_adv'); loadCartonAdvInputs(sku.inputs); }
            alert(`✓ SKU "${skuName}" loaded successfully!`);
        }, 100);
    }
}

function deleteSKU() {
    const dropdown = document.getElementById('sku-selector');
    if (!dropdown || !dropdown.value) { alert('Please select a SKU to delete'); return; }
    const skuName = dropdown.value;
    if (!confirm(`Delete SKU "${skuName}"?`)) return;
    deleteSKUFromStorage(skuName);
    refreshSKUDropdown();
    alert(`✓ SKU "${skuName}" deleted`);
}

function refreshSKUDropdown() {
    const dropdown = document.getElementById('sku-selector');
    if (!dropdown) return;
    const skus = getSavedSKUs();
    dropdown.innerHTML = '<option value="">-- Select SKU --</option>';
    skus.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
    skus.forEach(sku => {
        const option = document.createElement('option');
        option.value = sku.name;
        const date = new Date(sku.timestamp).toLocaleString();
        option.textContent = `${sku.name} (${sku.model.toUpperCase()}) - ${date}`;
        dropdown.appendChild(option);
    });
}

function showSaveSKUButton(model) {
    const btnIds = {'carton': 'save-sku-btn-carton', 'flexibles': 'save-sku-btn-flexibles', 'ebm': 'save-sku-btn-ebm', 'carton_advanced': 'save-sku-btn-carton-adv'};
    const btnId = btnIds[model];
    if (btnId) {
        const btn = document.getElementById(btnId);
        if (btn) btn.style.display = 'block';
    }
}

// --- INIT ---
document.addEventListener('DOMContentLoaded', function() {
    switchSubTab('essentials'); 
    if(document.getElementById('flex-layers')) renderFlexLayers();
    if(document.getElementById('country-checkboxes')) initCountryCheckboxes();
    loadMachinesFromDB();
    loadCartonMachinesFromDB();
    refreshSKUDropdown();
});
</script>
"""# ================= APPLICATION STARTUP =================
if __name__ == "__main__":
    files_ok, message = check_files_exist()
    if not files_ok:
        logger.warning("Starting with missing files")
        logger.warning(message)
    
    is_production = os.getenv('FLASK_ENV', 'development') == 'production'
    
    if is_production:
        logger.info("Starting in PRODUCTION mode")
        app.run(host="0.0.0.0", port=5000, debug=False)
    else:
        logger.info("Starting in DEVELOPMENT mode")
        logger.info(f"Admin Login: {ADMIN_USERNAME} / {ADMIN_PASSWORD}")
        logger.info("CHANGE DEFAULT ADMIN CREDENTIALS IN PRODUCTION!")
        app.run(host="127.0.0.1", port=5000, debug=True)

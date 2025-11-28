import base64
import hashlib
from cryptography.fernet import Fernet
import secrets
from dataclasses import dataclass, field
from enum import Enum, auto
from pathlib import Path
from typing import Dict, List, Optional, TypedDict, Literal, Set, Any

# --- Application Version (Semantic Versioning) ---
APP_VERSION = "1.0.0"

# --- Enums for Type Safety ---
class ThemeColor(str, Enum):
    """Available theme colors for the application."""
    BLUE = "blue"
    GREEN = "green"
    RED = "red"
    OCHRE = "ochre"
    ORCHID = "orchid"

class ReportFormat(str, Enum):
    """Available report output formats."""
    PDF = "PDF"
    EXCEL = "Excel"

class SecurityLayer(int, Enum):
    """Security layer identifiers for multi-layer encryption."""
    FERNET = 1
    XOR = 2
    SCRAMBLE = 3

# --- Report Settings ---
REPORT_FORMATS: List[ReportFormat] = [ReportFormat.PDF, ReportFormat.EXCEL]
REPORT_FORMAT: ReportFormat = ReportFormat.EXCEL  # Default format

# --- Internal Security Version (not exposed to users) ---
_SECURITY_VERSION = "v5.1_triple_layer"

def create_multi_layer_key(salt: bytes, layer: SecurityLayer) -> bytes:
    return hashlib.pbkdf2_hmac('sha256', salt + str(layer).encode(), b'DOLE_DIA_v5.1', 150000)

def protect_sensitive_data(text: str) -> str:
    # Layer 1: Fernet encryption
    layer1_key = create_multi_layer_key(ENCRYPTION_SALT, SecurityLayer.FERNET)
    fernet = Fernet(base64.urlsafe_b64encode(layer1_key[:32]))
    layer1_encrypted = fernet.encrypt(text.encode())
    
    # Layer 2: Custom XOR with derived key
    layer2_key = create_multi_layer_key(ENCRYPTION_SALT, SecurityLayer.XOR)
    layer2_data = bytearray(layer1_encrypted)
    for i in range(len(layer2_data)):
        layer2_data[i] ^= layer2_key[i % len(layer2_key)]
    
    # Layer 3: Final scrambling with base64 encoding
    layer3_key = create_multi_layer_key(ENCRYPTION_SALT, SecurityLayer.SCRAMBLE)
    scrambled = bytearray(layer2_data)
    for i in range(len(scrambled)):
        scrambled[i] = (scrambled[i] + layer3_key[i % len(layer3_key)]) % 256
    
    return base64.urlsafe_b64encode(scrambled).decode()

def unprotect_sensitive_data(encrypted_text: str) -> str:
    # First try: Legacy Fernet decryption (for existing data)
    try:
        # Try to decrypt as legacy Fernet (old system) - using the same key derivation as data_utils
        from cryptography.hazmat.primitives import hashes
        from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
        kdf = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=ENCRYPTION_SALT, iterations=480000)
        legacy_key = base64.urlsafe_b64encode(kdf.derive("doleadmin".encode()))
        fernet = Fernet(legacy_key)
        decrypted = fernet.decrypt(encrypted_text)
        return decrypted.decode()
    except Exception:
        pass
    
    # Second try: New triple-layer decryption
    try:
        # Layer 3: Reverse scrambling
        layer3_key = create_multi_layer_key(ENCRYPTION_SALT, SecurityLayer.SCRAMBLE)
        scrambled = bytearray(base64.urlsafe_b64decode(encrypted_text))
        for i in range(len(scrambled)):
            scrambled[i] = (scrambled[i] - layer3_key[i % len(layer3_key)]) % 256
        
        # Layer 2: Reverse XOR
        layer2_key = create_multi_layer_key(ENCRYPTION_SALT, SecurityLayer.XOR)
        layer2_data = bytearray(scrambled)
        for i in range(len(layer2_data)):
            layer2_data[i] ^= layer2_key[i % len(layer2_key)]
        
        # Layer 1: Fernet decryption
        layer1_key = create_multi_layer_key(ENCRYPTION_SALT, SecurityLayer.FERNET)
        fernet = Fernet(base64.urlsafe_b64encode(layer1_key[:32]))
        decrypted = fernet.decrypt(bytes(layer2_data))
        
        return decrypted.decode()
    except Exception as e:
        raise ValueError(f"Failed to decrypt sensitive data with both legacy and new methods: {e}")

# --- Security & Encryption ---
ENCRYPTION_SALT = b'\x1a\x8b\xcd\xef\x01\x23\x45\x67\x89\xab\xcd\xef\x01\x23\x45\x67'
HIDDEN_PASSWORD = protect_sensitive_data("doleadmin")

# --- Encrypted GitHub Personal Access Token (PAT) ---
ENCRYPTED_GITHUB_TOKEN = 'gAAAAABojBq2InlIwzUrRpB7mCyeRjHPWvx2v2lF-_h3I64gpFGx0FfbAcMEeoP5Zfwu3x9VUInAlTaACxe4MM5U62LIE_zzH1dwF6-Zzq1uDy8oe6rVfKUZB_8CM5oEC2PJ3Ujk04Ui'

# --- Simple Threshold Adjustment Configuration ---
ADAPTIVE_MATCHING_CONFIG = {
    "enable_adaptive_mode": True,  # Set to True to enable conservative adjustments for datasets without birthdate/sex
    
    # Simple threshold adjustments (start small and increase gradually)
    "threshold_adjustments": {
        "strict_adjustment": 0,       # Add to original 198 (start with 0, increase as needed)
        "standard_adjustment": -20,   # Adjusted for name-only datasets (no birthdate, sex, city): 110-20=90
        "lenient_adjustment": -5      # Adjusted for name-only datasets: 95-5=90
    },
    
    # Your proven baseline thresholds
    "baseline_thresholds": {
        "strict_threshold": 198,
        "standard_threshold": 110,    # Your carefully tuned threshold
        "lenient_threshold": 95
    }
}

# --- Global Configuration (Shared by all provinces) ---
GLOBAL_CONFIG = {
    "NICKNAME_CSV_URL": "https://raw.githubusercontent.com/DOLE-MIMAROPA/MIMAROPA-DATABASE/main/Nicknames.csv",
    "TEMPLATE_CSV_URL": "https://raw.githubusercontent.com/DOLE-MIMAROPA/MIMAROPA-DATABASE/main/Template.csv"
}

# --- Type Definitions ---
@dataclass(frozen=True)
class ThemeColors:
    """Theme color scheme configuration."""
    fg_color: List[str]
    hover_color: List[str]
    hyperlink_color: str

@dataclass(frozen=True)
class ProvinceURLs:
    """URLs for a province's data sources."""
    master_db: str
    officials: str
    
    def __post_init__(self):
        """Validate URLs on initialization."""
        if not all(url.startswith('http') for url in [self.master_db, self.officials]):
            raise ValueError("All URLs must be valid HTTP/HTTPS links")

@dataclass
class ProvinceConfig:
    """Configuration for a specific province."""
    name: str
    title: str
    theme: 'ThemeColor'
    urls: ProvinceURLs
    
    def __post_init__(self):
        """Convert theme to ThemeColor enum if it's a string."""
        if isinstance(self.theme, str):
            self.theme = ThemeColor(self.theme.lower())

# --- THEME DEFINITIONS ---
# Convert theme colors to ThemeColors objects
THEME_COLORS: Dict[ThemeColor, ThemeColors] = {
    ThemeColor.BLUE: ThemeColors(
        fg_color=["#3B8ED0", "#1F6AA5"],
        hover_color=["#36719F", "#144870"],
        hyperlink_color="#FFC107"
    ),
    ThemeColor.GREEN: ThemeColors(
        fg_color=["#2CC985", "#2FA572"],
        hover_color=["#17A76A", "#14754B"],
        hyperlink_color="#FFC107"
    ),
    ThemeColor.RED: ThemeColors(
        fg_color=["#D32F2F", "#B71C1C"],
        hover_color=["#E57373", "#C62828"],
        hyperlink_color="#FFFFE0"
    ),
    ThemeColor.OCHRE: ThemeColors(
        fg_color=["#D9882B", "#B56D19"],
        hover_color=["#E8A14A", "#D9882B"],
        hyperlink_color="#FFFFFF"
    ),
    ThemeColor.ORCHID: ThemeColors(
        fg_color=["#C042D8", "#A42ABF"],
        hover_color=["#D863EE", "#C042D8"],
        hyperlink_color="#FFFFFF"
    )
}

def validate_province_config(config: dict) -> bool:
    """
    Validate that a province configuration has all required fields.
    
    Args:
        config: The province configuration dictionary to validate
        
    Returns:
        bool: True if the configuration is valid, False otherwise
    """
    required_keys = {"title", "theme", "urls"}
    required_urls = {"master_db", "officials"}
    
    # Check for required top-level keys
    if not all(key in config for key in required_keys):
        return False
    
    # Check for required URLs
    if not all(url in config["urls"] for url in required_urls):
        return False
    
    # Validate theme is a valid ThemeColor
    if not isinstance(config["theme"], ThemeColor):
        try:
            config["theme"] = ThemeColor(config["theme"])
        except ValueError:
            return False
            
    return True

def create_province_configs() -> Dict[str, ProvinceConfig]:
    """Create and validate province configurations."""
    raw_configs = {
        "Oriental Mindoro": {
            "title": f"Deduplication and Official Linkage Engine v{APP_VERSION} (Oriental Mindoro)",
            "theme": ThemeColor.BLUE,
            "urls": {
                "master_db": "https://raw.githubusercontent.com/DOLE-MIMAROPA/MIMAROPA-DATABASE/main/Database-Oriental-Mindoro.csv",
                "officials": "https://raw.githubusercontent.com/DOLE-MIMAROPA/MIMAROPA-DATABASE/main/Officials-Oriental-Mindoro.csv"
            }
        },
        "Occidental Mindoro": {
            "title": f"Deduplication and Official Linkage Engine v{APP_VERSION} (Occidental Mindoro)",
            "theme": ThemeColor.GREEN,
            "urls": {
                "master_db": "https://raw.githubusercontent.com/DOLE-MIMAROPA/MIMAROPA-DATABASE/main/Database-Occidental-Mindoro.csv",
                "officials": "https://raw.githubusercontent.com/DOLE-MIMAROPA/MIMAROPA-DATABASE/main/Officials-Occidental-Mindoro.csv"
            }
        },
        "Marinduque": {
            "title": f"Deduplication and Official Linkage Engine v{APP_VERSION} (Marinduque)",
            "theme": ThemeColor.RED,
            "urls": {
                "master_db": "https://raw.githubusercontent.com/DOLE-MIMAROPA/MIMAROPA-DATABASE/main/Database-Marinduque.csv",
                "officials": "https://raw.githubusercontent.com/DOLE-MIMAROPA/MIMAROPA-DATABASE/main/Officials-Marinduque.csv"
            }
        },
        "Romblon": {
            "title": f"Deduplication and Official Linkage Engine v{APP_VERSION} (Romblon)",
            "theme": ThemeColor.OCHRE,
            "urls": {
                "master_db": "https://raw.githubusercontent.com/DOLE-MIMAROPA/MIMAROPA-DATABASE/main/Database-Romblon.csv",
                "officials": "https://raw.githubusercontent.com/DOLE-MIMAROPA/MIMAROPA-DATABASE/main/Officials-Romblon.csv"
            }
        },
        "Palawan": {
            "title": f"Deduplication and Official Linkage Engine v{APP_VERSION} (Palawan)",
            "theme": ThemeColor.ORCHID,
            "urls": {
                "master_db": "https://raw.githubusercontent.com/DOLE-MIMAROPA/MIMAROPA-DATABASE/main/Database-Palawan.csv",
                "officials": "https://raw.githubusercontent.com/DOLE-MIMAROPA/MIMAROPA-DATABASE/main/Officials-Palawan.csv"
            }
        }
    }
    
    # Convert raw configs to ProvinceConfig objects
    configs = {}
    for name, config in raw_configs.items():
        try:
            urls = ProvinceURLs(**config["urls"])
            configs[name] = ProvinceConfig(
                name=name,
                title=config["title"],
                theme=config["theme"],
                urls=urls
            )
        except Exception as e:
            raise ValueError(f"Invalid configuration for province {name}: {e}")
    
    return configs

# --- Province-Specific Profiles ---
PROVINCE_PROFILES = create_province_configs()

# Validation is now handled by the dataclass constructors

# --- Data Structure Constants ---
MATCHING_COLS = ["First Name", "Middle Name", "Last Name", "Suffix", "Birthdate", "City", "Sex"]
INTENDED_COLS = ["First Name", "Middle Name", "Last Name", "Suffix", "Birthdate", "City", "Sex", "Contact Number"]
COLUMN_ALIASES = {
    "First Name": {"firstname", "first_name", "first name", "fname"},
    "Middle Name": {"middlename", "middle_name", "middle name", "mname"},
    "Last Name": {"lastname", "last_name", "last name", "lname"},
    "Suffix": {"suffix"},
    "Birthdate": {"birthdate", "dob", "dateofbirth", "date_of_birth"},
    "City": {
        "city",
        # City/Municipality variants
        "city/municipality", "city-municipality", "citymunicipality",
        "municipality/city", "municipality-city", "municipalitycity",
        # Municipality-only variants and abbreviations
        # Note: "MUNICIPALITY" (new header) normalizes to "municipality" which matches below
        "municipality", "muncipality", "municipal", "mun", "mun.", "city/mun", "city/mun.", "citymun",
        # Town variants sometimes used in place of city
        "town", "city/town", "town/city", "city-town", "town-city"
    },
    "Sex": {"sex", "gender"},
    "Contact Number": {
        "contactnumber", "contact_number", "contact no", "contactno", "contact",
        "phone", "phonenumber", "landline",
        "cellphone", "cellphoneno", "cellphonenumber",
        "mobile", "mobileno", "mobilenumber",
        "cpno", "cpnumber",
        "telephone", "telephoneno", "telephonenumber"
    },
    "FullName": {"name", "fullname", "full_name"},
    "Position": {"position"},
    "Barangay": {"barangay"},
    # Batch name is informational only (not used for matching), but we should detect many variants
    # Note: "BATCH No." (new header) normalizes to "batchno." which matches below
    "Batch Name": {
        "batchname", "batch name",
        "batchnumber", "batch number", "batchnum",
        "batchno", "batch no", "batchno.",
        "batch#", "batch", "batchid", "batch id"
    },
}
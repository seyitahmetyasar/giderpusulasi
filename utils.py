import re
from typing import Any, List, Tuple

TR_MAP = str.maketrans("çğıöşüÇĞİÖŞÜ","cgiosuCGIOSU")

def norm(s: Any) -> str:
    return ("" if s is None else str(s)).strip()

def nlow(s: Any) -> str:
    return norm(s).lower().translate(TR_MAP)

def nup(s: Any) -> str:
    return norm(s).upper().translate(TR_MAP)

IMEI_RE_STRICT = re.compile(r'(?<!\d)\d{15}(?!\d)')


def _luhn_ok_imei(s15: str) -> bool:
    if len(s15) != 15 or not s15.isdigit():
        return False
    total = 0
    for i, ch in enumerate(s15):
        n = int(ch)
        if i % 2 == 1:
            n *= 2
            if n > 9:
                n -= 9
        total += n
    return (total % 10) == 0

def extract_imeis(text: str) -> List[str]:
    t = norm(text)
    if not t:
        return []
    out = set()
    for m in IMEI_RE_STRICT.finditer(t):
        s = m.group(0)
        if _luhn_ok_imei(s):
            out.add(s)
    return sorted(out)


def safe_filename(name: str) -> str:
    s = norm(name)
    s = re.sub(r'[\\/*?:"<>|]', "_", s)
    s = re.sub(r"\s+", " ", s).strip()
    return (s or "DOSYA")[:120]


def _cp(p: str) -> re.Pattern:
    return re.compile(p, re.I)

_BRAND_PATTERNS: List[Tuple[str, re.Pattern]] = [
    ("APPLE",   _cp(r"\bAPPLE\b|\bIPHONE?\b|\bAPLE\b|\bI ?PHONE\b")),
    ("SAMSUNG", _cp(r"\bSAMSUNG\b|\bGALAXY\b|\bSM[-\s]")),
    ("XIAOMI",  _cp(r"\bXIAOM[Iİ]\b|\bREDMI\b|\bPOCO\b|\bMI[-\s]")),
    ("HUAWEI",  _cp(r"\bHUAWEI\b|\bHUAWE\b|\bP\d{2}\b|\bMATE\b(?!.*HONOR)")),
    ("HONOR",   _cp(r"\bHONOR\b")),
    ("OPPO",    _cp(r"\bOPPO\b")),
    ("REALME",  _cp(r"\bREALME\b")),
    ("VIVO",    _cp(r"\bVIVO\b")),
    ("TECNO",   _cp(r"\bTECNO\b")),
    ("NOKIA",   _cp(r"\bNOKIA\b")),
    ("CASPER",  _cp(r"\bCASPER\b")),
    ("GENERAL MOBILE", _cp(r"\bGENERAL\s*MOBILE\b|\bGM\s?\d+\b")),
    ("INFINIX", _cp(r"\bINFINIX\b")),
    ("REEDER",  _cp(r"\bREEDER\b")),
]

def brand_from_text(t: str) -> str:
    U = nup(t)
    for brand, pat in _BRAND_PATTERNS:
        if pat.search(U):
            if brand == "HUAWEI" and "HONOR" in U:
                continue
            return brand
    return "Bilinmeyen"

KEY_2EL  = re.compile(r"\b(2\.?\s*EL|İKİNCİ\s*EL|IKINCI\s*EL|SECOND\s*HAND)\b", re.I)
KEY_REF  = re.compile(r"\b(YENİLENMİŞ|YENILENMIS|REFURB)\b", re.I)
DOCNO_RE = re.compile(r"\bE(?:AR|FR)\d{13}\b", re.I)

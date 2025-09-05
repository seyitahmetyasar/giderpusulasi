import math
import threading
from typing import Any, Dict, List, Optional

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# NES API endpoints
EINV_IN_LIST   = "https://api.nes.com.tr/einvoice/v1/incoming/invoices"
EINV_OUT_LIST  = "https://api.nes.com.tr/einvoice/v1/outgoing/invoices"
EARCH_OUT_LIST = "https://api.nes.com.tr/earchive/v1/invoices"

EINV_IN_DOC    = "https://api.nes.com.tr/einvoice/v1/incoming/invoices/{id}"
EINV_OUT_DOC   = "https://api.nes.com.tr/einvoice/v1/outgoing/invoices/{id}"
EARCH_OUT_DOC  = "https://api.nes.com.tr/earchive/v1/invoices/{id}"

PAGE_SIZE = 50

# HTTP session configuration
TIMEOUT_CONNECT = 15
TIMEOUT_READ = 90
RETRIES = 4
BACKOFF = 0.6
SESSION: Optional[requests.Session] = None


def make_session(retries: int = None, backoff: float = None) -> requests.Session:
    r = retries if retries is not None else RETRIES
    b = backoff if backoff is not None else BACKOFF
    s = requests.Session()
    retry = Retry(
        total=r, connect=r, read=r, status=r,
        backoff_factor=b,
        status_forcelist=(429, 500, 502, 503, 504),
        allowed_methods=frozenset(["GET", "POST"]),
        raise_on_status=False,
    )
    adapter = HTTPAdapter(max_retries=retry, pool_connections=100, pool_maxsize=100)
    s.mount("https://", adapter)
    s.mount("http://", adapter)
    return s


def get_session() -> requests.Session:
    global SESSION
    if SESSION is None:
        SESSION = make_session()
    return SESSION


def headers(token: str, typ: str = "json") -> Dict[str, str]:
    h = {
        "Authorization": f"Bearer {token.strip()}",
        "Accept": "application/json",
        "Content-Type": "application/json",
        "User-Agent": "IMEI-NES-Client/10.3",
    }
    if typ == "xml":
        h["Accept"] = "application/xml"
    if typ == "pdf":
        h["Accept"] = "application/pdf"
    if typ == "bin":
        h["Accept"] = "*/*"
    return h


def http_get(
    url: str,
    *,
    token: Optional[str],
    typ: str = "json",
    params=None,
    log=None,
    stop_evt: Optional[threading.Event] = None,
):
    if stop_evt is not None and stop_evt.is_set():
        return None
    try:
        sess = get_session()
        hdrs = headers(token, typ) if token is not None else {"User-Agent": "IMEI-NES-Client/10.3"}
        r = sess.get(url, headers=hdrs, params=params, timeout=(TIMEOUT_CONNECT, TIMEOUT_READ))
        if r.status_code == 200:
            return r
        if log:
            log(f"[HTTP] Hata {r.status_code}: {url} → {r.text[:300]}")
        return None
    except requests.Timeout:
        if log:
            log(f"[HTTP] Zaman aşımı: {url}")
        return None
    except requests.RequestException as e:
        if log:
            log(f"[HTTP] İstek hatası: {e}")
        return None


def paged_list(
    url: str,
    token: str,
    start: str,
    end: str,
    log,
    stop_evt: threading.Event,
    *,
    archived: Optional[bool] = None,
    section_name: str = "NES",
) -> List[Dict[str, Any]]:
    out: List[Dict[str, Any]] = []
    page = 1
    total = None
    while True:
        if stop_evt.is_set():
            break
        params = {"sort": "createdAt desc", "page": page, "pageSize": PAGE_SIZE}
        if archived is not None:
            params["archived"] = "true" if archived else "false"
        if start:
            params["startDate"] = f"{start}T00:00:00+03:00"
        if end:
            params["endDate"] = f"{end}T23:59:59+03:00"
        r = http_get(url, token=token, typ="json", params=params, log=log, stop_evt=stop_evt)
        if r is None:
            log(f"[{section_name}] İstek başarısız/timeout. Döngü sonlandırıldı.")
            break
        try:
            data = r.json() or {}
        except ValueError:
            log(f"[{section_name}] JSON çözümlenemedi.")
            break
        if total is None:
            tc = data.get("totalCount") or 0
            total = max(1, math.ceil(tc / PAGE_SIZE)) if isinstance(tc, int) else 1
        batch = data.get("data") or data.get("invoices") or []
        if not batch:
            break
        log(f"[{section_name}] Sayfa {page}/{total} → {len(batch)} kayıt")
        out.extend(batch)
        if page >= total:
            break
        page += 1
    return out


def list_both_archived(
    url: str,
    token: str,
    start: str,
    end: str,
    log,
    stop_evt: threading.Event,
    section_name: str,
) -> List[Dict[str, Any]]:
    items: List[Dict[str, Any]] = []
    log(f"[{section_name}] Arşivsiz çekiliyor...")
    items += paged_list(url, token, start, end, log, stop_evt, archived=False, section_name=section_name)
    log(f"[{section_name}] Arşivli çekiliyor...")
    items += paged_list(url, token, start, end, log, stop_evt, archived=True, section_name=section_name)
    seen, uniq = set(), []
    for m in items:
        mid = str(m.get("id") or "")
        if mid and mid not in seen:
            seen.add(mid)
            uniq.append(m)
    log(f"[{section_name}] Birleştirildi (tekil): {len(uniq)} kayıt")
    return uniq


def fetch_xml_by(url_tpl: str, token: str, inv_id: str) -> Optional[bytes]:
    r = http_get(f"{url_tpl.format(id=inv_id)}/xml", token=token, typ="xml")
    return r.content if r is not None else None


def fetch_pdf_by(url_tpl: str, token: str, inv_id: str) -> Optional[bytes]:
    r = http_get(f"{url_tpl.format(id=inv_id)}/pdf", token=token, typ="pdf")
    return r.content if r is not None else None

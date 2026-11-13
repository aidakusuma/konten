"""
APP KONTEN — satu Streamlit untuk:
  📊 SC BW Video            (explode SC + TAP/AFF, kirim BRAND UNIQ NOT IN TAP - RANGE ke Google Sheets)
  📡 SC BW Live             (explode SC Livestream + TAP LS/AFF, BRAND UNIQUE NOT IN TAP - RANGE)
  🧩 Gabungan SC + TAP      (TAP range + SC not in TAP -> satu tabel layout TAP, SV dan LS)
  🎬 TAP Short Video / Live (sheet per brand untuk range terpilih + GABUNGAN semua range)

Jalankan:  streamlit run app_konten.py
Syarat  :  di folder yang sama ada
             - script engine SC BW (berisi fungsi enrich_sc)        -> untuk menu SC BW Video
             - script engine SC Live (berisi fungsi enrich_sc_live)  -> untuk menu SC BW Live
             - tap_utils.py (berisi dedupe_source_rows, dst.)         -> untuk menu TAP
Install :  pip install streamlit pandas openpyxl
           (kirim ke Google Sheets: pip install google-api-python-client google-auth)
"""

from __future__ import annotations

import contextlib
import importlib.util
import io
import math
import re
import sys
import traceback
from datetime import date, datetime, timedelta
from pathlib import Path

import pandas as pd
import streamlit as st

BASE_DIR = Path(__file__).resolve().parent
EXCEL_EXT = {".xlsx", ".xlsm", ".xls", ".csv", ".txt"}
ID_HEADERS = {"videoid", "productid", "contentid", "orderid", "liveid", "livestreamroomid"}

st.set_page_config(page_title="App Konten", page_icon="📊", layout="wide")

# Streamlit Cloud: buat credentials.json dari st.secrets["gcp_service_account"] bila belum ada
from _cloud_secrets import ensure_credentials_file  # noqa: E402
ensure_credentials_file(BASE_DIR)

SETTINGS_FILE = BASE_DIR / ".app_konten_settings.json"


def _load_settings() -> dict:
    import json
    try:
        return json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
    except Exception:  # noqa: BLE001
        return {}


def _save_setting(key: str, value: str) -> None:
    """Ingat isian (mis. link mapping BW) supaya tidak perlu diisi lagi tiap buka app."""
    import json
    cur = _load_settings()
    if value and cur.get(key) != value:
        cur[key] = value
        try:
            SETTINGS_FILE.write_text(json.dumps(cur, ensure_ascii=False, indent=2), encoding="utf-8")
        except OSError:
            pass


if "bw_link_shared" not in st.session_state:
    st.session_state["bw_link_shared"] = _load_settings().get("bw_link") or (
        "https://docs.google.com/spreadsheets/d/1Wi0kcm7apS4gbLDcHJdz8P1iRohF62aGzM38Oy_f7gM/edit?gid=0#gid=0")


# ============================================================================
# MENU + HALAMAN TAP (Short Video / Livestream)
# ============================================================================
with st.sidebar:
    st.title("Menu")
    PAGE = st.radio(
        "Pilih proses",
        ["📊 SC BW Video", "📡 SC BW Live", "🎬 TAP Short Video / Livestream", "🧩 Gabungan SC + TAP", "🧱 Gabung Konten + Label", "📋 Rekap per Brand", "🔎 Cek Data di Google Sheets"],
        label_visibility="collapsed",
    )
    st.divider()


def _is_streamlit_app(text: str) -> bool:
    """File aplikasi Streamlit (mis. app_sc_bw.py lama) TIDAK BOLEH diimpor sebagai engine/modul."""
    return bool(re.search(r"^\s*(import streamlit|from streamlit)", text, re.M))


def _find_tap_module():
    for p in sorted(BASE_DIR.glob("*.py")):
        if p.resolve() == Path(__file__).resolve():
            continue
        try:
            text = p.read_text(encoding="utf-8", errors="ignore")
        except OSError:
            continue
        if _is_streamlit_app(text):
            continue
        if (re.search(r"^def dedupe_source_rows\(", text, re.M)
                and re.search(r"^def aggregate_bw_brand_content\(", text, re.M)):
            return p
    return None


@st.cache_resource
def _load_tap(path: str):
    spec = importlib.util.spec_from_file_location("tap_utils_mod", path)
    mod = importlib.util.module_from_spec(spec)
    sys.modules["tap_utils_mod"] = mod
    spec.loader.exec_module(mod)
    return mod


def _tap_config(tap) -> dict:
    est, act, items = "Estimated affiliate partner commission ", "Actual affiliate partner commission", "Creator-attributed items sold"
    return {
        "sv": {
            "content": "Video ID", "date_src": "Post time", "date_col": "Post date",
            "gmv": "Affiliate video-attributed GMV",
            "sum": ["Affiliate video-attributed GMV", "Creator video-attributed orders",
                    "Affiliate video orders", est, act, items],
            "per_date": ["Video views", "Video likes"], "max": ["Creator follower count"],
            "name": "Video name", "suffix": " SV",
            "types": {"Comparison date": tap.TYPE_TEXT, "Campaign duration": tap.TYPE_TEXT,
                      "Post time": tap.TYPE_TEXT},
        },
        "ls": {
            "content": "Livestream room ID", "date_src": "LIVE time info", "date_col": "Date Live",
            "gmv": "Creator LIVE-attributed GMV",
            "sum": ["Creator LIVE-attributed GMV", "Creator LIVE-attributed orders",
                    "Affiliate LIVE orders", est, act, items],
            "per_date": ["LIVE views", "LIVE likes"], "max": ["Creator follower count"],
            "name": "Livestream name", "suffix": " LS",
            "types": {"Comparison date": tap.TYPE_TEXT, "Campaign duration": tap.TYPE_TEXT,
                      "LIVE time info": tap.TYPE_TEXT},
        },
    }


def _tap_read_raw(src):
    """src = (nama, bytes | path). Baca sheet 'Custom report' tanpa percaya dimensi sheet."""
    import openpyxl

    name, data = src
    target = io.BytesIO(data) if isinstance(data, (bytes, bytearray)) else data
    wb = openpyxl.load_workbook(target, read_only=True, data_only=True)
    ws = wb["Custom report"] if "Custom report" in wb.sheetnames else wb.worksheets[0]
    ws.reset_dimensions()
    it = ws.iter_rows(values_only=True)
    header = [str(v).strip() if v is not None else "" for v in next(it)]
    while header and not header[-1]:
        header.pop()
    rows = [list(r)[: len(header)] for r in it]
    wb.close()
    return header, rows


def _tap_gsheet_rows(url: str):
    import csv
    import urllib.error
    import urllib.request

    url = url.strip().strip("\"'<>").split()[0]
    m = re.search(r"/d/([a-zA-Z0-9_-]{15,})", url) or re.search(r"[?&]id=([a-zA-Z0-9_-]{15,})", url)
    if not m:
        raise ValueError(f"Link mapping tidak dikenali sebagai Google Sheets: {url[:120]!r}")
    export = f"https://docs.google.com/spreadsheets/d/{m.group(1)}/export?format=csv"
    g = re.search(r"[#&?]gid=(\d+)", url)
    if g:
        export += f"&gid={g.group(1)}"
    try:
        with urllib.request.urlopen(export, timeout=60) as resp:
            ctype, data = resp.headers.get("Content-Type", ""), resp.read()
    except urllib.error.HTTPError as e:
        raise RuntimeError(f"Gagal mengambil mapping (HTTP {e.code}). Set akses sheet 'Anyone with the link' → Viewer.") from e
    if "text/html" in ctype.lower():
        raise RuntimeError("Google meminta login. Set akses sheet mapping 'Anyone with the link' → Viewer.")
    return [list(r) for r in csv.reader(io.StringIO(data.decode("utf-8-sig")))]


def _tap_mapping(tap, text: str):
    import csv
    import openpyxl

    text = (text or "").strip().strip('"')
    if not text:
        raise ValueError("Mapping BW belum diisi (link Google Sheets atau path file).")
    if text.lower().startswith("http"):
        table = _tap_gsheet_rows(text)
    else:
        p = Path(text)
        if not p.is_file():
            raise FileNotFoundError(f"File mapping tidak ditemukan: {text}")
        if p.suffix.lower() in (".csv", ".txt"):
            with open(p, newline="", encoding="utf-8-sig") as f:
                table = [list(r) for r in csv.reader(f)]
        else:
            wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
            ws = wb.worksheets[0]
            ws.reset_dimensions()
            table = [list(r) for r in ws.iter_rows(values_only=True)]
    want = [tap.normalize_header(x) for x in ("Product ID", "Brand", "Sub Product")]
    for i, row in enumerate(table[:20]):
        norm = [tap.normalize_header(v) for v in row]
        if all(w in norm for w in want):
            return tap.build_bw_mapping_from_rows(table[i + 1:], *[norm.index(w) for w in want])
    raise ValueError("Header mapping tidak ditemukan (harus ada Product ID | Brand | Sub Product).")


TAP_PUSH_LINK = "https://docs.google.com/spreadsheets/d/1lX7Yuo50Rh6nHKv7s3bYZFcNrogauVXTr8EG3SHzS1E/edit?usp=sharing"


def _nk(v) -> str:
    return re.sub(r"[^a-z0-9]", "", str(v or "").lower())


def _tap_push_spec(cfg) -> dict:
    """Urutan kolom sheet tujuan. (label header, alias header lain, kolom sumber).
    'batas' dan 'cek creator paid' sengaja kosong (diisi manual di sheet)."""
    if cfg["suffix"] == " SV":
        f = [("batas", [], None), ("cek creator paid", [], None), ("Date", [], "@date"),
             ("Campaign name", [], "Campaign name"), ("Creator name", [], "Creator name"),
             ("creator follower", ["followers", "creatorfollowercount"], "Creator follower count"),
             ("Video id", [], "Video ID"), ("Video name", [], "Video name"),
             ("Post time", [], "Post time"), ("Date Post", ["postdate"], "Post date"),
             ("Product ID", [], "Product ID"), ("Product name", [], "Product name"),
             ("Shop name", [], "Shop name"), ("Brand", [], "Brand"),
             ("Sub Brand", ["subproduct", "subbrand"], "Sub Product"),
             ("Affiliate video GMV", [], "Affiliate video-attributed GMV"),
             ("Video orders", [], "Creator video-attributed orders"),
             ("Duration", [], "Duration"), ("Video views", [], "Video views"),
             ("Video likes", [], "Video likes"), ("Items sold", [], "Creator-attributed items sold")]
        return {"kind": "sv", "fields": f, "id_alias": "videoid"}
    f = [("batas", [], None), ("cek creator paid", [], None), ("Date", [], "@date"),
         ("Campaign name", [], "Campaign name"), ("Creator name", [], "Creator name"),
         ("Followers", ["creatorfollower", "creatorfollowercount"], "Creator follower count"),
         ("Livestream room ID", [], "Livestream room ID"), ("Livestream name", [], "Livestream name"),
         ("LIVE time info", [], "LIVE time info"), ("Start", [], "@start"), ("End", [], "@end"),
         ("Date Live", [], "Date Live"), ("Durasi", [], "@durasi"),
         ("Product ID", [], "Product ID"), ("Product name", [], "Product name"),
         ("Shop name", [], "Shop name"), ("Brand", [], "Brand"),
         ("Sub Brand", ["subproduct", "subbrand", "su brand"], "Sub Product"),
         ("Affiliate LIVE GMV", [], "Creator LIVE-attributed GMV"),
         ("LIVE orders", [], "Creator LIVE-attributed orders"),
         ("Duration", [], "Duration"), ("LIVE views", [], "LIVE views"),
         ("LIVE likes", [], "LIVE likes"), ("Items sold", [], "Creator-attributed items sold")]
    return {"kind": "ls", "fields": f, "id_alias": "livestreamroomid"}


def _live_parts(text):
    """'2026-10-21 19:27:16-2026-10-22 00:01:59' -> (start, end, durasi 'H:MM:SS')."""
    m = re.search(r"(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})\s*-\s*(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})", str(text or ""))
    if not m:
        return "", "", ""
    a, b = (datetime.strptime(x, "%Y-%m-%d %H:%M:%S") for x in m.groups())
    sec = max(int((b - a).total_seconds()), 0)
    return m.group(1), m.group(2), f"{sec // 3600}:{sec % 3600 // 60:02d}:{sec % 60:02d}"


def _col_letter(i: int) -> str:
    i += 1
    out = ""
    while i:
        i, r_ = divmod(i - 1, 26)
        out = chr(65 + r_) + out
    return out


def _ensure_grid(svc, sid: str, title: str, need_rows: int, need_cols: int = 0) -> None:
    """Perbesar grid tab kalau baris/kolom yang akan ditulis melewati batas (default sheet baru hanya 1000 baris)."""
    meta = svc.spreadsheets().get(spreadsheetId=sid, fields="sheets(properties(sheetId,title,gridProperties))").execute()
    p_ = next((x["properties"] for x in meta.get("sheets", []) if x["properties"]["title"] == title), None)
    if p_ is None:
        return
    gp, reqs = p_.get("gridProperties", {}), []
    if need_rows > gp.get("rowCount", 0):
        reqs.append({"appendDimension": {"sheetId": p_["sheetId"], "dimension": "ROWS",
                                         "length": need_rows - gp.get("rowCount", 0)}})
    if need_cols > gp.get("columnCount", 0):
        reqs.append({"appendDimension": {"sheetId": p_["sheetId"], "dimension": "COLUMNS",
                                         "length": need_cols - gp.get("columnCount", 0)}})
    if reqs:
        svc.spreadsheets().batchUpdate(spreadsheetId=sid, body={"requests": reqs}).execute()


_RETRY_BASE = 4      # detik; jeda = base * 2^percobaan (4, 8, 16, 32, 64, 64, ...)
_RETRY_MAX = 8


def _retry_patch() -> None:
    """Semua request Google Sheets otomatis diulang dengan jeda kalau kena batas kuota (429 / 'rate limit') atau error
    sementara server (500/502/503/504). Kuota tulis Google Sheets = 60 request per menit per user."""
    try:
        from googleapiclient import http as _gh
    except ImportError:
        return
    if getattr(_gh.HttpRequest, "_app_retry", False):
        return
    import time as _t
    orig = _gh.HttpRequest.execute

    def execute(self, *a, **k):
        for attempt in range(_RETRY_MAX + 1):
            try:
                return orig(self, *a, **k)
            except Exception as e:  # noqa: BLE001
                status = getattr(getattr(e, "resp", None), "status", None)
                text = str(e).lower()
                transient = status in (429, 500, 502, 503, 504) or (
                    status == 403 and ("ratelimit" in text or "rate limit" in text or "quota" in text))
                if not transient or attempt == _RETRY_MAX:
                    raise
                _t.sleep(min(_RETRY_BASE * (2 ** attempt), 64))

    _gh.HttpRequest.execute = execute
    _gh.HttpRequest._app_retry = True


def _tap_push(push: dict, link: str, tab_name: str, dry_run: bool, cred_path: Path) -> dict:
    """Kirim GABUNGAN ke Google Sheets. Hanya MENAMBAH; semua baris dimasukkan (duplikat tidak dicek).
    Kolom disejajarkan dengan header yang sudah ada di tab (kolom 'batas'/'cek creator paid' tidak disentuh)."""
    from google.oauth2 import service_account
    from googleapiclient.discovery import build

    m = re.search(r"/spreadsheets/d/([a-zA-Z0-9_-]+)", link or "")
    if not m:
        raise ValueError("Link target bukan link Google Sheets.")
    sid = m.group(1)
    if not cred_path.is_file():
        raise FileNotFoundError(f"credentials.json tidak ditemukan: {cred_path}")
    creds = service_account.Credentials.from_service_account_file(
        str(cred_path), scopes=["https://www.googleapis.com/auth/spreadsheets"])
    _retry_patch()
    svc = build("sheets", "v4", credentials=creds, cache_discovery=False)

    props = [x["properties"] for x in svc.spreadsheets().get(
        spreadsheetId=sid, fields="sheets(properties(sheetId,title))").execute().get("sheets", [])]
    tab_name = (tab_name or "").strip()
    if not tab_name:
        raise ValueError("Nama tab tujuan kosong.")
    target = next((p for p in props if p["title"].strip().lower() == tab_name.lower()), None)
    if target is None:
        if not dry_run:
            svc.spreadsheets().batchUpdate(
                spreadsheetId=sid, body={"requests": [{"addSheet": {"properties": {"title": tab_name}}}]}).execute()
        title = tab_name
    else:
        title = target["title"]

    header, rows = push["header"], push["rows"]
    nkh = [_nk(x) for x in header]
    g_i = next(i for i, a in enumerate(nkh) if a.startswith("affiliate") and a.endswith("gmv"))
    c_i = nkh.index("creatorname")

    def _stat(rs_):
        return {"n": len(rs_), "gmv": sum(float(x[g_i] or 0) for x in rs_),
                "creator": len({str(x[c_i]).strip().lower() for x in rs_ if str(x[c_i]).strip()})}

    res_ = {"title": title, "total": len(rows), "added": len(rows), "skipped": 0,
            "existing": {"n": 0, "gmv": 0.0, "creator": 0}}
    chunk = 2000

    got = [] if target is None else svc.spreadsheets().values().get(
        spreadsheetId=sid, range=f"'{title}'!1:1").execute().get("values", [])  # error JANGAN ditelan
    ex_header = got[0] if got else []
    id_col = next(i for i, a in enumerate([_nk(x) for x in header]) if a == push["id_alias"])

    if not any(str(c).strip() for c in ex_header):  # tab kosong -> tulis header lengkap + data
        res_["header_written"] = True
        res_["add_stat"] = _stat(rows)
        if not dry_run:
            allrows = [header] + rows
            _ensure_grid(svc, sid, title, len(allrows), len(header))
            for i in range(0, len(allrows), chunk):
                svc.spreadsheets().values().update(
                    spreadsheetId=sid, range=f"'{title}'!A{i + 1}", valueInputOption="RAW",
                    body={"values": allrows[i:i + chunk]}).execute()
        return res_

    # sejajarkan dengan header yang sudah ada (pakai nama + alias)
    lookup = {}
    for i, (lab, al) in enumerate(zip(header, push["aliases"])):
        for name in [lab] + list(al):
            lookup.setdefault(_nk(name), i)
    ex_idx = [lookup.get(_nk(c)) for c in ex_header]

    # total yang SUDAH ada di tab (hanya untuk "Cek dulu"; tidak dipakai untuk menolak baris)
    e_id = ex_idx.index(id_col) if id_col in ex_idx else None
    e_g = ex_idx.index(g_i) if g_i in ex_idx else None
    e_c = ex_idx.index(c_i) if c_i in ex_idx else None
    want = [x for x in (e_id, e_g, e_c) if x is not None]
    vr = svc.spreadsheets().values().batchGet(
        spreadsheetId=sid, ranges=[f"'{title}'!{_col_letter(x)}2:{_col_letter(x)}" for x in want],
        valueRenderOption="UNFORMATTED_VALUE").execute().get("valueRanges", []) if want else []
    ex_vals = {x: [r[0] if r else "" for r in vr[k].get("values", [])] for k, x in enumerate(want)}
    ex_n = sum(1 for v in ex_vals.get(e_id, []) if str(v).strip())
    ex_gmv = sum(v for v in ex_vals.get(e_g, []) if isinstance(v, (int, float)) and not isinstance(v, bool))
    ex_cr = {str(v).strip().lower() for v in ex_vals.get(e_c, []) if str(v).strip()}
    nrows_now = max((len(v) for v in ex_vals.values()), default=0)

    # SEMUA baris ditambahkan (duplikat dibiarkan masuk)
    aligned = [[(rr[ix] if ix is not None else "") for ix in ex_idx] for rr in rows]
    res_.update({"added": len(aligned), "skipped": 0, "add_stat": _stat(rows),
                 "existing": {"n": ex_n, "gmv": ex_gmv, "creator": len(ex_cr)}})
    if not dry_run:
        _ensure_grid(svc, sid, title, nrows_now + 1 + len(aligned) + 5, len(ex_header))
        for i in range(0, len(aligned), chunk):
            svc.spreadsheets().values().append(
                spreadsheetId=sid, range=f"'{title}'!A1", valueInputOption="RAW",
                insertDataOption="INSERT_ROWS", body={"values": aligned[i:i + chunk]}).execute()
    return res_


LS_BRAND_LINK = "https://docs.google.com/spreadsheets/d/1ZtTXLHnV8VdIhy90WfhK7A6Q6NYej7VZdSGyZWK9XGY/edit?gid=1881741696#gid=1881741696"


def _norm_dt(v, as_date_only=False):
    """Nilai sel (angka serial / teks) -> 'YYYY-MM-DD HH:MM' (atau 'YYYY-MM-DD')."""
    if v is None or v == "":
        return ""
    dt = None
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        dt = datetime(1899, 12, 30) + timedelta(days=float(v))
    else:
        t = str(v).strip()
        for f in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d", "%m/%d/%Y %H:%M:%S", "%m/%d/%Y %H:%M",
                  "%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M", "%m/%d/%Y", "%d/%m/%Y"):
            try:
                dt = datetime.strptime(t, f)
                break
            except ValueError:
                continue
        if dt is None:
            return t
    return dt.strftime("%Y-%m-%d") if as_date_only else dt.strftime("%Y-%m-%d %H:%M")


def _tap_push_brands(brand_rows: dict, link: str, dry_run: bool, cred_path: Path) -> list[dict]:
    """LS per brand -> tab brand di sheet 'Content Unilever Beauty (LS)'. Tab yang namanya mengandung 'paid' dilewati.
    Hanya menulis kolom B (brand; PONDS = sub brand), C (username), D (LS ID), G (live date), L (start), M (end), O (GMV), P (nama live) di bawah data terakhir.
    Kolom lain tidak disentuh. Semua baris dimasukkan (duplikat tidak dicek)."""
    from google.oauth2 import service_account
    from googleapiclient.discovery import build

    m = re.search(r"/spreadsheets/d/([a-zA-Z0-9_-]+)", link or "")
    if not m:
        raise ValueError("Link target bukan link Google Sheets.")
    sid = m.group(1)
    if not cred_path.is_file():
        raise FileNotFoundError(f"credentials.json tidak ditemukan: {cred_path}")
    creds = service_account.Credentials.from_service_account_file(
        str(cred_path), scopes=["https://www.googleapis.com/auth/spreadsheets"])
    _retry_patch()
    svc = build("sheets", "v4", credentials=creds, cache_discovery=False)
    titles = [x["properties"]["title"] for x in svc.spreadsheets().get(
        spreadsheetId=sid, fields="sheets(properties(title))").execute().get("sheets", [])]
    usable = [t for t in titles if "paid" not in t.lower()]

    out = []
    for brand, rows in sorted(brand_rows.items()):
        b = brand.strip().lower()
        cands = [t for t in usable if b and b in t.lower()]
        cands.sort(key=lambda t: (not t.lower().startswith(b), len(t)))
        if not cands:
            out.append({"brand": brand, "tab": None, "added": 0, "skipped": 0, "total": len(rows),
                        "note": "tidak ada tab untuk brand ini (selain tab PAID)"})
            continue
        tab = cands[0]
        res = svc.spreadsheets().values().batchGet(
            spreadsheetId=sid, ranges=[f"'{tab}'!{c}:{c}" for c in "BCDGLMOP"],
            valueRenderOption="UNFORMATTED_VALUE").execute().get("valueRanges", [])
        col = {c: [r[0] if r else "" for r in vr.get("values", [])] for c, vr in zip("BCDGLMOP", res)}
        last = max(len(v) for v in col.values())
        new = list(rows)  # semua baris dimasukkan; duplikat tidak dicek
        info = {"brand": brand, "tab": tab, "added": len(new), "skipped": 0,
                "total": len(rows), "start_row": last + 1,
                "add_gmv": sum(r_["gmv"] for r_ in new), "add_creator": len({r_["user"].strip().lower() for r_ in new}),
                "ex_n": sum(1 for u in col["C"][1:] if str(u).strip()),
                "ex_gmv": sum(v for v in col["O"][1:] if isinstance(v, (int, float)) and not isinstance(v, bool)),
                "ex_creator": len({str(u).strip().lower() for u in col["C"][1:] if str(u).strip()})}
        if new and not dry_run:
            n = last + 1
            _ensure_grid(svc, sid, tab, n + len(new) + 1, 16)
            cols = {"B": [r_.get("brand_col", brand) for r_ in new], "C": [r_["user"] for r_ in new], "D": [r_.get("id", "") for r_ in new], "G": [r_["date"] for r_ in new],
                    "L": [r_["start"] for r_ in new], "M": [r_["end"] for r_ in new],
                    "O": [r_["gmv"] for r_ in new], "P": [r_["name"] for r_ in new]}
            for option, letters in (("USER_ENTERED", "GLM"), ("RAW", "BCDOP")):
                svc.spreadsheets().values().batchUpdate(spreadsheetId=sid, body={
                    "valueInputOption": option,
                    "data": [{"range": f"'{tab}'!{c}{n}:{c}{n + len(new) - 1}",
                              "values": [[v] for v in cols[c]]} for c in letters]}).execute()
        out.append(info)
    return out


def push_kind_is_ls(res) -> bool:
    return res.get("push", {}).get("kind") == "ls"


def _rp(v) -> str:
    return "Rp" + f"{v:,.0f}".replace(",", ".")


def _fmt_brand(x: dict) -> str:
    if not x["tab"]:
        return f"{x['brand']}: {x['note']} ({x['total']:,} live tidak dikirim)"
    return (f"{x['brand']} → tab '{x['tab']}': {x['added']:,} baris baru"
            + (f" mulai baris {x['start_row']}" if x["added"] else "")
            + f" | sudah ada di tab: {x['ex_n']:,} live, {_rp(x['ex_gmv'])}, {x['ex_creator']:,} creator"
            + f" | akan/sudah ditambah: {x['added']:,} live, {_rp(x['add_gmv'])}, {x['add_creator']:,} creator")


def _fmt_push(r_: dict, label: str) -> str:
    a, e = r_["add_stat"], r_["existing"]
    return (f"Tab '{r_['title']}' — sudah ada: {e['n']:,} {label}, {_rp(e['gmv'])}, {e['creator']:,} creator. "
            f"Ditambah: {a['n']:,} {label}, {_rp(a['gmv'])}, {a['creator']:,} creator.")


def _tap_run(tap, cfg, files, mapping_text, rs, re_, log):
    """Pipeline TAP: raw -> bersih -> brand -> unik per Brand+Content ID -> sheet per brand (range) + gabungan."""
    import openpyxl
    from collections import Counter

    mapping, mstats = _tap_mapping(tap, mapping_text)
    log(f"Mapping: {mstats['mapping_count']:,} Product ID"
        + (f" | variasi huruf brand digabung: {mstats['brand_case_variants']}" if mstats["brand_case_variants"] else ""))

    validator = tap.make_tap_row_validator("Date", [cfg["content"], "Product ID"])
    key_cols = ["Date", cfg["content"], "Product ID", "Campaign ID", "Creator name"]

    header, per_file, summaries, rejected = None, [], [], Counter()
    raw_total, dup_total, file_rows = 0, 0, []
    for src in files:
        h, rows = _tap_read_raw(src)
        if header is None:
            header = h
        elif [tap.normalize_header(x) for x in h] != [tap.normalize_header(x) for x in header]:
            raise ValueError(f"Header {src[0]} berbeda dengan file pertama; pakai file dengan kolom yang sama.")
        valid, summary = [], None
        for n, row in enumerate(rows, start=2):
            if str(tap.cell(row, 0) or "").strip().lower() == "summary":
                summary = dict(zip(h, row))
            ok, reason = validator(row, n, h)
            if ok:
                valid.append(row)
            else:
                rejected[reason] += 1
        valid, dropped, conflicts = tap.dedupe_source_rows(h, valid, key_cols)
        raw_total += len(rows)
        dup_total += dropped
        file_rows.append({"File": src[0], "Baris raw": len(rows), "Dipakai": len(valid)})
        win = tap.source_date_window(h, valid, "Date")
        log(f"{src[0]}: {len(rows):,} baris | tanggal data {win[0]} s/d {win[1]} | duplikat persis dibuang {dropped}"
            + (f" | kunci sama tapi nilai beda: {len(conflicts)}" if conflicts else ""))
        per_file.append(valid)
        summaries.append(summary)

    # Tumpang tindih antar file: baris dengan kunci harian yang sama -> file TERAKHIR yang dipakai
    replaced = 0
    if len(per_file) > 1:
        idx = [tap.find_header_index(header, k) for k in key_cols]
        later_keys, kept_per_file = set(), [None] * len(per_file)
        for i in range(len(per_file) - 1, -1, -1):
            kept, mine = [], set()
            for row in per_file[i]:
                k = tuple(str(tap.cell(row, j) if tap.cell(row, j) is not None else "").strip() for j in idx)
                mine.add(k)
                if k in later_keys:
                    replaced += 1
                else:
                    kept.append(row)
            kept_per_file[i] = kept
            later_keys |= mine
        per_file = kept_per_file
        if replaced:
            log(f"Tumpang tindih antar file: {replaced:,} baris harian diganti oleh file yang lebih akhir (tidak dihitung dua kali).")
    valid = [row for part in per_file for row in part]
    log(f"Baris dipakai: {len(valid):,}")
    recon = [("File dibaca", len(files)), ("Baris raw (semua file)", raw_total)]
    recon += [(f"Ditolak: {k}", v) for k, v in rejected.items()]
    recon += [("Duplikat persis dibuang", dup_total), ("Diganti file yang lebih akhir (tumpang tindih)", replaced),
              ("Baris harian dipakai", len(valid))]

    keep = ["Date", "Campaign ID", "Campaign name", "Creator name", "Product ID", "Product name",
            "Shop name", "Duration", cfg["content"], cfg["date_src"], cfg["name"]] + cfg["sum"] + cfg["per_date"] + cfg["max"]
    seen, keep_u = set(), []
    for k in keep:
        nk = tap.normalize_header(k)
        if nk not in seen:
            seen.add(nk)
            keep_u.append(k)
    idx = [tap.find_header_index(header, k) for k in keep_u]
    h1 = [header[i] for i in idx]
    r1 = [[tap.cell(r, i) for i in idx] for r in valid]
    log("Konversi tipe data…")
    r2, _types, _ts, _errs = tap.coerce_rows_by_schema(h1, r1, column_types=cfg["types"], strict=True)

    si = tap.find_header_index(h1, cfg["date_src"])
    h2 = list(h1) + [cfg["date_col"]]
    r3 = [row + [tap.normalize_date_text(str(row[si])[:10])] for row in r2]

    h, r, status, missing, _bc, _sc = tap.add_brand_subproduct_from_product_id(h2, r3, mapping)
    log(f"Klasifikasi brand: {dict(status)}")

    log("Agregasi per Brand + ID konten…")
    groups, accepted, _blank, _invalid = tap.aggregate_bw_brand_content(
        h, r, content_id_header=cfg["content"], date_header=cfg["date_col"], sum_columns=cfg["sum"],
        per_date_sum_columns=cfg["per_date"], daily_header="Date", max_columns=cfg["max"])
    checks = tap.reconcile_sum_columns(h, accepted, groups, cfg["sum"])
    if not all(c["ok"] for c in checks):
        raise ValueError("Rekonsiliasi agregasi gagal.")
    log(f"Unik Brand + {cfg['content']}: {len(groups):,}")

    qc_rows = []
    have_summary = all(s for s in summaries)
    for c in checks:
        row = {"Kolom": c["column"].strip(), "Hasil agregasi": c["aggregate"]}
        if have_summary and not replaced:
            total = sum(tap.to_number(s.get(c["column"], 0)) for s in summaries)
            row["Summary TAP"] = total
            row["Status"] = "OK" if tap.numbers_equal(total, c["aggregate"]) else "BEDA"
        else:
            row["Summary TAP"] = None
            row["Status"] = "n/a (ada tumpang tindih file)" if replaced else "n/a"
        qc_rows.append(row)

    cols = [cfg["date_col"], cfg["content"], cfg["name"], "Creator name", "Creator follower count",
            "Brand", "Sub Product", "Product ID", "Product name", "Campaign name"] + cfg["sum"] + cfg["per_date"]
    cidx = [tap.find_header_index(h, c) for c in cols]
    header_out = [h[i] for i in cidx] + ["Jenis"]
    gi = tap.find_header_index(h, cfg["gmv"])

    def project(rows_with_jenis):
        return [[x[i] for i in cidx] + [x[-1]] for x in rows_with_jenis]

    if rs is None:
        in_range = groups
    else:
        in_range, _bad = tap.filter_groups_by_date_range(h, groups, cfg["date_col"], rs.isoformat(), re_.isoformat())
    tabs = tap.build_bw_brand_tabs(in_range, gi, cfg["suffix"])
    combined = tap.build_bw_combined_from_groups(groups, h, cfg["gmv"])

    sheets = [(t, header_out, project(rows)) for t, rows in tabs.items()]
    sheets.append(("GABUNGAN (semua range)", header_out, project(combined)))

    wb = openpyxl.Workbook(write_only=True)
    used = set()
    for title, hd, rows in sheets:
        t = "".join(c for c in title if c not in '[]:*?/\\')[:31] or "Sheet"
        base, k = t, 2
        while t.lower() in used:
            t = f"{base[:28]}~{k}"
            k += 1
        used.add(t.lower())
        ws = wb.create_sheet(t)
        ws.append(hd)
        for row in rows:
            ws.append(row)
    qc_ws = wb.create_sheet("QC")
    qc_ws.append(["Kolom", "Summary TAP", "Hasil agregasi", "Status"])
    for q in qc_rows:
        qc_ws.append([q["Kolom"], q["Summary TAP"], q["Hasil agregasi"], q["Status"]])
    policy = tap.aggregation_policy_report(h, [cfg["content"], "Brand"], cfg["sum"], cfg["per_date"], cfg["max"])
    qc_ws.append([])
    qc_ws.append(["Aturan agregasi", "Kolom"])
    for k, v in policy.items():
        qc_ws.append([k, ", ".join(v)])
    bio = io.BytesIO()
    wb.save(bio)

    gmv_i = tap.find_header_index(header_out, cfg["gmv"])
    brand_i = tap.find_header_index(header_out, "Brand")
    all_by_brand, rng_by_brand = Counter(), Counter()
    all_gmv, rng_gmv = Counter(), Counter()
    for row in project(combined):
        all_by_brand[row[brand_i]] += 1
        all_gmv[row[brand_i]] += row[gmv_i]
    for rows in tabs.values():
        for row in rows:
            b = row[tap.find_header_index(h, "Brand")]
            rng_by_brand[b] += 1
            rng_gmv[b] += tap.to_number(row[gi])
    brands = sorted(all_by_brand)
    ci_ = tap.find_header_index(h, "Creator name")
    cr_all, cr_rng = {}, {}
    for x in combined:
        cr_all.setdefault(str(x[brand_i_h := tap.find_header_index(h, "Brand")]), set()).add(str(x[ci_]).strip().lower())
    for rows in tabs.values():
        for x in rows:
            cr_rng.setdefault(str(x[brand_i_h]), set()).add(str(x[ci_]).strip().lower())
    summary_df = pd.DataFrame({
        "Brand": brands,
        "Konten di range": [rng_by_brand.get(b, 0) for b in brands],
        "GMV di range": [rng_gmv.get(b, 0) for b in brands],
        "Creator di range": [len(cr_rng.get(b, ())) for b in brands],
        "Konten (semua range)": [all_by_brand[b] for b in brands],
        "GMV (semua range)": [all_gmv[b] for b in brands],
        "Creator (semua range)": [len(cr_all.get(b, ())) for b in brands],
    })
    totals = {
        "range": {"konten": sum(rng_by_brand.values()), "gmv": sum(rng_gmv.values()),
                  "creator": len(set().union(*cr_rng.values())) if cr_rng else 0},
        "all": {"konten": sum(all_by_brand.values()), "gmv": sum(all_gmv.values()),
                "creator": len(set().union(*cr_all.values())) if cr_all else 0},
    }
    preview = pd.DataFrame(project(combined)[:50], columns=header_out)
    recon += [("Konten unik Brand + ID (GABUNGAN)", len(groups))]

    # Product ID yang belum ada di mapping (N/A) -> pengiriman ke Google Sheets diblok
    miss_df = pd.DataFrame(columns=["Product ID", "Product name", "Jumlah baris raw"])
    if missing:
        pi_, pn_, bi_ = (tap.find_header_index(h, n) for n in ("Product ID", "Product name", "Brand"))
        first_name = {}
        for x in r:
            if str(x[bi_]).strip().upper() == "MISSING":
                first_name.setdefault(tap.normalize_id(x[pi_]) or "<BLANK_PRODUCT_ID>", str(x[pn_] or ""))
        miss_df = pd.DataFrame([{"Product ID": k, "Product name": first_name.get(k, ""), "Jumlah baris raw": v}
                                for k, v in sorted(missing.items(), key=lambda kv: -kv[1])])

    # ---- baris untuk Google Sheets (GABUNGAN, semua range) ----
    ci, bi = tap.find_header_index(h, cfg["content"]), tap.find_header_index(h, "Brand")
    di = tap.find_header_index(h, "Date")
    first_date = {}
    for x in r:
        d_ = tap.normalize_date_text(str(x[di])[:10])
        k = (str(x[bi]).strip().lower(), str(x[ci]).strip())
        if d_ and (k not in first_date or d_ < first_date[k]):
            first_date[k] = d_
    spec = _tap_push_spec(cfg)
    hidx = {tap.normalize_header(c): i for i, c in enumerate(h)}
    def _rec(x):
        key = (str(x[bi]).strip().lower(), str(x[ci]).strip())
        rec = []
        for _label, _alias, src in spec["fields"]:
            if src == "@date":
                v = first_date.get(key, "")
            elif src in ("@start", "@end", "@durasi"):
                v = _live_parts(x[hidx[tap.normalize_header(cfg["date_src"])]])[("@start", "@end", "@durasi").index(src)]
            elif src is None:
                v = ""
            else:
                v = x[hidx[tap.normalize_header(src)]]
            rec.append("" if v is None else v)
        return rec

    push_rows, skipped_missing = [], 0
    for x in combined:
        if str(x[bi]).strip().upper() == "MISSING":  # Product ID belum ada di mapping -> jangan masuk sheet
            skipped_missing += 1
            continue
        push_rows.append(_rec(x))
    range_push = [_rec(x) for rows_ in tabs.values() for x in rows_ if str(x[bi]).strip().upper() != "MISSING"]
    brand_rows = {}
    if spec["kind"] == "ls":
        ui_, di_, ti_ = (hidx[tap.normalize_header(n)] for n in ("Creator name", "Date Live", "LIVE time info"))
        gi_, ni_, br_ = (hidx[tap.normalize_header(n)] for n in (cfg["gmv"], cfg["name"], "Brand"))
        id_ = hidx[tap.normalize_header(cfg["content"])]
        sb_ = hidx.get(tap.normalize_header("Sub Product"))
        for rows_ in tabs.values():
            for x in rows_:
                st_, en_, _du = _live_parts(x[ti_])
                _b = str(x[br_])
                _sub = str(x[sb_] or "").strip() if sb_ is not None else ""
                brand_rows.setdefault(_b, []).append({
                    "brand_col": (_sub or _b) if _b.strip().lower() == "ponds" else _b,
                    "id": ("" if x[id_] is None else str(x[id_]).strip()),
                    "user": str(x[ui_] or ""), "date": str(x[di_] or ""), "start": st_, "end": en_,
                    "gmv": tap.to_number(x[gi_]), "name": str(x[ni_] or "")})
        for v in brand_rows.values():
            v.sort(key=lambda z: (z["start"], z["user"]))
    push = {"kind": spec["kind"], "brand_rows": brand_rows, "skipped_missing": skipped_missing, "header": [f[0] for f in spec["fields"]],
            "aliases": [f[1] for f in spec["fields"]], "id_alias": spec["id_alias"], "rows": push_rows}
    return {
        "push": push, "range_push": range_push, "range": (rs, re_), "totals": totals, "label": "Video" if cfg["suffix"] == " SV" else "Livestream",
        "bytes": bio.getvalue(), "summary": summary_df, "qc": pd.DataFrame(qc_rows), "preview": preview,
        "missing": len(missing), "missing_df": miss_df,
        "recon": pd.DataFrame(recon, columns=["Tahap", "Jumlah"]), "file_rows": pd.DataFrame(file_rows), "n_all": len(groups), "n_range": len(in_range),
    }


def render_tap_page():
    st.title("🎬 TAP — Short Video & Livestream")
    st.caption("Raw TAP 'Custom report' → klasifikasi brand (mapping BW) → unik per Brand + ID konten → "
               "**sheet per brand (hanya range yang dipilih)** + **sheet GABUNGAN (semua range)** + sheet QC.")

    tap_path = _find_tap_module()
    if tap_path is None:
        st.error(
            f"Modul TAP (`tap_utils.py`, yang berisi `dedupe_source_rows` dan `aggregate_bw_brand_content`) "
            f"tidak ditemukan di folder `{BASE_DIR}`. Taruh di folder yang sama dengan `app_konten.py`.")
        st.write("File .py di folder ini:", [p.name for p in BASE_DIR.glob("*.py")])
        st.stop()
    tap = _load_tap(str(tap_path))
    st.sidebar.caption(f"Modul TAP: `{tap_path.name}`")
    configs = _tap_config(tap)

    kind = st.radio("Jenis data", ["Short Video (SV)", "Livestream (LS)"], horizontal=True)
    cfg = configs["sv" if kind.startswith("Short") else "ls"]
    st.info(f"ID konten: **{cfg['content']}** · kolom tanggal untuk range: **{cfg['date_col']}** "
            f"(dari `{cfg['date_src']}`) · baris Summary di atas otomatis dibuang.")

    c1, c2 = st.columns(2)
    with c1:
        st.subheader("Raw TAP")
        ups = st.file_uploader("Upload export TAP (.xlsx, boleh lebih dari satu file/periode)", type=["xlsx"],
                               accept_multiple_files=True, key="tap_ups")
        paths = st.text_area("atau path lokal (satu per baris; boleh path folder)", key="tap_paths", height=80,
                             placeholder=r"C:\Users\...\TAP_VID_18-24_OCT_2026.xlsx")
        st.caption("File besar (puluhan MB) lebih cepat lewat path lokal. Baris harian yang sama di dua file "
                   "tidak dihitung dua kali — file yang lebih akhir dipakai.")
    with c2:
        st.subheader("Mapping BW")
        mapping_text = st.text_input(
            "Link Google Sheets mapping / path file (Product ID | Brand | Sub Product)",
            value=st.session_state.get("bw_link_shared", ""), key="tap_map",
            placeholder="https://docs.google.com/spreadsheets/d/.../edit?gid=0")
        st.session_state["bw_link_shared"] = mapping_text
        _save_setting("bw_link", mapping_text.strip())
        st.subheader("Range sheet per brand")
        today = date.today()
        use_range = st.checkbox("Batasi sheet per brand ke range tertentu", value=True)
        rng = st.date_input("Range (kolom tanggal di atas)", value=(today - timedelta(days=6), today),
                            disabled=not use_range, key="tap_range")
        st.caption("Sheet GABUNGAN selalu berisi semua tanggal dari seluruh file.")

    out_dir = st.text_input("Simpan juga ke folder lokal (opsional)", key="tap_outdir",
                            placeholder=r"C:\Users\...\Output")
    default_name = f"TAP_{'SV' if cfg['suffix'] == ' SV' else 'LS'}_hasil.xlsx"
    out_name = st.text_input("Nama file output", value=default_name, key="tap_outname")

    st.subheader("Kirim GABUNGAN ke Google Sheets")
    tap_link = st.text_input("Link Google Sheets tujuan", value=TAP_PUSH_LINK, key="tap_push_link")
    tap_tab = st.text_input("Nama tab tujuan", value="sv" if cfg["suffix"] == " SV" else "ls",
                            key=f"tap_push_tab_{cfg['suffix'].strip()}")
    tap_auto = st.checkbox("Langsung kirim ke Google Sheets setelah proses selesai", value=False, key="tap_auto_push",
                           help="Menambah di bawah data yang ada. Semua baris dimasukkan, duplikat tidak dicek. "
                                "Tidak jalan kalau masih ada Product ID yang belum ada di mapping.")
    tap_cred = BASE_DIR / "credentials.json"
    brand_link, brand_auto = None, False
    if cfg["suffix"] == " LS":
        st.subheader("Kirim LS per brand → Content Unilever Beauty (LS)")
        brand_link = st.text_input("Link Google Sheets 'Content Unilever Beauty (LS)'", value=LS_BRAND_LINK,
                                   key="ls_brand_link")
        brand_auto = st.checkbox("Langsung kirim LS per brand setelah proses selesai", value=False, key="ls_brand_auto",
                                 help="Menambah baris di bawah data terakhir tab brand (selain tab PAID). "
                                      "Kolom B, C, D, G, L, M, O, P saja. Semua baris dimasukkan, duplikat tidak dicek. "
                                      "Tidak jalan kalau masih ada Product ID yang belum ada di mapping.")

    if st.button("▶️ Jalankan", type="primary", key="tap_run"):
        st.session_state.pop("tap_result", None)
        st.session_state.pop("tap_push_msg", None)
        try:
            files = [(u_.name, u_.getvalue()) for u_ in (ups or [])]
            for line in (paths or "").splitlines():
                line = line.strip().strip('"').strip("'")
                if not line:
                    continue
                p = Path(line)
                if p.is_dir():  # semua .xlsx di folder
                    found = sorted(x for x in p.glob("*.xlsx") if not x.name.startswith("~$"))
                    if not found:
                        raise FileNotFoundError(f"Tidak ada .xlsx di folder: {line}")
                    files.extend((x.name, str(x)) for x in found)
                    continue
                if not p.is_file():
                    raise FileNotFoundError(f"File tidak ditemukan: {line}")
                files.append((p.name, str(p)))
            if not files:
                raise ValueError("Belum ada file TAP (upload atau isi path).")
            rs = re_ = None
            if use_range:
                if not (isinstance(rng, (list, tuple)) and len(rng) == 2):
                    raise ValueError("Pilih tanggal awal DAN akhir range.")
                rs, re_ = rng
            with st.status("Memproses… (file besar bisa 2–3 menit)", expanded=True) as status:
                res = _tap_run(tap, cfg, files, mapping_text, rs, re_, lambda m: status.write(m))
                status.update(label="Selesai", state="complete")
            if out_dir.strip():
                folder = Path(out_dir.strip().strip('"'))
                folder.mkdir(parents=True, exist_ok=True)
                (folder / out_name).write_bytes(res["bytes"])
                res["saved_to"] = str(folder / out_name)
            res["name"] = out_name
            st.session_state["tap_result"] = res
            st.session_state[f"tap_range_{res['push']['kind']}"] = {
                "header": res["push"]["header"], "rows": res["range_push"], "range": res["range"],
                "missing": res["missing"], "files": [n_ for n_, _ in files]}
            if (brand_auto or tap_auto) and res["missing"]:
                st.session_state["tap_push_blocked"] = True
            else:
                st.session_state.pop("tap_push_blocked", None)
            if brand_auto and res["push"]["kind"] == "ls" and not res["missing"]:
                rb = _tap_push_brands(res["push"]["brand_rows"], brand_link, False, tap_cred)
                st.session_state["ls_brand_msg"] = rb
            if tap_auto and not res["missing"]:
                r_ = _tap_push(res["push"], tap_link, tap_tab, False, tap_cred)
                st.session_state["tap_push_msg"] = "Terkirim. " + _fmt_push(r_, res["label"].lower())
        except Exception as e:  # noqa: BLE001
            st.error(f"{type(e).__name__}: {e}")
            with st.expander("Detail error"):
                st.code(traceback.format_exc())

    res = st.session_state.get("tap_result")
    if res:
        st.success(f"Selesai — {res['n_range']:,} konten di range, {res['n_all']:,} konten di GABUNGAN.")
        if res.get("saved_to"):
            st.caption(f"Tersimpan di: `{res['saved_to']}`")
        blocked = bool(res["missing"])
        if blocked:
            st.error(f"⛔ {res['missing']:,} Product ID belum ada di mapping produk (N/A). "
                     "Pengiriman ke Google Sheets DIBLOK sampai mapping dilengkapi. "
                     "Lengkapi mapping lalu jalankan ulang.")
            st.dataframe(res["missing_df"], hide_index=True, use_container_width=True)
            st.download_button("⬇️ Download daftar Product ID yang belum ada di mapping (.csv)",
                               res["missing_df"].to_csv(index=False).encode("utf-8-sig"),
                               file_name="product_id_belum_ada_di_mapping.csv", mime="text/csv")
        elif st.session_state.pop("tap_push_blocked", None):
            pass
        st.download_button("⬇️ Download hasil (.xlsx)", res["bytes"], file_name=res["name"], type="primary",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.subheader("Cek total")
        for title_, key_ in (("Di range yang dipilih (sheet per brand)", "range"), ("Semua range (GABUNGAN)", "all")):
            t_ = res["totals"][key_]
            st.caption(title_)
            m1, m2, m3 = st.columns(3)
            m1.metric(f"Total {res['label']}", f"{t_['konten']:,}")
            m2.metric("Total GMV", f"Rp{t_['gmv']:,.0f}".replace(",", "."))
            m3.metric("Creator unik", f"{t_['creator']:,}")
        st.caption("Creator unik dihitung lintas brand (satu creator di dua brand dihitung sekali); "
                   "kolom Creator di tabel per brand dihitung per brand.")
        with st.expander("Rekonsiliasi baris (dari raw sampai konten unik)"):
            st.dataframe(res["recon"], hide_index=True, use_container_width=True)
            st.dataframe(res["file_rows"], hide_index=True, use_container_width=True)
            st.caption("Satu konten (Brand + ID) = satu baris di GABUNGAN; raw berisi satu baris per hari × produk × kampanye, "
                       "jadi jumlah konten unik jauh lebih kecil dari baris raw.")
        st.subheader("Ringkasan per brand")
        st.dataframe(res["summary"], hide_index=True, use_container_width=True)
        st.subheader("QC: hasil vs baris Summary TAP")
        st.dataframe(res["qc"], hide_index=True, use_container_width=True)
        with st.expander("Preview GABUNGAN (50 baris pertama)"):
            st.dataframe(res["preview"], hide_index=True, use_container_width=True)

        if push_kind_is_ls(res):
            st.subheader("Kirim LS per brand → Content Unilever Beauty (LS)")
            st.caption("Dari sheet per brand (range yang dipilih). Ditambah di bawah data terakhir tab brand "
                       "(tab PAID dilewati): B=brand (Ponds: sub brand), C=username, D=LS ID, G=live date, L=start, M=end, O=GMV, P=nama live.")
            for rb in st.session_state.pop("ls_brand_msg", None) or []:
                st.success(_fmt_brand(rb))
            lb1, lb2 = st.columns(2)
            for btn, dry in ((lb1, True), (lb2, False)):
                if btn.button("🔍 Cek dulu (tidak menulis)" if dry else "📤 Kirim LS per brand",
                              key=f"ls_brand_{'dry' if dry else 'go'}", type="secondary" if dry else "primary",
                              disabled=blocked):
                    try:
                        rb = _tap_push_brands(res["push"]["brand_rows"], st.session_state.get("ls_brand_link", LS_BRAND_LINK),
                                              dry, tap_cred)
                        for x in rb:
                            (st.info if dry else st.success)(("Cek: " if dry else "") + _fmt_brand(x))
                    except Exception as e:  # noqa: BLE001
                        st.error(f"{type(e).__name__}: {e}")
                        st.caption("Pastikan sheet di-share sebagai **Editor** ke email service account di credentials.json.")

        st.subheader("Kirim GABUNGAN → Google Sheets")
        push = res["push"]
        if push["kind"] != ("sv" if cfg["suffix"] == " SV" else "ls"):
            st.warning("Hasil di atas dari jenis data yang berbeda dengan pilihan sekarang; "
                       "tab & kolom mengikuti hasil terakhir yang dijalankan.")
        st.caption(f"{len(push['rows']):,} baris GABUNGAN (semua range). Kolom `batas` & `cek creator paid` dikosongkan. "
                   "Hanya menambah di bawah; semua baris dimasukkan (duplikat tidak dicek).")
        msg = st.session_state.get("tap_push_msg")
        if msg:
            st.success(msg)
        b1, b2 = st.columns(2)
        for btn, dry in ((b1, True), (b2, False)):
            label = "🔍 Cek dulu (tidak menulis)" if dry else "📤 Kirim ke Google Sheets"
            if btn.button(label, key=f"tap_push_{'dry' if dry else 'go'}", type="secondary" if dry else "primary",
                          disabled=blocked):
                try:
                    r_ = _tap_push(push, tap_link, tap_tab, dry, tap_cred)
                    note = " (tab kosong: header ikut ditulis)" if r_.get("header_written") else ""
                    if dry:
                        st.info("Cek (belum menulis): " + _fmt_push(r_, res["label"].lower()) + note)
                    else:
                        st.session_state["tap_push_msg"] = "Terkirim. " + _fmt_push(r_, res["label"].lower()) + note
                        st.rerun()
                except Exception as e:  # noqa: BLE001
                    st.error(f"{type(e).__name__}: {e}")
                    st.caption("Pastikan sheet di-share sebagai **Editor** ke email service account di credentials.json, "
                               "dan `pip install google-api-python-client google-auth`.")


def _gs_service(cred_path: Path):
    from google.oauth2 import service_account
    from googleapiclient.discovery import build

    if not cred_path.is_file():
        raise FileNotFoundError(f"credentials.json tidak ditemukan: {cred_path}")
    creds = service_account.Credentials.from_service_account_file(
        str(cred_path), scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"])
    _retry_patch()
    return build("sheets", "v4", credentials=creds, cache_discovery=False)


def _sheet_id(link: str) -> str:
    m = re.search(r"/spreadsheets/d/([a-zA-Z0-9_-]+)", link or "")
    if not m:
        raise ValueError("Link bukan link Google Sheets.")
    return m.group(1)


def _fnum(v) -> float:
    return float(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else 0.0


def _check_gabungan(link: str, tab: str, cred_path: Path) -> dict:
    """Baca seluruh tab SV/LS gabungan lalu hitung total konten, GMV, creator, per brand."""
    svc = _gs_service(cred_path)
    sid = _sheet_id(link)
    rows = svc.spreadsheets().values().get(spreadsheetId=sid, range=f"'{tab}'",
                                           valueRenderOption="UNFORMATTED_VALUE").execute().get("values", [])
    if not rows or not any(str(c).strip() for c in rows[0]):
        return {"empty": True}
    hdr = [_nk(c) for c in rows[0]]

    def ix(pred):
        return next((i for i, a in enumerate(hdr) if pred(a)), None)

    i_id = ix(lambda a: a in ("videoid", "livestreamroomid"))
    i_br = ix(lambda a: a == "brand")
    i_cr = ix(lambda a: a == "creatorname")
    i_gmv = ix(lambda a: a.startswith("affiliate") and a.endswith("gmv"))
    i_dt = ix(lambda a: a == "date")
    if None in (i_id, i_cr, i_gmv):
        raise ValueError("Header tab tidak dikenali (butuh ID konten, Creator name, dan Affiliate ... GMV).")

    def g(r, i):
        return r[i] if i is not None and i < len(r) else ""

    data = [r for r in rows[1:] if str(g(r, i_id)).strip()]
    per, seen, dup = {}, set(), 0
    for r in data:
        b = str(g(r, i_br)).strip() or "(tanpa brand)"
        d = per.setdefault(b, {"Konten": 0, "GMV": 0.0, "cr": set()})
        d["Konten"] += 1
        d["GMV"] += _fnum(g(r, i_gmv))
        c = str(g(r, i_cr)).strip().lower()
        if c:
            d["cr"].add(c)
        k = (str(g(r, i_id)).strip(), b.lower())
        dup += k in seen
        seen.add(k)
    dates = sorted(str(g(r, i_dt))[:10] for r in data if str(g(r, i_dt)).strip())
    table = pd.DataFrame([{"Brand": b, "Konten": d["Konten"], "GMV": d["GMV"], "Creator": len(d["cr"])}
                          for b, d in sorted(per.items())])
    return {"empty": False, "konten": len(data), "gmv": sum(d["GMV"] for d in per.values()),
            "creator": len(set().union(*[d["cr"] for d in per.values()])) if per else 0,
            "dup": dup, "dates": (dates[0], dates[-1]) if dates else None, "table": table}


def _check_brand_tabs(link: str, cred_path: Path) -> dict:
    """Total per tab brand di sheet 'Content Unilever Beauty (LS)' (kolom C username, G tanggal, O GMV); tab PAID dilewati."""
    svc = _gs_service(cred_path)
    sid = _sheet_id(link)
    titles = [x["properties"]["title"] for x in svc.spreadsheets().get(
        spreadsheetId=sid, fields="sheets(properties(title))").execute().get("sheets", [])]
    usable = [t for t in titles if "paid" not in t.lower()]
    recs, allc = [], set()
    for t in usable:
        vr = svc.spreadsheets().values().batchGet(
            spreadsheetId=sid, ranges=[f"'{t}'!{c}:{c}" for c in "CGO"],
            valueRenderOption="UNFORMATTED_VALUE").execute().get("valueRanges", [])
        c_, g_, o_ = ([r[0] if r else "" for r in v.get("values", [])] for v in vr)
        users = [str(u).strip().lower() for u in c_[1:] if str(u).strip()]
        if not users:
            continue
        allc |= set(users)
        recs.append({"Tab": t, "Live": len(users), "GMV": sum(_fnum(v) for v in o_[1:]), "Creator": len(set(users))})
    return {"table": pd.DataFrame(recs), "live": sum(r["Live"] for r in recs),
            "gmv": sum(r["GMV"] for r in recs), "creator": len(allc), "skipped": [t for t in titles if t not in usable]}


def render_cek_page():
    st.title("🔎 Cek Data di Google Sheets")
    st.caption("Lihat isi sheet SEKARANG (total video / LS, GMV, creator) sebelum memasukkan data baru. Hanya membaca, tidak menulis.")
    cred = BASE_DIR / "credentials.json"
    what = st.radio("Cek apa", ["TAP Short Video (tab sv)", "TAP Livestream (tab ls)",
                                "LS per brand (Content Unilever Beauty)"], horizontal=True, key="cek_what")
    try:
        if what.startswith("LS per brand"):
            link = st.text_input("Link Google Sheets", value=LS_BRAND_LINK, key="cek_link_brand")
            if st.button("🔎 Cek", type="primary", key="cek_go_brand"):
                r_ = _check_brand_tabs(link, cred)
                a, b, c = st.columns(3)
                a.metric("Total Live", f"{r_['live']:,}")
                b.metric("Total GMV", _rp(r_["gmv"]))
                c.metric("Creator unik", f"{r_['creator']:,}")
                st.dataframe(r_["table"], hide_index=True, use_container_width=True)
                if r_["skipped"]:
                    st.caption("Tab dilewati (PAID): " + ", ".join(r_["skipped"]))
        else:
            sv = what.startswith("TAP Short")
            link = st.text_input("Link Google Sheets", value=TAP_PUSH_LINK, key="cek_link_tap")
            tab = st.text_input("Nama tab", value="sv" if sv else "ls", key=f"cek_tab_{'sv' if sv else 'ls'}")
            if st.button("🔎 Cek", type="primary", key="cek_go_tap"):
                r_ = _check_gabungan(link, tab, cred)
                if r_["empty"]:
                    st.info(f"Tab '{tab}' masih kosong.")
                else:
                    a, b, c = st.columns(3)
                    a.metric("Total Video" if sv else "Total Livestream", f"{r_['konten']:,}")
                    b.metric("Total GMV", _rp(r_["gmv"]))
                    c.metric("Creator unik", f"{r_['creator']:,}")
                    if r_["dates"]:
                        st.caption(f"Kolom Date dari {r_['dates'][0]} s/d {r_['dates'][1]}")
                    if r_["dup"]:
                        st.warning(f"{r_['dup']:,} baris dobel (ID konten + Brand sama).")
                    st.dataframe(r_["table"], hide_index=True, use_container_width=True)
    except Exception as e:  # noqa: BLE001
        st.error(f"{type(e).__name__}: {e}")
        st.caption("Pastikan sheet di-share ke email service account di credentials.json dan "
                   "`pip install google-api-python-client google-auth`.")


# ============================================================================
# HALAMAN GABUNGAN: hasil TAP (range) + hasil SC (tidak ada di TAP) -> satu tabel layout TAP
# ============================================================================
def _dur_text(a, b) -> str:
    try:
        da = datetime.strptime(str(a)[:19], "%Y-%m-%d %H:%M:%S")
        db = datetime.strptime(str(b)[:19], "%Y-%m-%d %H:%M:%S")
    except ValueError:
        return ""
    sec = max(int((db - da).total_seconds()), 0)
    return f"{sec // 3600}:{sec % 3600 // 60:02d}:{sec % 60:02d}"


def sc_to_tap_layout(kind: str, values: list) -> list[list]:
    """Baris hasil SC (header di values[0]) -> urutan kolom TAP SV/LS (_tap_push_spec). Kolom yang tidak ada di SC dikosongkan."""
    fields = [f[0] for f in _tap_push_spec({"suffix": " SV" if kind == "sv" else " LS"})["fields"]]
    hdr = values[0]
    hi = {h: i for i, h in enumerate(hdr)}
    out = []
    for r in values[1:]:
        def g(k):
            v = r[hi[k]] if k in hi else ""
            return "" if v is None else v
        if kind == "sv":
            post = str(g("Post date"))
            rec = {"Date": post[:10], "Creator name": g("Creator name"), "Video id": g("Video ID"),
                   "Video name": g("Video title"), "Post time": post, "Date Post": post[:10],
                   "Product ID": g("Product ID"), "Product name": g("sku name"), "Brand": g("Brand"),
                   "Sub Brand": g("Sub Product"), "Affiliate video GMV": g("Creator video-attributed GMV"),
                   "Video orders": g("Video-attributed orders"), "Video views": g("Video views"),
                   "Video likes": g("Likes"), "Items sold": g("Video-attributed items sold")}
        else:
            st_, en_ = str(g("LIVE start time")), str(g("LIVE end time"))
            rec = {"Date": st_[:10], "Creator name": g("Creator name"), "Livestream room ID": g("LIVE ID"),
                   "Livestream name": g("LIVE title"), "LIVE time info": f"{st_}-{en_}" if st_ and en_ else "",
                   "Start": st_, "End": en_, "Date Live": st_[:10], "Durasi": _dur_text(st_, en_),
                   "Product ID": g("Product ID"), "Product name": g("sku name") or g("Product name"),
                   "Brand": g("Brand"), "Sub Brand": g("Sub Product"),
                   "Affiliate LIVE GMV": g("Creator LIVE-attributed GMV"), "LIVE orders": g("LIVE-attributed orders"),
                   "LIVE likes": g("Likes"), "Items sold": g("LIVE-attributed items sold")}
        out.append([rec.get(f, "") for f in fields])
    return out


def build_gabungan(kind: str, tap_entry: dict, sc_values: list):
    """TAP range (baris dulu) + SC not-in-TAP -> (header, baris). Kolom terakhir 'Sumber' = TAP / SC."""
    fields = [f[0] for f in _tap_push_spec({"suffix": " SV" if kind == "sv" else " LS"})["fields"]]
    rows = [list(r) + ["TAP"] for r in tap_entry["rows"]]
    rows += [r + ["SC"] for r in sc_to_tap_layout(kind, sc_values)]
    apply_vaseline_rule(fields, rows)
    return fields + ["Sumber"], rows


_PRODERMA_RE = re.compile(r"pro\s*derma", re.I)


def apply_vaseline_rule(header: list, rows: list) -> int:
    """Brand Vaseline: nama produk mengandung 'proderma' / 'pro derma' -> 'Vaseline Pro Derma'; selain itu 'Vaseline'.
    Brand lain tidak disentuh. Mengubah rows di tempat; mengembalikan jumlah baris Pro Derma."""
    bi, pi = header.index("Brand"), header.index("Product name")
    n = 0
    for r in rows:
        if str(r[bi]).strip().lower().startswith("vaseline"):
            if _PRODERMA_RE.search(str(r[pi] or "")):
                r[bi] = "Vaseline Pro Derma"
                n += 1
            else:
                r[bi] = "Vaseline"
    return n


def read_gabungan_sheet(link: str, tabs: list, cred_path: Path):
    """Baca tab gabungan (mis. sv, ls) langsung dari Google Sheets -> ({tab: (header, rows)}, catatan).
    Hanya kolom Brand, Creator name, Product name yang diambil. Aturan Vaseline Pro Derma diterapkan."""
    from google.oauth2 import service_account
    from googleapiclient.discovery import build

    m = re.search(r"/spreadsheets/d/([a-zA-Z0-9_-]+)", link or "")
    if not m:
        raise ValueError("Link gabungan bukan link Google Sheets.")
    sid = m.group(1)
    if not cred_path.is_file():
        raise FileNotFoundError(f"credentials.json tidak ditemukan: {cred_path}")
    creds = service_account.Credentials.from_service_account_file(
        str(cred_path), scopes=["https://www.googleapis.com/auth/spreadsheets"])
    _retry_patch()
    svc = build("sheets", "v4", credentials=creds, cache_discovery=False)
    titles = [x["properties"]["title"] for x in svc.spreadsheets().get(
        spreadsheetId=sid, fields="sheets(properties(title))").execute().get("sheets", [])]
    by_lower = {t.strip().lower(): t for t in titles}
    norm = lambda x: re.sub(r"[^a-z0-9]", "", str(x).lower())  # noqa: E731
    out, notes = {}, []
    for tab in tabs:
        if not tab.strip():
            continue
        title = by_lower.get(tab.strip().lower())
        if title is None:
            notes.append(f"tab '{tab}' tidak ada di sheet gabungan")
            continue
        got = svc.spreadsheets().values().get(spreadsheetId=sid, range=f"'{title}'!1:1").execute().get("values", [])
        idx = {norm(h): i for i, h in enumerate(got[0] if got else [])}
        if "brand" not in idx or "creatorname" not in idx:
            notes.append(f"tab '{title}': kolom 'Brand' / 'Creator name' tidak ditemukan di baris 1")
            continue
        names, cols = ["Brand", "Creator name"], [idx["brand"], idx["creatorname"]]
        if "productname" in idx:
            names.append("Product name")
            cols.append(idx["productname"])
        vr = svc.spreadsheets().values().batchGet(
            spreadsheetId=sid, ranges=[f"'{title}'!{_col_letter(i)}2:{_col_letter(i)}" for i in cols],
            valueRenderOption="UNFORMATTED_VALUE").execute().get("valueRanges", [])
        columns = [[r[0] if r else "" for r in v.get("values", [])] for v in vr]
        n = max((len(c) for c in columns), default=0)
        rows = [[c[i] if i < len(c) else "" for c in columns] for i in range(n)]
        if "Product name" in names:
            apply_vaseline_rule(names, rows)
        else:
            notes.append(f"tab '{title}': tidak ada kolom 'Product name' → aturan Vaseline Pro Derma tidak diterapkan")
        out[title] = (names, rows)
    if not out:
        raise ValueError("Tidak ada tab gabungan yang bisa dibaca. " + "; ".join(notes))
    return out, notes


def usn_by_brand(ready: dict) -> dict:
    """{brand: [username unik]} dari tabel gabungan (SV dan/atau LS), tanpa duplikat per brand (tidak peka huruf besar/kecil)."""
    out, seen = {}, {}
    for _k, (hdr, rows) in ready.items():
        bi, ci = hdr.index("Brand"), hdr.index("Creator name")
        for r in rows:
            b, u = str(r[bi]).strip(), str(r[ci]).strip()
            if not b or not u or b.upper() == "MISSING":
                continue
            if u.lower() in seen.setdefault(b, set()):
                continue
            seen[b].add(u.lower())
            out.setdefault(b, []).append(u)
    return {b: sorted(v, key=str.lower) for b, v in sorted(out.items(), key=lambda kv: kv[0].lower())}


USN_START_ROW = 12   # kolom C, mulai baris 12


def push_usn(usn: dict, link: str, create_missing: bool, dry_run: bool, cred_path: Path) -> list:
    """Satu tab per brand (nama tab = nama brand). Username unik ditulis ke kolom C mulai baris 12, di bawah isi yang sudah ada.
    Username yang sudah ada di kolom C (dari baris 12) tidak ditulis lagi."""
    from google.oauth2 import service_account
    from googleapiclient.discovery import build

    m = re.search(r"/spreadsheets/d/([a-zA-Z0-9_-]+)", link or "")
    if not m:
        raise ValueError("Link target USN bukan link Google Sheets.")
    sid = m.group(1)
    if not cred_path.is_file():
        raise FileNotFoundError(f"credentials.json tidak ditemukan: {cred_path}")
    creds = service_account.Credentials.from_service_account_file(
        str(cred_path), scopes=["https://www.googleapis.com/auth/spreadsheets"])
    _retry_patch()
    svc = build("sheets", "v4", credentials=creds, cache_discovery=False)
    titles = [x["properties"]["title"] for x in svc.spreadsheets().get(
        spreadsheetId=sid, fields="sheets(properties(title))").execute().get("sheets", [])]
    by_lower = {t.strip().lower(): t for t in titles}
    out = []
    for brand, users in usn.items():
        tab = by_lower.get(brand.strip().lower())
        created = False
        if tab is None:
            if not create_missing:
                out.append({"brand": brand, "tab": None, "total": len(users), "added": 0, "skipped": 0,
                            "note": "tab dengan nama brand ini tidak ada"})
                continue
            tab, created = brand, True
            if not dry_run:
                svc.spreadsheets().batchUpdate(spreadsheetId=sid, body={"requests": [
                    {"addSheet": {"properties": {"title": tab}}}]}).execute()
        got = [] if created else svc.spreadsheets().values().get(
            spreadsheetId=sid, range=f"'{tab}'!C{USN_START_ROW}:C", valueRenderOption="UNFORMATTED_VALUE"
        ).execute().get("values", [])
        have = {str(r[0]).strip().lower() for r in got if r and str(r[0]).strip()}
        new = [u for u in users if u.lower() not in have]
        start = USN_START_ROW + len(got)
        if new and not dry_run:
            _ensure_grid(svc, sid, tab, start + len(new) + 5, 3)
            for i in range(0, len(new), 2000):
                svc.spreadsheets().values().update(
                    spreadsheetId=sid, range=f"'{tab}'!C{start + i}", valueInputOption="RAW",
                    body={"values": [[u] for u in new[i:i + 2000]]}).execute()
        out.append({"brand": brand, "tab": tab, "total": len(users), "added": len(new), "skipped": len(users) - len(new),
                    "start_row": start, "created": created})
    return out


def _fmt_usn(x: dict) -> str:
    if not x["tab"]:
        return f"{x['brand']}: {x['note']} ({x['total']:,} username tidak dikirim)"
    return (f"{x['brand']} → tab '{x['tab']}'" + (" (baru)" if x.get("created") else "")
            + f": {x['added']:,} username baru"
            + (f" mulai C{x['start_row']}" if x["added"] else "")
            + (f", {x['skipped']:,} sudah ada" if x["skipped"] else ""))


def _gab_stats(kind: str, header: list, rows: list) -> dict:
    gi = header.index("Affiliate video GMV" if kind == "sv" else "Affiliate LIVE GMV")
    ci = header.index("Creator name")

    def num(v):
        try:
            return float(v)
        except (TypeError, ValueError):
            return 0.0
    return {"n": len(rows), "gmv": sum(num(r[gi]) for r in rows),
            "creator": len({str(r[ci]).strip().lower() for r in rows if str(r[ci]).strip()})}


def render_gabung_page():
    st.title("🧩 Gabungan SC + TAP")
    st.caption("Hasil **TAP (range)** + hasil **SC yang tidak ada di TAP** → satu tabel dengan kolom layout TAP, "
               "satu sheet untuk SV dan satu untuk LS. Kolom terakhir **Sumber** menandai TAP / SC. "
               "Bisa langsung dikirim otomatis ke Google Sheets (tab sv / ls) dan/atau di-download .xlsx.")
    st.info("Data diambil dari hasil yang sudah diproses di sesi ini. Jalankan dulu **TAP** (SV dan/atau LS) dan "
            "**SC BW Video / SC BW Live** dengan RANGE yang sama, lalu buka halaman ini. Jangan refresh tab browser "
            "di antaranya (hasil hilang kalau halaman di-refresh).")
    ss = st.session_state
    done = ss.setdefault("gab_done", {})
    src = {
        "sv": (ss.get("tap_range_sv"), ss.get("sc_result"), "SV: TAP Short Video + SC BW Video"),
        "ls": (ss.get("tap_range_ls"), ss.get("lv_result"), "LS: TAP Livestream + SC BW Live"),
    }
    status = []
    for k, (t, s_, label) in src.items():
        status.append({"Jenis": label,
                       "TAP range": (f"{len(t['rows']):,} baris ({t['range'][0]} → {t['range'][1]})" if t and t["range"][0]
                                     else (f"{len(t['rows']):,} baris (semua tanggal)" if t else "belum diproses")),
                       "SC not in TAP": (f"{len(s_['push']) - 1:,} baris ({s_['rng'][0]} → {s_['rng'][1]})" if s_ else "belum diproses")})
    st.dataframe(pd.DataFrame(status), hide_index=True, use_container_width=True)

    ready = {}
    for k, (t, s_, label) in src.items():
        if not (t and s_):
            continue
        st.subheader(label)
        problems = []
        if t["missing"]:
            problems.append(f"TAP: {t['missing']:,} Product ID belum ada di mapping (N/A).")
        if len(s_["missing_df"]):
            problems.append(f"SC: {len(s_['missing_df']):,} Product ID / sku name belum ada di mapping (N/A).")
        if problems:
            st.error("⛔ " + " ".join(problems) + " Lengkapi mapping lalu proses ulang; gabungan tidak dibuat.")
            continue
        if t["range"][0] and (t["range"][0], t["range"][1]) != tuple(s_["rng"]):
            st.warning(f"RANGE TAP ({t['range'][0]} → {t['range'][1]}) berbeda dengan RANGE SC "
                       f"({s_['rng'][0]} → {s_['rng'][1]}).")
        hdr, rows = build_gabungan(k, t, s_["push"])
        n_tap = len(t["rows"])
        a, b, c = st.columns(3)
        stt = _gab_stats(k, hdr, rows)
        a.metric("Total baris", f"{stt['n']:,}", help=f"TAP {n_tap:,} + SC {stt['n'] - n_tap:,}")
        b.metric("Total GMV", _rp(stt["gmv"]))
        c.metric("Creator unik", f"{stt['creator']:,}")
        st.caption(f"TAP {n_tap:,} baris + SC {stt['n'] - n_tap:,} baris.")
        with st.expander("Preview (10 baris SC pertama)"):
            st.dataframe(pd.DataFrame(rows[n_tap:n_tap + 10], columns=hdr), hide_index=True, use_container_width=True)
        ready[k] = (hdr, rows)

    if ready:
        st.subheader("Kirim ke Google Sheets")
        gl = st.text_input("Link Google Sheets tujuan", value=TAP_PUSH_LINK, key="gab_link")
        g1, g2 = st.columns(2)
        gtabs = {"sv": g1.text_input("Nama tab SV", value="sv", key="gab_tab_sv"),
                 "ls": g2.text_input("Nama tab LS", value="ls", key="gab_tab_ls")}
        gauto = st.checkbox("Langsung kirim otomatis begitu gabungan siap (sekali per hasil proses)", value=True, key="gab_auto")
        st.caption("Menambah di bawah data yang sudah ada di tab; kolom disejajarkan dengan header tab (kolom 'Sumber' tidak ikut). "
                   "Semua baris masuk, duplikat tidak dicek. Jangan kirim GABUNGAN dari menu TAP ke tab yang sama untuk range ini "
                   "supaya tidak dobel.")
        gcred = BASE_DIR / "credentials.json"
        done = ss.setdefault("gab_done", {})

        def do_push(k, dry):
            hdr_, rows_ = ready[k]
            spec_ = _tap_push_spec({"suffix": " SV" if k == "sv" else " LS"})
            push_ = {"kind": k, "header": [f[0] for f in spec_["fields"]], "aliases": [f[1] for f in spec_["fields"]],
                     "id_alias": spec_["id_alias"], "rows": [r[:-1] for r in rows_]}
            return _tap_push(push_, gl, gtabs[k], dry, gcred)

        for k in ready:
            sig = (k, id(src[k][0]), id(src[k][1]))
            if gauto and sig not in done:
                try:
                    with st.spinner(f"Mengirim {k.upper()} gabungan…"):
                        done[sig] = ("ok", "Terkirim otomatis. " + _fmt_push(do_push(k, False), "baris"))
                except Exception as e:  # noqa: BLE001
                    done[sig] = ("err", f"GAGAL kirim otomatis: {type(e).__name__}: {e}")
            if sig in done:
                (st.success if done[sig][0] == "ok" else st.error)(f"{k.upper()}: {done[sig][1]}")
            b1, b2 = st.columns(2)
            if b1.button(f"🔍 Cek dulu {k.upper()} (tidak menulis)", key=f"gab_dry_{k}"):
                try:
                    st.info("Cek (belum menulis): " + _fmt_push(do_push(k, True), "baris"))
                except Exception as e:  # noqa: BLE001
                    st.error(f"{type(e).__name__}: {e}")
            if b2.button(f"📤 Kirim {k.upper()} sekarang" + (" (lagi)" if sig in done else ""), key=f"gab_go_{k}"):
                try:
                    done[sig] = ("ok", "Terkirim. " + _fmt_push(do_push(k, False), "baris"))
                except Exception as e:  # noqa: BLE001
                    done[sig] = ("err", f"GAGAL kirim: {type(e).__name__}: {e}")
                (st.success if done[sig][0] == "ok" else st.error)(f"{k.upper()}: {done[sig][1]}")

    st.subheader("Daftar USN unik per brand → spreadsheet lain")
    st.caption("Username unik per brand diambil langsung dari **link Google Sheets gabungan** (tab sv / ls), tanpa upload file "
               "dan tanpa memproses TAP + SC dulu. Satu tab per brand di spreadsheet tujuan (nama tab = nama brand), username "
               f"ditulis ke kolom C mulai baris {USN_START_ROW}, di bawah isi yang sudah ada; username yang sudah ada dilewati. "
               "Brand Vaseline: nama produk mengandung 'proderma' / 'pro derma' → **Vaseline Pro Derma**, selain itu **Vaseline**.")
    usn_mode = st.radio("Sumber username", ["Link Google Sheets gabungan (tab sv / ls)", "Hasil proses TAP + SC di sesi ini"],
                        horizontal=True, key="gab_usn_mode")
    from_sheet = usn_mode.startswith("Link")
    if from_sheet:
        s1, s2, s3 = st.columns([3, 1, 1])
        src_link = s1.text_input("Link Google Sheets gabungan (sumber)", value=TAP_PUSH_LINK, key="gab_usn_src_link")
        src_tabs = [s2.text_input("Tab SV", value="sv", key="gab_usn_src_sv"),
                    s3.text_input("Tab LS", value="ls", key="gab_usn_src_ls")]
    usn_link = st.text_input("Link Google Sheets tujuan USN", value=_load_settings().get("usn_link", ""),
                             key="gab_usn_link", placeholder="https://docs.google.com/spreadsheets/d/.../edit")
    _save_setting("usn_link", usn_link.strip())
    u1, u2 = st.columns(2)
    usn_create = u1.checkbox("Buat tab kalau belum ada", value=False, key="gab_usn_create",
                             help="Mati: brand yang tabnya tidak ada dilewati dan dilaporkan.")
    usn_auto = u2.checkbox("Langsung kirim otomatis (sekali per hasil proses)", value=True, key="gab_usn_auto",
                           disabled=from_sheet, help="Hanya untuk sumber 'hasil proses di sesi ini'. Dari link gabungan, "
                                                      "pengiriman lewat tombol supaya sheet besar tidak dibaca tiap halaman dibuka.")
    gcred2 = BASE_DIR / "credentials.json"

    def show(kind_, msgs_):
        (st.success if kind_ == "ok" else st.error)("USN: \n\n" + "\n\n".join(msgs_))

    if from_sheet:
        ub1, ub2 = st.columns(2)
        want = "dry" if ub1.button("🔍 Ambil dari link gabungan & cek (tidak menulis)", key="gab_usn_dry",
                                   disabled=not usn_link.strip()) else (
               "go" if ub2.button("📤 Ambil dari link gabungan & kirim USN", type="primary", key="gab_usn_go",
                                  disabled=not usn_link.strip()) else None)
        if want:
            try:
                with st.spinner("Membaca sheet gabungan…"):
                    sheets_, notes_ = read_gabungan_sheet(src_link, src_tabs, gcred2)
                usn = usn_by_brand(sheets_)
                for n_ in notes_:
                    st.warning(n_)
                st.dataframe(pd.DataFrame([{"Brand": b_, "Username unik": len(v_)} for b_, v_ in usn.items()]),
                             hide_index=True, use_container_width=True)
                with st.spinner("Memeriksa/mengirim USN…"):
                    res_u = push_usn(usn, usn_link, usn_create, want == "dry", gcred2)
                msgs = [_fmt_usn(x) for x in res_u]
                if want == "dry":
                    st.info("Cek (belum menulis): \n\n" + "\n\n".join(msgs))
                else:
                    show("ok", msgs)
            except Exception as e:  # noqa: BLE001
                st.error(f"{type(e).__name__}: {e}")
        else:
            st.caption("Klik tombol di atas untuk membaca tab sv / ls dari link gabungan lalu mengirim USN per brand.")
    else:
        usn = usn_by_brand(ready) if ready else {}
        if usn:
            st.dataframe(pd.DataFrame([{"Brand": b_, "Username unik": len(v_)} for b_, v_ in usn.items()]),
                         hide_index=True, use_container_width=True)
        else:
            st.info("Belum ada hasil TAP + SC di sesi ini. Pilih sumber 'Link Google Sheets gabungan' untuk tanpa proses.")
        usn_sig = ("usn",) + tuple((k, id(src[k][0]), id(src[k][1])) for k in ready)

        def run_usn(dry):
            return push_usn(usn, usn_link, usn_create, dry, gcred2)

        can_usn = bool(usn) and bool(usn_link.strip())
        if can_usn and usn_auto and usn_sig not in done:
            try:
                with st.spinner("Mengirim USN per brand…"):
                    done[usn_sig] = ("ok", [_fmt_usn(x) for x in run_usn(False)])
            except Exception as e:  # noqa: BLE001
                done[usn_sig] = ("err", [f"GAGAL kirim USN otomatis: {type(e).__name__}: {e}"])
        if usn_sig in done:
            show(*done[usn_sig])
        ub1, ub2 = st.columns(2)
        if ub1.button("🔍 Cek dulu USN (tidak menulis)", key="gab_usn_dry2", disabled=not can_usn):
            try:
                st.info("Cek (belum menulis): \n\n" + "\n\n".join(_fmt_usn(x) for x in run_usn(True)))
            except Exception as e:  # noqa: BLE001
                st.error(f"{type(e).__name__}: {e}")
        if ub2.button("📤 Kirim USN sekarang", key="gab_usn_go2", disabled=not can_usn):
            try:
                done[usn_sig] = ("ok", [_fmt_usn(x) for x in run_usn(False)])
            except Exception as e:  # noqa: BLE001
                done[usn_sig] = ("err", [f"GAGAL kirim USN: {type(e).__name__}: {e}"])
            show(*done[usn_sig])

    if ready and st.button("📦 Siapkan file gabungan (.xlsx)", type="primary", key="gab_build"):
        import openpyxl
        wb = openpyxl.Workbook(write_only=True)
        for k, (hdr, rows) in ready.items():
            ws = wb.create_sheet("SV GABUNGAN" if k == "sv" else "LS GABUNGAN")
            ws.append(hdr)
            for r in rows:
                ws.append(["" if v is None else (str(v) if isinstance(v, (date, datetime)) else v) for v in r])
        bio = io.BytesIO()
        wb.save(bio)
        ss["gab_xlsx"] = bio.getvalue()
        ss["gab_kinds"] = list(ready)
    if ss.get("gab_xlsx"):
        st.download_button("⬇️ Download gabungan (.xlsx)", ss["gab_xlsx"],
                           file_name="GABUNGAN_SC_TAP_" + "_".join(k.upper() for k in ss.get("gab_kinds", [])) + ".xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="gab_dl")
    if not ready:
        st.caption("Belum ada pasangan TAP + SC yang siap digabung (hanya memengaruhi gabungan & kirim ke tab sv/ls; "
                   "daftar USN tetap bisa diambil langsung dari link gabungan).")


# ============================================================================
# GABUNG KONTEN + LABEL KEGIATAN (dua sheet RAW: SV & LS)
# ============================================================================
RAW_SV_LINK = "https://docs.google.com/spreadsheets/d/1HDhxZtvUIhanLZ-lyTjXaWw4gYnsDoxhTnsvmdIfECs/edit"
RAW_LS_LINK = "https://docs.google.com/spreadsheets/d/1ltIPVD0ds8NIqcGH1s5jpcUe955xr63dSN4qHcAepxs/edit"

KONTEN_SV_HEADER = ["Post date", "USN", "Video name", "Video ID", "Brand", "Sub Brand", "SKU name", "Week", "Nama kegiatan", "Jenis kegiatan", "Sumber"]
KONTEN_LS_HEADER = ["Live date", "USN", "Live ID", "Brand", "Sub Brand", "SKU name", "Start time", "End time", "Duration", "GMV",
                    "Live name", "Week", "Nama kegiatan", "Jenis kegiatan", "Sumber"]


def _kn(x) -> str:
    return re.sub(r"[^a-z0-9]", "", str(x).lower())


def _to_date(v):
    """'2026-11-02', '2026-11-02 06:40:39', '11/2/2026', serial angka -> date atau None."""
    if v is None or str(v).strip() == "":
        return None
    s = str(v).strip()
    if re.match(r"^\d{4}-\d{2}-\d{2}", s):
        try:
            return date(int(s[:4]), int(s[5:7]), int(s[8:10]))
        except ValueError:
            return None
    if re.fullmatch(r"\d{5}(\.\d+)?", s):
        return date(1899, 12, 30) + timedelta(days=int(float(s)))
    try:
        d = pd.to_datetime(s, dayfirst=False, errors="coerce")
        return None if pd.isna(d) else d.date()
    except Exception:  # noqa: BLE001
        return None


def _to_num(v) -> float:
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return float(v)
    s = re.sub(r"[^0-9,.\-]", "", str(v or ""))
    if not s or s in "-.,":
        return 0.0
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".") if s.rfind(",") > s.rfind(".") else s.replace(",", "")
    elif s.count(".") > 1:
        s = s.replace(".", "")
    elif s.count(",") > 1:
        s = s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".") if len(s.split(",")[1]) != 3 else s.replace(",", "")
    try:
        return float(s)
    except ValueError:
        return 0.0


def week_of(d) -> int | str:
    """Week 1 = tgl 1-7, 2 = 8-14, 3 = 15-21, 4 = 22-28, 5 = 29-akhir bulan."""
    return min((d.day - 1) // 7 + 1, 5) if d else ""


def _hash_re(tag: str):
    t = str(tag).strip().lstrip("#").strip()
    return re.compile(r"#" + re.escape(t) + r"(?![0-9A-Za-z_])", re.I) if t else None


def label_kegiatan(text, d, rules: list):
    """(Nama kegiatan, Jenis) dari hashtag di teks + rentang tanggal. Beberapa aturan cocok -> digabung ' | '."""
    names, jenis = [], []
    for r in rules:
        rx = r.get("_rx")
        if rx is None or not rx.search(str(text or "")):
            continue
        a, b = r.get("start"), r.get("end")
        if (a or b) and d is None:
            continue
        if (a and d < a) or (b and d > b):
            continue
        if r["name"] not in names:
            names.append(r["name"])
        if r.get("jenis") and r["jenis"] not in jenis:
            jenis.append(r["jenis"])
    return " | ".join(names), " | ".join(jenis)


def prep_rules(df) -> list:
    rules = []
    for _, x in df.iterrows():
        tag, name = str(x.get("Hashtag") or "").strip(), str(x.get("Nama kegiatan") or "").strip()
        if not tag or tag.lower() == "nan" or not name or name.lower() == "nan":
            continue
        a, b = _to_date(x.get("Tanggal mulai")), _to_date(x.get("Tanggal akhir"))
        jn = str(x.get("Jenis") or "").strip()
        rules.append({"tag": tag, "name": name, "jenis": "" if jn.lower() == "nan" else jn,
                      "start": a, "end": b, "_rx": _hash_re(tag)})
    return rules


def _read_tab_cols(svc, sid: str, title: str, wanted: dict) -> dict:
    """wanted: {nama_kanonik: [alias dinormalisasi...]} -> {nama_kanonik: list nilai (tanpa header)}; kolom hilang -> None."""
    got = svc.spreadsheets().values().get(spreadsheetId=sid, range=f"'{title}'!1:1").execute().get("values", [])
    idx = {_kn(h): i for i, h in enumerate(got[0] if got else [])}
    pick = {}
    for canon, aliases in wanted.items():
        pick[canon] = next((idx[a] for a in aliases if a in idx), None)
    cols = [(c, i) for c, i in pick.items() if i is not None]
    res = {c: None for c in wanted}
    if not cols:
        return res
    vr = svc.spreadsheets().values().batchGet(
        spreadsheetId=sid, ranges=[f"'{title}'!{_col_letter(i)}2:{_col_letter(i)}" for _, i in cols],
        valueRenderOption="FORMATTED_VALUE").execute().get("valueRanges", [])
    for (c, _), v in zip(cols, vr):
        res[c] = [r[0] if r else "" for r in v.get("values", [])]
    n = max([len(v) for v in res.values() if v is not None] or [0])
    for c in res:
        if res[c] is not None and len(res[c]) < n:
            res[c] = res[c] + [""] * (n - len(res[c]))
    res["_n"] = n
    return res


def _svc_rw(cred_path: Path):
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
    if not cred_path.is_file():
        raise FileNotFoundError(f"credentials.json tidak ditemukan: {cred_path}")
    creds = service_account.Credentials.from_service_account_file(
        str(cred_path), scopes=["https://www.googleapis.com/auth/spreadsheets"])
    _retry_patch()
    return build("sheets", "v4", credentials=creds, cache_discovery=False)


def _find_title(svc, sid: str, name: str):
    titles = [x["properties"]["title"] for x in svc.spreadsheets().get(
        spreadsheetId=sid, fields="sheets(properties(title))").execute().get("sheets", [])]
    return next((t for t in titles if t.strip().lower() == name.strip().lower()), None), titles


def _vaseline(brand, prod) -> str:
    if str(brand).strip().lower().startswith("vaseline"):
        return "Vaseline Pro Derma" if _PRODERMA_RE.search(str(prod or "")) else "Vaseline"
    return str(brand).strip()


def _dur_from(a, b) -> str:
    try:
        s = int((pd.to_datetime(b) - pd.to_datetime(a)).total_seconds())
    except Exception:  # noqa: BLE001
        return ""
    return f"{s // 3600}:{s % 3600 // 60:02d}:{s % 60:02d}" if s >= 0 else ""


def read_raw_konten(kind: str, link: str, tap_tab: str, sc_tab: str, cred_path: Path):
    """Baca tab TAP + SC dari sheet RAW -> daftar record (dict) unik, TAP diutamakan atas SC. Return (records, notes)."""
    svc = _svc_rw(cred_path)
    sid = _sheet_id(link)
    notes, recs, seen = [], [], set()
    if kind == "sv":
        spec = [
            (tap_tab, "TAP", {"date": ["dateposted", "datepost", "postdate"], "usn": ["creatorname"], "name": ["videoname"],
                              "id": ["videoid"], "brand": ["brand"], "sub": ["subbrand"], "prod": ["productname"],
                              "sku": ["skuname", "productname"]}),
            (sc_tab, "SC", {"date": ["postdate", "dateposted"], "usn": ["creatorname"], "name": ["videotitle", "videoname"],
                            "id": ["videoid"], "brand": ["brand"], "sub": ["subproduct", "subbrand"], "prod": ["skuname", "productname"],
                            "sku": ["skuname", "productname"]}),
        ]
    else:
        spec = [
            (tap_tab, "TAP", {"date": ["datelive"], "usn": ["creatorname"], "id": ["livestreamroomid", "liveid"],
                              "brand": ["brand"], "sub": ["subbrand"], "start": ["start"], "end": ["end"], "dur": ["durasi"],
                              "gmv": ["affiliatelivegmv"], "name": ["livestreamname", "livetitle"], "prod": ["productname"],
                              "sku": ["skuname", "productname"], "foll": ["followers"]}),
            (sc_tab, "SC", {"id": ["liveid", "livestreamroomid"], "usn": ["creatorname"], "brand": ["brand"],
                            "sub": ["subproduct", "subbrand"], "start": ["livestarttime", "start"], "end": ["liveendtime", "end"],
                            "gmv": ["creatorliveattributedgmv", "affiliatelivegmv"], "name": ["livetitle", "livestreamname"],
                            "prod": ["productname", "skuname"], "sku": ["skuname", "productname"]}),
        ]
    for tab, src, wanted in spec:
        title, titles = _find_title(svc, sid, tab)
        if title is None:
            notes.append(f"tab '{tab}' tidak ada (tab yang ada: {', '.join(titles)})")
            continue
        cols = _read_tab_cols(svc, sid, title, wanted)
        miss = [k for k in ("usn", "id", "brand") if cols.get(k) is None]
        if miss or not cols.get("_n"):
            notes.append(f"tab '{title}': kolom {', '.join(miss) or 'data'} tidak ditemukan / kosong")
            continue
        n = cols["_n"]
        g = lambda k, i: (cols[k][i] if cols.get(k) is not None else "")  # noqa: E731
        add, by_key = {}, {}
        for i in range(n):
            vid = str(g("id", i)).strip()
            if not vid:
                continue
            brand = _vaseline(g("brand", i), g("prod", i))
            sub = str(g("sub", i)).strip()
            key = (vid, brand.lower(), sub.lower())
            sku = str(g("sku", i)).strip()
            if kind == "ls":
                e = add.setdefault(key, {"gmv": 0.0, "i": i, "skus": []})
                e["gmv"] += _to_num(g("gmv", i))
                if sku and sku not in e["skus"]:
                    e["skus"].append(sku)
            elif key not in seen:
                seen.add(key)
                rec = {"key": key, "src": src, "date": _to_date(g("date", i)), "usn": str(g("usn", i)).strip(),
                       "name": str(g("name", i)), "id": vid, "brand": brand, "sub": sub, "skus": [sku] if sku else []}
                recs.append(rec)
                by_key[key] = rec
            elif by_key.get(key, {}).get("src") == src and sku and sku not in by_key[key]["skus"]:
                by_key[key]["skus"].append(sku)
        if kind == "ls":
            for key, a in add.items():
                if key in seen:
                    continue
                seen.add(key)
                i = a["i"]
                st_, en_ = str(g("start", i)).strip(), str(g("end", i)).strip()
                d = _to_date(g("date", i)) if cols.get("date") is not None else None
                d = d or _to_date(st_)
                dur = str(g("dur", i)).strip() or _dur_from(st_, en_)
                recs.append({"key": key, "src": src, "date": d, "usn": str(g("usn", i)).strip(), "name": str(g("name", i)),
                             "id": key[0], "brand": _vaseline(g("brand", i), g("prod", i)), "sub": str(g("sub", i)).strip(),
                             "start": st_, "end": en_, "dur": dur, "gmv": a["gmv"], "skus": a["skus"],
                             "foll": (_to_num(g("foll", i)) or "") if cols.get("foll") is not None else ""})
        notes.append(f"{src} '{title}': {n:,} baris dibaca")
    return recs, notes


def build_konten(kind: str, recs: list, rules: list):
    rows = []
    for r in recs:
        d = r["date"]
        ds = d.isoformat() if d else ""
        lab, jn = label_kegiatan(r["name"], d, rules)
        if kind == "sv":
            rows.append([ds, r["usn"], r["name"], r["id"], r["brand"], r["sub"], " | ".join(r.get("skus") or []),
                         week_of(d), lab, jn, r["src"]])
        else:
            rows.append([ds, r["usn"], r["id"], r["brand"], r["sub"], " | ".join(r.get("skus") or []), r["start"], r["end"], r["dur"], r["gmv"],
                         r["name"], week_of(d), lab, jn, r["src"]])
    rows.sort(key=lambda x: (x[0], x[1]))
    return (KONTEN_SV_HEADER if kind == "sv" else KONTEN_LS_HEADER), rows


def write_konten_tab(header: list, rows: list, link: str, tab: str, cred_path: Path) -> str:
    """Tulis hasil ke tab (dibuat kalau belum ada; isi lama dikosongkan lalu ditulis ulang)."""
    svc = _svc_rw(cred_path)
    sid = _sheet_id(link)
    title, _ = _find_title(svc, sid, tab)
    if title is None:
        svc.spreadsheets().batchUpdate(spreadsheetId=sid, body={"requests": [
            {"addSheet": {"properties": {"title": tab}}}]}).execute()
        title = tab
    else:
        svc.spreadsheets().values().clear(spreadsheetId=sid, range=f"'{title}'", body={}).execute()
    _ensure_grid(svc, sid, title, len(rows) + 5, len(header))
    data = [header] + rows
    for i in range(0, len(data), 5000):
        svc.spreadsheets().values().update(
            spreadsheetId=sid, range=f"'{title}'!A{i + 1}", valueInputOption="RAW",
            body={"values": data[i:i + 5000]}).execute()
    return f"{len(rows):,} baris ditulis ke tab '{title}'"


BRAND_REKAP_ORDER = ["Vaseline", "Vaseline Pro Derma", "Ponds", "GAL", "Tresemme", "Clear", "Sunsilk", "Dove"]


def _brand_group(brand, sub) -> str:
    """Vaseline dengan Sub Brand / produk 'proderma' / 'pro derma' -> 'Vaseline Pro Derma'; Vaseline lain -> 'Vaseline'."""
    b = str(brand).strip()
    if b.lower().startswith("vaseline"):
        if b.lower() == "vaseline pro derma" or _PRODERMA_RE.search(str(sub or "")):
            return "Vaseline Pro Derma"
        return "Vaseline"
    return b


def _kategori(jenis) -> str:
    j = str(jenis or "").lower()
    return "Kompetisi" if "kompetisi" in j else ("Event" if "event" in j else "Mandatory")


def rekap_per_brand(kind: str, header: list, rows: list, tiers: dict | None = None) -> dict:
    """Tabel konten (hasil gabung / build_konten) -> {brand: (header, rows)}. Per USN unik per brand:
    LS = jumlah LS per Week 1-5 + total; SV = per Week dan total: Mandatory (= Total - Kompetisi - Event), Kompetisi, Event, Total.
    Konten yang cocok dengan Kompetisi dan Event dihitung sebagai Kompetisi. Baris TOTAL + 'USN unik' di bawah."""
    ix = {_kn(h): i for i, h in enumerate(header)}
    idk = "videoid" if kind == "sv" else "liveid"
    for need in ("usn", idk, "brand", "week"):
        if need not in ix:
            raise ValueError(f"kolom '{need}' tidak ditemukan di tabel gabungan")
    si, ji = ix.get("subbrand"), ix.get("jeniskegiatan")
    seen, data, names = set(), {}, {}
    for r in rows:
        g = lambda i: r[i] if i is not None and i < len(r) else ""  # noqa: E731
        usn, cid = str(g(ix["usn"])).strip(), str(g(ix[idk])).strip()
        if not usn or not cid:
            continue
        br = _brand_group(g(ix["brand"]), g(si))
        if not br or (br.lower(), cid) in seen:
            continue
        seen.add((br.lower(), cid))
        try:
            wk = int(float(g(ix["week"])))
        except (TypeError, ValueError):
            wk = 0
        wk = wk if 1 <= wk <= 5 else 0
        cat = _kategori(g(ji)) if kind == "sv" else "Total"
        uk = usn.lower()
        names.setdefault((br, uk), usn)
        data.setdefault(br, {}).setdefault(uk, {}).setdefault((wk, cat), 0)
        data[br][uk][(wk, cat)] += 1
    cats = ["Mandatory", "Kompetisi", "Event"] if kind == "sv" else []
    out = {}
    order = {b.lower(): i for i, b in enumerate(BRAND_REKAP_ORDER)}
    for br in sorted(data, key=lambda b: (order.get(b.lower(), 99), b)):
        if kind == "sv":
            hdr = ["USN"] + [f"W{w} {c}" for w in range(1, 6) for c in cats + ["Total"]] + cats + ["Total"]
        else:
            hdr = ["USN"] + [f"Week {w}" for w in range(1, 6)] + ["Total LS"]
        has0 = any(k[0] == 0 for u in data[br].values() for k in u)
        if has0 and kind == "ls":
            hdr.insert(6, "Tanpa week")
        lines = []
        for u, cnt in data[br].items():
            u = names[(br, u)]
            if kind == "sv":
                row = [u]
                for w in range(1, 6):
                    vs = [cnt.get((w, c), 0) for c in cats]
                    row += vs + [sum(vs)]
                tot = [sum(v for (w_, c_), v in cnt.items() if c_ == c) for c in cats]
                row += tot + [sum(tot)]
            else:
                row = [u] + [cnt.get((w, "Total"), 0) for w in range(1, 6)]
                if has0:
                    row.append(cnt.get((0, "Total"), 0))
                row.append(sum(cnt.values()))
            lines.append(row)
        lines.sort(key=lambda x: (-x[-1], x[0].lower()))
        if tiers is not None:
            tmap = tiers.get(_kn(br), {})
            for ln in lines:
                ln.insert(1, tmap.get(ln[0].strip().lower(), "MASS"))
            hdr.insert(1, "Tier")
        c0 = 2 if tiers is not None else 1
        lines.append(["TOTAL"] + [""] * (c0 - 1) + [sum(l[i] for l in lines) for i in range(c0, len(hdr))])
        if tiers is not None:
            body = lines[:-1]
            lines.append(["USN unik", "", len(body)] + [""] * (len(hdr) - 3))
            lines.append([""] * len(hdr))
            for t in ("TOP", "MID", "MASS"):
                sel = [l for l in body if l[1] == t]
                lines.append([f"{t} ({len(sel)} USN)", t] + [sum(l[i] for l in sel) for i in range(2, len(hdr))])
        else:
            lines.append(["USN unik", len(lines) - 1] + [""] * (len(hdr) - 2))
        out[br] = (hdr, lines)
    return out


def read_tab_table(link: str, tab: str, cred_path: Path):
    svc = _svc_rw(cred_path)
    sid = _sheet_id(link)
    title, titles = _find_title(svc, sid, tab)
    if title is None:
        raise ValueError(f"tab '{tab}' tidak ada (tab yang ada: {', '.join(titles)})")
    vals = svc.spreadsheets().values().get(spreadsheetId=sid, range=f"'{title}'",
                                           valueRenderOption="FORMATTED_VALUE").execute().get("values", [])
    if not vals:
        raise ValueError(f"tab '{title}' kosong")
    return vals[0], vals[1:]


SV_DEST_LINK = "https://docs.google.com/spreadsheets/d/1XyLbR5fvye34D5xSPW55X_Y7rio5gAnjIYoReePVrXU/edit"
LS_DEST_LINK = "https://docs.google.com/spreadsheets/d/1d8q0AbefAB-4Ng4ykUJz0XdtSJoPMjczHXeiLMH0VA4/edit"
LS_TAB_MAP = {"vaseline": "LS VASELINE", "vaseline pro derma": "LS VASELINE PRO DERMA", "ponds": "LS PONDS", "gal": "LS GAL", "dove": "LS DOVE",
              "tresemme": "LS TRESEMME", "clear": "LS CLEAR", "sunsilk": "LS Sunsilk"}
LS_FIRST_ROW = 19   # baris data pertama di tab "LS <BRAND>" template (header di baris 18, rumus contoh di baris 19)


def read_tiers(link: str, tab: str, cred_path: Path) -> dict:
    """Tab TOP MID: baris 1 = nama brand di kolom A, C, E, ... ; di bawahnya pasangan (username, TOP/MID).
    -> {kunci_brand(_kn): {username_lowercase: 'TOP'|'MID'}}"""
    svc = _svc_rw(cred_path)
    sid = _sheet_id(link)
    title, titles = _find_title(svc, sid, tab)
    if title is None:
        raise ValueError(f"tab '{tab}' tidak ada (tab yang ada: {', '.join(titles)})")
    vals = svc.spreadsheets().values().get(spreadsheetId=sid, range=f"'{title}'",
                                           valueRenderOption="FORMATTED_VALUE").execute().get("values", [])
    if not vals:
        raise ValueError(f"tab '{title}' kosong")
    out = {}
    for ci, name in enumerate(vals[0]):
        if not str(name).strip():
            continue
        m = out.setdefault(_kn(name), {})
        for r in vals[1:]:
            u = str(r[ci]).strip().lower() if ci < len(r) else ""
            t = str(r[ci + 1]).strip().upper() if ci + 1 < len(r) else ""
            if u and t in ("TOP", "MID"):
                m[u] = t
    return out


def _serial_dt(v) -> float | str:
    dt = pd.to_datetime(str(v or "").strip(), errors="coerce")
    if pd.isna(dt):
        return ""
    delta = dt.to_pydatetime().replace(tzinfo=None) - datetime(1899, 12, 30)
    return round(delta.days + delta.seconds / 86400, 9)


def _ls_template_rows(recs: list):
    """recs LS -> ({tab: [baris input]}, {brand_tanpa_tab: jumlah}). Ponds: kolom Brand diisi Sub Brand."""
    groups, skipped = {}, {}
    for r in recs:
        g = _brand_group(r["brand"], r["sub"])
        tab = LS_TAB_MAP.get(g.lower())
        if tab is None:
            skipped[g or "(kosong)"] = skipped.get(g or "(kosong)", 0) + 1
            continue
        if g.lower() == "ponds":
            bval = r["sub"] or "Ponds"
        elif g.lower() == "vaseline":
            bval = "VASELINE"
        elif g.lower() == "vaseline pro derma":
            bval = "VASELINE PRO DERMA"
        else:
            bval = g
        d = r["date"]
        groups.setdefault(tab, []).append((
            (d or date(1900, 1, 1)), str(r["start"]),
            [bval, r["usn"], str(r["id"]), (d - date(1899, 12, 30)).days if d else "", _serial_dt(r["start"]),
             _serial_dt(r["end"]), float(r["gmv"] or 0), r["name"], r["src"], r.get("foll", "")]))
    out = {t: [x[2] for x in sorted(v, key=lambda z: (z[0], z[1]))] for t, v in groups.items()}
    return out, skipped


def fill_ls_template(link: str, recs: list, dry_run: bool, cred_path: Path, replace: bool = False) -> dict:
    """Isi kolom input tab 'LS <BRAND>' di template dari data LS. Default: TAMBAH saja — Livestream Room (LS ID) yang sudah ada di
    kolom D tab itu dilewati, yang baru ditaruh tepat di bawah isi yang ada. replace=True: isi lama (baris 19 ke bawah) diganti semua.
    Rumus (kolom A, E, F, H-K, N, R) disalin dari baris 19 ke baris yang diisi, jadi tab TOP MID / RECAP ikut ter-update."""
    rows_by_tab, skipped = _ls_template_rows(recs)
    svc = _svc_rw(cred_path)
    sid = _sheet_id(link)
    meta = svc.spreadsheets().get(spreadsheetId=sid, fields="sheets(properties(sheetId,title,index,gridProperties))").execute()
    props = {x["properties"]["title"].strip().lower(): x["properties"] for x in meta.get("sheets", [])}
    res, notes = [], []
    blank_inputs = [f"B{LS_FIRST_ROW}:D{LS_FIRST_ROW}", f"G{LS_FIRST_ROW}", f"L{LS_FIRST_ROW}:M{LS_FIRST_ROW}",
                    f"O{LS_FIRST_ROW}:Q{LS_FIRST_ROW}", f"S{LS_FIRST_ROW}"]
    for tab in LS_TAB_MAP.values():
        rows = rows_by_tab.get(tab, [])
        pr = props.get(tab.lower())
        fresh = False
        if pr is None and tab == "LS VASELINE PRO DERMA" and "ls vaseline" in props:
            if dry_run:
                res.append({"tab": tab + " (akan dibuat dari salinan LS VASELINE)", "rows": len(rows), "skip": 0})
                continue
            if not rows:
                continue
            src = props["ls vaseline"]
            rp = svc.spreadsheets().batchUpdate(spreadsheetId=sid, body={"requests": [{"duplicateSheet": {
                "sourceSheetId": src["sheetId"], "newSheetName": tab, "insertSheetIndex": src.get("index", 0) + 1}}]}).execute()
            pr = rp["replies"][0]["duplicateSheet"]["properties"]
            props[tab.lower()] = pr
            fresh = True
            notes.append(f"tab '{tab}' dibuat dari salinan 'LS VASELINE' (isi datanya dikosongkan dulu)")
        if pr is None:
            notes.append(f"tab '{tab}' tidak ada di sheet tujuan")
            continue
        title, shid = pr["title"], pr["sheetId"]
        if fresh or replace:
            if not dry_run:
                svc.spreadsheets().values().clear(spreadsheetId=sid, range=f"'{title}'!A{LS_FIRST_ROW + 1}:S", body={}).execute()
                for rg in blank_inputs:
                    svc.spreadsheets().values().clear(spreadsheetId=sid, range=f"'{title}'!{rg}", body={}).execute()
            have, start = set(), LS_FIRST_ROW
        else:
            got = svc.spreadsheets().values().get(spreadsheetId=sid, range=f"'{title}'!D{LS_FIRST_ROW}:D",
                                                  valueRenderOption="FORMATTED_VALUE").execute().get("values", [])
            have = {str(r[0]).strip() for r in got if r and str(r[0]).strip()}
            start = LS_FIRST_ROW + len(got)
            # baris 19 masih kosong (template belum pernah diisi) -> mulai dari 19
            if not have:
                start = LS_FIRST_ROW
        new, seen_ids = [], set()
        for r in rows:
            lid = str(r[2]).strip()
            if lid in have or lid in seen_ids:
                continue
            seen_ids.add(lid)
            new.append(r)
        n = len(new)
        res.append({"tab": title, "rows": n, "skip": len(rows) - n, "start": start})
        if dry_run or not n:
            continue
        last = start + n - 1
        _ensure_grid(svc, sid, title, last + 5, 19)
        c0 = max(start, LS_FIRST_ROW + 1)
        if last >= c0:
            svc.spreadsheets().batchUpdate(spreadsheetId=sid, body={"requests": [{"copyPaste": {
                "source": {"sheetId": shid, "startRowIndex": LS_FIRST_ROW - 1, "endRowIndex": LS_FIRST_ROW,
                           "startColumnIndex": 0, "endColumnIndex": 19},
                "destination": {"sheetId": shid, "startRowIndex": c0 - 1, "endRowIndex": last,
                                "startColumnIndex": 0, "endColumnIndex": 19},
                "pasteType": "PASTE_NORMAL"}}]}).execute()
        for i in range(0, n, 4000):
            part, r0 = new[i:i + 4000], start + i
            r1 = r0 + len(part) - 1
            data = [
                {"range": f"'{title}'!B{r0}:D{r1}", "values": [x[0:3] for x in part]},
                {"range": f"'{title}'!G{r0}:G{r1}", "values": [[x[3]] for x in part]},
                {"range": f"'{title}'!L{r0}:M{r1}", "values": [x[4:6] for x in part]},
                {"range": f"'{title}'!O{r0}:Q{r1}", "values": [x[6:9] for x in part]},
                {"range": f"'{title}'!S{r0}:S{r1}", "values": [[x[9]] for x in part]},
            ]
            svc.spreadsheets().values().batchUpdate(spreadsheetId=sid, body={"valueInputOption": "RAW", "data": data}).execute()
    return {"tabs": res, "skipped": skipped, "notes": notes}


def _rekap_ui(kind: str, link: str, header: list, rows: list, cred: Path, recs: list | None = None):
    k = kind.upper()
    ss = st.session_state
    st.subheader("Rekap per brand")
    tiers, dest = None, link
    if kind == "sv":
        st.caption("Per brand: USN unik dan jumlah video per Week 1–5 + total, dipisah **Mandatory / Kompetisi / Event / Total**. "
                   "Kompetisi dan Event dari Jenis di aturan label; **Mandatory = Total − Kompetisi − Event** (konten tanpa label). "
                   "Konten yang kena Kompetisi sekaligus Event dihitung Kompetisi.")
        dest = st.text_input("Link Google Sheets tujuan rekap (tab 'TOP MID' ada di sheet ini)", value=SV_DEST_LINK, key="rk_dest_sv")
        t1, t2 = st.columns([1, 2])
        use_t = t1.checkbox("Hitung TOP / MID / MASS", value=True, key="rk_use_tier")
        tier_tab = t2.text_input("Tab daftar TOP MID", value="TOP MID", key="rk_tier_tab")
        if use_t:
            ck = (dest, tier_tab)
            if ss.get("rk_tier_key") != ck or t1.button("🔄 Muat ulang TOP MID", key="rk_tier_reload"):
                try:
                    with st.spinner("Membaca tab TOP MID…"):
                        ss["rk_tiers"] = read_tiers(dest, tier_tab, cred)
                    ss["rk_tier_key"] = ck
                except Exception as e:  # noqa: BLE001
                    ss["rk_tiers"], ss["rk_tier_key"] = None, ck
                    st.error(f"Gagal membaca TOP MID: {type(e).__name__}: {e}")
            tiers = ss.get("rk_tiers")
            if tiers:
                st.caption("TOP/MID per username dari tab TOP MID (dicocokkan per brand, huruf besar/kecil diabaikan); username yang tidak ada "
                           "di daftar brand itu dihitung **MASS**. Jumlah username TOP/MID: "
                           + ", ".join(f"{b} {sum(1 for v in m.values() if v == 'TOP')}/{sum(1 for v in m.values() if v == 'MID')}"
                                       for b, m in tiers.items()))
    else:
        st.caption("Per brand: USN unik, jumlah LS (Live ID unik) per Week 1–5 dan total.")
    st.caption("Satu tab per brand di sheet tujuan (nama tab = 'Rekap " + k + " - <brand>'), isi lama tab rekap dikosongkan lalu ditulis ulang. "
               "Vaseline hanya menghitung yang BUKAN proderma; proderma masuk tab 'Vaseline Pro Derma'. Konten tanpa tanggal valid "
               "tidak masuk kolom week tapi tetap masuk total.")
    try:
        rk = rekap_per_brand(kind, header, rows, tiers)
    except Exception as e:  # noqa: BLE001
        st.error(f"{type(e).__name__}: {e}")
        return
    if not rk:
        st.warning("Tidak ada data untuk direkap.")
        return
    summ = []
    for b, (h_, r_) in rk.items():
        tot = next(x for x in r_ if x[0] == "TOTAL")
        x = {"Brand": b, "USN unik": next(x_ for x_ in r_ if x_[0] == "USN unik")[2 if tiers is not None else 1], "Total": tot[-1]}
        if tiers is not None:
            for t in ("TOP", "MID", "MASS"):
                x[f"USN {t}"] = int(next(z for z in r_ if str(z[0]).startswith(t + " (")) [0].split("(")[1].split()[0])
        summ.append(x)
    st.dataframe(pd.DataFrame(summ), hide_index=True, use_container_width=True)
    pick = st.selectbox("Lihat rekap brand", list(rk), key=f"rk_pick_{kind}")
    st.dataframe(pd.DataFrame(rk[pick][1], columns=rk[pick][0]), hide_index=True, use_container_width=True)
    c1, c2 = st.columns(2)
    if kind == "sv" and c1.button("📤 Tulis semua tab rekap ke Google Sheets", key=f"rk_go_{kind}"):
        try:
            with st.spinner("Menulis tab rekap…"):
                for b, (h_, r_) in rk.items():
                    st.write("• " + b + ": " + write_konten_tab(h_, r_, dest, f"Rekap {k} - {b}", cred))
            st.success("Selesai.")
        except Exception as e:  # noqa: BLE001
            st.error(f"{type(e).__name__}: {e}")
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as xw:
        for b, (h_, r_) in rk.items():
            pd.DataFrame(r_, columns=h_).to_excel(xw, sheet_name=b[:31], index=False)
    c2.download_button("⬇️ Download rekap (.xlsx)", bio.getvalue(), file_name=f"REKAP_{k}_PER_BRAND.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=f"rk_dl_{kind}")

    if kind == "ls" and recs is not None:
        st.divider()
        st.subheader("Isi template LS (auto-update)")
        st.caption("Mengisi kolom input tab **LS VASELINE / PONDS / GAL / DOVE / TRESEMME / CLEAR / Sunsilk** di sheet template mulai baris 19 "
                   "(Brand, Creator Username, Livestream Room, Live Date, Start Hour, Stop Hour, Live GMV, Livestream name, Source, FOLLS). "
                   "Rumus di baris 19 disalin ke baris yang diisi, jadi tab TOP MID, RECAP dan lainnya ter-update otomatis. **LS ID yang sudah "
                   "ada di kolom Livestream Room tab itu dilewati (tidak masuk dua kali)**; yang baru ditaruh di bawah isi yang ada. Ponds: kolom Brand diisi Sub Brand. Tab **LS VASELINE PRO DERMA** dibuat otomatis (salinan LS VASELINE) kalau belum ada; LS Vaseline dengan proderma masuk ke sana, sisanya ke LS VASELINE.")
        ldest = st.text_input("Link Google Sheets template LS (NOV)", value=LS_DEST_LINK, key="rk_dest_ls")
        repl = st.checkbox("Ganti semua isi tab LS per brand (hapus data lama dulu) — biasanya JANGAN dicentang", value=False,
                           key="rk_ls_replace")
        b1, b2 = st.columns(2)
        for btn, dry, lbl in ((b1, True, "🔍 Cek dulu (tidak menulis)"), (b2, False, "📤 Isi tab LS sekarang")):
            if btn.button(lbl, key=f"rk_ls_{dry}"):
                try:
                    with st.spinner("Memproses…"):
                        r = fill_ls_template(ldest, recs, dry, cred, repl)
                    (st.info if dry else st.success)(("Cek (belum menulis): " if dry else "Selesai. ")
                                                     + "; ".join(f"{x['tab']}: {x['rows']:,} baru"
                                                                 + (f", {x['skip']:,} sudah ada" if x.get("skip") else "")
                                                                 for x in r["tabs"]))
                    if r["skipped"]:
                        st.warning("Tidak punya tab di template: " + ", ".join(f"{b} ({n:,} LS)" for b, n in r["skipped"].items()))
                    for nt in r["notes"]:
                        st.warning(nt)
                except Exception as e:  # noqa: BLE001
                    st.error(f"{type(e).__name__}: {e}")


def top_hashtags(texts: list, n: int = 40) -> pd.DataFrame:
    from collections import Counter
    c = Counter()
    for t in texts:
        for h in set(x.lower() for x in re.findall(r"#[0-9A-Za-z_]+", str(t))):
            c[h] += 1
    return pd.DataFrame([{"Hashtag": h, "Jumlah konten": k} for h, k in c.most_common(n)])


def label_existing_tab(kind: str, link: str, tab: str, rules: list, overwrite: bool, dry_run: bool, cred_path: Path) -> dict:
    """Beri label (Week, Nama kegiatan, Jenis kegiatan) langsung pada tab gabungan yang SUDAH ada, tanpa menggabung ulang.
    Hanya 3 kolom label yang ditulis; kolom lain tidak disentuh. Label terisi (mis. diketik manual) tidak ditimpa
    kecuali overwrite=True. Kolom label yang belum ada ditambahkan di ujung kanan header."""
    svc = _svc_rw(cred_path)
    sid = _sheet_id(link)
    title, titles = _find_title(svc, sid, tab)
    if title is None:
        raise ValueError(f"tab '{tab}' tidak ada (tab yang ada: {', '.join(titles)})")
    vals = svc.spreadsheets().values().get(spreadsheetId=sid, range=f"'{title}'",
                                           valueRenderOption="FORMATTED_VALUE").execute().get("values", [])
    if not vals:
        raise ValueError(f"tab '{title}' kosong")
    hdr = list(vals[0])
    idx = {_kn(h): i for i, h in enumerate(hdr)}
    tcol = idx.get("videoname" if kind == "sv" else "livename")
    dcol = idx.get("postdate" if kind == "sv" else "livedate")
    if tcol is None or dcol is None:
        raise ValueError("kolom " + ("'Video name' / 'Post date'" if kind == "sv" else "'Live name' / 'Live date'")
                         + f" tidak ditemukan di baris 1 tab '{title}'")
    out_cols, added = {}, []
    for nm in ("Week", "Nama kegiatan", "Jenis kegiatan"):
        i = idx.get(_kn(nm))
        if i is None:
            i = len(hdr)
            hdr.append(nm)
            idx[_kn(nm)] = i
            added.append(nm)
        out_cols[nm] = i
    body = vals[1:]
    new = {nm: [] for nm in out_cols}
    changed = labeled = 0
    for r in body:
        g = lambda i: r[i] if i < len(r) else ""  # noqa: E731
        d = _to_date(g(dcol))
        lab, jn = label_kegiatan(g(tcol), d, rules)
        cur = {nm: g(i) for nm, i in out_cols.items()}
        res = {"Week": week_of(d) if d else cur["Week"],
               "Nama kegiatan": lab if (overwrite or not str(cur["Nama kegiatan"]).strip()) else cur["Nama kegiatan"],
               "Jenis kegiatan": jn if (overwrite or not str(cur["Jenis kegiatan"]).strip()) else cur["Jenis kegiatan"]}
        if not overwrite:
            if str(cur["Week"]).strip():
                res["Week"] = cur["Week"]
        if any(str(res[k]) != str(cur[k]) for k in res):
            changed += 1
        if res["Nama kegiatan"]:
            labeled += 1
        for k in out_cols:
            new[k].append(res[k] if res[k] != "" else "")
    info = {"rows": len(body), "changed": changed, "labeled": labeled, "added_cols": added, "tab": title}
    if dry_run or not body:
        return info
    _ensure_grid(svc, sid, title, len(body) + 2, len(hdr))
    for nm, i in out_cols.items():
        cl = _col_letter(i)
        if nm in added:
            svc.spreadsheets().values().update(spreadsheetId=sid, range=f"'{title}'!{cl}1", valueInputOption="RAW",
                                               body={"values": [[nm]]}).execute()
        col = [[v] for v in new[nm]]
        for k in range(0, len(col), 5000):
            svc.spreadsheets().values().update(spreadsheetId=sid, range=f"'{title}'!{cl}{k + 2}",
                                               valueInputOption="RAW", body={"values": col[k:k + 5000]}).execute()
    return info


def _rules_editor(kind: str, texts: list):
    import json
    st.subheader("Label Nama kegiatan")
    st.caption("Satu baris = satu hashtag → satu nama kegiatan, **Jenis** (Kompetisi / Event), plus rentang tanggal (posting / live). "
               "Kosongkan tanggal = berlaku untuk semua tanggal. Hashtag yang sama boleh dipakai di beberapa baris dengan rentang "
               "tanggal berbeda. Konten yang cocok dengan >1 aturan diberi beberapa nama dipisah ' | '.")
    ss = st.session_state
    cols = ["Hashtag", "Nama kegiatan", "Jenis", "Tanggal mulai", "Tanggal akhir"]
    if "kn_rules" not in ss:
        try:
            saved = json.loads(_load_settings().get("kegiatan_rules") or "null")
        except Exception:  # noqa: BLE001
            saved = None
        df = pd.DataFrame(saved if isinstance(saved, list) and saved else [{}])
        ss["kn_rules"] = df.reindex(columns=cols).fillna("")
    df0 = ss["kn_rules"].copy()
    for c in ("Tanggal mulai", "Tanggal akhir"):
        df0[c] = df0[c].map(_to_date)
    edited = st.data_editor(
        df0, num_rows="dynamic", use_container_width=True, key=f"kn_ed_{kind}",
        column_config={"Hashtag": st.column_config.TextColumn("Hashtag (mis. #stayclearstayultra)"),
                       "Nama kegiatan": st.column_config.TextColumn("Nama kegiatan"),
                       "Jenis": st.column_config.SelectboxColumn("Jenis", options=["Kompetisi", "Event"]),
                       "Tanggal mulai": st.column_config.DateColumn("Tanggal mulai"),
                       "Tanggal akhir": st.column_config.DateColumn("Tanggal akhir")})
    rules = prep_rules(edited)
    ser = edited.reindex(columns=cols).copy()
    for c in ("Tanggal mulai", "Tanggal akhir"):
        ser[c] = ser[c].map(lambda v: (_to_date(v).isoformat() if _to_date(v) else ""))
    if st.button("💾 Simpan aturan label", key=f"kn_save_{kind}"):
        try:
            cur = _load_settings()
            cur["kegiatan_rules"] = json.dumps(ser.fillna("").to_dict("records"), ensure_ascii=False)
            SETTINGS_FILE.write_text(json.dumps(cur, ensure_ascii=False, indent=2), encoding="utf-8")
            ss["kn_rules"] = ser.fillna("")
            st.success("Aturan tersimpan (dipakai lagi saat app dibuka berikutnya, juga untuk SV dan LS).")
        except OSError as e:
            st.error(f"Gagal menyimpan: {e}")
    if texts:
        with st.expander("Hashtag paling sering muncul (bantu isi aturan)"):
            st.dataframe(top_hashtags(texts), hide_index=True, use_container_width=True)
    return rules


REKAP_MODE = "Rekap per brand (langsung dari TAP + SC, tanpa gabung dulu)"


def render_konten_page(force_mode: str = ""):
    st.title("📋 Rekap per Brand" if force_mode else "🧱 Gabung Konten + Label Kegiatan")
    if force_mode:
        st.caption("USN unik per brand + jumlah konten per week, langsung dari tab TAP + SC di sheet RAW — tidak perlu menggabung "
                   "atau menulis tab gabungan dulu.")
    ss = st.session_state
    kind_lbl = st.radio("Jenis", ["Short Video (SV)", "Livestream (LS)"], horizontal=True, key="kn_kind")
    kind = "sv" if kind_lbl.startswith("Short") else "ls"
    mode = force_mode or st.radio(
        "Mode", ["Gabung dari tab TAP + SC (lalu beri label)", "Beri label di tab gabungan yang sudah ada", REKAP_MODE],
        horizontal=True, key="kn_mode")
    link = st.text_input("Link Google Sheets RAW", value=RAW_SV_LINK if kind == "sv" else RAW_LS_LINK, key=f"kn_link_{kind}")
    cred = BASE_DIR / "credentials.json"
    st.caption("Week: 1 = tgl 1–7, 2 = 8–14, 3 = 15–21, 4 = 22–28, 5 = 29–akhir bulan. Kolom **Jenis kegiatan** = Kompetisi / Event "
               "sesuai aturan label. Sheet harus di-share (Editor) ke service account.")

    if mode.startswith("Beri label"):
        out_tab = st.text_input("Tab gabungan yang mau dilabeli", value="Gabungan SV" if kind == "sv" else "Gabungan LS",
                                key=f"kn_ex_tab_{kind}")
        rules = _rules_editor(kind, [])
        ow = st.checkbox("Timpa label yang sudah terisi (kalau tidak dicentang, sel yang sudah ada isinya — termasuk yang "
                         "diketik manual — dibiarkan)", value=False, key=f"kn_ow_{kind}")
        st.caption("Hanya kolom Week, Nama kegiatan, Jenis kegiatan yang ditulis; kolom lain tidak disentuh. Kolom yang belum ada "
                   "ditambahkan di ujung kanan. Membutuhkan kolom " + ("'Video name' dan 'Post date'." if kind == "sv"
                                                                       else "'Live name' dan 'Live date'."))
        b1, b2 = st.columns(2)
        for btn, dry, lbl in ((b1, True, "🔍 Cek dulu (tidak menulis)"), (b2, False, "🏷️ Beri label sekarang")):
            if btn.button(lbl, key=f"kn_ex_{dry}_{kind}"):
                try:
                    with st.spinner("Memproses…"):
                        r = label_existing_tab(kind, link, out_tab, rules, ow, dry, cred)
                    msg = (f"Tab '{r['tab']}': {r['rows']:,} baris, {r['labeled']:,} berlabel, {r['changed']:,} baris berubah"
                           + (f"; kolom baru: {', '.join(r['added_cols'])}" if r["added_cols"] else ""))
                    (st.info if dry else st.success)(("Cek (belum menulis): " if dry else "Selesai. ") + msg)
                except Exception as e:  # noqa: BLE001
                    st.error(f"{type(e).__name__}: {e}")
        return

    c2, c3 = st.columns(2)
    tap_tab = c2.text_input("Tab TAP", value="SV TAP" if kind == "sv" else "LS TAP", key=f"kn_tap_{kind}")
    sc_tab = c3.text_input("Tab SC", value="SC VID Range" if kind == "sv" else "SC LS Range", key=f"kn_sc_{kind}")
    st.caption("Membaca tab TAP dan SC, menggabungkannya jadi daftar unik (TAP diutamakan bila konten sama ada di dua tab). "
               "Brand Vaseline: nama produk mengandung 'proderma' / 'pro derma' → Vaseline Pro Derma, selain itu Vaseline. "
               "Konten unik per " + ("Video ID + Brand + Sub Brand." if kind == "sv" else "Live ID + Brand + Sub Brand (GMV dijumlah per produk)."))
    if st.button("📥 Baca data dari Google Sheets", type="primary", key=f"kn_read_{kind}"):
        try:
            with st.spinner("Membaca tab TAP & SC (bisa beberapa menit untuk data besar)…"):
                recs, notes = read_raw_konten(kind, link, tap_tab, sc_tab, cred)
            ss[f"kn_recs_{kind}"] = {"recs": recs, "notes": notes}
        except Exception as e:  # noqa: BLE001
            st.error(f"{type(e).__name__}: {e}")
    got = ss.get(f"kn_recs_{kind}")
    if not got:
        st.info("Klik **Baca data dari Google Sheets** dulu. Kalau tab gabungan sudah ada dan tinggal dilabeli, pakai mode "
                "'Beri label di tab gabungan yang sudah ada'.")
        return
    for nt in got["notes"]:
        st.write("• " + nt)
    recs = got["recs"]
    if not recs:
        st.warning("Tidak ada konten terbaca.")
        return
    today = date.today()
    m1 = today.replace(day=1)
    m2 = (m1.replace(day=28) + timedelta(days=4)).replace(day=1) - timedelta(days=1)
    lbl_d = "Post date" if kind == "sv" else "Live date"
    f1, f2, f3 = st.columns([1, 1, 2])
    use_f = f3.checkbox(f"Hanya {lbl_d} dalam rentang ini", value=True, key=f"kn_usef_{kind}")
    d_from = f1.date_input(f"{lbl_d} dari", value=m1, key=f"kn_df_{kind}")
    d_to = f2.date_input(f"{lbl_d} sampai", value=m2, key=f"kn_dt_{kind}")
    if use_f:
        recs_f = [r for r in recs if r["date"] and d_from <= r["date"] <= d_to]
        st.caption(f"{len(recs_f):,} dari {len(recs):,} konten masuk rentang {d_from} → {d_to}. Konten yang diposting / live di luar "
                   "rentang (walau punya GMV di periode ini) dan yang tanggalnya kosong tidak ikut.")
    else:
        recs_f = recs
    if not recs_f:
        st.warning("Tidak ada konten dalam rentang tanggal itu.")
        return
    rules = _rules_editor(kind, [r["name"] for r in recs_f])

    header, rows = build_konten(kind, recs_f, rules)
    if mode.startswith("Rekap"):
        _rekap_ui(kind, link, header, rows, cred, recs_f)
        return
    li = header.index("Nama kegiatan")
    n_lab = sum(1 for r in rows if r[li])
    a, b, c = st.columns(3)
    a.metric("Total konten unik", f"{len(rows):,}")
    b.metric("Sudah berlabel", f"{n_lab:,}")
    c.metric("Belum berlabel", f"{len(rows) - n_lab:,}")
    if kind == "ls":
        gi = header.index("GMV")
        st.metric("Total GMV", _rp(sum(r[gi] for r in rows)))
    st.dataframe(pd.DataFrame(rows[:200], columns=header), hide_index=True, use_container_width=True)
    st.caption("Preview 200 baris pertama.")

    st.subheader("Simpan hasil")
    d1, d2 = st.columns(2)
    out_tab = d1.text_input("Nama tab tujuan (di sheet RAW yang sama)", value="Gabungan SV" if kind == "sv" else "Gabungan LS",
                            key=f"kn_out_{kind}")
    d1.caption("Tab dibuat kalau belum ada; **isi lama tab ini dikosongkan lalu ditulis ulang** (label manual di tab itu ikut hilang — "
               "kalau sudah ada isinya dan tinggal dilabeli, pakai mode 'Beri label di tab gabungan yang sudah ada'). "
               "Tab TAP / SC tidak disentuh.")
    if d1.button("📤 Tulis ke Google Sheets", key=f"kn_write_{kind}"):
        try:
            with st.spinner("Menulis…"):
                st.success(write_konten_tab(header, rows, link, out_tab, cred))
        except Exception as e:  # noqa: BLE001
            st.error(f"{type(e).__name__}: {e}")
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as xw:
        pd.DataFrame(rows, columns=header).to_excel(xw, sheet_name="Gabungan " + kind.upper(), index=False)
    d2.download_button("⬇️ Download .xlsx", bio.getvalue(), file_name=f"KONTEN_GABUNGAN_{kind.upper()}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=f"kn_dl_{kind}")


if PAGE.startswith("🧩"):
    render_gabung_page()
    st.stop()
if PAGE.startswith("🧱"):
    render_konten_page()
    st.stop()
if PAGE.startswith("📋"):
    render_konten_page(REKAP_MODE)
    st.stop()
if PAGE.startswith("🔎"):
    render_cek_page()
    st.stop()
if PAGE.startswith("🎬"):
    render_tap_page()
    st.stop()


# ----------------------------------------------------------------------------
# Engine loader
# ----------------------------------------------------------------------------
def find_engine_scripts() -> list[Path]:
    found = []
    for p in sorted(BASE_DIR.glob("*.py")):
        if p.resolve() == Path(__file__).resolve():
            continue
        try:
            text = p.read_text(encoding="utf-8", errors="ignore")
        except OSError:
            continue
        if _is_streamlit_app(text):
            continue
        if (re.search(r"^def enrich_sc\(", text, re.M)
                and re.search(r"^def build_brand_unique_rows\(", text, re.M)):
            found.append(p)
    return found


@st.cache_resource
def load_module(path: str):
    spec = importlib.util.spec_from_file_location("sc_bw_engine", path)
    mod = importlib.util.module_from_spec(spec)
    sys.modules["sc_bw_engine"] = mod
    spec.loader.exec_module(mod)
    return mod


with st.sidebar:
    st.header("⚙️ Opsi")
    with st.expander("Lanjutan"):
        engine_override = st.text_input("Path script engine (opsional)", value="")
        cred_override = st.text_input(
            "Path credentials.json (hanya untuk fitur kirim ke Google Sheets)",
            value="", placeholder="kosong = credentials.json di folder app")

if engine_override.strip():
    SCRIPT_PATH = Path(engine_override.strip().strip('"'))
    if not SCRIPT_PATH.exists():
        st.error(f"Script engine tidak ada: {SCRIPT_PATH}")
        st.stop()
    if _is_streamlit_app(SCRIPT_PATH.read_text(encoding="utf-8", errors="ignore")):
        st.error(f"`{SCRIPT_PATH.name}` adalah aplikasi Streamlit, bukan script engine. Pilih `sc_bw_video_ebu.py`.")
        st.stop()
else:
    candidates = find_engine_scripts()
    if not candidates:
        st.error(
            f"Script engine tidak ditemukan di folder `{BASE_DIR}`.\n\n"
            "Taruh script SC BW (yang berisi fungsi `enrich_sc`) di folder yang sama "
            "dengan `app_konten.py`, atau isi path-nya di sidebar → Lanjutan."
        )
        st.write("File .py di folder ini:", [p.name for p in BASE_DIR.glob("*.py")])
        st.stop()
    elif len(candidates) == 1:
        SCRIPT_PATH = candidates[0]
    else:
        SCRIPT_PATH = Path(st.sidebar.selectbox("Pilih script engine", [str(p) for p in candidates]))

engine = load_module(str(SCRIPT_PATH))
st.sidebar.caption(f"Engine: `{SCRIPT_PATH.name}`")


# Excel bisa memberi angka native (float). parse_percent bawaan mengubahnya ke string
# dulu sehingga 0.123 bisa salah dibaca jadi 123. Patch aman untuk input numerik.
_orig_parse_percent = engine.parse_percent


def _safe_parse_percent(value):
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return 0.0 if (isinstance(value, float) and math.isnan(value)) else float(value)
    return _orig_parse_percent(value)


engine.parse_percent = _safe_parse_percent

SKU_COL = "sku name"
PUSH_COLUMNS = [
    "Video title", "Video ID", "Post date", "Video link", "Creator name", "Product ID",
    SKU_COL, "Brand", "Sub Product", "Creator video-attributed GMV", "Video-attributed orders",
    "AOV", "Video-attributed items sold", "Likes", "Comments", "Shares", "Video views", "CTR",
    "Est. commission", "GMV Affiliate Order (Est. Commission Base)", "Jumlah Order Affiliate",
]
DEFAULT_BW_LINK = "https://docs.google.com/spreadsheets/d/1Wi0kcm7apS4gbLDcHJdz8P1iRohF62aGzM38Oy_f7gM/edit?gid=0#gid=0"
DEFAULT_TARGET_LINK = (
    "https://docs.google.com/spreadsheets/d/1Xlo85BMLtZZF8gYSu8434KAru80U5A6KAzxfUktOF_U/edit?usp=sharing"
)


# ----------------------------------------------------------------------------
# Reader lokal (pengganti detect_and_read_source Google Sheets)
# ----------------------------------------------------------------------------
def _is_blank(v) -> bool:
    if v is None:
        return True
    if isinstance(v, str):
        return v.strip() == ""
    try:
        return bool(pd.isna(v))
    except (TypeError, ValueError):
        return False


def _cell(v, is_id: bool):
    if _is_blank(v):
        return ""
    if isinstance(v, (pd.Timestamp, datetime, date)):
        return v.strftime("%Y-%m-%d %H:%M:%S")
    if hasattr(v, "item") and not isinstance(v, str):
        v = v.item()
    if is_id:
        if isinstance(v, bool):
            return str(v)
        if isinstance(v, int):
            return str(v)
        if isinstance(v, float):
            if abs(v) >= 1e15:
                raise ValueError(
                    f"ID {v!r} tersimpan sebagai ANGKA di Excel sehingga digit terpotong. "
                    "Pastikan kolom ID di file sumber berformat Text."
                )
            return str(int(v)) if v.is_integer() else repr(v)
        return str(v).strip()
    return v


def load_table(src, required: list[str], label: str):
    """src = (nama_file, bytes | path_str). Auto-detect sheet + baris header."""
    name, data = src
    target = io.BytesIO(data) if isinstance(data, (bytes, bytearray)) else data
    if Path(name).suffix.lower() in (".csv", ".txt"):
        is_txt = Path(name).suffix.lower() == ".txt"
        sheets = {"csv": pd.read_csv(target, header=None, dtype=str, keep_default_na=False,
                                     encoding="utf-8-sig", sep=None if is_txt else ",",
                                     engine="python" if is_txt else "c")}
    else:
        sheets = pd.read_excel(target, sheet_name=None, header=None, dtype=object)

    required_norm = None if callable(required) else {engine.normalize_header(x) for x in required}

    for sheet_name, df in sheets.items():
        for i in range(min(20, len(df))):
            raw = df.iloc[i].tolist()
            header = [("" if _is_blank(v) else str(v).strip()) for v in raw]
            while header and not header[-1]:
                header.pop()
            if not header:
                continue
            if callable(required):
                if not required(header):
                    continue
            elif not required_norm.issubset({engine.normalize_header(h) for h in header if h}):
                continue

            width = len(header)
            id_flags = [engine.normalize_header(h) in ID_HEADERS for h in header]
            rows = []
            body = df.iloc[i + 1:, :width].astype(object).values.tolist()
            for r in body:
                if all(_is_blank(v) for v in r):
                    continue
                r = list(r) + [""] * (width - len(r))
                rows.append({header[j]: _cell(r[j], id_flags[j]) for j in range(width) if header[j]})
            return header, rows

    raise ValueError(f"{label} ({name}): tidak menemukan header yang sesuai. Required: {required}")


# ----------------------------------------------------------------------------
# Input helpers
# ----------------------------------------------------------------------------
def collect(uploads, path_text: str) -> list[tuple[str, object]]:
    items: list[tuple[str, object]] = []
    for up in uploads or []:
        items.append((up.name, up.getvalue()))
    for line in (path_text or "").splitlines():
        line = line.strip().strip('"').strip("'")
        if not line:
            continue
        p = Path(line)
        if p.is_dir():
            for f in sorted(p.iterdir()):
                if f.suffix.lower() in EXCEL_EXT and not f.name.startswith("~$"):
                    items.append((f.name, str(f)))
        elif p.is_file():
            items.append((p.name, str(p)))
        else:
            raise FileNotFoundError(f"Path tidak ditemukan: {line}")
    return items


def fetch_gsheet_csv(url: str) -> tuple[str, bytes]:
    """Ambil 1 tab Google Sheets (yang dibagikan 'Anyone with the link') sebagai CSV."""
    import urllib.error
    import urllib.request

    url = url.strip().strip("\"'<>").split()[0] if url.strip() else url
    direct_csv = bool(re.search(r"(output|format)=csv", url))
    m = re.search(r"/d/([a-zA-Z0-9_-]{15,})", url) or re.search(r"[?&]id=([a-zA-Z0-9_-]{15,})", url)
    if direct_csv and "docs.google.com" in url and not m:
        export, sid = url, "pub"   # link "Publish to web → CSV"
    elif not m:
        raise ValueError(
            "Link tidak dikenali sebagai Google Sheets.\n"
            f"Yang terbaca: {url[:120]!r}\n"
            "Pakai link dari tombol Share / address bar, contoh: "
            "https://docs.google.com/spreadsheets/d/<ID>/edit?gid=0  "
            "(link file Drive/Excel, link pendek, atau link yang dipotong tidak bisa dipakai)."
        )
    else:
        sid = m.group(1)
        g = re.search(r"[#&?]gid=(\d+)", url)
        export = f"https://docs.google.com/spreadsheets/d/{sid}/export?format=csv"
        if g:
            export += f"&gid={g.group(1)}"
    import time as _time
    for attempt in range(3):   # DNS/koneksi kadang putus sesaat -> coba lagi
        try:
            with urllib.request.urlopen(export, timeout=60) as resp:
                ctype = resp.headers.get("Content-Type", "")
                data = resp.read()
            break
        except urllib.error.HTTPError as e:
            raise RuntimeError(
                f"Gagal mengambil Google Sheets (HTTP {e.code}). Pastikan akses sheet diset "
                "'Anyone with the link' → Viewer."
            ) from e
        except urllib.error.URLError as e:
            if attempt < 2:
                _time.sleep(2 * (attempt + 1))
                continue
            raise RuntimeError(
                "Komputer ini tidak bisa terhubung ke docs.google.com (internet/DNS putus, VPN atau proxy kantor memblokir). "
                f"Detail: {e.reason}. Cek koneksi internet lalu jalankan ulang. Kalau koneksinya memang bermasalah, "
                "download tab mapping BW sebagai .xlsx/.csv lalu isi kolom mapping dengan path file itu."
            ) from e
    if "text/html" in ctype.lower():
        raise RuntimeError(
            "Google meminta login. Ubah akses sheet mapping jadi 'Anyone with the link' → Viewer."
        )
    return f"bw_mapping_{sid[:6]}.csv", data


def get_bw_source(text: str):
    text = (text or "").strip().strip('"')
    if not text:
        return None
    if text.lower().startswith("http"):
        return fetch_gsheet_csv(text)
    p = Path(text)
    if not p.is_file():
        raise FileNotFoundError(f"File mapping tidak ditemukan: {text}")
    return (p.name, str(p))


class LiveWriter:
    def __init__(self, placeholder):
        self.placeholder = placeholder
        self.buf = ""

    def write(self, s):
        self.buf += s
        self.placeholder.code(self.buf[-20000:], language="text")
        return len(s)

    def flush(self):
        pass


# ----------------------------------------------------------------------------
# Pipeline (sama dengan main() engine, tanpa Google Sheets)
# ----------------------------------------------------------------------------
def cred_file() -> Path:
    if cred_override.strip():
        return Path(cred_override.strip().strip('"'))
    return BASE_DIR / "credentials.json"


def service_account_email() -> str | None:
    import json
    try:
        return json.loads(cred_file().read_text(encoding="utf-8")).get("client_email")
    except Exception:  # noqa: BLE001
        return None


def sheets_service():
    from google.oauth2 import service_account
    from googleapiclient.discovery import build

    creds = service_account.Credentials.from_service_account_file(
        str(cred_file()), scopes=["https://www.googleapis.com/auth/spreadsheets"])
    _retry_patch()
    return build("sheets", "v4", credentials=creds, cache_discovery=False)


def _col_letter(i: int) -> str:
    i += 1
    out = ""
    while i:
        i, r = divmod(i - 1, 26)
        out = chr(65 + r) + out
    return out



# ----------------------------------------------------------------------------
# VID RAW (SC): copy tab Template -> isi per grup brand
# ----------------------------------------------------------------------------
VIDRAW_LINK = "https://docs.google.com/spreadsheets/d/1iO3NHHND-9skchWespfyRyDECYrvIXuiMq9M-AD6dBI/edit?gid=1244677668#gid=1244677668"
VIDRAW_GROUPS = {"VSL": "Vaseline", "UHBS": "Clear, Dove, Sunsilk, Tresemme", "PONDS": "Ponds", "GAL": "GAL"}
# header tab (dinormalisasi) -> kolom di PUSH_COLUMNS (None = dikosongkan)
VIDRAW_MAP = {
    "brand": "Brand", "videoname": "Video title", "videolink": "Video link", "idvid": "Video ID",
    "videopostdate": "Post date", "creatorusername": "Creator name", "creatorcheck": None,
    "gmvcreatorpage": "Creator video-attributed GMV", "order": "Video-attributed orders",
    "orders": "Video-attributed orders", "videoattributedorders": "Video-attributed orders",
    "aov": "Video-attributed orders", "itemsold": "Video-attributed items sold", "views": "Video views",
    "ctr": "CTR", "comment": "Comments", "comments": "Comments", "likes": "Likes",
}
VIDRAW_NUM = {"Creator video-attributed GMV", "Video-attributed orders", "Video-attributed items sold",
              "Video views", "CTR", "Comments", "Likes"}
_MON = ["JAN", "FEB", "MAR", "APR", "MEI", "JUN", "JUL", "AGU", "SEP", "OKT", "NOV", "DES"]


def vidraw_range_label(rs, re_) -> str:
    if rs.year == re_.year and rs.month == re_.month:
        return f"{rs.day}-{re_.day}"
    return f"{rs.day} {_MON[rs.month - 1]}-{re_.day} {_MON[re_.month - 1]}"


def _nk2(v) -> str:
    return re.sub(r"[^a-z0-9]", "", str(v or "").lower())


def push_vidraw(push_rows, link, groups: dict, label: str, dry_run: bool) -> list[dict]:
    """push_rows[0] = header PUSH_COLUMNS. groups = {"UHBS": ["clear", ...]}.
    Tiap grup: copy tab 'Template <grup>' -> '<GRUP> <label>' lalu isi mulai baris 4.
    Kalau tab hasilnya sudah ada: tambah di bawah. Semua baris dimasukkan (duplikat tidak dicek)."""
    m = re.search(r"/spreadsheets/d/([a-zA-Z0-9_-]+)", link or "")
    if not m:
        raise ValueError("Link bukan link Google Sheets.")
    sid = m.group(1)
    svc = sheets_service()
    props = [x["properties"] for x in svc.spreadsheets().get(
        spreadsheetId=sid, fields="sheets(properties(sheetId,title,gridProperties(rowCount)))").execute().get("sheets", [])]
    by_title = {p["title"].strip().lower(): p for p in props}

    hdr, body = push_rows[0], push_rows[1:]
    hi = {h: i for i, h in enumerate(hdr)}
    assigned = set()
    out = []
    for grp, brands in groups.items():
        bset = {b.strip().lower() for b in brands if b.strip()}
        mine = [r for r in body if str(r[hi["Brand"]]).strip().lower() in bset]
        assigned |= bset
        title = f"{grp} {label}".strip()
        info = {"group": grp, "title": title, "rows": len(mine), "added": 0, "skipped": 0, "new_tab": False}
        tpl = next((p for p in props if p["title"].lower().startswith("template") and grp.lower() in p["title"].lower()), None)
        exist = by_title.get(title.lower())
        if exist is None and tpl is None:
            info["error"] = f"tab template untuk '{grp}' tidak ditemukan"
            out.append(info)
            continue

        def cell(r, push_col):
            if push_col is None:
                return ""
            v = r[hi[push_col]]
            if push_col in VIDRAW_NUM:
                if v in ("", None):
                    return ""
                try:
                    f = float(str(v).replace(",", ""))
                    return int(f) if f.is_integer() and push_col != "CTR" else f
                except ValueError:
                    return v
            return "" if v is None else str(v)

        if exist is None:
            info["new_tab"] = True
            if dry_run:
                info["added"] = len(mine)
                out.append(info)
                continue
            new = svc.spreadsheets().sheets().copyTo(
                spreadsheetId=sid, sheetId=tpl["sheetId"], body={"destinationSpreadsheetId": sid}).execute()
            svc.spreadsheets().batchUpdate(spreadsheetId=sid, body={"requests": [{"updateSheetProperties": {
                "properties": {"sheetId": new["sheetId"], "title": title}, "fields": "title"}}]}).execute()
            sheet_id, row_count = new["sheetId"], tpl["gridProperties"]["rowCount"]
        else:
            title, sheet_id, row_count = exist["title"], exist["sheetId"], exist["gridProperties"]["rowCount"]
            info["title"] = title

        head = svc.spreadsheets().values().get(spreadsheetId=sid, range=f"'{title}'!3:3").execute().get("values", [[]])[0]
        cols = [VIDRAW_MAP.get(_nk2(h)) if _nk2(h) in VIDRAW_MAP else None for h in head]
        nh = [_nk2(h) for h in head]
        if "idvid" not in nh or "brand" not in nh:
            raise ValueError(f"Tab '{title}' baris 3 harus berisi header 'ID Vid' dan 'Brand'.")
        c_id, c_br = nh.index("idvid"), nh.index("brand")

        start = 4
        keep = mine
        if exist is not None:
            got = svc.spreadsheets().values().batchGet(
                spreadsheetId=sid, ranges=[f"'{title}'!{_col_letter(c_id)}4:{_col_letter(c_id)}",
                                           f"'{title}'!{_col_letter(c_br)}4:{_col_letter(c_br)}"],
                valueRenderOption="FORMATTED_VALUE").execute().get("valueRanges", [])
            ids = [r[0] if r else "" for r in got[0].get("values", [])]
            last = max((i for i, v in enumerate(ids) if str(v).strip()), default=-1)
            start = 4 + last + 1
            # semua baris ditambahkan; duplikat dibiarkan masuk
        info["added"] = len(keep)
        if dry_run or not keep:
            out.append(info)
            continue
        values = [[cell(r, c) for c in cols] for r in keep]
        need = start + len(values) - 1
        if need > row_count:
            svc.spreadsheets().batchUpdate(spreadsheetId=sid, body={"requests": [{"appendDimension": {
                "sheetId": sheet_id, "dimension": "ROWS", "length": need - row_count}}]}).execute()
        for i in range(0, len(values), 2000):
            svc.spreadsheets().values().update(
                spreadsheetId=sid, range=f"'{title}'!A{start + i}", valueInputOption="RAW",
                body={"values": values[i:i + 2000]}).execute()
        ci = next((i for i, c in enumerate(cols) if c == "CTR"), None)
        if ci is not None:
            svc.spreadsheets().batchUpdate(spreadsheetId=sid, body={"requests": [{"repeatCell": {
                "range": {"sheetId": sheet_id, "startRowIndex": start - 1, "endRowIndex": start - 1 + len(values),
                          "startColumnIndex": ci, "endColumnIndex": ci + 1},
                "cell": {"userEnteredFormat": {"numberFormat": {"type": "PERCENT", "pattern": "0%"}}},
                "fields": "userEnteredFormat.numberFormat"}}]}).execute()
        out.append(info)
    other = sorted({str(r[hi["Brand"]]).strip() for r in body if str(r[hi["Brand"]]).strip().lower() not in assigned})
    if out:
        out[0]["unassigned"] = other
    return out



# ----------------------------------------------------------------------------
# Pipeline SC BW (versi ringkas: hanya hasil BRAND UNIQ NOT IN TAP - RANGE)
# ----------------------------------------------------------------------------
def _first_pid_only(eng, header, rows):
    """Tanpa explode: Product ID di tiap baris SC dipangkas jadi ID pertama saja."""
    h = eng.resolve_header(header, "Product ID")
    for r in rows:
        ids = eng.split_product_ids(r.get(h, ""))
        r[h] = ids[0] if ids else ""
    return rows


def run_sc(sc, taps, affs, bw, metric, rng, explode: bool = True):
    ms, me = metric
    rs, re_ = rng
    if SKU_COL in engine.OUTPUT_COLUMNS:  # bersihkan sisa run sebelumnya
        engine.OUTPUT_COLUMNS.remove(SKU_COL)

    print("Membaca file...")
    sc_header, sc_rows = load_table(sc, engine.SC_REQUIRED, "SC VIDEO")
    print(f"  SC VIDEO ({sc[0]}): {len(sc_rows):,} baris")
    if not explode:
        sc_rows = _first_pid_only(engine, sc_header, sc_rows)
        print("  Tanpa explode: hanya Product ID pertama tiap baris SC yang dipakai.")
    tap_datasets = []
    for i, s_ in enumerate(taps, 1):
        h, r = load_table(s_, engine.TAP_REQUIRED, f"TAP {s_[0]}")
        print(f"  TAP {i} ({s_[0]}): {len(r):,} baris")
        tap_datasets.append((s_[0], h, r))
    aff_datasets = []
    for i, s_ in enumerate(affs, 1):
        h, r = load_table(s_, engine.AFF_REQUIRED, f"AFF {s_[0]}")
        print(f"  AFF {i} ({s_[0]}): {len(r):,} baris")
        aff_datasets.append((s_[0], h, r))
    bw_header, bw_rows = load_table(bw, engine.BW_MAPPING_REQUIRED, "BW PRODUCT MAPPING")
    print(f"  BW MAPPING ({bw[0]}): {len(bw_rows):,} baris")

    bw_map, bw_stats = engine.build_bw_product_mapping(bw_header, bw_rows)
    engine.configure_brand_output_order(bw_map)

    # sku name = kolom Product Name di mapping BW (per Product ID)
    sku_lookup: dict[str, str] = {}
    sku_h = engine.resolve_first_header(bw_header, ["Product Name", "Product name", "Nama Produk", "SKU name", "SKU"])
    if sku_h is None:
        raise ValueError("Kolom 'Product Name' tidak ada di mapping BW, sku name tidak bisa diisi.")
    pid_h = engine.resolve_header(bw_header, "Product ID")
    for r in bw_rows:
        pid = engine.normalize_id(r.get(pid_h, ""))
        name = str(r.get(sku_h, "") or "").strip()
        if pid and name and pid not in sku_lookup:
            sku_lookup[pid] = name
    print(f"  SKU name ('{sku_h}'): {len(sku_lookup):,} Product ID")

    print("Membangun TAP + AFF mapping...")
    tap_map, _ = engine.build_tap_mapping(tap_datasets, ms, me)
    aff_map, _ = engine.build_aff_mapping(aff_datasets, ms, me)
    master_map, _ = engine.build_master_mapping(tap_map, aff_map)
    pid_names = engine.build_pid_product_name_lookup(tap_map, aff_map)
    tap_vb_pairs = engine.build_tap_video_brand_pair_set_raw(tap_datasets, bw_map)

    print("Explode + enrich SC Video...")
    (all_rows, sc_stats, _o, _e, _b, _m) = engine.enrich_sc(
        sc_header, sc_rows, master_map, aff_map, pid_names, bw_map)
    for row in all_rows:
        row[SKU_COL] = sku_lookup.get(engine.normalize_id(row.get("Product ID", "")), "")
    engine.OUTPUT_COLUMNS.insert(engine.OUTPUT_COLUMNS.index("Product ID") + 1, SKU_COL)

    range_rows, invalid_range = engine.filter_output_range(all_rows, rs, re_)
    nit_range = engine.filter_sc_not_in_tap(range_rows, tap_vb_pairs)
    bu_nit_range = engine.build_brand_unique_rows(nit_range)
    print(f"Baris explode: {len(all_rows):,} | di RANGE: {len(range_rows):,} | tidak ada di TAP: {len(nit_range):,} "
          f"| unik Brand+Video (yang dikirim): {len(bu_nit_range):,}")
    if invalid_range:
        print(f"PERINGATAN: {invalid_range:,} baris dengan Post date tidak valid tidak masuk RANGE.")

    # Product ID / sku name yang belum ada di mapping produk (N/A) -> pengiriman diblok
    miss: dict[str, dict] = {}
    for row in bu_nit_range:
        pid = engine.normalize_id(row.get("Product ID", "")) or "<KOSONG>"
        brand = str(row.get("Brand", "") or "").strip().upper()
        sku = str(row.get(SKU_COL, "") or "").strip()
        reason = ""
        if brand in ("", "MISSING"):
            reason = "Product ID tidak ada di mapping (Brand N/A)"
        elif not sku:
            reason = "sku name (Product Name) kosong di mapping"
        if reason:
            d = miss.setdefault(pid, {"Product ID": pid, "Product name": str(row.get("Product name", "") or ""),
                                      "Alasan": reason, "Jumlah video": 0})
            d["Jumlah video"] += 1
    missing_df = pd.DataFrame(sorted(miss.values(), key=lambda d: -d["Jumlah video"]),
                              columns=["Product ID", "Product name", "Alasan", "Jumlah video"])

    return {
        "push": [PUSH_COLUMNS] + [[row.get(c, "") for c in PUSH_COLUMNS] for row in bu_nit_range],
        "missing_df": missing_df,
        "stats": {"explode": len(all_rows), "range": len(range_rows), "not_in_tap": len(nit_range),
                  "unik": len(bu_nit_range)},
    }


def _sc_stats(values, gmv_col: str = "Creator video-attributed GMV") -> dict:
    hdr, rows = values[0], values[1:]
    hi = {h: i for i, h in enumerate(hdr)}

    def num(v):
        try:
            return float(v)
        except (TypeError, ValueError):
            return 0.0

    return {"n": len(rows), "gmv": sum(num(r[hi[gmv_col]]) for r in rows),
            "creator": len({str(r[hi["Creator name"]]).strip().lower() for r in rows if str(r[hi["Creator name"]]).strip()})}


def push_sc_range(values, link, tab_name, dry_run, gmv_col: str = "Creator video-attributed GMV") -> dict:
    """Tambah semua baris ke tab SC VID Range (di bawah data yang ada). Duplikat tidak dicek.
    Kolom disejajarkan dengan header tab; tab kosong -> header ikut ditulis."""
    m = re.search(r"/spreadsheets/d/([a-zA-Z0-9_-]+)", link or "")
    if not m:
        raise ValueError("Link target bukan link Google Sheets.")
    sid = m.group(1)
    tab_name = (tab_name or "").strip()
    if not tab_name:
        raise ValueError("Nama tab SC VID Range kosong.")
    svc = sheets_service()
    props = [x["properties"] for x in svc.spreadsheets().get(
        spreadsheetId=sid, fields="sheets(properties(sheetId,title))").execute().get("sheets", [])]
    target = next((p for p in props if p["title"].strip().lower() == tab_name.lower()), None)
    if target is None and not dry_run:
        svc.spreadsheets().batchUpdate(spreadsheetId=sid, body={"requests": [
            {"addSheet": {"properties": {"title": tab_name}}}]}).execute()
    title = target["title"] if target else tab_name

    header, rows = values[0], values[1:]
    got = [] if target is None else svc.spreadsheets().values().get(
        spreadsheetId=sid, range=f"'{title}'!1:1").execute().get("values", [])  # error jangan ditelan
    ex_header = got[0] if got else []
    res_ = {"title": title, "added": len(rows), "stat": _sc_stats(values, gmv_col), "header_written": False}
    if not any(str(c).strip() for c in ex_header):
        res_["header_written"] = True
        if not dry_run:
            allrows = [header] + rows
            _ensure_grid(svc, sid, title, len(allrows), len(header))
            for i in range(0, len(allrows), 2000):
                svc.spreadsheets().values().update(
                    spreadsheetId=sid, range=f"'{title}'!A{i + 1}", valueInputOption="RAW",
                    body={"values": allrows[i:i + 2000]}).execute()
        return res_
    ours = {engine.normalize_header(h): i for i, h in enumerate(header)}
    idx = [ours.get(engine.normalize_header(h)) for h in ex_header]
    aligned = [[(r[i] if i is not None else "") for i in idx] for r in rows]
    if not dry_run:
        used = svc.spreadsheets().values().get(spreadsheetId=sid, range=f"'{title}'!A:A").execute().get("values", [])
        _ensure_grid(svc, sid, title, len(used) + len(aligned) + 5, len(ex_header))
        for i in range(0, len(aligned), 2000):
            svc.spreadsheets().values().append(
                spreadsheetId=sid, range=f"'{title}'!A1", valueInputOption="RAW",
                insertDataOption="INSERT_ROWS", body={"values": aligned[i:i + 2000]}).execute()
    return res_


# ----------------------------------------------------------------------------
# UI
# ----------------------------------------------------------------------------
def _d(s):
    try:
        return date.fromisoformat(s) if s else None
    except ValueError:
        return None


def _pair(val):
    if isinstance(val, (list, tuple)):
        if len(val) == 2:
            return val[0], val[1]
        # 1 tanggal saja = baru pilih awal, akhir belum dipilih -> belum valid (jangan diam-diam jadi rentang 1 hari)
    return None, None


# ----------------------------------------------------------------------------
# SC BW LIVE (Livestream): SC LS di-explode per Product ID + TAP LS + AFF
# ----------------------------------------------------------------------------
# nama kolom engine -> nama alternatif di export SC Livestream (export lama tanpa awalan "Avg. daily")
LIVE_ALIASES = {
    "Avg. daily GPM": ["Show GPM"],
    "Avg. daily products sold": ["Products sold"],
    "Avg. daily engagement": ["Engagement"],
    "Avg. daily LIVE product CTR": ["CTR", "LIVE product CTR"],
    "Avg. daily unique viewers": ["LIVE product viewers"],
    "Avg. daily tap-through rate": ["Tap-through rate"],
}
LIVE_SETTING_LINK = "ls_target_link"


def find_live_engine_scripts() -> list[Path]:
    found = []
    for p in sorted(BASE_DIR.glob("*.py")):
        if p.resolve() == Path(__file__).resolve():
            continue
        try:
            text = p.read_text(encoding="utf-8", errors="ignore")
        except OSError:
            continue
        if _is_streamlit_app(text):
            continue
        if (re.search(r"^def enrich_sc_live\(", text, re.M)
                and re.search(r"^def build_brand_unique_rows\(", text, re.M)):
            found.append(p)
    return found


@st.cache_resource
def load_live_module(path: str):
    spec = importlib.util.spec_from_file_location("sc_bw_live_engine", path)
    mod = importlib.util.module_from_spec(spec)
    sys.modules["sc_bw_live_engine"] = mod
    spec.loader.exec_module(mod)
    return mod


def _patch_live_engine(ls) -> None:
    if getattr(ls, "_app_patched", False):
        return
    orig = ls.parse_percent

    def safe(value):
        if isinstance(value, (int, float)) and not isinstance(value, bool):
            return 0.0 if (isinstance(value, float) and math.isnan(value)) else float(value)
        return orig(value)

    ls.parse_percent = safe
    ls._app_patched = True


def _apply_live_aliases(ls, header, rows):
    have = {ls.normalize_header(h): h for h in header if h}
    ren = {}
    for out_name, alts in LIVE_ALIASES.items():
        if ls.normalize_header(out_name) in have:
            continue
        for a in alts:
            src = have.get(ls.normalize_header(a))
            if src:
                ren[src] = out_name
                break
    if not ren:
        return header, rows, {}
    return [ren.get(h, h) for h in header], [{ren.get(k, k): v for k, v in r.items()} for r in rows], ren


def run_sc_live(ls, sc, taps, affs, bw, metric, rng, explode: bool = True):
    ms, me = metric
    rs, re_ = rng
    _patch_live_engine(ls)

    def tap_ok(header):
        try:
            ls.resolve_tap_layout(header)
            return True
        except (ValueError, KeyError):
            return False

    print("Membaca file...")
    sc_header, sc_rows = load_table(sc, ls.SC_REQUIRED, "SC LIVE")
    sc_header, sc_rows, ren = _apply_live_aliases(ls, sc_header, sc_rows)
    print(f"  SC LIVE ({sc[0]}): {len(sc_rows):,} baris")
    if not explode:
        sc_rows = _first_pid_only(ls, sc_header, sc_rows)
        print("  Tanpa explode: hanya Product ID pertama tiap baris SC yang dipakai.")
    if ren:
        print("  Nama kolom SC disesuaikan: " + ", ".join(f"'{a}' → '{b}'" for a, b in ren.items()))
    tap_datasets = []
    for i, s_ in enumerate(taps, 1):
        h, r = load_table(s_, tap_ok, f"TAP LS {s_[0]}")
        print(f"  TAP LS {i} ({s_[0]}): {len(r):,} baris")
        tap_datasets.append((s_[0], h, r))
    aff_datasets = []
    for i, s_ in enumerate(affs, 1):
        h, r = load_table(s_, ls.AFF_REQUIRED, f"AFF {s_[0]}")
        print(f"  AFF {i} ({s_[0]}): {len(r):,} baris")
        aff_datasets.append((s_[0], h, r))
    bw_header, bw_rows = load_table(bw, ls.BW_MAPPING_REQUIRED, "BW PRODUCT MAPPING")
    print(f"  BW MAPPING ({bw[0]}): {len(bw_rows):,} baris")

    bw_map, _ = ls.build_bw_product_mapping(bw_header, bw_rows)
    ls.configure_brand_output_order(bw_map)

    sku_lookup: dict[str, str] = {}
    sku_h = ls.resolve_first_header(bw_header, ["Product Name", "Product name", "Nama Produk", "SKU name", "SKU"])
    if sku_h is None:
        raise ValueError("Kolom 'Product Name' tidak ada di mapping BW, sku name tidak bisa diisi.")
    pid_h = ls.resolve_header(bw_header, "Product ID")
    for r in bw_rows:
        pid = ls.normalize_id(r.get(pid_h, ""))
        name = str(r.get(sku_h, "") or "").strip()
        if pid and name and pid not in sku_lookup:
            sku_lookup[pid] = name
    print(f"  SKU name ('{sku_h}'): {len(sku_lookup):,} Product ID")

    print("Membangun TAP LS + AFF mapping...")
    tap_map, tap_stats = ls.build_tap_ls_mapping(tap_datasets, ms, me)
    aff_map, _ = ls.build_aff_ls_mapping(aff_datasets, ms, me)
    master_map, _ = ls.build_master_mapping(tap_map, aff_map)
    pid_names = ls.build_pid_product_name_lookup(tap_map, aff_map)
    pairs = ls.build_tap_brand_pair_set_from_mapping(tap_map, bw_map)
    print(f"  TAP LS: {tap_stats['raw_rows']:,} baris → {tap_stats['pairs']:,} pasangan LIVE ID+Product ID "
          f"(di luar METRIC: {tap_stats['period_skipped']:,}, duplikat: {tap_stats['duplicate_rows']:,})")

    print("Explode + enrich SC Live...")
    (all_rows, sc_stats, _o, _e, _b, _m) = ls.enrich_sc_live(
        sc_header, sc_rows, master_map, aff_map, pid_names, bw_map)
    for row in all_rows:
        row[SKU_COL] = sku_lookup.get(ls.normalize_id(row.get("Product ID", "")), "")
    cols = list(ls.OUTPUT_COLUMNS)
    if SKU_COL not in cols:
        cols.insert(cols.index("Product ID") + 1, SKU_COL)

    range_rows, invalid_range = ls.filter_output_range(all_rows, rs, re_)
    nit_range = ls.filter_sc_ls_not_in_tap(range_rows, pairs)
    bu_nit_range = ls.build_brand_unique_rows(nit_range)
    print(f"Baris SC: {sc_stats['raw_rows'] - sc_stats.get('invalid_live_id', 0):,} | explode: {len(all_rows):,} | di RANGE: {len(range_rows):,} "
          f"| tidak ada di TAP: {len(nit_range):,} | unik LIVE ID+Brand (yang dikirim): {len(bu_nit_range):,}")
    if sc_stats.get("invalid_live_id"):
        print(f"  Baris tanpa LIVE ID dilewati: {sc_stats['invalid_live_id']:,} (mis. baris keterangan metrik)")
    if invalid_range:
        print(f"PERINGATAN: {invalid_range:,} baris dengan LIVE start time tidak valid tidak masuk RANGE.")

    miss: dict[str, dict] = {}
    for row in bu_nit_range:
        pid = ls.normalize_id(row.get("Product ID", "")) or "<KOSONG>"
        brand = str(row.get("Brand", "") or "").strip().upper()
        sku = str(row.get(SKU_COL, "") or "").strip()
        reason = ""
        if brand in ("", "MISSING"):
            reason = "Product ID tidak ada di mapping (Brand N/A)"
        elif not sku:
            reason = "sku name (Product Name) kosong di mapping"
        if reason:
            d = miss.setdefault(pid, {"Product ID": pid, "Product name": str(row.get("Product name", "") or ""),
                                      "Alasan": reason, "Jumlah LS": 0})
            d["Jumlah LS"] += 1
    missing_df = pd.DataFrame(sorted(miss.values(), key=lambda d: -d["Jumlah LS"]),
                              columns=["Product ID", "Product name", "Alasan", "Jumlah LS"])

    push = [cols] + [[row.get(c, "") for c in cols] for row in bu_nit_range]
    df = pd.DataFrame(push[1:], columns=push[0])
    for c_ in ("LIVE ID", "Product ID"):
        if c_ in df:
            df[c_] = df[c_].astype(str)
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as xw:
        df.to_excel(xw, index=False, sheet_name="SC LS RANGE")
    return {
        "push": push, "missing_df": missing_df, "xlsx": bio.getvalue(),
        "stats": {"sc": sc_stats["raw_rows"] - sc_stats.get("invalid_live_id", 0), "explode": len(all_rows), "range": len(range_rows),
                  "not_in_tap": len(nit_range), "unik": len(bu_nit_range)},
    }


def render_sc_live_page():
    st.title("📡 SC BW Live")
    st.caption("SC Livestream di-explode per Product ID → diperkaya TAP LS + AFF → hasil **BRAND UNIQUE NOT IN TAP - RANGE** "
               "(1 baris = LIVE ID + Brand). Bisa di-download dan/atau dikirim ke Google Sheets.")
    with st.expander("Lanjutan (engine SC Live)"):
        ov = st.text_input("Path script engine SC Live (opsional)", value="", key="lv_engine_override")
    if ov.strip():
        script = Path(ov.strip().strip('"'))
        if not script.exists():
            st.error(f"Script engine tidak ada: {script}")
            st.stop()
        if _is_streamlit_app(script.read_text(encoding="utf-8", errors="ignore")):
            st.error(f"`{script.name}` adalah aplikasi Streamlit, bukan script engine. Pilih `sc_bw_live_ebu.py`.")
            st.stop()
    else:
        cands = find_live_engine_scripts()
        if not cands:
            st.error(f"Script engine SC Live tidak ditemukan di folder `{BASE_DIR}`.\n\n"
                     "Taruh `sc_bw_live_ebu.py` (yang berisi fungsi `enrich_sc_live`) di folder yang sama dengan "
                     "`app_konten.py`, atau isi path-nya di Lanjutan.")
            st.stop()
        script = cands[0] if len(cands) == 1 else Path(st.selectbox("Pilih script engine SC Live", [str(p) for p in cands]))
    ls = load_live_module(str(script))
    st.caption(f"Engine: `{script.name}`")

    st.subheader("1 · File input")
    c1, c2 = st.columns(2)
    with c1:
        sc_up = st.file_uploader("RAW SC Livestream", type=["xlsx", "xlsm", "xls", "csv", "txt"], key="lv_sc_up")
        sc_path = st.text_input("atau path lokal SC (1 file)", key="lv_sc_path", placeholder=r"C:\Users\...\SC_LS_PONDS.xlsx")
        bw_link = st.text_input(
            "BW Product Mapping (link Google Sheets / path file)", value=st.session_state.get("bw_link_shared", ""),
            key="lv_bw_link", help="Link yang sama dengan SC Video; tersimpan.")
        st.session_state["bw_link_shared"] = bw_link
        _save_setting("bw_link", bw_link.strip())
    with c2:
        tap_up = st.file_uploader("RAW TAP Livestream (boleh banyak)", type=["xlsx", "xlsm", "xls", "csv", "txt"],
                                  accept_multiple_files=True, key="lv_tap_up")
        tap_path = st.text_area("atau path lokal TAP LS (satu per baris; boleh folder)", key="lv_tap_path", height=70)
        aff_up = st.file_uploader("RAW Affiliate Order (boleh banyak)", type=["xlsx", "xlsm", "xls", "csv", "txt"],
                                  accept_multiple_files=True, key="lv_aff_up")
        aff_path = st.text_area("atau path lokal AFF (satu per baris; boleh folder)", key="lv_aff_path", height=70)
    _sz = sum(getattr(u_, "size", 0) for u_ in ([sc_up] if sc_up else []) + list(tap_up or []) + list(aff_up or []))
    if _sz > 100 * 1024 * 1024:
        st.warning(f"Total file upload {_sz / 1024 / 1024:,.0f} MB. Untuk file besar, KOSONGKAN upload dan pakai **path lokal**.")
    st.info("File SC boleh memuat baris keterangan metrik di bawah header (otomatis dilewati). "
            "ID (LIVE ID / Product ID) sebaiknya berformat **Text**.")

    st.subheader("2 · Periode")
    today = date.today()
    r0 = r1 = m0 = m1 = None  # tidak memakai tanggal bawaan di script engine; atur langsung di sini
    with st.form("lv_periode_form", border=False):
        range_val = st.date_input("RANGE — rentang LIVE start time SC", value=(r0 or today - timedelta(days=6), r1 or today),
                                  key="lv_range_dates")
        metric_diff = st.checkbox("METRIC (filter TAP LIVE start time & AFF Time Created) pakai rentang berbeda dari RANGE",
                                  value=bool(m0 and (m0, m1) != (r0, r1)), key="lv_metric_diff")
        metric_val = st.date_input("Rentang METRIC (hanya dipakai kalau dicentang)",
                                   value=(m0 or today - timedelta(days=6), m1 or today), key="lv_metric_dates")
        st.form_submit_button("✅ Terapkan periode")
        apply_run = st.form_submit_button("▶️ Terapkan periode & Jalankan", type="primary")
    rs, re_ = _pair(range_val)
    ms, me = _pair(metric_val) if metric_diff else (rs, re_)
    if rs and re_:
        st.caption(f"✅ Yang dipakai: RANGE {rs} → {re_} · METRIC {ms} → {me}  (tanggal yang diubah baru berlaku setelah "
               "menekan *Terapkan periode* atau *Terapkan periode & Jalankan*)")
    else:
        st.warning("Pilih tanggal awal DAN akhir RANGE, lalu tekan Terapkan periode.")

    explode_lv = st.checkbox("Explode per Product ID (1 baris SC → 1 baris per Product ID)", value=True, key="lv_explode",
                             help="Dimatikan: hanya Product ID pertama tiap baris SC yang dipakai; metrik SC dipakai apa adanya.")
    st.subheader("3 · Tujuan Google Sheets (opsional)")
    t1, t2 = st.columns(2)
    with t1:
        target_link = st.text_input("Link Google Sheets", value=_load_settings().get(LIVE_SETTING_LINK, ""),
                                    key="lv_target_link", placeholder="kosong = hanya download .xlsx")
        _save_setting(LIVE_SETTING_LINK, target_link.strip())
    with t2:
        target_tab = st.text_input("Nama tab", value="SC LS Range", key="lv_target_tab")
    sa_email = service_account_email()
    if target_link.strip():
        if sa_email:
            st.info(f"Share sheet tujuan sebagai **Editor** ke: `{sa_email}`")
        else:
            st.warning(f"credentials.json tidak ditemukan/terbaca di `{cred_file()}`.")

    st.subheader("4 · Jalankan")
    if (st.button("▶️ Jalankan", disabled=not (rs and re_), key="lv_run") or apply_run) and (rs and re_):
        st.session_state.pop("lv_result", None)
        placeholder = st.empty()
        writer = LiveWriter(placeholder)
        try:
            sc_items = collect([sc_up] if sc_up else [], sc_path)
            bw_src = get_bw_source(bw_link)
            tap_items = collect(tap_up, tap_path)
            aff_items = collect(aff_up, aff_path)
            problems = []
            if len(sc_items) != 1:
                problems.append(f"RAW SC harus tepat 1 file (sekarang {len(sc_items)}).")
            if not bw_src:
                problems.append("BW Product Mapping belum diisi.")
            if not tap_items:
                problems.append("RAW TAP Livestream belum diisi.")
            if not aff_items:
                problems.append("RAW Affiliate Order belum diisi.")
            if problems:
                for p_ in problems:
                    st.error(p_)
            else:
                with st.spinner("Memproses… (data besar bisa beberapa menit)"):
                    with contextlib.redirect_stdout(writer):
                        out = run_sc_live(ls, sc_items[0], tap_items, aff_items, bw_src, (ms, me), (rs, re_), explode_lv)
                out["log"] = writer.buf
                out["rng"] = (rs, re_)
                st.session_state["lv_result"] = out
        except SystemExit as e:
            st.error(f"Berhenti: {e.code}")
        except Exception as e:  # noqa: BLE001
            st.error(f"Error: {type(e).__name__}: {e}")
            with st.expander("Detail error"):
                st.code(traceback.format_exc())

    res = st.session_state.get("lv_result")
    if not res:
        return
    gcol = "Creator LIVE-attributed GMV"
    n_rows = len(res["push"]) - 1
    stt = res["stats"]
    miss_df = res["missing_df"]
    st.success(f"Selesai — {n_rows:,} baris (unik LIVE ID + Brand, tidak ada di TAP, RANGE "
               f"{res['rng'][0]} → {res['rng'][1]}).")
    tot = _sc_stats(res["push"], gcol) if n_rows else {"n": 0, "gmv": 0, "creator": 0}
    k1, k2, k3 = st.columns(3)
    k1.metric("Total LS", f"{tot['n']:,}")
    k2.metric("Total GMV", _rp(tot["gmv"]))
    k3.metric("Creator unik", f"{tot['creator']:,}")
    st.dataframe(pd.DataFrame([
        {"Tahap": "Baris SC (yang punya LIVE ID)", "Jumlah": stt["sc"]},
        {"Tahap": "Setelah explode per Product ID", "Jumlah": stt["explode"]},
        {"Tahap": "Di RANGE", "Jumlah": stt["range"]},
        {"Tahap": "Tidak ada di TAP (LIVE ID + Brand)", "Jumlah": stt["not_in_tap"]},
        {"Tahap": "Unik LIVE ID + Brand (hasil)", "Jumlah": stt["unik"]},
    ]), hide_index=True, use_container_width=True)
    with st.expander("Log proses"):
        st.code(res["log"], language="text")
    with st.expander("Preview (10 baris pertama)"):
        st.dataframe(pd.DataFrame(res["push"][1:11], columns=res["push"][0]), hide_index=True, use_container_width=True)
    st.download_button("⬇️ Download hasil (.xlsx)", res["xlsx"], file_name="SC_LS_BRAND_UNIQUE_NOT_IN_TAP_RANGE.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="lv_dl")

    blocked = len(miss_df) > 0
    if blocked:
        st.error(f"⛔ {len(miss_df):,} Product ID belum ada / sku name kosong di mapping produk (N/A). "
                 "Pengiriman ke Google Sheets DIBLOK. Lengkapi mapping BW lalu jalankan ulang.")
        st.dataframe(miss_df, hide_index=True, use_container_width=True)
        st.download_button("⬇️ Download daftar yang belum ada di mapping (.csv)",
                           miss_df.to_csv(index=False).encode("utf-8-sig"),
                           file_name="product_id_belum_ada_di_mapping_ls.csv", mime="text/csv", key="lv_dl_miss")

    def send(dry: bool):
        try:
            with st.spinner("Mengirim…"):
                o = push_sc_range(res["push"], target_link, target_tab, dry, gmv_col=gcol)
            a = o["stat"]
            st.success(("Cek (belum menulis): " if dry else "Terkirim. ") +
                       f"Tab '{o['title']}': {a['n']:,} baris, {_rp(a['gmv'])}, {a['creator']:,} creator"
                       + (" (tab kosong: header ikut ditulis)" if o["header_written"] else "") + ".")
        except Exception as e:  # noqa: BLE001
            st.error(f"{type(e).__name__}: {e}")

    st.subheader("5 · Kirim ke Google Sheets")
    can_send = (not blocked) and bool(sa_email) and n_rows > 0 and bool(target_link.strip()) and bool(target_tab.strip())
    if not target_link.strip():
        st.caption("Isi link Google Sheets di bagian 3 untuk mengaktifkan pengiriman.")
    b1, b2 = st.columns(2)
    if b1.button("🔍 Cek dulu (tidak menulis)", disabled=not can_send, key="lv_dry"):
        send(True)
    if b2.button("📤 Kirim", type="primary", disabled=not can_send, key="lv_go"):
        send(False)


if PAGE.startswith("📡"):
    render_sc_live_page()
    st.stop()


st.title("📊 SC BW Video")
st.caption("SC + TAP + AFF → hasil **BRAND UNIQ NOT IN TAP - RANGE** dikirim ke tab SC VID Range. "
           "Tidak ada output lain.")

st.subheader("1 · File input")
c1, c2 = st.columns(2)
with c1:
    sc_up = st.file_uploader("RAW SC Video", type=["xlsx", "xlsm", "xls", "csv", "txt"], key="sc_up")
    sc_path = st.text_input("atau path lokal SC (1 file)", key="sc_path", placeholder=r"C:\Users\...\SC Video.xlsx")
    bw_link = st.text_input(
        "BW Product Mapping (link Google Sheets / path file)", value=st.session_state.get("bw_link_shared", ""),
        key="bw_link", placeholder="https://docs.google.com/spreadsheets/d/.../edit?gid=0",
        help="Sheet harus 'Anyone with the link' (Viewer). Link tersimpan, tidak perlu diisi ulang.")
    st.session_state["bw_link_shared"] = bw_link
    _save_setting("bw_link", bw_link.strip())
with c2:
    tap_up = st.file_uploader("RAW TAP Short Video (boleh banyak)", type=["xlsx", "xlsm", "xls", "csv", "txt"],
                              accept_multiple_files=True, key="tap_up")
    tap_path = st.text_area("atau path lokal TAP (satu per baris; boleh folder)", key="tap_path", height=70)
    aff_up = st.file_uploader("RAW Affiliate Order (boleh banyak)", type=["xlsx", "xlsm", "xls", "csv", "txt"],
                              accept_multiple_files=True, key="aff_up")
    aff_path = st.text_area("atau path lokal AFF (satu per baris; boleh folder)", key="aff_path", height=70)
_sz = sum(getattr(u_, "size", 0) for u_ in ([sc_up] if sc_up else []) + list(tap_up or []) + list(aff_up or []))
if _sz > 100 * 1024 * 1024:
    st.warning(f"Total file upload {_sz / 1024 / 1024:,.0f} MB. Upload disimpan di memori dan disalin tiap halaman "
               "dimuat ulang, jadi bisa MemoryError. Untuk file besar, KOSONGKAN upload dan pakai **path lokal**.")
st.info("ID (Video ID / Product ID) di file sumber sebaiknya berformat **Text** agar 19 digitnya tidak terpotong.")

st.subheader("2 · Periode")
today = date.today()
r0 = r1 = m0 = m1 = None  # tidak memakai tanggal bawaan di script engine (mis. 2026-10-18); atur langsung di sini
with st.form("periode_form", border=False):
    range_val = st.date_input("RANGE — rentang Post date SC", value=(r0 or today - timedelta(days=6), r1 or today),
                              key="range_dates")
    metric_diff = st.checkbox("METRIC (filter TAP Date & AFF Time Created) pakai rentang berbeda dari RANGE",
                              value=bool(m0 and (m0, m1) != (r0, r1)))
    metric_val = st.date_input("Rentang METRIC (hanya dipakai kalau dicentang)",
                               value=(m0 or today - timedelta(days=6), m1 or today), key="metric_dates")
    st.form_submit_button("✅ Terapkan periode")
    apply_run = st.form_submit_button("▶️ Terapkan periode & Jalankan", type="primary")
rs, re_ = _pair(range_val)
ms, me = _pair(metric_val) if metric_diff else (rs, re_)
if rs and re_:
    st.caption(f"✅ Yang dipakai: RANGE {rs} → {re_} · METRIC {ms} → {me}  (tanggal yang diubah baru berlaku setelah "
               "menekan *Terapkan periode* atau *Terapkan periode & Jalankan*)")
else:
    st.warning("Pilih tanggal awal DAN akhir RANGE, lalu tekan Terapkan periode.")

explode_vid = st.checkbox("Explode per Product ID (1 baris SC → 1 baris per Product ID)", value=True, key="sc_explode",
                          help="Dimatikan: hanya Product ID pertama tiap baris SC yang dipakai; metrik SC dipakai apa adanya.")
st.subheader("3 · Tujuan")
st.markdown("**SC VID Range**")
target_link = st.text_input("Link Google Sheets", value=DEFAULT_TARGET_LINK, key="target_link")
sc_tab = st.text_input("Nama tab", value="SC VID Range", key="tab_range")

sa_email = service_account_email()
if sa_email:
    st.info(f"Share sheet tujuan sebagai **Editor** ke: `{sa_email}`")
else:
    st.warning(f"credentials.json tidak ditemukan/terbaca di `{cred_file()}`. Taruh di folder app atau isi path-nya di sidebar → Lanjutan.")

auto_push = st.checkbox("Langsung kirim ke SC VID Range setelah proses selesai", value=False, key="auto_push",
                        help="Tidak jalan kalau masih ada Product ID / sku name yang belum ada di mapping produk.")


def send_all(values, dry: bool):
    """Kirim ke SC VID Range; tampilkan hasil."""
    try:
        with st.spinner("SC VID Range…"):
            o1 = push_sc_range(values, target_link, sc_tab, dry)
        a = o1["stat"]
        st.success(("Cek (belum menulis): " if dry else "Terkirim. ") +
                   f"SC VID Range → tab '{o1['title']}': {a['n']:,} baris, {_rp(a['gmv'])}, {a['creator']:,} creator"
                   + (" (tab kosong: header ikut ditulis)" if o1["header_written"] else "") + ".")
    except Exception as e:  # noqa: BLE001
        st.error(f"{type(e).__name__}: {e}")
        st.caption("Pastikan sheet di-share sebagai **Editor** ke email service account di credentials.json, "
                   "dan `pip install google-api-python-client google-auth`.")


st.subheader("4 · Jalankan")
if (st.button("▶️ Jalankan", disabled=not (rs and re_)) or apply_run) and (rs and re_):
    st.session_state.pop("sc_result", None)
    placeholder = st.empty()
    writer = LiveWriter(placeholder)
    try:
        sc_items = collect([sc_up] if sc_up else [], sc_path)
        bw_src = get_bw_source(bw_link)
        bw_items = [bw_src] if bw_src else []
        tap_items = collect(tap_up, tap_path)
        aff_items = collect(aff_up, aff_path)
        problems = []
        if len(sc_items) != 1:
            problems.append(f"RAW SC harus tepat 1 file (sekarang {len(sc_items)}).")
        if len(bw_items) != 1:
            problems.append("BW Product Mapping belum diisi.")
        if not tap_items:
            problems.append("RAW TAP belum diisi.")
        if not aff_items:
            problems.append("RAW Affiliate Order belum diisi.")
        if problems:
            for p_ in problems:
                st.error(p_)
        else:
            with st.spinner("Memproses… (data besar bisa beberapa menit)"):
                with contextlib.redirect_stdout(writer):
                    out = run_sc(sc_items[0], tap_items, aff_items, bw_items[0], (ms, me), (rs, re_), explode_vid)
            out["log"] = writer.buf
            out["rng"] = (rs, re_)
            st.session_state["sc_result"] = out
            if auto_push:
                if len(out["missing_df"]):
                    st.error("Auto-kirim DIBATALKAN: masih ada Product ID / sku name yang belum ada di mapping (lihat di bawah).")
                elif not sa_email:
                    st.error("Auto-kirim dilewati: credentials.json tidak terbaca.")
                else:
                    send_all(out["push"], False)
    except SystemExit as e:
        st.error(f"Berhenti: {e.code}")
    except Exception as e:  # noqa: BLE001
        st.error(f"Error: {type(e).__name__}: {e}")
        with st.expander("Detail error"):
            st.code(traceback.format_exc())

res_sc = st.session_state.get("sc_result")
if res_sc:
    n_rows = len(res_sc["push"]) - 1
    stt = res_sc["stats"]
    miss_df = res_sc["missing_df"]
    st.success(f"Selesai — {n_rows:,} baris (unik Brand + Video, tidak ada di TAP, RANGE "
               f"{res_sc['rng'][0]} → {res_sc['rng'][1]}).")
    tot = _sc_stats(res_sc["push"]) if n_rows else {"n": 0, "gmv": 0, "creator": 0}
    k1, k2, k3 = st.columns(3)
    k1.metric("Total video", f"{tot['n']:,}")
    k2.metric("Total GMV", _rp(tot["gmv"]))
    k3.metric("Creator unik", f"{tot['creator']:,}")
    with st.expander("Log proses"):
        st.code(res_sc["log"], language="text")
    with st.expander(f"Preview (10 baris pertama)"):
        st.dataframe(pd.DataFrame(res_sc["push"][1:11], columns=res_sc["push"][0]), hide_index=True,
                     use_container_width=True)

    blocked = len(miss_df) > 0
    if blocked:
        st.error(f"⛔ {len(miss_df):,} Product ID belum ada / sku name kosong di mapping produk (N/A). "
                 "Pengiriman DIBLOK. Lengkapi mapping BW lalu jalankan ulang.")
        st.dataframe(miss_df, hide_index=True, use_container_width=True)
        st.download_button("⬇️ Download daftar yang belum ada di mapping (.csv)",
                           miss_df.to_csv(index=False).encode("utf-8-sig"),
                           file_name="product_id_belum_ada_di_mapping.csv", mime="text/csv")
    st.subheader("5 · Kirim")
    can_send = (not blocked) and bool(sa_email) and n_rows > 0
    b1, b2 = st.columns(2)
    if b1.button("🔍 Cek dulu (tidak menulis)", disabled=not can_send, key="sc_dry"):
        send_all(res_sc["push"], True)
    if b2.button("📤 Kirim ke SC VID Range", type="primary", disabled=not can_send, key="sc_go"):
        send_all(res_sc["push"], False)
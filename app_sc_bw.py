"""
Streamlit UI (versi LOKAL / EXCEL) untuk script SC BW Video.
Input : file Excel/CSV (upload atau path lokal)  -> tanpa Google Sheets, tanpa credentials
Output: file Excel (download / simpan ke folder lokal)

Jalankan:  streamlit run app_sc_bw.py
Syarat  :  script engine SC BW (yang berisi fungsi enrich_sc) ada di folder yang sama.
Install :  pip install streamlit pandas openpyxl
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
ID_HEADERS = {"videoid", "productid", "contentid", "orderid"}

st.set_page_config(page_title="SC BW Video (Excel)", page_icon="📊", layout="wide")

# Streamlit Cloud: buat credentials.json dari st.secrets["gcp_service_account"] bila belum ada
sys.path.insert(0, str(BASE_DIR))
from _cloud_secrets import ensure_credentials_file  # noqa: E402
ensure_credentials_file(BASE_DIR)


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
        if "def enrich_sc" in text and "def build_brand_unique_rows" in text:
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
    write_brand_tabs = st.toggle("Buat sheet per Brand", value=False)
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
else:
    candidates = find_engine_scripts()
    if not candidates:
        st.error(
            f"Script engine tidak ditemukan di folder `{BASE_DIR}`.\n\n"
            "Taruh script SC BW (yang berisi fungsi `enrich_sc`) di folder yang sama "
            "dengan `app_sc_bw.py`, atau isi path-nya di sidebar → Lanjutan."
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

    required_norm = {engine.normalize_header(x) for x in required}

    for sheet_name, df in sheets.items():
        for i in range(min(20, len(df))):
            raw = df.iloc[i].tolist()
            header = [("" if _is_blank(v) else str(v).strip()) for v in raw]
            while header and not header[-1]:
                header.pop()
            if not header:
                continue
            if not required_norm.issubset({engine.normalize_header(h) for h in header if h}):
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
    try:
        with urllib.request.urlopen(export, timeout=60) as resp:
            ctype = resp.headers.get("Content-Type", "")
            data = resp.read()
    except urllib.error.HTTPError as e:
        raise RuntimeError(
            f"Gagal mengambil Google Sheets (HTTP {e.code}). Pastikan akses sheet diset "
            "'Anyone with the link' → Viewer."
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
# Output Excel
# ----------------------------------------------------------------------------
def short_name(title: str, used: set[str]) -> str:
    short = {
        engine.SHEET_ALL: "EXPLODED - ALL",
        engine.SHEET_RANGE: "EXPLODED - RANGE",
        engine.SC_NOT_IN_TAP_ALL_SHEET: "EXPL NOT IN TAP - ALL",
        engine.SC_NOT_IN_TAP_RANGE_SHEET: "EXPL NOT IN TAP - RANGE",
        engine.BRAND_UNIQUE_ALL_SHEET: "BRAND UNIQUE - ALL",
        engine.BRAND_UNIQUE_RANGE_SHEET: "BRAND UNIQUE - RANGE",
        engine.BRAND_UNIQUE_NOT_IN_TAP_ALL_SHEET: "BRAND UNIQ NOT IN TAP - ALL",
        engine.BRAND_UNIQUE_NOT_IN_TAP_RANGE_SHEET: "BRAND UNIQ NOT IN TAP - RANGE",
        engine.SUMMARY_SHEET: "BRAND SUMMARY",
        engine.MISSING_SHEET: "MISSING TAP-AFF PAIR",
    }
    if title in short:
        n = short[title]
    elif title.endswith(engine.BRAND_ALL_SUFFIX):
        n = f"{title[:-len(engine.BRAND_ALL_SUFFIX)][:22]} SC ALL"
    elif title.endswith(engine.BRAND_RANGE_SUFFIX):
        n = f"{title[:-len(engine.BRAND_RANGE_SUFFIX)][:22]} SC RANGE"
    else:
        n = title
    n = re.sub(r"[\[\]:*?/\\]", "-", n)[:31]
    base, k = n, 2
    while n.lower() in used:
        suf = f"~{k}"
        n = base[: 31 - len(suf)] + suf
        k += 1
    used.add(n.lower())
    return n


def tables_to_excel(tables: dict[str, list[list]]):
    bio = io.BytesIO()
    used: set[str] = set()
    summary = []
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        for title, values in tables.items():
            sname = short_name(title, used)
            df = pd.DataFrame(values[1:], columns=values[0])
            df.to_excel(writer, sheet_name=sname, index=False)
            writer.sheets[sname].freeze_panes = "A2"
            summary.append({"Sheet": sname, "Baris": len(df), "Kolom": len(df.columns)})
    return bio.getvalue(), pd.DataFrame(summary)


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
    return build("sheets", "v4", credentials=creds, cache_discovery=False)


def _id_norm(v) -> str:
    return str(v).strip().replace(",", "")


def _id_approx(v) -> str:
    """Kunci kasar (15 digit signifikan) untuk ID di sheet yang terlanjur jadi angka/sci-notation."""
    s_ = _id_norm(v)
    try:
        return f"{float(s_):.14e}"
    except ValueError:
        return s_


def _col_letter(i: int) -> str:
    i += 1
    out = ""
    while i:
        i, r = divmod(i - 1, 26)
        out = chr(65 + r) + out
    return out


NUMERIC_PUSH_COLS = PUSH_COLUMNS[9:]


def _num(v):
    v = str(v).strip().replace(",", "") if v is not None else ""
    if v == "":
        return ""
    try:
        f = float(v)
    except ValueError:
        return v
    return int(f) if f.is_integer() and abs(f) < 1e15 else f


def values_from_file(up) -> list[list]:
    """Baca CSV/TXT/XLSX hasil (kolom = PUSH_COLUMNS) menjadi values siap kirim."""
    name = up.name.lower()
    if name.endswith((".csv", ".txt")):
        df = pd.read_csv(io.BytesIO(up.getvalue()), dtype=str, keep_default_na=False,
                         encoding="utf-8-sig", sep=None, engine="python")
    else:
        df = pd.read_excel(io.BytesIO(up.getvalue()), dtype=str).fillna("")
    norm = engine.normalize_header
    have = {norm(c): c for c in df.columns}
    missing = [c for c in PUSH_COLUMNS if norm(c) not in have]
    if missing:
        raise ValueError(f"Kolom tidak ditemukan di file: {missing}")
    df = df[[have[norm(c)] for c in PUSH_COLUMNS]]
    df.columns = PUSH_COLUMNS
    rows = []
    for rec in df.values.tolist():
        rows.append([_num(x) if PUSH_COLUMNS[i] in NUMERIC_PUSH_COLS else str(x)
                     for i, x in enumerate(rec)])
    return [PUSH_COLUMNS] + rows


def push_to_sheet(values, link, tab_name, replace, dry_run=False):
    """Kirim values (baris pertama = header) ke Google Sheets target.
    Mode tambah: baris dengan (Video ID + Brand) yang sudah ada di tab tujuan SELALU dilewati."""
    m = re.search(r"/spreadsheets/d/([a-zA-Z0-9_-]+)", link or "")
    if not m:
        raise ValueError("Link target bukan link Google Sheets.")
    sid = m.group(1)
    g = re.search(r"[#&?]gid=(\d+)", link)
    svc = sheets_service()

    meta = svc.spreadsheets().get(
        spreadsheetId=sid, fields="sheets(properties(sheetId,title))").execute()
    props = [x["properties"] for x in meta.get("sheets", [])]

    tab_name = (tab_name or "").strip()
    if tab_name:
        target = next((p for p in props if p["title"].strip().lower() == tab_name.lower()), None)
        if target is None:
            if dry_run:
                title = tab_name
            else:
                svc.spreadsheets().batchUpdate(
                    spreadsheetId=sid,
                    body={"requests": [{"addSheet": {"properties": {"title": tab_name}}}]}).execute()
                title = tab_name
        else:
            title = target["title"]
    elif g:
        target = next((p for p in props if str(p["sheetId"]) == g.group(1)), None)
        if target is None:
            raise ValueError(f"gid {g.group(1)} tidak ditemukan di spreadsheet target.")
        title = target["title"]
    else:
        title = props[0]["title"]

    header, new_rows = values[0], values[1:]
    chunk = 2000
    total = len(new_rows)

    def _write_all(rows_with_header):
        for i in range(0, len(rows_with_header), chunk):
            svc.spreadsheets().values().update(
                spreadsheetId=sid, range=f"'{title}'!A{i + 1}",
                valueInputOption="RAW", body={"values": rows_with_header[i:i + chunk]}).execute()

    res_ = {"title": title, "sid": sid, "total": total, "skipped": 0, "added": total}

    # JANGAN telan error di sini: kalau gagal baca, berhenti (jangan dianggap tab kosong lalu menimpa)
    got = svc.spreadsheets().values().get(
        spreadsheetId=sid, range=f"'{title}'!1:1").execute().get("values", [])
    ex_header = got[0] if got else []
    sheet_empty = not any(str(c).strip() for c in ex_header)

    if replace:
        if not dry_run:
            svc.spreadsheets().values().clear(spreadsheetId=sid, range=f"'{title}'", body={}).execute()
            _write_all([header] + new_rows)
        return res_

    if sheet_empty:
        if not dry_run:
            _write_all([header] + new_rows)
        return res_

    norm = engine.normalize_header
    ours = {norm(h): i for i, h in enumerate(header)}
    ex_norm = [norm(h) for h in ex_header]
    if "videoid" not in ex_norm or "brand" not in ex_norm or "videoid" not in ours or "brand" not in ours:
        raise ValueError("Tab tujuan dan data harus punya kolom Video ID dan Brand untuk cek duplikat.")
    vcol = _col_letter(ex_norm.index("videoid"))
    bcol = _col_letter(ex_norm.index("brand"))

    cols = svc.spreadsheets().values().batchGet(
        spreadsheetId=sid, ranges=[f"'{title}'!{vcol}2:{vcol}", f"'{title}'!{bcol}2:{bcol}"],
        valueRenderOption="FORMATTED_VALUE").execute().get("valueRanges", [])
    va = cols[0].get("values", []) if len(cols) > 0 else []
    ba = cols[1].get("values", []) if len(cols) > 1 else []
    keys, approx_keys = set(), set()
    for i in range(max(len(va), len(ba))):
        v = _id_norm(va[i][0]) if i < len(va) and va[i] else ""
        b = str(ba[i][0]).strip().lower() if i < len(ba) and ba[i] else ""
        if not (v or b):
            continue
        if v.isdigit():
            keys.add((v, b))
        else:  # mis. 7.69232522369765E+18 -> bandingkan secara kasar
            approx_keys.add((_id_approx(v), b))

    ov, ob = ours["videoid"], ours["brand"]
    kept, skipped = [], 0
    for r in new_rows:
        vid, br = _id_norm(r[ov]), str(r[ob]).strip().lower()
        if (vid, br) in keys or (_id_approx(vid), br) in approx_keys:
            skipped += 1
            continue
        keys.add((vid, br))
        kept.append(r)

    aligned_idx = [ours.get(n) for n in ex_norm]
    aligned = [[(r[ix] if ix is not None else "") for ix in aligned_idx] for r in kept]
    res_.update({"added": len(aligned), "skipped": skipped})
    if dry_run:
        return res_
    for i in range(0, len(aligned), chunk):
        svc.spreadsheets().values().append(
            spreadsheetId=sid, range=f"'{title}'!A1", valueInputOption="RAW",
            insertDataOption="INSERT_ROWS", body={"values": aligned[i:i + chunk]}).execute()
    return res_


def run_pipeline(sc, taps, affs, bw, metric, rng, brand_tabs):
    ms, me = metric
    rs, re_ = rng
    if SKU_COL in engine.OUTPUT_COLUMNS:  # bersihkan sisa run sebelumnya
        engine.OUTPUT_COLUMNS.remove(SKU_COL)

    print("Membaca file...")
    sc_header, sc_rows = load_table(sc, engine.SC_REQUIRED, "SC VIDEO")
    print(f"  SC VIDEO ({sc[0]}): {len(sc_rows):,} baris")

    tap_datasets = []
    for i, s in enumerate(taps, 1):
        h, r = load_table(s, engine.TAP_REQUIRED, f"TAP {s[0]}")
        print(f"  TAP {i} ({s[0]}): {len(r):,} baris")
        tap_datasets.append((s[0], h, r))

    aff_datasets = []
    for i, s in enumerate(affs, 1):
        h, r = load_table(s, engine.AFF_REQUIRED, f"AFF {s[0]}")
        print(f"  AFF {i} ({s[0]}): {len(r):,} baris")
        aff_datasets.append((s[0], h, r))

    bw_header, bw_rows = load_table(bw, engine.BW_MAPPING_REQUIRED, "BW PRODUCT MAPPING")
    print(f"  BW MAPPING ({bw[0]}): {len(bw_rows):,} baris")

    bw_map, bw_stats = engine.build_bw_product_mapping(bw_header, bw_rows)
    engine.configure_brand_output_order(bw_map)
    engine.print_bw_mapping_summary(bw_stats)

    # sku name dari mapping BW (per Product ID)
    sku_lookup: dict[str, str] = {}
    sku_h = engine.resolve_first_header(bw_header, ["Product Name", "Product name", "Nama Produk", "SKU name", "SKU"])
    if sku_h is None:
        print("  PERINGATAN: kolom 'Product Name' tidak ada di mapping BW -> kolom sku name dikosongkan.")
    else:
        pid_h = engine.resolve_header(bw_header, "Product ID")
        for r in bw_rows:
            pid = engine.normalize_id(r.get(pid_h, ""))
            name = str(r.get(sku_h, "") or "").strip()
            if pid and name and pid not in sku_lookup:
                sku_lookup[pid] = name
        print(f"  SKU name ('{sku_h}')       : {len(sku_lookup):,} Product ID")

    tap_video_ids, tap_exist_stats = engine.build_tap_video_id_set(tap_datasets)

    print("\nMembangun TAP mapping...")
    tap_map, tap_stats = engine.build_tap_mapping(tap_datasets, ms, me)
    print("Membangun Affiliate Order mapping...")
    aff_map, aff_stats = engine.build_aff_mapping(aff_datasets, ms, me)
    print("Menggabungkan TAP + AFF...")
    master_map, master_counts = engine.build_master_mapping(tap_map, aff_map)
    pid_names = engine.build_pid_product_name_lookup(tap_map, aff_map)
    tap_vb_pairs = engine.build_tap_video_brand_pair_set_raw(tap_datasets, bw_map)

    print("Explode + enrich SC Video...")
    (all_rows, sc_stats, orig_tot, enr_tot, brand_counter, missing) = engine.enrich_sc(
        sc_header, sc_rows, master_map, aff_map, pid_names, bw_map
    )
    for row in all_rows:
        row[SKU_COL] = sku_lookup.get(engine.normalize_id(row.get("Product ID", "")), "")
    engine.OUTPUT_COLUMNS.insert(engine.OUTPUT_COLUMNS.index("Product ID") + 1, SKU_COL)

    range_rows, invalid_range = engine.filter_output_range(all_rows, rs, re_)

    nit_all = engine.filter_sc_not_in_tap(all_rows, tap_vb_pairs)
    nit_range = engine.filter_sc_not_in_tap(range_rows, tap_vb_pairs)

    bu_all = engine.build_brand_unique_rows(all_rows)
    bu_range = engine.build_brand_unique_rows(range_rows)
    bu_nit_all = engine.build_brand_unique_rows(nit_all)
    bu_nit_range = engine.build_brand_unique_rows(nit_range)

    tables = engine.build_output_tables(all_rows, range_rows, nit_all, nit_range, missing, brand_tabs)
    tables[engine.BRAND_UNIQUE_ALL_SHEET] = engine.data_values(bu_all)
    tables[engine.BRAND_UNIQUE_RANGE_SHEET] = engine.data_values(bu_range)
    tables[engine.BRAND_UNIQUE_NOT_IN_TAP_ALL_SHEET] = engine.data_values(bu_nit_all)
    tables[engine.BRAND_UNIQUE_NOT_IN_TAP_RANGE_SHEET] = engine.data_values(bu_nit_range)
    tables[engine.SUMMARY_SHEET] = engine.build_brand_summary(bu_all, bu_range)

    # supaya blok "Output" di summary tidak menyebut Google Sheets / dry run
    engine.WRITE = True
    engine.REPLACE = False
    engine.OUTPUT_SPREADSHEET_ID = "(file Excel lokal)"
    engine.print_summary(
        tap_stats, aff_stats, master_counts, sc_stats, orig_tot, enr_tot, brand_counter,
        all_rows, range_rows, nit_all, nit_range, tap_video_ids, tap_exist_stats,
        invalid_range, rs, re_, brand_tabs, missing,
    )
    push_values = {
        "range": [PUSH_COLUMNS] + [[row.get(c, "") for c in PUSH_COLUMNS] for row in bu_nit_range],
    }
    return tables, push_values


# ----------------------------------------------------------------------------
# UI
# ----------------------------------------------------------------------------
st.title("📊 SC BW Video — Explode + Brand Unique")
st.caption("Input Excel/CSV lokal → hasil Excel. Proses tidak butuh credentials; "
           "credentials.json hanya dipakai kalau kamu menekan tombol kirim ke Google Sheets.")

tab_in, tab_period, tab_run = st.tabs(["1 · File Input", "2 · Periode", "3 · Jalankan"])

with tab_in:
    st.markdown(
        "Untuk tiap input: **upload file** atau isi **path lokal** "
        "(path file, atau path folder untuk ambil semua Excel/CSV di dalamnya). "
        "Tab dan baris header dideteksi otomatis."
    )
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("RAW SC Video")
        sc_up = st.file_uploader("Upload SC", type=["xlsx", "xlsm", "xls", "csv", "txt"], key="sc_up")
        sc_path = st.text_input("atau path lokal (1 file)", key="sc_path",
                                placeholder=r"C:\Users\...\SC Video.xlsx")
        st.subheader("BW Product Mapping")
        bw_link = st.text_input(
            "Link Google Sheets mapping (Product ID | Brand | Sub Product)",
            key="bw_link",
            placeholder="https://docs.google.com/spreadsheets/d/.../edit?gid=0",
            help="Akses sheet harus 'Anyone with the link' (Viewer). "
                 "Link dengan #gid=... akan mengambil tab itu; tanpa gid = tab pertama. "
                 "Path file lokal juga boleh.",
        )
    with c2:
        st.subheader("RAW TAP Short Video")
        tap_up = st.file_uploader("Upload TAP (boleh banyak)", type=["xlsx", "xlsm", "xls", "csv", "txt"],
                                  accept_multiple_files=True, key="tap_up")
        tap_path = st.text_area("atau path lokal (satu per baris; boleh folder)", key="tap_path", height=70)
        st.subheader("RAW Affiliate Order")
        aff_up = st.file_uploader("Upload AFF (boleh banyak)", type=["xlsx", "xlsm", "xls", "csv", "txt"],
                                  accept_multiple_files=True, key="aff_up")
        aff_path = st.text_area("atau path lokal (satu per baris; boleh folder)", key="aff_path", height=70)

    st.info("ID (Video ID / Product ID / Content ID) di file sumber sebaiknya berformat **Text** "
            "agar 19 digitnya tidak terpotong.")


def _d(s):
    try:
        return date.fromisoformat(s) if s else None
    except ValueError:
        return None


def _pair(val):
    if isinstance(val, (list, tuple)):
        if len(val) == 2:
            return val[0], val[1]
        if len(val) == 1:
            return val[0], val[0]
    return None, None


with tab_period:
    today = date.today()
    r0, r1 = _d(engine.RANGE_START_DATE), _d(engine.RANGE_END_DATE)
    m0, m1 = _d(engine.METRIC_START_DATE), _d(engine.METRIC_END_DATE)

    st.subheader("RANGE — filter Post date SC")
    use_range = st.checkbox("Aktifkan filter RANGE", value=bool(r0 or r1))
    range_val = st.date_input("Rentang Post date",
                              value=(r0 or today - timedelta(days=6), r1 or today),
                              disabled=not use_range, key="range_dates")

    st.subheader("METRIC — filter TAP Date & AFF Time Created")
    same = st.checkbox("Samakan dengan RANGE", value=(m0 == r0 and m1 == r1))
    use_metric = st.checkbox("Aktifkan filter METRIC", value=bool(m0 or m1), disabled=same)
    metric_val = st.date_input("Rentang metric",
                               value=(m0 or today - timedelta(days=6), m1 or today),
                               disabled=same or not use_metric, key="metric_dates")
    st.caption("ALL = seluruh SC hasil explode. RANGE = subset berdasarkan Post date SC.")

with tab_run:
    rs, re_ = _pair(range_val) if use_range else (None, None)
    if same:
        ms, me = rs, re_
    elif use_metric:
        ms, me = _pair(metric_val)
    else:
        ms, me = None, None

    st.dataframe(
        pd.DataFrame(
            [["RANGE", f"{rs} → {re_}" if rs else "= ALL"],
             ["METRIC", f"{ms} → {me}" if ms else "Semua baris TAP/AFF"],
             ["Sheet per Brand", "Ya" if write_brand_tabs else "Tidak"]],
            columns=["Parameter", "Nilai"]),
        hide_index=True, use_container_width=True)

    out_dir = st.text_input("Simpan juga ke folder lokal (opsional)", placeholder=r"C:\Users\...\Output")
    default_name = f"SC_BW_VIDEO_{(rs or today).strftime('%Y%m%d')}_{(re_ or today).strftime('%Y%m%d')}.xlsx"
    file_name = st.text_input("Nama file output", value=default_name)

    auto_push = st.checkbox(
        "Langsung kirim ke Google Sheets setelah proses selesai (mode tambah, baris yang sudah ada dilewati)",
        value=False, key="auto_push",
        help="Memakai link dan nama tab di bagian 'Kirim ke Google Sheets' di bawah. "
             "Butuh credentials.json dan sheet di-share ke service account.")

    if st.button("▶️ Jalankan", type="primary"):
        st.session_state.pop("result", None)
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
                problems.append(f"BW Product Mapping harus tepat 1 file (sekarang {len(bw_items)}).")
            if not tap_items:
                problems.append("RAW TAP belum diisi.")
            if not aff_items:
                problems.append("RAW Affiliate Order belum diisi.")
            if use_range and (rs is None or re_ is None):
                problems.append("Rentang RANGE belum lengkap.")
            if problems:
                for p in problems:
                    st.error(p)
            else:
                with st.spinner("Memproses… (data besar bisa beberapa menit)"):
                    with contextlib.redirect_stdout(writer):
                        tables, push_values = run_pipeline(
                            sc_items[0], tap_items, aff_items, bw_items[0],
                            (ms, me), (rs, re_), write_brand_tabs)
                    xlsx_bytes, summary_df = tables_to_excel(tables)

                saved_to = None
                if out_dir.strip():
                    folder = Path(out_dir.strip().strip('"'))
                    folder.mkdir(parents=True, exist_ok=True)
                    saved_to = folder / file_name
                    saved_to.write_bytes(xlsx_bytes)

                st.session_state["result"] = {
                    "bytes": xlsx_bytes, "name": file_name, "summary": summary_df,
                    "log": writer.buf, "saved_to": str(saved_to) if saved_to else None,
                    "push": push_values,
                }

                if auto_push:
                    if not service_account_email():
                        st.error(f"Auto-kirim dilewati: credentials.json tidak terbaca di `{cred_file()}`.")
                    else:
                        tab_auto = (st.session_state.get("tab_range") or "SC VID Range").strip()
                        link_auto = st.session_state.get("target_link") or DEFAULT_TARGET_LINK
                        try:
                            with st.spinner(f"Mengirim ke tab '{tab_auto}' …"):
                                out = push_to_sheet(push_values["range"], link_auto, tab_auto, False)
                            st.success(f"Auto-kirim → tab '{out['title']}': {out['added']:,} baris ditambahkan"
                                       + (f", {out['skipped']:,} dilewati (sudah ada)." if out["skipped"] else "."))
                        except Exception as e:  # noqa: BLE001
                            st.error(f"Auto-kirim gagal: {type(e).__name__}: {e}")
        except SystemExit as e:
            st.error(f"Berhenti: {e.code}")
        except Exception as e:  # noqa: BLE001
            st.error(f"Error: {type(e).__name__}: {e}")
            with st.expander("Detail error"):
                st.code(traceback.format_exc())

    res = st.session_state.get("result")
    if res:
        st.success("Selesai.")
        if res["saved_to"]:
            st.caption(f"Tersimpan di: `{res['saved_to']}`")
        st.download_button(
            "⬇️ Download hasil (.xlsx)", res["bytes"], file_name=res["name"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
        st.subheader("Isi file output")
        st.dataframe(res["summary"], hide_index=True, use_container_width=True)
        with st.expander("Log QC lengkap"):
            st.code(res["log"], language="text")
            st.download_button("Download log", res["log"], file_name="sc_bw_log.txt")

    st.divider()
    st.subheader("📤 Kirim ke Google Sheets")
    st.caption("Data hanya DITAMBAH di bawah. Baris dengan Video ID + Brand yang sudah ada di tab tujuan "
               "otomatis dilewati, jadi aman dikirim berulang kali.")

    TARGETS = [
        ("range", "BRAND UNIQ NOT IN TAP - RANGE", "SC VID Range"),
    ]
    src_mode = st.radio(
        "Sumber data",
        ["Hasil proses di atas", "File CSV/TXT/XLSX (kolom sama dengan tabel tujuan)"],
        horizontal=True)

    datasets: dict[str, list[list]] = {}
    if src_mode.startswith("Hasil"):
        datasets = dict(res["push"]) if res and res.get("push") else {}
        if not datasets:
            st.info("Jalankan proses dulu, atau pilih sumber file.")
    else:
        for key, label, _ in TARGETS:
            up = st.file_uploader(f"File untuk {label} (opsional)", type=["csv", "txt", "xlsx"],
                                  key=f"push_up_{key}")
            if up is not None:
                try:
                    datasets[key] = values_from_file(up)
                except Exception as e:  # noqa: BLE001
                    st.error(f"{label}: file tidak bisa dibaca: {e}")

    st.markdown("**Tujuan**")
    target_link = st.text_input("Link Google Sheets tujuan (bisa diganti)", value=DEFAULT_TARGET_LINK,
                                key="target_link")
    tcols = st.columns(len(TARGETS))
    tab_names = {
        key: col.text_input(f"Tab tujuan ← {label}", value=default_tab, key=f"tab_{key}")
        for col, (key, label, default_tab) in zip(tcols, TARGETS)
    }
    st.caption("Tab yang belum ada di spreadsheet tujuan dibuat otomatis.")

    if datasets:
        tabs_for: dict[str, str] = {}
        for key, label, _ in TARGETS:
            if key not in datasets:
                continue
            tabs_for[key] = tab_names[key]
            with st.expander(f"{label} — {len(datasets[key]) - 1:,} baris → tab '{tab_names[key]}' (preview 10 baris)"):
                st.dataframe(pd.DataFrame(datasets[key][1:11], columns=datasets[key][0]),
                             hide_index=True, use_container_width=True)

        mode_push = st.radio(
            "Mode", ["Tambah di bawah data yang ada", "Timpa seluruh tab"], horizontal=True)
        do_replace = mode_push.startswith("Timpa")

        sa_email = service_account_email()
        if sa_email:
            st.info(f"Share sheet tujuan sebagai **Editor** ke: `{sa_email}`")
        else:
            st.warning(f"credentials.json tidak ditemukan/terbaca di `{cred_file()}`. "
                       "Taruh di folder app atau isi path-nya di sidebar → Lanjutan.")

        ok_replace = True
        if do_replace:
            ok_replace = st.checkbox("Saya paham isi tab tujuan akan DIHAPUS lalu diganti.")

        names = [t.strip() for t in tabs_for.values()]
        tab_problem = None
        if any(not n for n in names):
            tab_problem = "Nama tab tujuan tidak boleh kosong."
        if tab_problem:
            st.error(tab_problem)

        def run_push(dry: bool):
            for key, label, _ in TARGETS:
                if key not in datasets:
                    continue
                try:
                    with st.spinner(f"{label} → '{tabs_for[key]}' …"):
                        out = push_to_sheet(datasets[key], target_link, tabs_for[key], do_replace, dry_run=dry)
                except Exception as e:  # noqa: BLE001
                    st.error(f"{label}: gagal — {type(e).__name__}: {e}")
                    continue
                if do_replace:
                    st.warning(f"{label}: tab '{out['title']}' " +
                               ("akan" if dry else "sudah") + f" diganti dengan {out['total']:,} baris.")
                elif dry:
                    st.success(f"{label} → '{out['title']}': {out['added']:,} baris BARU akan ditambahkan, "
                               f"{out['skipped']:,} sudah ada (dilewati).")
                else:
                    st.success(f"{label} → '{out['title']}': {out['added']:,} baris ditambahkan"
                               + (f", {out['skipped']:,} dilewati (sudah ada)." if out["skipped"] else "."))
            if not dry:
                sid_m = re.search(r"/spreadsheets/d/([a-zA-Z0-9_-]+)", target_link or "")
                if sid_m:
                    st.link_button("Buka spreadsheet",
                                   f"https://docs.google.com/spreadsheets/d/{sid_m.group(1)}")

        can_go = bool(sa_email) and not tab_problem
        b1, b2 = st.columns(2)
        if b1.button("🔍 Cek dulu (tidak menulis)", disabled=not can_go):
            run_push(True)
        if b2.button("🚀 Kirim ke Google Sheets", type="primary", disabled=not (can_go and ok_replace)):
            run_push(False)

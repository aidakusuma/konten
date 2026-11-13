"""
SC BW LIVE / LIVESTREAM - EXPLODED + BRAND UNIQUE + TAP LS/AFF ENRICHMENT
============================================================================

Flow utama
----------
1. RAW SC LIVE dibaca dari Google Sheets.
2. Product ID SC yang comma-separated di-explode menjadi 1 row per Product ID.
3. Dari hasil exploded dibuat output BRAND UNIQUE:
       1 row = LIVE ID + Brand
       Product ID = FIRST PID di dalam LIVE ID + Brand
       GMV / Orders / Items = SUM metric UNIQUE PID di dalam LIVE ID + Brand.
4. Satu atau lebih RAW TAP LS digabung menjadi satu mapping.
   Struktur TAP saat ini memakai exact key:
       LIVE ID + Product ID
4. AFF Livestream mapping memakai exact key:
       Content ID + Product ID
5. Model A:
   - Single Product ID:
       GMV / Orders / Items tetap dari SC original.
   - Multi Product ID:
       metric memakai exact TAP pair;
       jika TAP exact tidak ada, AFF exact sebagai fallback;
       jika pair tidak ditemukan, metric commerce = 0.
6. AOV dihitung ulang dari mapped GMV / mapped Orders.
7. Product name:
       exact TAP/AFF -> PID lookup global -> fallback SC product name.
8. Brand + Sub Product:
       HANYA dari Product ID spreadsheet mapping BW.
       Product name tidak dipakai untuk klasifikasi.
9. Product ID yang tidak ada di mapping BW:
       Brand = MISSING
       Sub Product = MISSING
10. TAP metric period memakai LIVE start time.
11. RANGE SC memakai LIVE start time.
12. WRITE_BRAND_TABS hanya mengontrol penulisan tab per Brand.

Run
---
Dry run:
    python sc_bw_live_exploded_multi_tap_flexible.py

Write:
    python sc_bw_live_exploded_multi_tap_flexible.py --write

Replace managed output tabs:
    python sc_bw_live_exploded_multi_tap_flexible.py --write --replace

Brand tabs:
    python sc_bw_live_exploded_multi_tap_flexible.py --write --brand-tabs
    python sc_bw_live_exploded_multi_tap_flexible.py --write --no-brand-tabs

Dependency
----------
    pip install google-api-python-client google-auth
"""


from __future__ import annotations

import math
import re
import sys
import time
from collections import Counter, defaultdict
from datetime import date, datetime
from pathlib import Path


# =============================================================================
# CONFIG - EDIT BAGIAN INI
# =============================================================================

BASE_DIR = Path(__file__).resolve().parent

CREDENTIALS_CANDIDATES = [
    BASE_DIR / "credentials.json",
    BASE_DIR / "credentials_sc.json",
]
CREDENTIALS_FILE = next(
    (path for path in CREDENTIALS_CANDIDATES if path.exists()),
    CREDENTIALS_CANDIDATES[0],
)

# RAW SC LIVE / Seller Center Livestream.
SC_LS_SPREADSHEET_ID = "1byE9MptNeq4x30TV6YYJg-SmYqPlNPT7F0Y--O-kYwc"
SC_LS_SHEET_GID = None
SC_LS_SHEET_NAME = None

# RAW TAP Livestream.
# Bisa satu atau lebih source, sama seperti AFF_SOURCES.
# Struktur TAP yang dipakai sekarang:
#   LIVE ID
#   Product ID
#   LIVE start time
#   Creator LIVE-attributed GMV
#   LIVE-attributed orders
#   LIVE-attributed items sold
# Product name OPTIONAL.
# Semua source TAP digabung berdasarkan LIVE ID + Product ID.
# Exact duplicate row antar/source di-skip agar tidak double count.
TAP_LS_SOURCES = [
    {
        "spreadsheet_id": "1wtCGvvIJir5gzwCjNkKawcTSyB36YHJL1i4hkR1NY08",
        "sheet_gid": None,
        "sheet_name": "Custom report",
        "label": "TAP LS 1",
    },
    {
        "spreadsheet_id": "1hTsYPp0AUo5XiSX_1FDNr2GuVACskFULwnfCwPzbPDw",
        "sheet_gid": None,
        "sheet_name": "Custom report",
        "label": "TAP LS 2",
    },
]

# BW Product Mapping.
# Harus memiliki header: Product ID | Brand | Sub Product
BW_MAPPING_SPREADSHEET_ID = "1Wi0kcm7apS4gbLDcHJdz8P1iRohF62aGzM38Oy_f7gM"
BW_MAPPING_SHEET_GID = None
BW_MAPPING_SHEET_NAME = None

# AFF sama dengan flow SC Video sebelumnya.
AFF_SOURCES = [
    {
        "spreadsheet_id": "1hph0ZOeFBy7eVjo6dtMy904FmklbAL_5vWImwQceT1A",
        "sheet_gid": None,
        "sheet_name": None,
        "label": "AFF 1",
    },
    {
        "spreadsheet_id": "1ws8skmRKvzus8I56Ba8kOwR_lJW0Vc3zEfPgiTbbV2U",
        "sheet_gid": None,
        "sheet_name": None,
        "label": "AFF 2",
    },
]

OUTPUT_SPREADSHEET_ID = "1fQxmq0d_LcCTI41VBSXiPO0VjpNFC4rtu4i-qaGimtc"

# Metric mapping period.
# TAP LS  -> LIVE start time.
# AFF     -> Time Created.
METRIC_START_DATE = "2026-10-01"
METRIC_END_DATE = "2026-10-10"

# SC output RANGE -> LIVE start time.
RANGE_START_DATE = "2026-10-01"
RANGE_END_DATE = "2026-10-10"

WRITE_BRAND_TABS = False

SOURCE_SCAN_RANGE = "A1:AZ20"
WRITE_CHUNK_ROWS = 1500
MAX_CELLS = 20_000_000
MAX_RETRIES = 5
RETRYABLE_HTTP_STATUS = {429, 500, 502, 503, 504}

SHEET_ALL = "GABUNGAN SC BW LS EXPLODED - ALL"
SHEET_RANGE = "GABUNGAN SC BW LS EXPLODED - RANGE"
SC_NOT_IN_TAP_ALL_SHEET = "SC BW LS EXPLODED NOT IN TAP - ALL"
SC_NOT_IN_TAP_RANGE_SHEET = "SC BW LS EXPLODED NOT IN TAP - RANGE"
BRAND_UNIQUE_ALL_SHEET = "GABUNGAN SC BW LS BRAND UNIQUE - ALL"
BRAND_UNIQUE_RANGE_SHEET = "GABUNGAN SC BW LS BRAND UNIQUE - RANGE"
BRAND_UNIQUE_NOT_IN_TAP_ALL_SHEET = "SC BW LS BRAND UNIQUE NOT IN TAP - ALL"
BRAND_UNIQUE_NOT_IN_TAP_RANGE_SHEET = "SC BW LS BRAND UNIQUE NOT IN TAP - RANGE"
MAPPING_QC_ALL_SHEET = "MAPPING QC BW LS EXPLODED - ALL"
MAPPING_QC_RANGE_SHEET = "MAPPING QC BW LS EXPLODED - RANGE"
SUMMARY_SHEET = "BW BRAND SUMMARY LS"
MISSING_SHEET = "BW LS EXPLODED MISSING TAP-AFF PAIR"
BRAND_ALL_SUFFIX = " - SC BW LS EXPLODED ALL"
BRAND_RANGE_SUFFIX = " - SC BW LS EXPLODED RANGE"

CLEANING_VERSION = "bw-live-exploded-product-mapping"


# =============================================================================
# MODE
# =============================================================================

WRITE = "--write" in sys.argv
REPLACE = "--replace" in sys.argv

if REPLACE and not WRITE:
    raise SystemExit("--replace hanya boleh dipakai bersama --write")

if "--brand-tabs" in sys.argv and "--no-brand-tabs" in sys.argv:
    raise SystemExit("Pilih salah satu: --brand-tabs atau --no-brand-tabs")

if "--brand-tabs" in sys.argv:
    WRITE_BRAND_TABS_EFFECTIVE = True
elif "--no-brand-tabs" in sys.argv:
    WRITE_BRAND_TABS_EFFECTIVE = False
else:
    WRITE_BRAND_TABS_EFFECTIVE = WRITE_BRAND_TABS


# =============================================================================
# OUTPUT SCHEMA
# =============================================================================

OUTPUT_COLUMNS = [
    "LIVE title",
    "LIVE ID",
    "LIVE start time",
    "LIVE end time",
    "Creator name",
    "Product ID",
    "Product name",
    "Brand",
    "Sub Product",
    "Creator LIVE-attributed GMV",
    "LIVE-attributed items sold",
    "Refunds",
    "LIVE-attributed orders",
    "AOV",
    "Avg. viewing duration",
    "Likes",
    "Comments",
    "Shares",
    "LIVE product impressions",
    "Product clicks",
    "Impressions",
    "Avg. daily GPM",
    "Avg. daily products sold",
    "Avg. GMV per customer",
    "Avg. daily engagement",
    "Avg. daily LIVE product CTR",
    "Avg. daily unique viewers",
    "Avg. daily tap-through rate",
    "Est. commission",
    "GMV Affiliate Order (Est. Commission Base)",
    "Jumlah Order Affiliate",
]

SC_REQUIRED = [
    "LIVE title",
    "LIVE ID",
    "LIVE start time",
    "Product ID",
    "Creator LIVE-attributed GMV",
    "LIVE-attributed items sold",
    "LIVE-attributed orders",
]

# TAP source boleh memakai struktur BARU, LAMA, atau hybrid.
# Header akan di-resolve melalui alias, jadi source TAP yang berbeda
# tidak wajib memiliki nama kolom yang identik.
TAP_HEADER_ALIASES = {
    "live_id": ["LIVE ID", "Livestream room ID"],
    "product_id": ["Product ID"],
    "product_name": ["Product name", "Product Name"],
    "live_start": ["LIVE start time"],
    "live_time_info": ["LIVE time info"],
    "gmv": ["Creator LIVE-attributed GMV"],
    "orders": ["LIVE-attributed orders", "Creator LIVE-attributed orders"],
    "items": ["LIVE-attributed items sold", "Creator-attributed items sold"],
}

AFF_REQUIRED = [
    "Order ID",
    "Product ID",
    "Content Type",
    "Content ID",
    "Est. Commission Base",
    "Quantity",
    "Time Created",
]

INTEGER_COLUMNS = {
    "LIVE-attributed items sold",
    "LIVE-attributed orders",
    "Likes",
    "Comments",
    "Shares",
    "LIVE product impressions",
    "Product clicks",
    "Impressions",
    "Avg. daily products sold",
    "Avg. daily unique viewers",
}

NUMERIC_COLUMNS = {
    "Creator LIVE-attributed GMV",
    "Refunds",
    "AOV",
    "Avg. daily GPM",
    "Avg. GMV per customer",
    "Est. commission",
}

PERCENT_COLUMNS = {
    "Avg. daily engagement",
    "Avg. daily LIVE product CTR",
    "Avg. daily tap-through rate",
}


# =============================================================================
# BW PRODUCT CLASSIFICATION
# =============================================================================

MISSING_BRAND = "MISSING"
MISSING_SUB_PRODUCT = "MISSING"

# Diisi secara dinamis dari spreadsheet mapping BW pada saat runtime.
BRAND_OUTPUT_ORDER = [MISSING_BRAND]

BW_MAPPING_REQUIRED = [
    "Product ID",
    "Brand",
    "Sub Product",
]


# =============================================================================
# GENERAL HELPERS
# =============================================================================

def normalize_header(value: object) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(value or "").strip().lower())


def normalize_id(value: object) -> str:
    if value is None:
        return ""
    s = str(value).strip()
    if not s or s.lower() in {"nan", "none", "null", "--", "-"}:
        return ""
    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]
    if re.fullmatch(r"[+-]?\d+(?:\.\d+)?[eE][+-]?\d+", s):
        raise ValueError(
            f"ID dalam scientific notation tidak aman: {s!r}. "
            "Simpan ID di Google Sheets sebagai text / digit utuh."
        )
    return s


def split_product_ids(value: object) -> list[str]:
    s = str(value or "").strip()
    if not s:
        return []
    result = []
    seen = set()
    for piece in s.split(","):
        pid = normalize_id(piece)
        if pid and pid not in seen:
            result.append(pid)
            seen.add(pid)
    return result


def first_product_id(value: object) -> str:
    """Ambil Product ID pertama saja dari field comma-separated SC LIVE."""
    product_ids = split_product_ids(value)
    return product_ids[0] if product_ids else ""


def parse_number(value: object) -> float:
    if value is None:
        return 0.0
    if isinstance(value, bool):
        return 0.0
    if isinstance(value, (int, float)):
        if isinstance(value, float) and math.isnan(value):
            return 0.0
        return float(value)

    s = str(value).strip()
    if not s or s.upper() in {"-", "--", "N/A", "NA", "NULL", "NONE"}:
        return 0.0

    s = re.sub(r"^(?:RP|IDR)\s*", "", s, flags=re.I).replace(" ", "")
    negative = False
    if s.startswith("(") and s.endswith(")"):
        negative = True
        s = s[1:-1]

    if re.fullmatch(r"[+-]?\d{1,3}(?:\.\d{3})+", s):
        s = s.replace(".", "")
    elif re.fullmatch(r"[+-]?\d{1,3}(?:,\d{3})+", s):
        s = s.replace(",", "")
    elif re.fullmatch(r"[+-]?\d{1,3}(?:\.\d{3})*,\d+", s):
        s = s.replace(".", "").replace(",", ".")
    elif "." in s and "," in s:
        if s.rfind(".") > s.rfind(","):
            s = s.replace(",", "")
        else:
            s = s.replace(".", "").replace(",", ".")
    else:
        s = s.replace(",", "")

    try:
        out = float(s)
    except ValueError as exc:
        raise ValueError(f"Angka tidak valid: {value!r}") from exc

    return -out if negative else out


def parse_int(value: object) -> int:
    return int(round(parse_number(value)))


def parse_percent(value: object) -> float:
    if value is None:
        return 0.0
    s = str(value).strip()
    if not s:
        return 0.0
    if s.endswith("%"):
        return parse_number(s[:-1]) / 100.0
    return parse_number(s)


def clean_numeric(value: float) -> int | float:
    if math.isclose(float(value), round(float(value)), rel_tol=0, abs_tol=1e-9):
        return int(round(float(value)))
    return round(float(value), 6)


def parse_iso_date(value: str | None) -> date | None:
    if value is None or str(value).strip() == "":
        return None
    return datetime.strptime(str(value).strip(), "%Y-%m-%d").date()


def parse_sc_live_datetime(value: object) -> datetime | None:
    # SC LIVE export = MDY.
    s = str(value or "").strip()
    if not s or s in {"-", "--"}:
        return None

    for fmt in (
        "%m/%d/%Y %H:%M:%S",
        "%m/%d/%Y %H:%M",
        "%m/%d/%Y",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
        "%Y-%m-%d",
    ):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            pass

    raise ValueError(f"SC LIVE datetime tidak valid: {value!r}")


LIVE_TIME_RE = re.compile(
    r"^\s*"
    r"(\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2})"
    r"\s*-\s*"
    r"(\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2})"
    r"\s*$"
)


def parse_tap_live_time_info(value: object) -> tuple[datetime, datetime] | None:
    s = str(value or "").strip()
    if not s or s in {"-", "--"}:
        return None

    match = LIVE_TIME_RE.fullmatch(s)
    if not match:
        raise ValueError(f"TAP LIVE time info tidak valid: {value!r}")

    start_dt = datetime.strptime(match.group(1), "%Y-%m-%d %H:%M:%S")
    end_dt = datetime.strptime(match.group(2), "%Y-%m-%d %H:%M:%S")
    return start_dt, end_dt


def parse_aff_datetime(value: object) -> datetime | None:
    # AFF export = DD/MM/YYYY.
    s = str(value or "").strip()
    if not s or s in {"-", "--"}:
        return None

    for fmt in (
        "%d/%m/%Y %H:%M:%S",
        "%d/%m/%Y %H:%M",
        "%d/%m/%Y",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
        "%Y-%m-%d",
    ):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            pass

    raise ValueError(f"AFF Time Created tidak valid: {value!r}")


def normalize_sc_live_datetime(value: object) -> str:
    try:
        dt = parse_sc_live_datetime(value)
    except ValueError:
        return str(value or "").strip()
    return dt.strftime("%Y-%m-%d %H:%M:%S") if dt else ""


def parse_output_live_date(value: object) -> date | None:
    s = str(value or "").strip()
    if not s or s in {"-", "--"}:
        return None

    for fmt in (
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
        "%Y-%m-%d",
        "%m/%d/%Y %H:%M:%S",
        "%m/%d/%Y %H:%M",
        "%m/%d/%Y",
    ):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return None


def in_period(d: date | None, start: date | None, end: date | None) -> bool:
    if d is None:
        return False if (start or end) else True
    if start and d < start:
        return False
    if end and d > end:
        return False
    return True


def resolve_header(header: list[str], wanted: str) -> str:
    target = normalize_header(wanted)
    for name in header:
        if normalize_header(name) == target:
            return name
    raise KeyError(wanted)


def resolve_first_header(header: list[str], candidates: list[str]) -> str | None:
    normalized = {normalize_header(name): name for name in header}
    for candidate in candidates:
        found = normalized.get(normalize_header(candidate))
        if found is not None:
            return found
    return None


def validate_headers(header: list[str], required: list[str], label: str) -> None:
    missing = []
    for wanted in required:
        try:
            resolve_header(header, wanted)
        except KeyError:
            missing.append(wanted)
    if missing:
        raise ValueError(
            f"{label}: kolom wajib tidak ditemukan: {', '.join(missing)}\n"
            f"Header tersedia: {header}"
        )


def column_letter(index_1_based: int) -> str:
    result = ""
    n = index_1_based
    while n:
        n, rem = divmod(n - 1, 26)
        result = chr(65 + rem) + result
    return result or "A"


def trim_header(row: list[object]) -> list[str]:
    values = [str(v or "").strip() for v in row]
    last = -1
    for i, value in enumerate(values):
        if value:
            last = i
    return values[: last + 1] if last >= 0 else []


def rows_to_dicts(
    header: list[str],
    rows: list[list[object]],
) -> list[dict[str, object]]:
    result = []
    width = len(header)
    for raw in rows:
        padded = list(raw[:width]) + [""] * max(0, width - len(raw))
        result.append({header[i]: padded[i] for i in range(width)})
    return result


# =============================================================================
# GOOGLE SHEETS
# =============================================================================

def get_sheets_service():
    try:
        from google.oauth2 import service_account
        from googleapiclient.discovery import build
    except ImportError as exc:
        raise RuntimeError(
            "Google library belum terinstall. Jalankan:\n"
            "  pip install google-api-python-client google-auth"
        ) from exc

    scope = (
        "https://www.googleapis.com/auth/spreadsheets"
        if WRITE
        else "https://www.googleapis.com/auth/spreadsheets.readonly"
    )

    credentials = service_account.Credentials.from_service_account_file(
        str(CREDENTIALS_FILE),
        scopes=[scope],
    )

    return build(
        "sheets",
        "v4",
        credentials=credentials,
        cache_discovery=False,
    )


def execute_with_retry(request):
    for attempt in range(MAX_RETRIES):
        try:
            return request.execute()
        except Exception as exc:
            status = getattr(getattr(exc, "resp", None), "status", None)
            if status not in RETRYABLE_HTTP_STATUS or attempt == MAX_RETRIES - 1:
                raise
            wait_seconds = 2 ** attempt
            print(
                f"Google API error {status}; retry {attempt + 1}/{MAX_RETRIES} "
                f"dalam {wait_seconds}s..."
            )
            time.sleep(wait_seconds)


def get_sheet_properties(service, spreadsheet_id: str):
    response = execute_with_retry(
        service.spreadsheets().get(
            spreadsheetId=spreadsheet_id,
            fields="sheets(properties(sheetId,title,gridProperties(rowCount,columnCount)))",
        )
    )
    return [sheet["properties"] for sheet in response.get("sheets", [])]


def detect_and_read_source(
    service,
    *,
    spreadsheet_id: str,
    sheet_gid: int | None,
    sheet_name: str | None,
    required_headers: list[str],
    label: str,
):
    properties = get_sheet_properties(service, spreadsheet_id)
    if not properties:
        raise ValueError(f"{label}: spreadsheet tidak memiliki sheet.")

    titles = [p["title"] for p in properties]

    if sheet_name is not None:
        if sheet_name not in titles:
            raise ValueError(
                f'{label}: sheet_name "{sheet_name}" tidak ditemukan. Available: {titles}'
            )
        candidates = [p for p in properties if p["title"] == sheet_name]
    elif sheet_gid is not None:
        candidates = [p for p in properties if int(p["sheetId"]) == int(sheet_gid)]
        if not candidates:
            available = [(p["sheetId"], p["title"]) for p in properties]
            raise ValueError(
                f"{label}: sheet_gid {sheet_gid} tidak ditemukan. Available: {available}"
            )
    else:
        candidates = properties

    required_norm = {normalize_header(x) for x in required_headers}

    for prop in candidates:
        title = prop["title"]
        scan = execute_with_retry(
            service.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id,
                range=f"'{title}'!{SOURCE_SCAN_RANGE}",
                valueRenderOption="FORMATTED_VALUE",
            )
        )
        scan_rows = scan.get("values", [])

        for header_row_number, raw_header in enumerate(scan_rows, start=1):
            header = trim_header(raw_header)
            if not header:
                continue

            normalized = {normalize_header(v) for v in header if v}
            if not required_norm.issubset(normalized):
                continue

            end_col = column_letter(len(header))
            values = execute_with_retry(
                service.spreadsheets().values().get(
                    spreadsheetId=spreadsheet_id,
                    range=f"'{title}'!A{header_row_number}:{end_col}",
                    valueRenderOption="FORMATTED_VALUE",
                )
            ).get("values", [])

            if not values:
                raise ValueError(f'{label}: tab "{title}" kosong.')

            final_header = trim_header(values[0])
            data_rows = rows_to_dicts(final_header, values[1:])

            meta = {
                "spreadsheet_id": spreadsheet_id,
                "sheet_id": prop["sheetId"],
                "sheet_title": title,
                "header_row": header_row_number,
                "columns": len(final_header),
                "rows": len(data_rows),
            }
            return final_header, data_rows, meta

    raise ValueError(
        f"{label}: tidak menemukan header yang sesuai. Required: {required_headers}"
    )


def resolve_tap_layout(
    header: list[str],
    label: str = "TAP LS",
) -> dict[str, object]:
    """
    Resolve layout TAP LS secara fleksibel.

    Didukung:
    - NEW:
        LIVE ID
        Product ID
        LIVE start time
        Creator LIVE-attributed GMV
        LIVE-attributed orders
        LIVE-attributed items sold

    - OLD:
        Livestream room ID
        Product ID
        LIVE time info
        Creator LIVE-attributed GMV
        Creator LIVE-attributed orders
        Creator-attributed items sold

    - Hybrid kombinasi alias di atas.
    """
    def pick(key: str, required: bool = True):
        found = resolve_first_header(
            header,
            TAP_HEADER_ALIASES[key],
        )
        if required and found is None:
            raise ValueError(
                f"{label}: header TAP untuk '{key}' tidak ditemukan. "
                f"Alias diterima: {TAP_HEADER_ALIASES[key]}"
            )
        return found

    live_id = pick("live_id")
    product_id = pick("product_id")
    product_name = pick("product_name", required=False)
    live_start = pick("live_start", required=False)
    live_time_info = pick("live_time_info", required=False)
    gmv = pick("gmv")
    orders = pick("orders")
    items = pick("items")

    if live_start is None and live_time_info is None:
        raise ValueError(
            f"{label}: harus punya salah satu kolom waktu TAP: "
            f"{TAP_HEADER_ALIASES['live_start']} atau "
            f"{TAP_HEADER_ALIASES['live_time_info']}"
        )

    if live_start is not None:
        time_kind = "LIVE_START_TIME"
        time_header = live_start
    else:
        time_kind = "LIVE_TIME_INFO"
        time_header = live_time_info

    if normalize_header(live_id) == normalize_header("Livestream room ID"):
        id_kind = "LIVESTREAM_ROOM_ID"
    else:
        id_kind = "LIVE_ID"

    return {
        "live_id": live_id,
        "product_id": product_id,
        "product_name": product_name,
        "time_header": time_header,
        "time_kind": time_kind,
        "gmv": gmv,
        "orders": orders,
        "items": items,
        "schema": f"{id_kind} + {time_kind}",
    }


def parse_tap_start_datetime(
    row: dict[str, object],
    layout: dict[str, object],
) -> datetime | None:
    """
    Ambil start datetime dari TAP, baik schema baru maupun lama.
    """
    raw = row.get(str(layout["time_header"]), "")

    if layout["time_kind"] == "LIVE_START_TIME":
        return parse_sc_live_datetime(raw)

    live_period = parse_tap_live_time_info(raw)
    return live_period[0] if live_period else None


def detect_and_read_tap_source(
    service,
    *,
    spreadsheet_id: str,
    sheet_gid: int | None,
    sheet_name: str | None,
    label: str,
):
    """
    Detect TAP sheet tanpa memaksa satu set header exact.

    Sheet dianggap TAP valid jika resolve_tap_layout() berhasil.
    """
    properties = get_sheet_properties(
        service,
        spreadsheet_id,
    )

    if not properties:
        raise ValueError(
            f"{label}: spreadsheet tidak memiliki sheet."
        )

    titles = [p["title"] for p in properties]

    if sheet_name is not None:
        if sheet_name not in titles:
            raise ValueError(
                f'{label}: sheet_name "{sheet_name}" tidak ditemukan. '
                f"Available: {titles}"
            )
        candidates = [
            p for p in properties
            if p["title"] == sheet_name
        ]
    elif sheet_gid is not None:
        candidates = [
            p for p in properties
            if int(p["sheetId"]) == int(sheet_gid)
        ]
        if not candidates:
            available = [
                (p["sheetId"], p["title"])
                for p in properties
            ]
            raise ValueError(
                f"{label}: sheet_gid {sheet_gid} tidak ditemukan. "
                f"Available: {available}"
            )
    else:
        candidates = properties

    inspected = []

    for prop in candidates:
        title = prop["title"]

        scan = execute_with_retry(
            service.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id,
                range=f"'{title}'!{SOURCE_SCAN_RANGE}",
                valueRenderOption="FORMATTED_VALUE",
            )
        )

        scan_rows = scan.get("values", [])

        for header_row_number, raw_header in enumerate(
            scan_rows,
            start=1,
        ):
            header = trim_header(raw_header)

            if not header:
                continue

            try:
                layout = resolve_tap_layout(
                    header,
                    label=label,
                )
            except ValueError:
                if len(inspected) < 8:
                    inspected.append(
                        (
                            title,
                            header_row_number,
                            header,
                        )
                    )
                continue

            end_col = column_letter(len(header))

            values = execute_with_retry(
                service.spreadsheets().values().get(
                    spreadsheetId=spreadsheet_id,
                    range=(
                        f"'{title}'!"
                        f"A{header_row_number}:{end_col}"
                    ),
                    valueRenderOption="FORMATTED_VALUE",
                )
            ).get("values", [])

            if not values:
                raise ValueError(
                    f'{label}: tab "{title}" kosong.'
                )

            final_header = trim_header(values[0])
            final_layout = resolve_tap_layout(
                final_header,
                label=label,
            )

            data_rows = rows_to_dicts(
                final_header,
                values[1:],
            )

            meta = {
                "spreadsheet_id": spreadsheet_id,
                "sheet_id": prop["sheetId"],
                "sheet_title": title,
                "header_row": header_row_number,
                "columns": len(final_header),
                "rows": len(data_rows),
                "tap_schema": final_layout["schema"],
            }

            return (
                final_header,
                data_rows,
                meta,
            )

    examples = "\n".join(
        f'  - tab="{title}", row={row_no}, header={header}'
        for title, row_no, header in inspected[:5]
    )

    raise ValueError(
        f"{label}: tidak menemukan header TAP yang kompatibel.\n"
        "Script menerima alias:\n"
        "  LIVE ID / Livestream room ID\n"
        "  Product ID\n"
        "  LIVE start time / LIVE time info\n"
        "  Creator LIVE-attributed GMV\n"
        "  LIVE-attributed orders / Creator LIVE-attributed orders\n"
        "  LIVE-attributed items sold / Creator-attributed items sold\n"
        "  Product name OPTIONAL\n"
        + (
            "\nHeader yang sempat terdeteksi:\n"
            + examples
            if examples
            else ""
        )
    )


# =============================================================================
# BW PRODUCT ID MAPPING
# =============================================================================

def build_bw_product_mapping(
    header: list[str],
    rows: list[dict[str, object]],
):
    """
    Build Product ID -> Brand + Sub Product.

    Source classification untuk BW HANYA spreadsheet mapping ini.
    Product name tidak dipakai untuk menentukan Brand/Sub Product.

    Duplicate Product ID diperbolehkan hanya jika Brand + Sub Product identik.
    Jika satu Product ID punya mapping berbeda, script berhenti agar tidak ambigu.
    """
    validate_headers(
        header,
        BW_MAPPING_REQUIRED,
        "BW PRODUCT MAPPING",
    )

    h_pid = resolve_header(header, "Product ID")
    h_brand = resolve_header(header, "Brand")
    h_sub = resolve_header(header, "Sub Product")

    mapping: dict[str, dict[str, str]] = {}
    conflicts = defaultdict(set)

    stats = {
        "raw_rows": len(rows),
        "mapped_product_ids": 0,
        "blank_product_id": 0,
        "blank_brand": 0,
        "blank_sub_product": 0,
        "duplicate_same": 0,
    }

    for row in rows:
        pid = normalize_id(row.get(h_pid, ""))
        if not pid:
            stats["blank_product_id"] += 1
            continue

        brand = str(row.get(h_brand, "") or "").strip()
        sub_product = str(row.get(h_sub, "") or "").strip()

        if not brand:
            brand = MISSING_BRAND
            stats["blank_brand"] += 1

        if not sub_product:
            sub_product = MISSING_SUB_PRODUCT
            stats["blank_sub_product"] += 1

        candidate = {
            "brand": brand,
            "sub_product": sub_product,
        }

        if pid in mapping:
            current = mapping[pid]
            if current == candidate:
                stats["duplicate_same"] += 1
                continue

            conflicts[pid].add(
                (current["brand"], current["sub_product"])
            )
            conflicts[pid].add((brand, sub_product))
            continue

        mapping[pid] = candidate

    if conflicts:
        examples = []
        for pid, values in list(conflicts.items())[:20]:
            rendered = sorted(
                f"{brand} | {sub_product}"
                for brand, sub_product in values
            )
            examples.append(f"{pid}: {rendered}")

        raise ValueError(
            "BW PRODUCT MAPPING CONFLICT. "
            "Satu Product ID punya Brand/Sub Product berbeda:\n  "
            + "\n  ".join(examples)
        )

    stats["mapped_product_ids"] = len(mapping)
    stats["brands"] = len({rec["brand"] for rec in mapping.values()})
    stats["sub_products"] = len(
        {rec["sub_product"] for rec in mapping.values()}
    )

    return mapping, stats


def configure_brand_output_order(
    bw_product_mapping: dict[str, dict[str, str]],
):
    """Gunakan semua Brand mapping BW sebagai universe tab/summary."""
    global BRAND_OUTPUT_ORDER

    brands = sorted(
        {
            str(rec.get("brand", "") or "").strip()
            for rec in bw_product_mapping.values()
            if str(rec.get("brand", "") or "").strip()
            and str(rec.get("brand", "") or "").strip() != MISSING_BRAND
        },
        key=lambda value: value.casefold(),
    )

    BRAND_OUTPUT_ORDER = brands + [MISSING_BRAND]


def lookup_bw_classification(
    product_id: object,
    bw_product_mapping: dict[str, dict[str, str]],
) -> tuple[str, str, bool]:
    pid = normalize_id(product_id)
    mapped = bw_product_mapping.get(pid)

    if mapped is None:
        return MISSING_BRAND, MISSING_SUB_PRODUCT, False

    brand = str(mapped.get("brand", "") or "").strip() or MISSING_BRAND
    sub_product = (
        str(mapped.get("sub_product", "") or "").strip()
        or MISSING_SUB_PRODUCT
    )

    return brand, sub_product, True


def print_bw_mapping_summary(stats: dict):
    print("\nBW Product ID mapping:")
    print(f"  Raw mapping row          : {stats['raw_rows']:,}")
    print(f"  Unique Product ID        : {stats['mapped_product_ids']:,}")
    print(f"  Brand                    : {stats['brands']:,}")
    print(f"  Sub Product              : {stats['sub_products']:,}")
    print(f"  Blank Product ID         : {stats['blank_product_id']:,}")
    print(f"  Blank Brand -> MISSING   : {stats['blank_brand']:,}")
    print(f"  Blank Sub Product -> MISS: {stats['blank_sub_product']:,}")
    print(f"  Duplicate same mapping   : {stats['duplicate_same']:,}")


def most_common_text(counter: Counter) -> str:
    """
    Return text dengan frekuensi tertinggi dari Counter.
    Dipakai untuk memilih Product name paling umum dari TAP/AFF
    ketika satu Product ID punya beberapa variasi nama.
    """
    return counter.most_common(1)[0][0] if counter else ""


# =============================================================================
# TAP LS MAPPING
# =============================================================================

def build_tap_ls_mapping(
    datasets: list[
        tuple[
            str,
            list[str],
            list[dict[str, object]],
        ]
    ],
    metric_start: date | None,
    metric_end: date | None,
):
    """
    Gabungkan semua TAP LS source, walaupun schema antar-source berbeda.

    Key:
        LIVE ID / Livestream room ID + Product ID

    Metric period:
        - schema baru -> LIVE start time
        - schema lama -> START dari LIVE time info

    Metric alias:
        Orders:
            LIVE-attributed orders
            Creator LIVE-attributed orders
        Items:
            LIVE-attributed items sold
            Creator-attributed items sold

    Product name optional.
    """
    mapping = defaultdict(
        lambda: {
            "gmv": 0.0,
            "orders": 0,
            "items": 0,
            "rows": 0,
            "product_names": Counter(),
        }
    )

    total_raw = 0
    invalid_key = 0
    invalid_live_start = 0
    period_skipped = 0
    duplicate_rows = 0
    seen_fingerprints = set()
    schema_counter = Counter()

    for label, header, rows in datasets:
        layout = resolve_tap_layout(
            header,
            label=label,
        )
        schema_counter[
            str(layout["schema"])
        ] += 1

        h_live = str(layout["live_id"])
        h_pid = str(layout["product_id"])
        h_name = layout["product_name"]
        h_gmv = str(layout["gmv"])
        h_orders = str(layout["orders"])
        h_items = str(layout["items"])

        total_raw += len(rows)

        for row in rows:
            live_id = normalize_id(
                row.get(h_live, "")
            )
            product_id = normalize_id(
                row.get(h_pid, "")
            )

            if not live_id or not product_id:
                invalid_key += 1
                continue

            try:
                live_start_dt = parse_tap_start_datetime(
                    row,
                    layout,
                )
            except ValueError:
                invalid_live_start += 1
                continue

            live_start_date = (
                live_start_dt.date()
                if live_start_dt is not None
                else None
            )

            if metric_start or metric_end:
                if not in_period(
                    live_start_date,
                    metric_start,
                    metric_end,
                ):
                    period_skipped += 1
                    continue

            gmv = parse_number(
                row.get(h_gmv, "")
            )
            orders = parse_int(
                row.get(h_orders, "")
            )
            items = parse_int(
                row.get(h_items, "")
            )

            # Logical fingerprint, bukan full-row fingerprint.
            # Ini memungkinkan dedupe walaupun source 1 schema lama
            # dan source 2 schema baru.
            fingerprint = (
                live_id,
                product_id,
                (
                    live_start_dt.strftime(
                        "%Y-%m-%d %H:%M:%S"
                    )
                    if live_start_dt
                    else ""
                ),
                clean_numeric(gmv),
                orders,
                items,
            )

            key = (live_id, product_id)

            product_name = ""
            if h_name:
                product_name = str(
                    row.get(str(h_name), "") or ""
                ).strip()

            if fingerprint in seen_fingerprints:
                duplicate_rows += 1

                # Kalau duplicate source kedua punya product name
                # sedangkan source pertama tidak, tetap simpan namanya
                # tanpa menambah metric.
                if (
                    product_name
                    and key in mapping
                ):
                    mapping[key][
                        "product_names"
                    ][product_name] += 1

                continue

            seen_fingerprints.add(fingerprint)

            rec = mapping[key]
            rec["gmv"] += gmv
            rec["orders"] += orders
            rec["items"] += items
            rec["rows"] += 1

            if product_name:
                rec[
                    "product_names"
                ][product_name] += 1

    finalized = {}

    for key, rec in mapping.items():
        finalized[key] = {
            "gmv": rec["gmv"],
            "orders": rec["orders"],
            "items": rec["items"],
            "rows": rec["rows"],
            "product_name": most_common_text(
                rec["product_names"]
            ),
        }

    stats = {
        "sources": len(datasets),
        "raw_rows": total_raw,
        "pairs": len(finalized),
        "invalid_key": invalid_key,
        "invalid_live_start": invalid_live_start,
        "period_skipped": period_skipped,
        "duplicate_rows": duplicate_rows,
        "schemas": dict(schema_counter),
    }

    return finalized, stats


def build_tap_ls_live_id_set(
    datasets: list[
        tuple[
            str,
            list[str],
            list[dict[str, object]],
        ]
    ],
    metric_start: date | None,
    metric_end: date | None,
):
    """
    Union LIVE ID dari semua TAP source.

    Source boleh schema baru/lama/hybrid.
    Period selalu berdasarkan start live.
    """
    live_ids = set()

    stats = {
        "sources": len(datasets),
        "raw_rows": 0,
        "included_rows": 0,
        "blank_or_invalid_live_id": 0,
        "invalid_live_start": 0,
        "period_skipped": 0,
        "duplicate_rows": 0,
        "schemas": {},
    }

    schema_counter = Counter()
    seen_fingerprints = set()

    for label, header, rows in datasets:
        layout = resolve_tap_layout(
            header,
            label=label,
        )

        schema_counter[
            str(layout["schema"])
        ] += 1

        h_live = str(layout["live_id"])
        h_pid = str(layout["product_id"])
        h_gmv = str(layout["gmv"])
        h_orders = str(layout["orders"])
        h_items = str(layout["items"])

        stats["raw_rows"] += len(rows)

        for row in rows:
            live_id = normalize_id(
                row.get(h_live, "")
            )

            if not live_id:
                stats[
                    "blank_or_invalid_live_id"
                ] += 1
                continue

            try:
                live_start_dt = parse_tap_start_datetime(
                    row,
                    layout,
                )
            except ValueError:
                stats["invalid_live_start"] += 1
                continue

            live_start_date = (
                live_start_dt.date()
                if live_start_dt is not None
                else None
            )

            if metric_start or metric_end:
                if not in_period(
                    live_start_date,
                    metric_start,
                    metric_end,
                ):
                    stats["period_skipped"] += 1
                    continue

            product_id = normalize_id(
                row.get(h_pid, "")
            )

            fingerprint = (
                live_id,
                product_id,
                (
                    live_start_dt.strftime(
                        "%Y-%m-%d %H:%M:%S"
                    )
                    if live_start_dt
                    else ""
                ),
                clean_numeric(
                    parse_number(
                        row.get(h_gmv, "")
                    )
                ),
                parse_int(
                    row.get(h_orders, "")
                ),
                parse_int(
                    row.get(h_items, "")
                ),
            )

            if fingerprint in seen_fingerprints:
                stats["duplicate_rows"] += 1
                continue

            seen_fingerprints.add(fingerprint)

            stats["included_rows"] += 1
            live_ids.add(live_id)

    stats["schemas"] = dict(
        schema_counter
    )

    return live_ids, stats


# =============================================================================
# AFF LIVESTREAM MAPPING
# =============================================================================

def build_aff_ls_mapping(
    datasets: list[tuple[str, list[str], list[dict[str, object]]]],
    metric_start: date | None,
    metric_end: date | None,
):
    mapping = defaultdict(
        lambda: {
            "gmv": 0.0,
            "items": 0,
            "rows": 0,
            "product_names": Counter(),
        }
    )

    total_raw = 0
    non_livestream = 0
    invalid_key = 0
    invalid_date = 0
    period_skipped = 0
    duplicate_rows = 0
    seen_fingerprints = set()

    for label, header, rows in datasets:
        validate_headers(header, AFF_REQUIRED, label)

        h_order = resolve_header(header, "Order ID")
        h_pid = resolve_header(header, "Product ID")
        h_name = resolve_first_header(header, ["Product Name", "Product name"])
        h_type = resolve_header(header, "Content Type")
        h_content = resolve_header(header, "Content ID")
        h_gmv = resolve_header(header, "Est. Commission Base")
        h_qty = resolve_header(header, "Quantity")
        h_time = resolve_header(header, "Time Created")

        total_raw += len(rows)

        for row in rows:
            content_type = str(row.get(h_type, "") or "").strip().lower()
            if content_type != "livestream":
                non_livestream += 1
                continue

            live_id = normalize_id(row.get(h_content, ""))
            product_id = normalize_id(row.get(h_pid, ""))
            _order_id_raw = str(row.get(h_order, "") or "").strip()

            if not live_id or not product_id:
                invalid_key += 1
                continue

            if metric_start or metric_end:
                try:
                    dt = parse_aff_datetime(row.get(h_time, ""))
                except ValueError:
                    invalid_date += 1
                    continue

                d = dt.date() if dt else None
                if not in_period(d, metric_start, metric_end):
                    period_skipped += 1
                    continue

            fingerprint = tuple(
                str(row.get(name, "") or "").strip()
                for name in header
            )

            if fingerprint in seen_fingerprints:
                duplicate_rows += 1
                continue

            seen_fingerprints.add(fingerprint)

            key = (live_id, product_id)
            rec = mapping[key]
            rec["gmv"] += parse_number(row.get(h_gmv, ""))
            rec["items"] += parse_int(row.get(h_qty, ""))
            rec["rows"] += 1

            if h_name:
                product_name = str(row.get(h_name, "") or "").strip()
                if product_name:
                    rec["product_names"][product_name] += 1

    finalized = {}
    for key, rec in mapping.items():
        finalized[key] = {
            "gmv": rec["gmv"],
            "orders": rec["rows"],
            "items": rec["items"],
            "rows": rec["rows"],
            "product_name": most_common_text(rec["product_names"]),
        }

    stats = {
        "raw_rows": total_raw,
        "pairs": len(finalized),
        "non_livestream": non_livestream,
        "invalid_key": invalid_key,
        "invalid_date": invalid_date,
        "period_skipped": period_skipped,
        "duplicate_rows": duplicate_rows,
    }
    return finalized, stats


def build_master_mapping(tap_map: dict, aff_map: dict):
    master = {}
    counts = {"BOTH": 0, "TAP_ONLY": 0, "AFF_ONLY": 0}

    for key in set(tap_map) | set(aff_map):
        tap = tap_map.get(key)
        aff = aff_map.get(key)

        if tap is not None and aff is not None:
            status = "BOTH"
        elif tap is not None:
            status = "TAP_ONLY"
        else:
            status = "AFF_ONLY"

        counts[status] += 1

        # Model A: TAP metric prioritas, AFF fallback.
        chosen = tap if tap is not None else aff

        product_name = ""
        if tap is not None:
            product_name = str(tap.get("product_name", "") or "").strip()
        if not product_name and aff is not None:
            product_name = str(aff.get("product_name", "") or "").strip()

        master[key] = {
            "mapped_gmv": float(chosen["gmv"]),
            "mapped_orders": int(chosen["orders"]),
            "mapped_items": int(chosen["items"]),
            "metric_source": "TAP" if tap is not None else "AFF",
            "match_status": status,
            "product_name": product_name,
            "tap_gmv": float(tap["gmv"]) if tap is not None else 0.0,
            "tap_orders": int(tap["orders"]) if tap is not None else 0,
            "tap_items": int(tap["items"]) if tap is not None else 0,
            "aff_gmv": float(aff["gmv"]) if aff is not None else 0.0,
            "aff_orders": int(aff["orders"]) if aff is not None else 0,
            "aff_items": int(aff["items"]) if aff is not None else 0,
        }

    return master, counts


def build_pid_product_name_lookup(
    tap_map: dict,
    aff_map: dict,
) -> dict[str, str]:
    tap_names = defaultdict(Counter)
    aff_names = defaultdict(Counter)

    for (_live_id, pid), rec in tap_map.items():
        name = str(rec.get("product_name", "") or "").strip()
        if name:
            tap_names[pid][name] += 1

    for (_live_id, pid), rec in aff_map.items():
        name = str(rec.get("product_name", "") or "").strip()
        if name:
            aff_names[pid][name] += 1

    lookup = {}
    for pid in set(tap_names) | set(aff_names):
        if tap_names.get(pid):
            lookup[pid] = most_common_text(tap_names[pid])
        elif aff_names.get(pid):
            lookup[pid] = most_common_text(aff_names[pid])

    return lookup


def build_tap_brand_pair_set_from_mapping(
    tap_map: dict,
    bw_product_mapping: dict[str, dict[str, str]],
) -> set[tuple[str, str]]:
    """Build existence key TAP pada grain LIVE ID + Brand.

    TAP map di file ini sudah mengikuti METRIC period, jadi existence period
    tetap sama seperti logic NOT IN TAP sebelumnya; yang berubah hanya grain
    pembanding dari LIVE ID menjadi LIVE ID + Brand.
    """
    pairs: set[tuple[str, str]] = set()
    for (live_id, pid) in tap_map:
        live_id = normalize_id(live_id)
        if not live_id:
            continue
        brand, _sub_product, _mapped = lookup_bw_classification(
            pid, bw_product_mapping
        )
        pairs.add((live_id, brand))
    return pairs


# =============================================================================
# SC LIVE ENRICHMENT
# =============================================================================

def find_sc_source_column(
    sc_header: list[str],
    output_name: str,
) -> str | None:
    return resolve_first_header(sc_header, [output_name])


def clean_sc_value(output_name: str, value: object):
    if output_name in INTEGER_COLUMNS:
        return parse_int(value)

    if output_name in NUMERIC_COLUMNS:
        return clean_numeric(parse_number(value))

    if output_name in PERCENT_COLUMNS:
        return clean_numeric(parse_percent(value))

    if output_name == "LIVE ID":
        return normalize_id(value)

    if output_name in {"LIVE start time", "LIVE end time"}:
        return normalize_sc_live_datetime(value)

    return "" if value is None else str(value).strip()


def enrich_sc_live(
    header: list[str],
    rows: list[dict[str, object]],
    master_map: dict,
    aff_map: dict,
    pid_product_name_lookup: dict[str, str],
    bw_product_mapping: dict[str, dict[str, str]],
):
    """
    Explode SC LIVE menjadi 1 output row per Product ID.

    Commerce Model A:
    - Single Product ID:
        preserve SC original GMV / Orders / Items.
    - Multi Product ID:
        exact TAP LIVE ID + PID;
        AFF exact fallback bila TAP tidak ada;
        missing pair = 0.

    Brand + Sub Product SELALU berasal dari Product ID mapping BW.
    """
    validate_headers(header, SC_REQUIRED, "SC LIVE")

    derived_columns = {
        "Product name",
        "Brand",
        "Sub Product",
        "GMV Affiliate Order (Est. Commission Base)",
        "Jumlah Order Affiliate",
    }

    source_cols = {
        out_name: find_sc_source_column(header, out_name)
        for out_name in OUTPUT_COLUMNS
        if out_name not in derived_columns
    }

    required_output_sources = [
        name
        for name in OUTPUT_COLUMNS
        if name not in derived_columns
    ]

    missing_sources = [
        name
        for name in required_output_sources
        if source_cols.get(name) is None
    ]

    if missing_sources:
        raise ValueError(
            "SC LIVE tidak memiliki source untuk kolom output: "
            + ", ".join(missing_sources)
        )

    h_live = resolve_header(header, "LIVE ID")
    h_pid = resolve_header(header, "Product ID")
    h_sc_name = resolve_first_header(header, ["Product name", "Product Name"])
    h_sc_gmv = resolve_header(header, "Creator LIVE-attributed GMV")
    h_sc_orders = resolve_header(header, "LIVE-attributed orders")
    h_sc_items = resolve_header(header, "LIVE-attributed items sold")

    output_rows = []

    missing_details = defaultdict(
        lambda: {
            "count": 0,
            "example_title": "",
            "source_product_count": 0,
        }
    )

    stats = {
        "raw_rows": len(rows),
        "exploded_rows": 0,
        "multi_pid_source_rows": 0,
        "blank_pid_source_rows": 0,
        "multi_pid_mapped_tap": 0,
        "multi_pid_mapped_aff_only": 0,
        "single_pid_sc_preserved": 0,
        "unmapped_multi_pid": 0,
        "invalid_live_id": 0,
        "bw_product_mapped": 0,
        "bw_product_missing": 0,
    }

    original_totals = {
        "gmv": 0.0,
        "orders": 0,
        "items": 0,
    }

    enriched_totals = {
        "gmv": 0.0,
        "orders": 0,
        "items": 0,
    }

    brand_counter = Counter()

    for row in rows:
        live_id = normalize_id(row.get(h_live, ""))

        if not live_id:
            stats["invalid_live_id"] += 1
            continue

        product_ids = split_product_ids(row.get(h_pid, ""))

        if not product_ids:
            stats["blank_pid_source_rows"] += 1
            continue

        if len(product_ids) > 1:
            stats["multi_pid_source_rows"] += 1

        sc_original_gmv = parse_number(row.get(h_sc_gmv, ""))
        sc_original_orders = parse_int(row.get(h_sc_orders, ""))
        sc_original_items = parse_int(row.get(h_sc_items, ""))

        original_totals["gmv"] += sc_original_gmv
        original_totals["orders"] += sc_original_orders
        original_totals["items"] += sc_original_items

        sc_product_name = (
            str(row.get(h_sc_name, "") or "").strip()
            if h_sc_name
            else ""
        )

        source_title_col = source_cols.get("LIVE title")
        example_title = (
            str(row.get(source_title_col, "") or "").strip()
            if source_title_col
            else ""
        )

        for pid in product_ids:
            key = (live_id, pid)
            master = master_map.get(key)
            aff = aff_map.get(key)

            out = {}

            for out_name in OUTPUT_COLUMNS:
                if out_name in derived_columns:
                    continue

                if out_name == "Product ID":
                    out[out_name] = pid
                    continue

                source_name = source_cols.get(out_name)
                raw_value = row.get(source_name, "") if source_name else ""
                out[out_name] = clean_sc_value(out_name, raw_value)

            out["LIVE ID"] = live_id

            mapped_product_name = ""

            if master is not None:
                mapped_product_name = str(
                    master.get("product_name", "") or ""
                ).strip()

            if not mapped_product_name:
                mapped_product_name = str(
                    pid_product_name_lookup.get(pid, "") or ""
                ).strip()

            if not mapped_product_name:
                mapped_product_name = sc_product_name

            brand, sub_product, bw_mapped = lookup_bw_classification(
                pid,
                bw_product_mapping,
            )

            if bw_mapped:
                stats["bw_product_mapped"] += 1
            else:
                stats["bw_product_missing"] += 1

            out["Product name"] = mapped_product_name
            out["Brand"] = brand
            out["Sub Product"] = sub_product

            # -------------------------------------------------------------
            # MODEL A
            # -------------------------------------------------------------
            if len(product_ids) == 1:
                mapped_gmv = sc_original_gmv
                mapped_orders = sc_original_orders
                mapped_items = sc_original_items
                stats["single_pid_sc_preserved"] += 1

            elif master is not None:
                mapped_gmv = float(master["mapped_gmv"])
                mapped_orders = int(master["mapped_orders"])
                mapped_items = int(master["mapped_items"])

                if master["metric_source"] == "TAP":
                    stats["multi_pid_mapped_tap"] += 1
                else:
                    stats["multi_pid_mapped_aff_only"] += 1

            else:
                mapped_gmv = 0.0
                mapped_orders = 0
                mapped_items = 0
                stats["unmapped_multi_pid"] += 1

                detail = missing_details[key]
                detail["count"] += 1
                detail["source_product_count"] = len(product_ids)

                if not detail["example_title"]:
                    detail["example_title"] = example_title

            mapped_aov = (
                mapped_gmv / mapped_orders
                if mapped_orders
                else 0.0
            )

            out["Creator LIVE-attributed GMV"] = clean_numeric(
                mapped_gmv
            )
            out["LIVE-attributed orders"] = mapped_orders
            out["AOV"] = clean_numeric(mapped_aov)
            out["LIVE-attributed items sold"] = mapped_items

            # AFF supplemental tetap exact LIVE ID + Product ID.
            out[
                "GMV Affiliate Order (Est. Commission Base)"
            ] = clean_numeric(
                float(aff["gmv"])
                if aff is not None
                else 0.0
            )

            out["Jumlah Order Affiliate"] = (
                int(aff["orders"])
                if aff is not None
                else 0
            )

            out["_BRAND"] = brand
            out["_SUB_PRODUCT"] = sub_product

            # Exact mapping QC tetap berdasarkan exact pair TAP/AFF.
            if master is not None:
                out["_EXACT_MAPPING_STATUS"] = master["match_status"]
                out["_EXACT_MAPPING_SOURCE"] = master["metric_source"]
            else:
                out["_EXACT_MAPPING_STATUS"] = "MISSING"
                out["_EXACT_MAPPING_SOURCE"] = "MISSING"

            if len(product_ids) == 1:
                out["_MAPPING_SOURCE"] = "SC_SINGLE_PID_PRESERVED"
            elif master is not None:
                out["_MAPPING_SOURCE"] = master["metric_source"]
            else:
                out["_MAPPING_SOURCE"] = "MISSING"

            output_rows.append(out)

            stats["exploded_rows"] += 1
            brand_counter[brand] += 1

            enriched_totals["gmv"] += mapped_gmv
            enriched_totals["orders"] += mapped_orders
            enriched_totals["items"] += mapped_items

    return (
        output_rows,
        stats,
        original_totals,
        enriched_totals,
        brand_counter,
        missing_details,
    )


def filter_output_range(
    rows: list[dict[str, object]],
    range_start: date | None,
    range_end: date | None,
):
    if range_start is None and range_end is None:
        return list(rows), 0

    filtered = []
    invalid_date = 0

    for row in rows:
        d = parse_output_live_date(
            row.get("LIVE start time", "")
        )

        if d is None:
            invalid_date += 1
            continue

        if in_period(d, range_start, range_end):
            filtered.append(row)

    return filtered, invalid_date


def filter_sc_ls_not_in_tap(
    rows: list[dict[str, object]],
    tap_live_brand_pairs: set[tuple[str, str]],
):
    """NOT IN TAP pada grain LIVE ID + Brand, bukan LIVE ID saja."""
    result = []

    for row in rows:
        live_id = normalize_id(row.get("LIVE ID", ""))
        brand = str(
            row.get("_BRAND", row.get("Brand", MISSING_BRAND))
            or MISSING_BRAND
        ).strip()

        if live_id and (live_id, brand) not in tap_live_brand_pairs:
            result.append(row)

    return result


def build_brand_unique_rows(
    rows: list[dict[str, object]],
) -> list[dict[str, object]]:
    """
    Ubah exploded SC LIVE menjadi grain unik LIVE ID + Brand.

    Rule:
    - Representative Product ID = PID pertama yang muncul untuk LIVE ID + Brand.
    - Commerce metric = SUM per UNIQUE Product ID dalam LIVE ID + Brand.
    - PID yang sama tidak boleh menyumbang metric dua kali.
    - Product name dan Sub Product mengikuti representative PID.
    - Metric level-LIVE lain mengikuti row representative dan bersifat non-additive
      lintas Brand.
    """
    grouped: dict[tuple[str, str], dict[str, object]] = {}
    key_order: list[tuple[str, str]] = []

    for row in rows:
        live_id = normalize_id(row.get("LIVE ID", ""))
        if not live_id:
            continue

        brand = str(row.get("_BRAND", row.get("Brand", MISSING_BRAND)) or MISSING_BRAND).strip()
        if not brand:
            brand = MISSING_BRAND

        key = (live_id, brand)
        pid = normalize_id(row.get("Product ID", ""))

        if key not in grouped:
            representative = dict(row)
            grouped[key] = {
                "row": representative,
                "seen_pids": set(),
                "gmv": 0.0,
                "orders": 0,
                "items": 0,
                "aff_gmv": 0.0,
                "aff_orders": 0,
                "sub_products": [],
            }
            key_order.append(key)

        rec = grouped[key]

        sub_product = str(row.get("Sub Product", "") or "").strip()
        if sub_product and sub_product not in rec["sub_products"]:
            rec["sub_products"].append(sub_product)

        # Metric brand harus dijumlah dari PID unik, bukan jumlah exploded rows.
        if pid and pid not in rec["seen_pids"]:
            rec["seen_pids"].add(pid)
            rec["gmv"] += parse_number(row.get("Creator LIVE-attributed GMV", 0))
            rec["orders"] += parse_int(row.get("LIVE-attributed orders", 0))
            rec["items"] += parse_int(row.get("LIVE-attributed items sold", 0))
            rec["aff_gmv"] += parse_number(
                row.get("GMV Affiliate Order (Est. Commission Base)", 0)
            )
            rec["aff_orders"] += parse_int(row.get("Jumlah Order Affiliate", 0))

    result: list[dict[str, object]] = []

    for key in key_order:
        rec = grouped[key]
        out = rec["row"]
        gmv = float(rec["gmv"])
        orders = int(rec["orders"])
        items = int(rec["items"])

        out["Creator LIVE-attributed GMV"] = clean_numeric(gmv)
        out["LIVE-attributed orders"] = orders
        out["LIVE-attributed items sold"] = items
        out["AOV"] = clean_numeric(gmv / orders if orders else 0.0)
        out["GMV Affiliate Order (Est. Commission Base)"] = clean_numeric(
            float(rec["aff_gmv"])
        )
        out["Jumlah Order Affiliate"] = int(rec["aff_orders"])
        out["_BRAND"] = key[1]
        out["_MAPPING_SOURCE"] = "BRAND_SUM_UNIQUE_PID"
        out["_BRAND_PID_COUNT"] = len(rec["seen_pids"])
        out["_BRAND_PRODUCT_IDS"] = ",".join(sorted(rec["seen_pids"]))
        out["_BRAND_SUB_PRODUCTS"] = ",".join(rec["sub_products"])
        result.append(out)

    # Hard validation: final grain wajib benar-benar unik.
    seen_keys = set()
    for row in result:
        key = (
            normalize_id(row.get("LIVE ID", "")),
            str(row.get("_BRAND", MISSING_BRAND) or MISSING_BRAND),
        )
        if key in seen_keys:
            raise RuntimeError(f"Duplicate BRAND UNIQUE key terdeteksi: {key}")
        seen_keys.add(key)

    return result


def build_brand_tables(
    rows: list[dict[str, object]],
):
    grouped = {
        brand: []
        for brand in BRAND_OUTPUT_ORDER
    }

    for row in rows:
        brand = str(
            row.get("_BRAND", MISSING_BRAND)
            or MISSING_BRAND
        )

        if brand not in grouped:
            brand = MISSING_BRAND

        grouped[brand].append(row)

    return grouped



# =============================================================================
# MAPPING QC PER LIVE ID
# =============================================================================

def build_mapping_qc_values(
    rows: list[dict[str, object]],
) -> list[list[object]]:
    """
    Ringkas coverage exact mapping TAP LS + AFF per LIVE ID.

    MAPPED berarti exact pair:
        LIVE ID + Product ID
    tersedia di MASTER mapping (TAP LS dan/atau AFF Livestream).

    Single-PID yang metric-nya preserve SC tetap dinilai mapped/unmapped
    berdasarkan keberadaan exact pair, bukan berdasarkan metric source final.
    """
    grouped = {}

    for row in rows:
        live_id = normalize_id(row.get("LIVE ID", ""))
        if not live_id:
            continue

        rec = grouped.setdefault(
            live_id,
            {
                "title": str(row.get("LIVE title", "") or "").strip(),
                "start_time": str(row.get("LIVE start time", "") or "").strip(),
                "creator": str(row.get("Creator name", "") or "").strip(),
                "all_pids": set(),
                "mapped_pids": set(),
                "unmapped_pids": set(),
                "both_pids": set(),
                "tap_only_pids": set(),
                "aff_only_pids": set(),
            },
        )

        # Isi metadata pertama yang nonblank.
        if not rec["title"]:
            rec["title"] = str(row.get("LIVE title", "") or "").strip()
        if not rec["start_time"]:
            rec["start_time"] = str(row.get("LIVE start time", "") or "").strip()
        if not rec["creator"]:
            rec["creator"] = str(row.get("Creator name", "") or "").strip()

        pid = normalize_id(row.get("Product ID", ""))
        if not pid:
            continue

        rec["all_pids"].add(pid)

        status = str(
            row.get("_EXACT_MAPPING_STATUS", "MISSING")
            or "MISSING"
        ).strip().upper()

        if status == "BOTH":
            rec["mapped_pids"].add(pid)
            rec["both_pids"].add(pid)
        elif status == "TAP_ONLY":
            rec["mapped_pids"].add(pid)
            rec["tap_only_pids"].add(pid)
        elif status == "AFF_ONLY":
            rec["mapped_pids"].add(pid)
            rec["aff_only_pids"].add(pid)
        else:
            rec["unmapped_pids"].add(pid)

    header = [
        "LIVE ID",
        "LIVE title",
        "LIVE start time",
        "Creator name",
        "PID Count",
        "Mapped PID Count",
        "Unmapped PID Count",
        "Mapping Coverage",
        "BOTH PID Count",
        "TAP Only PID Count",
        "AFF Only PID Count",
        "Mapped Product IDs",
        "Unmapped Product IDs",
        "BOTH Product IDs",
        "TAP Only Product IDs",
        "AFF Only Product IDs",
        "Mapping Status",
    ]

    body = []

    for live_id, rec in grouped.items():
        all_pids = sorted(rec["all_pids"])
        mapped_pids = sorted(rec["mapped_pids"])
        unmapped_pids = sorted(rec["unmapped_pids"])
        both_pids = sorted(rec["both_pids"])
        tap_only_pids = sorted(rec["tap_only_pids"])
        aff_only_pids = sorted(rec["aff_only_pids"])

        total_count = len(all_pids)
        mapped_count = len(mapped_pids)
        unmapped_count = len(unmapped_pids)

        coverage = (
            mapped_count / total_count
            if total_count
            else 0.0
        )

        if total_count == 0:
            mapping_status = "NO PID"
        elif mapped_count == total_count:
            mapping_status = "FULLY MAPPED"
        elif mapped_count == 0:
            mapping_status = "NO MAPPING"
        else:
            mapping_status = "PARTIAL MAPPING"

        body.append([
            live_id,
            rec["title"],
            rec["start_time"],
            rec["creator"],
            total_count,
            mapped_count,
            unmapped_count,
            clean_numeric(coverage),
            len(both_pids),
            len(tap_only_pids),
            len(aff_only_pids),
            ",".join(mapped_pids),
            ",".join(unmapped_pids),
            ",".join(both_pids),
            ",".join(tap_only_pids),
            ",".join(aff_only_pids),
            mapping_status,
        ])

    # Prioritaskan LS yang paling bermasalah di atas:
    # unmapped paling banyak -> PID total paling banyak -> LIVE ID.
    body.sort(
        key=lambda row: (
            -int(row[6]),
            -int(row[4]),
            str(row[0]),
        )
    )

    return [header] + body


def summarize_mapping_qc(
    rows: list[dict[str, object]],
):
    """Summary agregat coverage mapping exact pair untuk console QC."""
    values = build_mapping_qc_values(rows)
    body = values[1:]

    summary = {
        "live_ids": len(body),
        "fully_mapped": 0,
        "partial_mapping": 0,
        "no_mapping": 0,
        "no_pid": 0,
        "pid_total": 0,
        "pid_mapped": 0,
        "pid_unmapped": 0,
    }

    for row in body:
        summary["pid_total"] += int(row[4])
        summary["pid_mapped"] += int(row[5])
        summary["pid_unmapped"] += int(row[6])

        status = str(row[16])
        if status == "FULLY MAPPED":
            summary["fully_mapped"] += 1
        elif status == "PARTIAL MAPPING":
            summary["partial_mapping"] += 1
        elif status == "NO MAPPING":
            summary["no_mapping"] += 1
        else:
            summary["no_pid"] += 1

    summary["pid_coverage"] = (
        summary["pid_mapped"] / summary["pid_total"]
        if summary["pid_total"]
        else 0.0
    )

    return summary


# =============================================================================
# OUTPUT TABLE BUILDERS
# =============================================================================

def data_values(
    rows: list[dict[str, object]],
) -> list[list[object]]:
    return [OUTPUT_COLUMNS] + [
        [row.get(col, "") for col in OUTPUT_COLUMNS]
        for row in rows
    ]


def build_brand_summary(
    all_rows,
    range_rows,
):
    all_counter = Counter(
        str(row.get("_BRAND", MISSING_BRAND))
        for row in all_rows
    )

    range_counter = Counter(
        str(row.get("_BRAND", MISSING_BRAND))
        for row in range_rows
    )

    values = [[
        "Brand",
        "Rows ALL",
        "Rows RANGE",
        "GMV ALL",
        "GMV RANGE",
    ]]

    for brand in BRAND_OUTPUT_ORDER:
        gmv_all = sum(
            parse_number(
                row.get(
                    "Creator LIVE-attributed GMV",
                    0,
                )
            )
            for row in all_rows
            if row.get("_BRAND") == brand
        )

        gmv_range = sum(
            parse_number(
                row.get(
                    "Creator LIVE-attributed GMV",
                    0,
                )
            )
            for row in range_rows
            if row.get("_BRAND") == brand
        )

        values.append([
            brand,
            all_counter.get(brand, 0),
            range_counter.get(brand, 0),
            clean_numeric(gmv_all),
            clean_numeric(gmv_range),
        ])

    return values


def build_missing_values(
    missing_details,
):
    values = [[
        "LIVE ID",
        "Product ID",
        "Occurrences",
        "Product Count in Source Row",
        "Example LIVE Title",
        "Reason",
    ]]

    for (
        live_id,
        product_id,
    ), detail in sorted(missing_details.items()):
        values.append([
            live_id,
            product_id,
            detail["count"],
            detail["source_product_count"],
            detail["example_title"],
            (
                "LIVE ID + Product ID tidak ditemukan "
                "di master TAP LS + AFF Livestream"
            ),
        ])

    return values


def build_output_tables(
    all_rows,
    range_rows,
    sc_not_in_tap_all_rows,
    sc_not_in_tap_range_rows,
    missing_details,
    write_brand_tabs,
):
    tables = {
        SHEET_ALL: data_values(all_rows),
        SHEET_RANGE: data_values(range_rows),
        SC_NOT_IN_TAP_ALL_SHEET: data_values(
            sc_not_in_tap_all_rows
        ),
        SC_NOT_IN_TAP_RANGE_SHEET: data_values(
            sc_not_in_tap_range_rows
        ),
        MAPPING_QC_ALL_SHEET: build_mapping_qc_values(
            all_rows
        ),
        MAPPING_QC_RANGE_SHEET: build_mapping_qc_values(
            range_rows
        ),
    }

    if write_brand_tabs:
        brand_all = build_brand_tables(all_rows)
        brand_range = build_brand_tables(range_rows)

        for brand in BRAND_OUTPUT_ORDER:
            tables[
                f"{brand}{BRAND_ALL_SUFFIX}"
            ] = data_values(
                brand_all[brand]
            )

        for brand in BRAND_OUTPUT_ORDER:
            tables[
                f"{brand}{BRAND_RANGE_SUFFIX}"
            ] = data_values(
                brand_range[brand]
            )

    tables[SUMMARY_SHEET] = build_brand_summary(
        all_rows,
        range_rows,
    )

    tables[MISSING_SHEET] = build_missing_values(
        missing_details
    )

    return tables


def all_managed_titles():
    titles = {
        SHEET_ALL,
        SHEET_RANGE,
        SC_NOT_IN_TAP_ALL_SHEET,
        SC_NOT_IN_TAP_RANGE_SHEET,
        MAPPING_QC_ALL_SHEET,
        MAPPING_QC_RANGE_SHEET,
        BRAND_UNIQUE_ALL_SHEET,
        BRAND_UNIQUE_RANGE_SHEET,
        BRAND_UNIQUE_NOT_IN_TAP_ALL_SHEET,
        BRAND_UNIQUE_NOT_IN_TAP_RANGE_SHEET,
        SUMMARY_SHEET,
        MISSING_SHEET,
    }

    for brand in BRAND_OUTPUT_ORDER:
        titles.add(
            f"{brand}{BRAND_ALL_SUFFIX}"
        )
        titles.add(
            f"{brand}{BRAND_RANGE_SUFFIX}"
        )

    return titles


# =============================================================================
# OUTPUT GOOGLE SHEETS
# =============================================================================

def get_output_grid_info(service):
    properties = get_sheet_properties(
        service,
        OUTPUT_SPREADSHEET_ID,
    )

    info = {}
    total_cells = 0

    for prop in properties:
        grid = prop.get(
            "gridProperties",
            {},
        )

        rows = int(
            grid.get("rowCount", 0)
            or 0
        )

        cols = int(
            grid.get("columnCount", 0)
            or 0
        )

        cells = rows * cols

        info[prop["title"]] = {
            "sheetId": prop["sheetId"],
            "rows": rows,
            "cols": cols,
            "cells": cells,
        }

        total_cells += cells

    return info, total_cells


def cells_for_values(values):
    row_count = max(len(values), 2)

    col_count = max(
        max(
            (len(row) for row in values),
            default=1,
        ),
        1,
    )

    return row_count * col_count


def validate_output_capacity(
    service,
    output_tables,
):
    info, existing_cells = get_output_grid_info(
        service
    )

    requested_titles = set(output_tables)

    managed_existing = (
        set(info)
        & all_managed_titles()
    )

    clashes = (
        requested_titles
        & set(info)
    )

    if clashes and not REPLACE:
        raise SystemExit(
            "Tab output sudah ada dan tidak akan ditimpa: "
            + ", ".join(sorted(clashes))
            + "\nJalankan dengan --write --replace "
              "jika memang mau mengganti output."
        )

    removed_titles = (
        managed_existing
        if REPLACE
        else set()
    )

    cells_removed = sum(
        info[t]["cells"]
        for t in removed_titles
    )

    new_cells = sum(
        cells_for_values(values)
        for values in output_tables.values()
    )

    projected = (
        existing_cells
        - cells_removed
        + new_cells
    )

    print("\nValidasi kapasitas output:")
    print(
        f"  Existing cells       : "
        f"{existing_cells:,}"
    )

    if REPLACE:
        print(
            f"  Managed cells removed: "
            f"-{cells_removed:,}"
        )

    print(
        f"  New output cells     : "
        f"{new_cells:,}"
    )

    print(
        f"  Projected total      : "
        f"{projected:,}"
    )

    print(
        f"  Safety limit         : "
        f"{MAX_CELLS:,}"
    )

    if projected > MAX_CELLS:
        raise SystemExit(
            "Projected output melebihi MAX_CELLS. "
            "Gunakan output spreadsheet lain "
            "atau nonaktifkan brand tabs."
        )

    return info


def create_temp_sheet_if_needed(
    service,
    info,
    delete_titles,
):
    existing_titles = set(info)

    if (
        not delete_titles
        or delete_titles != existing_titles
    ):
        return None

    temp_title = "_TMP_AUTOMATION_DO_NOT_USE"
    suffix = 1

    while temp_title in existing_titles:
        suffix += 1
        temp_title = (
            f"_TMP_AUTOMATION_DO_NOT_USE_"
            f"{suffix}"
        )

    response = execute_with_retry(
        service.spreadsheets().batchUpdate(
            spreadsheetId=OUTPUT_SPREADSHEET_ID,
            body={
                "requests": [
                    {
                        "addSheet": {
                            "properties": {
                                "title": temp_title
                            }
                        }
                    }
                ]
            },
        )
    )

    replies = response.get(
        "replies",
        [],
    )

    if replies:
        return replies[0][
            "addSheet"
        ][
            "properties"
        ][
            "sheetId"
        ]

    return None


def recreate_output_tabs(
    service,
    info,
    output_tables,
):
    delete_titles = (
        set(info)
        & all_managed_titles()
        if REPLACE
        else set()
    )

    temp_sheet_id = (
        create_temp_sheet_if_needed(
            service,
            info,
            delete_titles,
        )
    )

    requests = []

    for title in sorted(delete_titles):
        requests.append({
            "deleteSheet": {
                "sheetId": info[
                    title
                ][
                    "sheetId"
                ]
            }
        })

    for title, values in output_tables.items():
        row_count = max(
            len(values),
            2,
        )

        col_count = max(
            max(
                (
                    len(row)
                    for row in values
                ),
                default=1,
            ),
            1,
        )

        requests.append({
            "addSheet": {
                "properties": {
                    "title": title,
                    "gridProperties": {
                        "rowCount": row_count,
                        "columnCount": col_count,
                        "frozenRowCount": 1,
                    },
                }
            }
        })

    if requests:
        execute_with_retry(
            service.spreadsheets().batchUpdate(
                spreadsheetId=OUTPUT_SPREADSHEET_ID,
                body={"requests": requests},
            )
        )

    if temp_sheet_id is not None:
        execute_with_retry(
            service.spreadsheets().batchUpdate(
                spreadsheetId=OUTPUT_SPREADSHEET_ID,
                body={
                    "requests": [
                        {
                            "deleteSheet": {
                                "sheetId": temp_sheet_id
                            }
                        }
                    ]
                },
            )
        )


def write_values_chunked(
    service,
    title,
    values,
):
    for start in range(
        0,
        len(values),
        WRITE_CHUNK_ROWS,
    ):
        part = values[
            start:
            start + WRITE_CHUNK_ROWS
        ]

        execute_with_retry(
            service.spreadsheets().values().update(
                spreadsheetId=OUTPUT_SPREADSHEET_ID,
                range=(
                    f"'{title}'!"
                    f"A{start + 1}"
                ),
                valueInputOption="RAW",
                body={"values": part},
            )
        )

    print(
        f'  "{title}" selesai '
        f'({max(len(values) - 1, 0):,} row)'
    )


def write_output(
    service,
    output_tables,
    info,
):
    recreate_output_tabs(
        service,
        info,
        output_tables,
    )

    for title, values in output_tables.items():
        write_values_chunked(
            service,
            title,
            values,
        )


# =============================================================================
# DIAGNOSTICS
# =============================================================================

def pct_diff(
    mapped: float,
    original: float,
) -> float | None:
    if math.isclose(
        original,
        0.0,
        abs_tol=1e-12,
    ):
        return None

    return (
        (mapped - original)
        / original
        * 100.0
    )


def fmt_num(
    value: float | int,
) -> str:
    if (
        isinstance(value, int)
        or math.isclose(
            float(value),
            round(float(value)),
            abs_tol=1e-9,
        )
    ):
        return (
            f"{int(round(float(value))):,}"
        )

    return f"{float(value):,.2f}"


def print_source_meta(
    label,
    meta,
):
    print(f"\n{label} source:")
    print(
        f"  Spreadsheet ID : "
        f"{meta['spreadsheet_id']}"
    )
    print(
        f"  Sheet          : "
        f"{meta['sheet_title']}"
    )
    print(
        f"  GID            : "
        f"{meta['sheet_id']}"
    )
    print(
        f"  Header row     : "
        f"{meta['header_row']}"
    )
    print(
        f"  Columns        : "
        f"{meta['columns']}"
    )
    print(
        f"  Rows           : "
        f"{meta['rows']:,}"
    )

    if meta.get("tap_schema"):
        print(
            f"  TAP schema     : "
            f"{meta['tap_schema']}"
        )


def print_summary(
    tap_stats,
    aff_stats,
    master_counts,
    sc_stats,
    original_totals,
    enriched_totals,
    brand_counter,
    all_rows,
    range_rows,
    sc_not_in_tap_all_rows,
    sc_not_in_tap_range_rows,
    tap_live_ids,
    tap_existence_stats,
    invalid_range_dates,
    range_start,
    range_end,
    write_brand_tabs,
    missing_details,
):
    print("\n" + "=" * 78)
    print(
        "SC BW LIVE - EXPLODE + TAP LS/AFF "
        "PRODUCT MAPPING"
    )
    print("=" * 78)

    print("\nTAP LS mapping:")
    print(
        f"  Source count             : "
        f"{tap_stats['sources']:,}"
    )
    print(
        f"  Raw row                  : "
        f"{tap_stats['raw_rows']:,}"
    )
    print(
        f"  Exact duplicate skipped  : "
        f"{tap_stats['duplicate_rows']:,}"
    )
    print(
        f"  LIVE ID + Product ID     : "
        f"{tap_stats['pairs']:,}"
    )
    print(
        f"  Blank/invalid key        : "
        f"{tap_stats['invalid_key']:,}"
    )
    print(
        f"  Invalid LIVE start time  : "
        f"{tap_stats['invalid_live_start']:,}"
    )
    print(
        f"  Outside LIVE start period: "
        f"{tap_stats['period_skipped']:,}"
    )

    for schema_name, source_count in sorted(
        tap_stats.get("schemas", {}).items()
    ):
        print(
            f"  Schema {schema_name:<20}: "
            f"{source_count:,} source"
        )

    print("\nAffiliate Livestream mapping:")
    print(
        f"  Raw row                  : "
        f"{aff_stats['raw_rows']:,}"
    )
    print(
        f"  Non-Livestream skipped   : "
        f"{aff_stats['non_livestream']:,}"
    )
    print(
        f"  LIVE ID + Product ID     : "
        f"{aff_stats['pairs']:,}"
    )
    print(
        f"  Exact duplicate skipped  : "
        f"{aff_stats['duplicate_rows']:,}"
    )
    print(
        f"  Blank/invalid key        : "
        f"{aff_stats['invalid_key']:,}"
    )
    print(
        f"  Invalid Time Created     : "
        f"{aff_stats['invalid_date']:,}"
    )
    print(
        f"  Outside Time Created     : "
        f"{aff_stats['period_skipped']:,}"
    )

    print("\nMaster TAP LS + AFF mapping:")
    print(
        f"  BOTH                     : "
        f"{master_counts['BOTH']:,}"
    )
    print(
        f"  TAP only                 : "
        f"{master_counts['TAP_ONLY']:,}"
    )
    print(
        f"  AFF only                 : "
        f"{master_counts['AFF_ONLY']:,}"
    )
    print(
        f"  Total pair               : "
        f"{sum(master_counts.values()):,}"
    )

    print("\nSC LIVE explode + enrichment:")
    print(
        f"  Raw SC row               : "
        f"{sc_stats['raw_rows']:,}"
    )
    print(
        f"  Source row multi-PID     : "
        f"{sc_stats['multi_pid_source_rows']:,}"
    )
    print(
        f"  Exploded output row      : "
        f"{sc_stats['exploded_rows']:,}"
    )
    print(
        f"  Single-PID SC preserved  : "
        f"{sc_stats['single_pid_sc_preserved']:,}"
    )
    print(
        f"  Multi-PID mapped TAP     : "
        f"{sc_stats['multi_pid_mapped_tap']:,}"
    )
    print(
        f"  Multi-PID mapped AFF     : "
        f"{sc_stats['multi_pid_mapped_aff_only']:,}"
    )
    print(
        f"  Unmapped multi-PID row   : "
        f"{sc_stats['unmapped_multi_pid']:,}"
    )
    print(
        f"  Source row blank PID     : "
        f"{sc_stats['blank_pid_source_rows']:,}"
    )
    print(
        f"  Invalid LIVE ID          : "
        f"{sc_stats['invalid_live_id']:,}"
    )
    print(
        f"  Unique missing pair      : "
        f"{len(missing_details):,}"
    )
    print(
        f"  BW Product ID mapped     : "
        f"{sc_stats['bw_product_mapped']:,}"
    )
    print(
        f"  BW Product ID MISSING    : "
        f"{sc_stats['bw_product_missing']:,}"
    )

    print(
        "\nBW Brand from Product ID mapping:"
    )

    for brand in BRAND_OUTPUT_ORDER:
        print(
            f"  {brand:<15} "
            f"{brand_counter.get(brand, 0):>10,} row"
        )

    print("\nALL / RANGE:")

    if range_start or range_end:
        start_text = (
            range_start.isoformat()
            if range_start
            else "-inf"
        )

        end_text = (
            range_end.isoformat()
            if range_end
            else "+inf"
        )

        print(
            f"  RANGE                    : "
            f"{start_text} -> {end_text} "
            f"(SC LIVE start time)"
        )
    else:
        print(
            "  RANGE                    : "
            "same as ALL"
        )

    print(
        f"  {SHEET_ALL:<28}: "
        f"{len(all_rows):>10,} row"
    )

    print(
        f"  {SHEET_RANGE:<28}: "
        f"{len(range_rows):>10,} row"
    )

    print(
        f"  Invalid SC start RANGE   : "
        f"{invalid_range_dates:>10,}"
    )

    print(
        f"  Brand tabs               : "
        f"{'YES' if write_brand_tabs else 'NO'}"
    )

    missing_all_live_ids = {
        normalize_id(
            row.get("LIVE ID", "")
        )
        for row in sc_not_in_tap_all_rows
        if normalize_id(
            row.get("LIVE ID", "")
        )
    }

    missing_range_live_ids = {
        normalize_id(
            row.get("LIVE ID", "")
        )
        for row in sc_not_in_tap_range_rows
        if normalize_id(
            row.get("LIVE ID", "")
        )
    }

    print(
        "\nSC LIVE ID vs TAP LS "
        "(LIVE start time range; union semua TAP source):"
    )

    print(
        f"  TAP source count         : "
        f"{tap_existence_stats['sources']:>10,}"
    )

    print(
        f"  TAP raw rows             : "
        f"{tap_existence_stats['raw_rows']:>10,}"
    )

    print(
        f"  TAP exact dup skipped    : "
        f"{tap_existence_stats['duplicate_rows']:>10,}"
    )

    print(
        f"  TAP rows start in period : "
        f"{tap_existence_stats['included_rows']:>10,}"
    )

    print(
        f"  TAP outside start period : "
        f"{tap_existence_stats['period_skipped']:>10,}"
    )

    print(
        f"  TAP invalid start time   : "
        f"{tap_existence_stats['invalid_live_start']:>10,}"
    )

    print(
        f"  TAP blank LIVE ID        : "
        f"{tap_existence_stats['blank_or_invalid_live_id']:>10,}"
    )

    print(
        f"  Unique TAP LIVE ID RANGE : "
        f"{len(tap_live_ids):>10,}"
    )

    print(
        f"  Missing SC LIVE ID ALL   : "
        f"{len(missing_all_live_ids):>10,}"
    )

    print(
        f"  Missing SC rows ALL      : "
        f"{len(sc_not_in_tap_all_rows):>10,}"
    )

    print(
        f"  Missing SC LIVE ID RANGE : "
        f"{len(missing_range_live_ids):>10,}"
    )

    print(
        f"  Missing SC rows RANGE    : "
        f"{len(sc_not_in_tap_range_rows):>10,}"
    )

    print(
        "  Compare key              : "
        "LIVE ID vs TAP LIVE ID"
    )

    print(
        "  TAP existence basis      : "
        "LIVE start time"
    )

    qc_all = summarize_mapping_qc(all_rows)
    qc_range = summarize_mapping_qc(range_rows)

    print("\nExact Product-ID Mapping QC (TAP LS + AFF):")
    print(
        f"  LIVE ID ALL              : "
        f"{qc_all['live_ids']:>10,}"
    )
    print(
        f"  Fully mapped ALL         : "
        f"{qc_all['fully_mapped']:>10,}"
    )
    print(
        f"  Partial mapping ALL      : "
        f"{qc_all['partial_mapping']:>10,}"
    )
    print(
        f"  No mapping ALL           : "
        f"{qc_all['no_mapping']:>10,}"
    )
    print(
        f"  PID mapped / total ALL   : "
        f"{qc_all['pid_mapped']:>10,} / "
        f"{qc_all['pid_total']:,} "
        f"({qc_all['pid_coverage']:.2%})"
    )
    print(
        f"  PID unmapped ALL         : "
        f"{qc_all['pid_unmapped']:>10,}"
    )
    print(
        f"  LIVE ID RANGE            : "
        f"{qc_range['live_ids']:>10,}"
    )
    print(
        f"  Fully mapped RANGE       : "
        f"{qc_range['fully_mapped']:>10,}"
    )
    print(
        f"  Partial mapping RANGE    : "
        f"{qc_range['partial_mapping']:>10,}"
    )
    print(
        f"  No mapping RANGE         : "
        f"{qc_range['no_mapping']:>10,}"
    )
    print(
        f"  PID mapped / total RANGE : "
        f"{qc_range['pid_mapped']:>10,} / "
        f"{qc_range['pid_total']:,} "
        f"({qc_range['pid_coverage']:.2%})"
    )
    print(
        f"  PID unmapped RANGE       : "
        f"{qc_range['pid_unmapped']:>10,}"
    )
    print(
        f"  Detail                   : "
        f"{MAPPING_QC_ALL_SHEET} / "
        f"{MAPPING_QC_RANGE_SHEET}"
    )

    print("\nCommerce metric reconciliation:")

    labels = {
        "gmv": "GMV",
        "orders": "Orders",
        "items": "Items sold",
    }

    for key in (
        "gmv",
        "orders",
        "items",
    ):
        original = original_totals[key]
        enriched = enriched_totals[key]
        delta = enriched - original

        error = pct_diff(
            float(enriched),
            float(original),
        )

        error_text = (
            "N/A"
            if error is None
            else f"{error:+.2f}%"
        )

        print(
            f"  {labels[key]:<10} "
            f"SC original {fmt_num(original):>16} | "
            f"mapped {fmt_num(enriched):>16} | "
            f"delta {fmt_num(delta):>16} | "
            f"{error_text}"
        )

    print("\nOutput:")
    print(
        f"  Spreadsheet ID           : "
        f"{OUTPUT_SPREADSHEET_ID}"
    )
    print(
        f"  Output columns           : "
        f"{len(OUTPUT_COLUMNS)}"
    )
    print(
        f"  Mode                     : "
        f"{'WRITE' if WRITE else 'DRY RUN'}"
    )
    print(
        f"  Replace                  : "
        f"{'YES' if REPLACE else 'NO'}"
    )

    print("\nCATATAN:")
    print(
        "  TAP metric mapping: union semua TAP source; schema baru/lama auto-detect."
    )
    print(
        "  AFF metric mapping: Time Created."
    )
    print(
        "  SC RANGE: LIVE start time."
    )
    print(
        "  Brand + Sub Product: Product ID mapping BW."
    )
    print(
        "  SC LS NOT IN TAP: LIVE ID + Brand dibandingkan "
        "ke union TAP LIVE ID yang LIVE start time-nya "
        "masuk METRIC range."
    )
    print(
        "  Single-PID: metric commerce utama "
        "dipertahankan dari SC original."
    )
    print(
        "  Multi-PID: metric commerce utama "
        "direct mapping TAP/AFF (Model A)."
    )
    print(
        "  Metric level-LIVE seperti Likes/Comments/"
        "Shares/Impressions akan terulang pada setiap PID "
        "hasil explode; jangan di-SUM lintas Product ID."
    )


# =============================================================================
# CONFIG VALIDATION
# =============================================================================

def is_placeholder(value: object) -> bool:
    text = str(value or "").strip().upper()
    return (
        text.startswith("PASTE_")
        or text.startswith("ISI_")
        or not text
    )


def validate_config():
    missing = []

    if is_placeholder(SC_LS_SPREADSHEET_ID):
        missing.append(
            "SC_LS_SPREADSHEET_ID"
        )

    if not TAP_LS_SOURCES:
        missing.append(
            "TAP_LS_SOURCES"
        )
    else:
        for (
            i,
            source,
        ) in enumerate(
            TAP_LS_SOURCES,
            start=1,
        ):
            if is_placeholder(
                source.get("spreadsheet_id")
            ):
                missing.append(
                    f"TAP_LS_SOURCES[{i}]."
                    "spreadsheet_id"
                )

    if is_placeholder(BW_MAPPING_SPREADSHEET_ID):
        missing.append(
            "BW_MAPPING_SPREADSHEET_ID"
        )

    if is_placeholder(OUTPUT_SPREADSHEET_ID):
        missing.append(
            "OUTPUT_SPREADSHEET_ID"
        )

    if not AFF_SOURCES:
        missing.append(
            "AFF_SOURCES"
        )
    else:
        for (
            i,
            source,
        ) in enumerate(
            AFF_SOURCES,
            start=1,
        ):
            if is_placeholder(
                source.get("spreadsheet_id")
            ):
                missing.append(
                    f"AFF_SOURCES[{i}]."
                    "spreadsheet_id"
                )

    if missing:
        raise SystemExit(
            "Isi CONFIG spreadsheet ID dulu:\n  - "
            + "\n  - ".join(missing)
        )

    if not CREDENTIALS_FILE.exists():
        candidates = "\n  - ".join(
            str(path)
            for path in CREDENTIALS_CANDIDATES
        )

        raise SystemExit(
            "Credentials tidak ditemukan. "
            "Script mencoba:\n  - "
            + candidates
        )

    metric_start = parse_iso_date(
        METRIC_START_DATE
    )

    metric_end = parse_iso_date(
        METRIC_END_DATE
    )

    range_start = parse_iso_date(
        RANGE_START_DATE
    )

    range_end = parse_iso_date(
        RANGE_END_DATE
    )

    if (
        metric_start
        and metric_end
        and metric_start > metric_end
    ):
        raise SystemExit(
            "METRIC_START_DATE tidak boleh "
            "lebih besar dari METRIC_END_DATE"
        )

    if (
        range_start
        and range_end
        and range_start > range_end
    ):
        raise SystemExit(
            "RANGE_START_DATE tidak boleh "
            "lebih besar dari RANGE_END_DATE"
        )

    return (
        metric_start,
        metric_end,
        range_start,
        range_end,
    )


# =============================================================================
# MAIN
# =============================================================================

def main():
    (
        metric_start,
        metric_end,
        range_start,
        range_end,
    ) = validate_config()

    service = get_sheets_service()

    print("=" * 78)
    print(
        "SC BW LIVE CLEAN + TAP LS/AFF "
        "ENRICHMENT - GOOGLE SHEETS"
    )
    print("=" * 78)

    print(
        f"Mode             : "
        f"{'WRITE' if WRITE else 'DRY RUN'}"
    )

    print(
        f"Replace          : "
        f"{'YES' if REPLACE else 'NO'}"
    )

    print(
        f"Brand tabs       : "
        f"{'YES' if WRITE_BRAND_TABS_EFFECTIVE else 'NO'}"
    )

    print(
        f"Version          : "
        f"{CLEANING_VERSION}"
    )

    if metric_start or metric_end:
        print(
            "Metric period    : "
            f"{metric_start.isoformat() if metric_start else '-inf'} "
            "-> "
            f"{metric_end.isoformat() if metric_end else '+inf'}"
        )

        print(
            "  TAP basis      : LIVE start time"
        )

        print(
            "  AFF basis      : Time Created"
        )
    else:
        print(
            "Metric period    : "
            "ALL TAP/AFF source rows"
        )

    if range_start or range_end:
        print(
            "SC RANGE         : "
            f"{range_start.isoformat() if range_start else '-inf'} "
            "-> "
            f"{range_end.isoformat() if range_end else '+inf'} "
            "(LIVE start time)"
        )
    else:
        print(
            "SC RANGE         : same as ALL"
        )

    # -------------------------------------------------------------------------
    # Read SC LIVE
    # -------------------------------------------------------------------------
    (
        sc_header,
        sc_rows,
        sc_meta,
    ) = detect_and_read_source(
        service,
        spreadsheet_id=SC_LS_SPREADSHEET_ID,
        sheet_gid=SC_LS_SHEET_GID,
        sheet_name=SC_LS_SHEET_NAME,
        required_headers=SC_REQUIRED,
        label="SC LIVE",
    )

    print_source_meta(
        "SC LIVE",
        sc_meta,
    )

    # -------------------------------------------------------------------------
    # Read TAP LS source(s)
    # -------------------------------------------------------------------------
    tap_datasets = []

    for (
        index,
        source,
    ) in enumerate(
        TAP_LS_SOURCES,
        start=1,
    ):
        label = str(
            source.get("label")
            or f"TAP LS {index}"
        )

        (
            header,
            rows,
            meta,
        ) = detect_and_read_tap_source(
            service,
            spreadsheet_id=source[
                "spreadsheet_id"
            ],
            sheet_gid=source.get(
                "sheet_gid"
            ),
            sheet_name=source.get(
                "sheet_name"
            ),
            label=label,
        )

        print_source_meta(
            label,
            meta,
        )

        tap_datasets.append(
            (
                label,
                header,
                rows,
            )
        )

    # -------------------------------------------------------------------------
    # Read BW Product Mapping
    # -------------------------------------------------------------------------
    (
        bw_map_header,
        bw_map_rows,
        bw_map_meta,
    ) = detect_and_read_source(
        service,
        spreadsheet_id=BW_MAPPING_SPREADSHEET_ID,
        sheet_gid=BW_MAPPING_SHEET_GID,
        sheet_name=BW_MAPPING_SHEET_NAME,
        required_headers=BW_MAPPING_REQUIRED,
        label="BW PRODUCT MAPPING",
    )

    print_source_meta(
        "BW PRODUCT MAPPING",
        bw_map_meta,
    )

    bw_product_mapping, bw_mapping_stats = build_bw_product_mapping(
        bw_map_header,
        bw_map_rows,
    )
    configure_brand_output_order(bw_product_mapping)
    print_bw_mapping_summary(bw_mapping_stats)

    (
        tap_live_ids,
        tap_existence_stats,
    ) = build_tap_ls_live_id_set(
        tap_datasets,
        metric_start,
        metric_end,
    )

    # -------------------------------------------------------------------------
    # Read AFF
    # -------------------------------------------------------------------------
    aff_datasets = []

    for (
        index,
        source,
    ) in enumerate(
        AFF_SOURCES,
        start=1,
    ):
        label = str(
            source.get("label")
            or f"AFF {index}"
        )

        (
            header,
            rows,
            meta,
        ) = detect_and_read_source(
            service,
            spreadsheet_id=source[
                "spreadsheet_id"
            ],
            sheet_gid=source.get(
                "sheet_gid"
            ),
            sheet_name=source.get(
                "sheet_name"
            ),
            required_headers=AFF_REQUIRED,
            label=label,
        )

        print_source_meta(
            label,
            meta,
        )

        aff_datasets.append(
            (
                label,
                header,
                rows,
            )
        )

    # -------------------------------------------------------------------------
    # Mapping
    # -------------------------------------------------------------------------
    print(
        "\nMembangun TAP LS mapping dari semua source..."
    )

    (
        tap_map,
        tap_stats,
    ) = build_tap_ls_mapping(
        tap_datasets,
        metric_start,
        metric_end,
    )

    print(
        "Membangun Affiliate "
        "Livestream mapping..."
    )

    (
        aff_map,
        aff_stats,
    ) = build_aff_ls_mapping(
        aff_datasets,
        metric_start,
        metric_end,
    )

    print(
        "Menggabungkan TAP LS + AFF "
        "menjadi master mapping..."
    )

    (
        master_map,
        master_counts,
    ) = build_master_mapping(
        tap_map,
        aff_map,
    )

    pid_product_name_lookup = (
        build_pid_product_name_lookup(
            tap_map,
            aff_map,
        )
    )

    tap_live_brand_pairs = build_tap_brand_pair_set_from_mapping(
        tap_map,
        bw_product_mapping,
    )

    # -------------------------------------------------------------------------
    # SC explode + enrichment
    # -------------------------------------------------------------------------
    print(
        "Explode + enrich SC LIVE..."
    )

    (
        all_rows,
        sc_stats,
        original_totals,
        enriched_totals,
        brand_counter,
        missing_details,
    ) = enrich_sc_live(
        sc_header,
        sc_rows,
        master_map,
        aff_map,
        pid_product_name_lookup,
        bw_product_mapping,
    )

    (
        range_rows,
        invalid_range_dates,
    ) = filter_output_range(
        all_rows,
        range_start,
        range_end,
    )

    # -------------------------------------------------------------------------
    # SC LS NOT IN TAP
    # -------------------------------------------------------------------------
    sc_not_in_tap_all_rows = (
        filter_sc_ls_not_in_tap(
            all_rows,
            tap_live_brand_pairs,
        )
    )

    sc_not_in_tap_range_rows = (
        filter_sc_ls_not_in_tap(
            range_rows,
            tap_live_brand_pairs,
        )
    )

    # -------------------------------------------------------------------------
    # BRAND UNIQUE: FIRST PID per LIVE ID + Brand
    # Commerce metric = SUM unique PID dalam LIVE ID + Brand.
    # NOT IN TAP memakai grain LIVE ID + Brand.
    # -------------------------------------------------------------------------
    brand_unique_all_rows = build_brand_unique_rows(all_rows)
    brand_unique_range_rows = build_brand_unique_rows(range_rows)
    brand_unique_not_in_tap_all_rows = build_brand_unique_rows(
        sc_not_in_tap_all_rows
    )
    brand_unique_not_in_tap_range_rows = build_brand_unique_rows(
        sc_not_in_tap_range_rows
    )

    # -------------------------------------------------------------------------
    # Output
    # -------------------------------------------------------------------------
    output_tables = build_output_tables(
        all_rows,
        range_rows,
        sc_not_in_tap_all_rows,
        sc_not_in_tap_range_rows,
        missing_details,
        WRITE_BRAND_TABS_EFFECTIVE,
    )

    output_tables[BRAND_UNIQUE_ALL_SHEET] = data_values(brand_unique_all_rows)
    output_tables[BRAND_UNIQUE_RANGE_SHEET] = data_values(brand_unique_range_rows)
    output_tables[BRAND_UNIQUE_NOT_IN_TAP_ALL_SHEET] = data_values(
        brand_unique_not_in_tap_all_rows
    )
    output_tables[BRAND_UNIQUE_NOT_IN_TAP_RANGE_SHEET] = data_values(
        brand_unique_not_in_tap_range_rows
    )

    # Brand summary mengikuti grain final LIVE ID + Brand, bukan exploded rows.
    output_tables[SUMMARY_SHEET] = build_brand_summary(
        brand_unique_all_rows,
        brand_unique_range_rows,
    )

    output_info = validate_output_capacity(
        service,
        output_tables,
    )

    print(
        "\nBRAND UNIQUE OUTPUT (FIRST PID per LIVE ID + Brand):"
    )
    print(f"  {BRAND_UNIQUE_ALL_SHEET:<42}: {len(brand_unique_all_rows):>10,} row")
    print(f"  {BRAND_UNIQUE_RANGE_SHEET:<42}: {len(brand_unique_range_rows):>10,} row")
    print(f"  {BRAND_UNIQUE_NOT_IN_TAP_ALL_SHEET:<42}: {len(brand_unique_not_in_tap_all_rows):>10,} row")
    print(f"  {BRAND_UNIQUE_NOT_IN_TAP_RANGE_SHEET:<42}: {len(brand_unique_not_in_tap_range_rows):>10,} row")
    print("  NOT IN TAP key: LIVE ID + Brand.")
    print("  GMV/Orders/Items BRAND UNIQUE: SUM unique Product ID per LIVE ID + Brand.")

    print_summary(
        tap_stats,
        aff_stats,
        master_counts,
        sc_stats,
        original_totals,
        enriched_totals,
        brand_counter,
        all_rows,
        range_rows,
        sc_not_in_tap_all_rows,
        sc_not_in_tap_range_rows,
        tap_live_ids,
        tap_existence_stats,
        invalid_range_dates,
        range_start,
        range_end,
        WRITE_BRAND_TABS_EFFECTIVE,
        missing_details,
    )

    if not WRITE:
        print(
            "\nDRY RUN selesai. "
            "Tidak ada data ditulis."
        )

        print(
            "Jika QC sudah sesuai:"
        )

        print(
            "  python "
            "sc_bw_live_exploded_and_brand_unique.py "
            "--write"
        )

        print(
            "Jika tab output lama memang "
            "mau diganti:"
        )

        print(
            "  python "
            "sc_bw_live_exploded_and_brand_unique.py "
            "--write --replace"
        )

        return

    print(
        "\nMenulis output ke Google Sheets..."
    )

    write_output(
        service,
        output_tables,
        output_info,
    )

    print(
        "\nSELESAI."
    )


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print(
            "\nDibatalkan user.",
            file=sys.stderr,
        )
        raise SystemExit(130)
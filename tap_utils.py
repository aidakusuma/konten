"""Shared utilities for TAP BW Livestream and Short Video processing."""

from __future__ import annotations

import hashlib
import json
import math
import re
import time
from collections import Counter, defaultdict
from datetime import date, datetime, timedelta

try:
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
    from googleapiclient.errors import HttpError
except ImportError:  # pemakaian lokal (file Excel) tidak butuh library Google
    service_account = None
    build = None

    class HttpError(Exception):
        """Placeholder agar `except HttpError` tetap valid tanpa library Google."""


# =============================================================================
# BW classification constants
# =============================================================================

SOURCE_PRODUCT_ID_HEADER = "Product ID"
OUTPUT_BRAND_HEADER = "Brand"
OUTPUT_SUB_PRODUCT_HEADER = "Sub Product"

MISSING_BRAND = "MISSING"
MISSING_SUB_PRODUCT = "MISSING"

BW_MAPPING_PRODUCT_ID_HEADER = "Product ID"
BW_MAPPING_BRAND_HEADER = "Brand"
BW_MAPPING_SUB_PRODUCT_HEADER = "Sub Product"


# =============================================================================
# Google Sheets policy
# =============================================================================

MAX_CELLS = 20_000_000
WRITE_CHUNK_ROWS = 1500
MAX_RETRIES = 5
RETRYABLE_HTTP_STATUS = {429, 500, 502, 503, 504}


# =============================================================================
# Type constants
# =============================================================================

TYPE_TEXT = "text"
TYPE_DATE = "date"
TYPE_DATETIME = "datetime"
TYPE_INTEGER = "integer"
TYPE_NUMBER = "number"
TYPE_CURRENCY = "currency"
TYPE_PERCENT = "percent"

NULL_NUMERIC_TOKENS = {
    "",
    "-",
    "--",
    "—",
    "N/A",
    "NA",
    "NULL",
    "NONE",
    "#N/A",
}


# =============================================================================
# Google Sheets client
# =============================================================================

def get_sheets_service(credentials_file, write=False):
    scope = (
        "https://www.googleapis.com/auth/spreadsheets"
        if write
        else "https://www.googleapis.com/auth/spreadsheets.readonly"
    )

    if service_account is None or build is None:
        raise ImportError(
            "Library Google belum terpasang: pip install google-api-python-client google-auth"
        )

    credentials = service_account.Credentials.from_service_account_file(
        str(credentials_file),
        scopes=[scope],
    )

    return build(
        "sheets",
        "v4",
        credentials=credentials,
        cache_discovery=False,
    )


def execute_with_retry(request, max_retries=MAX_RETRIES):
    for attempt in range(max_retries):
        try:
            return request.execute()

        except HttpError as exc:
            status = getattr(exc.resp, "status", None)

            if (
                status not in RETRYABLE_HTTP_STATUS
                or attempt == max_retries - 1
            ):
                raise

            wait_seconds = 2 ** attempt

            print(
                f"Google API error {status}; "
                f"retry {attempt + 1}/{max_retries} "
                f"in {wait_seconds}s..."
            )

            time.sleep(wait_seconds)


# =============================================================================
# General helpers
# =============================================================================

def cell(row, index):
    return row[index] if index < len(row) else ""


def normalize_header(value):
    return re.sub(
        r"[^a-z0-9]+",
        "",
        str(value or "").strip().lower(),
    )


def find_header_index(header, name):
    target = normalize_header(name)

    for index, value in enumerate(header):
        if normalize_header(value) == target:
            return index

    raise ValueError(
        f'Kolom "{name}" tidak ditemukan. Header: {header}'
    )


def column_letter(index_zero_based):
    number = index_zero_based + 1
    result = ""

    while number:
        number, remainder = divmod(number - 1, 26)
        result = chr(65 + remainder) + result

    return result


# =============================================================================
# ID handling
# =============================================================================

def normalize_id(value):
    if value is None:
        return ""

    if isinstance(value, bool):
        return str(value)

    if isinstance(value, int):
        return str(value)

    if isinstance(value, float):
        if not math.isfinite(value):
            return ""

        if abs(value) >= 1e15:
            raise ValueError(
                f"ID terbaca sebagai FLOAT besar ({value}). "
                "Format source ID sebagai Plain text."
            )

        if value.is_integer():
            return format(value, ".0f")

        return (
            format(value, "f")
            .rstrip("0")
            .rstrip(".")
        )

    result = str(value).strip()

    if not result:
        return ""

    if re.fullmatch(
        r"[+-]?\d+(?:\.\d+)?[eE][+-]?\d+",
        result,
    ):
        raise ValueError(
            f"ID tampil dalam scientific notation ({result}). "
            "Format source ID sebagai Plain text."
        )

    if re.fullmatch(r"\d+\.0", result):
        result = result[:-2]

    return result


# =============================================================================
# Numeric conversion
# =============================================================================

def _is_blank_numeric(value):
    if value is None:
        return True

    return (
        str(value).strip().upper()
        in NULL_NUMERIC_TOKENS
    )


def _split_compact_suffix(value):
    text = value.strip()

    match = re.fullmatch(
        r"(.+?)([KMB])",
        text,
        flags=re.I,
    )

    if not match:
        return text, 1.0

    multiplier = {
        "K": 1_000.0,
        "M": 1_000_000.0,
        "B": 1_000_000_000.0,
    }[match.group(2).upper()]

    return match.group(1), multiplier


def _normalize_decimal_string(value, currency_mode=False):
    text = str(value).strip()

    text = re.sub(
        r"^(?:RP|IDR)\s*",
        "",
        text,
        flags=re.I,
    )

    text = text.replace(" ", "")
    text, multiplier = _split_compact_suffix(text)

    negative = (
        text.startswith("(")
        and text.endswith(")")
    )

    if negative:
        text = text[1:-1]

    text = text.replace("%", "")

    if re.fullmatch(
        r"[+-]?\d+(?:\.\d+)?[eE][+-]?\d+",
        text,
    ):
        number = float(text) * multiplier
        return -number if negative else number

    if not re.fullmatch(
        r"[+-]?[\d.,]+",
        text,
    ):
        raise ValueError(
            f"nilai numerik tidak valid: {value!r}"
        )

    sign = ""

    if text[:1] in {"+", "-"}:
        sign = text[0]
        text = text[1:]

    dot_count = text.count(".")
    comma_count = text.count(",")

    if dot_count and comma_count:
        decimal_separator = (
            "."
            if text.rfind(".") > text.rfind(",")
            else ","
        )

        thousands_separator = (
            ","
            if decimal_separator == "."
            else "."
        )

        text = text.replace(
            thousands_separator,
            "",
        )

        text = text.replace(
            decimal_separator,
            ".",
        )

    elif dot_count or comma_count:
        separator = "." if dot_count else ","
        count = text.count(separator)
        groups = text.split(separator)

        if count > 1:
            if all(
                len(group) == 3
                for group in groups[1:]
            ):
                text = "".join(groups)
            else:
                text = (
                    "".join(groups[:-1])
                    + "."
                    + groups[-1]
                )

        else:
            left, right = groups

            # "Rp1.500" -> ribuan (1500), tetapi "0,703" / "0.703" -> desimal
            if (
                currency_mode
                and len(right) == 3
                and left not in ("", "0")
            ):
                text = left + right
            else:
                text = left + "." + right

    number = float(sign + text) * multiplier

    if negative:
        number = -number

    return number


def parse_integer_value(value):
    if _is_blank_numeric(value):
        return ""

    if isinstance(value, bool):
        raise ValueError(
            f"boolean bukan integer metric: {value!r}"
        )

    if isinstance(value, int):
        return value

    if isinstance(value, float):
        if not math.isfinite(value):
            raise ValueError(
                f"nilai integer tidak finite: {value!r}"
            )

        if not value.is_integer():
            raise ValueError(
                f"nilai integer memiliki desimal: {value!r}"
            )

        return int(value)

    text = str(value).strip()

    if re.search(
        r"[KMB]$",
        text,
        flags=re.I,
    ):
        number = _normalize_decimal_string(
            text,
            currency_mode=False,
        )

    else:
        cleaned = re.sub(
            r"^(?:RP|IDR)\s*",
            "",
            text,
            flags=re.I,
        ).replace(" ", "")

        if re.fullmatch(
            r"[+-]?\d+(?:\.0+)?",
            cleaned,
        ):
            number = float(cleaned)

        elif re.fullmatch(
            r"[+-]?\d{1,3}(?:[.,]\d{3})+",
            cleaned,
        ):
            number = float(
                re.sub(r"[.,]", "", cleaned)
            )

        elif re.fullmatch(
            r"[+-]?\d+(?:\.\d+)?[eE][+-]?\d+",
            cleaned,
        ):
            number = float(cleaned)

        else:
            number = _normalize_decimal_string(
                cleaned,
                currency_mode=False,
            )

    if not float(number).is_integer():
        raise ValueError(
            f"nilai integer memiliki desimal: {value!r}"
        )

    return int(number)


def parse_number_value(value, currency_mode=False):
    if _is_blank_numeric(value):
        return ""

    if isinstance(value, bool):
        raise ValueError(
            f"boolean bukan numeric metric: {value!r}"
        )

    if isinstance(value, (int, float)):
        if (
            isinstance(value, float)
            and not math.isfinite(value)
        ):
            raise ValueError(
                f"nilai numerik tidak finite: {value!r}"
            )

        return value

    number = _normalize_decimal_string(
        value,
        currency_mode=currency_mode,
    )

    if float(number).is_integer():
        return int(number)

    return float(number)


def parse_percent_value(value):
    if _is_blank_numeric(value):
        return ""

    if isinstance(value, bool):
        raise ValueError(
            f"boolean bukan percentage metric: {value!r}"
        )

    if isinstance(value, (int, float)):
        return value

    text = str(value).strip()
    has_percent = text.endswith("%")

    if has_percent:
        text = text[:-1]

    number = parse_number_value(
        text,
        currency_mode=False,
    )

    if number == "":
        return ""

    return (
        float(number) / 100.0
        if has_percent
        else number
    )


def to_number(value, strict=True):
    """
    Convert a metric cell to a number.

    Blank / null tokens (-, --, N/A, ...) are 0. A genuinely invalid value
    (e.g. "abc") raises ValueError when strict=True (default) instead of being
    silently counted as 0. Pass strict=False only for best-effort display code.
    """
    if _is_blank_numeric(value):
        return 0

    if isinstance(value, bool):
        return 0

    if isinstance(value, (int, float)):
        if isinstance(value, float) and not math.isfinite(value):
            return 0
        return value

    try:
        return parse_number_value(
            value,
            currency_mode=True,
        )

    except ValueError:
        if strict:
            raise ValueError(
                f"nilai numerik tidak valid: {value!r}"
            ) from None
        return 0


def numbers_equal(left, right):
    return math.isclose(
        float(left),
        float(right),
        rel_tol=1e-9,
        abs_tol=1e-6,
    )


# =============================================================================
# Date and datetime conversion
# =============================================================================

def normalize_date_text(value):
    if (
        value is None
        or str(value).strip() in {"", "--", "-"}
    ):
        return ""

    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")

    if isinstance(value, date):
        return value.strftime("%Y-%m-%d")

    if (
        isinstance(value, (int, float))
        and not isinstance(value, bool)
    ):
        converted = (
            datetime(1899, 12, 30)
            + timedelta(days=float(value))
        )
        return converted.strftime("%Y-%m-%d")

    text = str(value).strip()

    formats = (
        "%Y-%m-%d",
        "%Y/%m/%d",
        "%d/%m/%Y",
        "%m/%d/%Y",
        "%d-%m-%Y",
    )

    for fmt in formats:
        try:
            return datetime.strptime(
                text[:10],
                fmt,
            ).strftime("%Y-%m-%d")

        except ValueError:
            pass

    raise ValueError(
        f"tanggal tidak valid: {value!r}"
    )


def normalize_datetime_text(value):
    if (
        value is None
        or str(value).strip() in {"", "--", "-"}
    ):
        return ""

    if isinstance(value, datetime):
        return value.strftime(
            "%Y-%m-%d %H:%M:%S"
        )

    if isinstance(value, date):
        return datetime.combine(
            value,
            datetime.min.time(),
        ).strftime(
            "%Y-%m-%d %H:%M:%S"
        )

    text = str(value).strip()

    formats = (
    "%Y-%m-%d %H:%M:%S",
    "%Y-%m-%d %H:%M",
    "%Y-%m-%d",

    # TAP export: MM/DD/YYYY
    "%m/%d/%Y %H:%M:%S",
    "%m/%d/%Y %H:%M",
    "%m/%d/%Y",

    # Optional fallback
    "%d/%m/%Y %H:%M:%S",
    "%d/%m/%Y %H:%M",
    "%d/%m/%Y",
    )

    for fmt in formats:
        try:
            converted = datetime.strptime(
                text,
                fmt,
            )

            return converted.strftime(
                "%Y-%m-%d %H:%M:%S"
            )

        except ValueError:
            pass

    raise ValueError(
        f"datetime tidak valid: {value!r}"
    )


def infer_column_type(header_name):
    """Infer a conservative type from complete header tokens."""
    raw = str(header_name or "").strip().lower()
    tokens = re.findall(r"[a-z0-9]+", raw)
    token_set = set(tokens)
    compact = normalize_header(header_name)

    if token_set.intersection({"id", "code"}):
        return TYPE_TEXT

    if "duration" in token_set or "duration" in compact:
        return TYPE_TEXT

    if "date" in token_set:
        return TYPE_DATE

    if "time" in token_set:
        return TYPE_DATETIME

    if token_set.intersection(
        {
            "likes",
            "views",
            "orders",
            "items",
            "follower",
            "followers",
            "count",
            "clicks",
            "impressions",
            "comments",
            "shares",
        }
    ):
        return TYPE_INTEGER

    if token_set.intersection(
        {
            "gmv",
            "commission",
            "rpm",
            "aov",
            "gpm",
            "revenue",
            "amount",
            "price",
        }
    ):
        return TYPE_CURRENCY

    if token_set.intersection(
        {
            "ctr",
            "rate",
            "engagement",
            "percentage",
            "percent",
        }
    ):
        return TYPE_PERCENT

    return TYPE_TEXT


def coerce_rows_by_schema(
    header,
    rows,
    column_types=None,
    strict=True,
    skip_invalid_rows=False,
):
    """
    Convert source values to stable Python types.

    When skip_invalid_rows is enabled, a row with any conversion error is
    excluded from downstream processing and reported through the errors result.
    """
    configured = {
        normalize_header(name): data_type
        for name, data_type
        in (column_types or {}).items()
    }

    resolved = [
        configured.get(
            normalize_header(name),
            infer_column_type(name),
        )
        for name in header
    ]

    converted_rows = []
    errors = []

    for row_number, row in enumerate(
        rows,
        start=2,
    ):
        output_row = [
            cell(row, index)
            for index in range(len(header))
        ]

        row_errors = []

        for index, (
            name,
            data_type,
        ) in enumerate(
            zip(header, resolved)
        ):
            value = output_row[index]

            try:
                if data_type == TYPE_TEXT:
                    output_row[index] = (
                        ""
                        if value is None
                        else str(value).strip()
                    )

                elif data_type == TYPE_DATE:
                    output_row[index] = (
                        normalize_date_text(value)
                    )

                elif data_type == TYPE_DATETIME:
                    output_row[index] = (
                        normalize_datetime_text(value)
                    )

                elif data_type == TYPE_INTEGER:
                    output_row[index] = (
                        parse_integer_value(value)
                    )

                elif data_type == TYPE_NUMBER:
                    output_row[index] = (
                        parse_number_value(
                            value,
                            currency_mode=False,
                        )
                    )

                elif data_type == TYPE_CURRENCY:
                    output_row[index] = (
                        parse_number_value(
                            value,
                            currency_mode=True,
                        )
                    )

                elif data_type == TYPE_PERCENT:
                    output_row[index] = (
                        parse_percent_value(value)
                    )

                else:
                    raise ValueError(
                        f"tipe kolom tidak dikenal: "
                        f"{data_type!r}"
                    )

            except ValueError as exc:
                row_errors.append(
                    (
                        row_number,
                        name,
                        value,
                        str(exc),
                    )
                )

        if row_errors:
            errors.extend(row_errors)

            if skip_invalid_rows:
                continue

        converted_rows.append(
            output_row
        )

    if errors and strict:
        detail = "\n".join(
            f"  row {row_number}, "
            f"{name}: {value!r} "
            f"({reason})"
            for (
                row_number,
                name,
                value,
                reason,
            ) in errors[:20]
        )

        extra = (
            f"\n  ... {len(errors) - 20:,} error lainnya."
            if len(errors) > 20
            else ""
        )

        raise ValueError(
            "Konversi tipe data gagal. "
            "Periksa source atau COLUMN_TYPES:\n"
            + detail
            + extra
        )

    type_summary = Counter(resolved)

    return (
        converted_rows,
        dict(zip(header, resolved)),
        type_summary,
        errors,
    )


# =============================================================================
# Reporting dates
# =============================================================================

def parse_config_date(value, label):
    text = str(value or "").strip()

    if not text:
        raise ValueError(
            f"{label} wajib diisi YYYY-MM-DD."
        )

    try:
        return datetime.strptime(
            text,
            "%Y-%m-%d",
        ).date()

    except ValueError as exc:
        raise ValueError(
            f"{label} harus format YYYY-MM-DD. "
            f"Dapat: {value!r}"
        ) from exc


def parse_date_value(value):
    if (
        value is None
        or value == ""
    ):
        return None

    if isinstance(value, datetime):
        return value.date()

    if isinstance(value, date):
        return value

    text = str(value).strip()

    if not text:
        return None

    if re.match(
        r"^\d{4}-\d{2}-\d{2}",
        text,
    ):
        try:
            return datetime.strptime(
                text[:10],
                "%Y-%m-%d",
            ).date()

        except ValueError:
            pass

    for fmt in (
        "%d/%m/%Y",
        "%m/%d/%Y",
        "%d-%m-%Y",
    ):
        try:
            return datetime.strptime(
                text[:10],
                fmt,
            ).date()

        except ValueError:
            pass

    return None


def filter_groups_by_date_range(
    header,
    groups,
    date_column,
    start_date,
    end_date,
):
    start = parse_config_date(
        start_date,
        "RANGE_START_DATE",
    )

    end = parse_config_date(
        end_date,
        "RANGE_END_DATE",
    )

    if start > end:
        raise ValueError(
            "RANGE_START_DATE tidak boleh "
            "lebih besar dari RANGE_END_DATE."
        )

    date_index = find_header_index(
        header,
        date_column,
    )

    filtered = {}
    invalid_count = 0

    for key, row in groups.items():
        parsed_date = parse_date_value(
            cell(
                row,
                date_index,
            )
        )

        if parsed_date is None:
            invalid_count += 1
            continue

        if start <= parsed_date <= end:
            filtered[key] = row

    return (
        filtered,
        invalid_count,
    )


def filter_groups_by_date_ranges(
    header,
    groups,
    date_column,
    ranges,
):
    """Filter aggregated groups by the union of several inclusive ranges."""
    date_index = find_header_index(
        header,
        date_column,
    )

    normalized_ranges = []

    for start_date, end_date in ranges:
        if isinstance(start_date, str):
            start_date = parse_config_date(start_date, "RANGE_START_DATE")
        if isinstance(end_date, str):
            end_date = parse_config_date(end_date, "RANGE_END_DATE")

        if start_date > end_date:
            raise ValueError(
                "Tanggal awal cumulative RANGE tidak boleh lebih besar dari tanggal akhir."
            )

        normalized_ranges.append((start_date, end_date))

    normalized_ranges = merge_date_ranges(normalized_ranges)

    filtered = {}
    invalid_count = 0

    for key, row in groups.items():
        parsed_date = parse_date_value(
            cell(row, date_index)
        )

        if parsed_date is None:
            invalid_count += 1
            continue

        if any(
            start_date <= parsed_date <= end_date
            for start_date, end_date in normalized_ranges
        ):
            filtered[key] = row

    return filtered, invalid_count


# =============================================================================
# Source reader
# =============================================================================

def read_source(
    service,
    spreadsheet_id,
    source_sheet="Custom report",
    source_range="A1:AZ",
):
    response = execute_with_retry(
        service
        .spreadsheets()
        .values()
        .get(
            spreadsheetId=spreadsheet_id,
            range=f"'{source_sheet}'!{source_range}",
            valueRenderOption="UNFORMATTED_VALUE",
            dateTimeRenderOption="FORMATTED_STRING",
        )
    )

    values = response.get(
        "values",
        [],
    )

    if not values:
        raise ValueError(
            f'Tab source "{source_sheet}" '
            "kosong / tidak ditemukan."
        )

    return (
        values[0],
        values[1:],
    )


def _canonical_date_text(value):
    """Return an ISO date / datetime text suitable for chronological comparison."""
    text = str(value if value is not None else "").strip()

    if not text:
        return ""

    try:
        normalized = normalize_datetime_text(text)
    except ValueError:
        raise ValueError(
            f"kolom tanggal agregasi tidak valid: {value!r}"
        ) from None

    if normalized.endswith(" 00:00:00") and ":" not in text:
        return normalized[:10]

    return normalized


def _update_earliest_date(base, date_index, value):
    """Keep the earliest date, comparing normalized values (never raw text)."""
    current = _canonical_date_text(value)
    stored = _canonical_date_text(cell(base, date_index))

    if current and (not stored or current < stored):
        base[date_index] = current
    elif stored:
        base[date_index] = stored


# =============================================================================
# Content-level aggregation
# =============================================================================

def aggregate_brand_content(
    header,
    rows,
    content_id_header,
    date_header,
    sum_columns,
    brand_header=OUTPUT_BRAND_HEADER,
    row_validator=None,
):
    """
    Aggregate source rows to the (Brand, Content ID) grain.

    Additive metrics are summed. Date keeps the earliest value. Other fields
    keep the first accepted row.
    """
    content_id_index = find_header_index(
        header,
        content_id_header,
    )

    brand_index = find_header_index(
        header,
        brand_header,
    )

    date_index = find_header_index(
        header,
        date_header,
    )

    sum_indexes = [
        find_header_index(
            header,
            column,
        )
        for column in sum_columns
    ]

    groups = {}
    accepted_rows = []
    skipped_blank = 0
    invalid_rows = []

    for row_number, row in enumerate(
        rows,
        start=2,
    ):
        content_id = normalize_id(
            cell(
                row,
                content_id_index,
            )
        )

        brand = str(
            cell(
                row,
                brand_index,
            )
            or ""
        ).strip().upper()

        if (
            not content_id
            or not brand
        ):
            skipped_blank += 1
            continue

        if row_validator is not None:
            (
                is_valid,
                reason,
            ) = row_validator(
                row,
                row_number,
                header,
            )

            if not is_valid:
                invalid_rows.append(
                    (
                        row_number,
                        reason,
                    )
                )
                continue

        accepted_rows.append(
            row
        )

        key = (
            brand,
            content_id,
        )

        if key not in groups:
            base = [
                cell(
                    row,
                    index,
                )
                for index in range(
                    len(header)
                )
            ]

            base[
                brand_index
            ] = brand

            base[
                content_id_index
            ] = content_id

            for index in sum_indexes:
                base[index] = 0

            groups[
                key
            ] = base

        base = groups[key]

        for index in sum_indexes:
            base[index] += to_number(
                cell(
                    row,
                    index,
                )
            )

        _update_earliest_date(base, date_index, cell(row, date_index))

    return (
        groups,
        accepted_rows,
        skipped_blank,
        invalid_rows,
    )


# =============================================================================
# Output builders
# =============================================================================

def build_brand_tabs(
    groups,
    i_gmv,
    tab_suffix,
):
    tabs = {}

    for (
        brand,
        _content_id,
    ), row in groups.items():
        jenis = (
            "GMV"
            if to_number(
                cell(
                    row,
                    i_gmv,
                )
            ) > 0
            else "NON GMV"
        )

        title = (
            f"{brand}"
            f"{tab_suffix}"
        )

        tabs.setdefault(
            title,
            [],
        ).append(
            list(row)
            + [jenis]
        )

    for title in tabs:
        tabs[
            title
        ].sort(
            key=lambda row: to_number(
                cell(
                    row,
                    i_gmv,
                )
            ),
            reverse=True,
        )

    return dict(
        sorted(
            tabs.items()
        )
    )


def build_gabungan(tabs):
    rows = []

    for tab_rows in tabs.values():
        rows.extend(
            tab_rows
        )

    return rows


def validate_brand_column(
    header,
    rows,
    brand_header=OUTPUT_BRAND_HEADER,
):
    brand_index = find_header_index(
        header,
        brand_header,
    )

    blank_count = sum(
        1
        for row in rows
        if not str(
            cell(
                row,
                brand_index,
            )
            or ""
        ).strip()
    )

    if blank_count:
        raise ValueError(
            f'Output memiliki {blank_count:,} row '
            f'dengan "{brand_header}" kosong.'
        )

    return brand_index


# =============================================================================
# Reconciliation
# =============================================================================

def reconcile_sum_columns(
    header,
    accepted_rows,
    groups,
    sum_columns,
):
    results = []

    for column in sum_columns:
        index = find_header_index(
            header,
            column,
        )

        source_total = sum(
            to_number(
                cell(
                    row,
                    index,
                )
            )
            for row in accepted_rows
        )

        aggregate_total = sum(
            to_number(
                cell(
                    row,
                    index,
                )
            )
            for row in groups.values()
        )

        results.append(
            {
                "column": column,
                "source": source_total,
                "aggregate": aggregate_total,
                "ok": numbers_equal(
                    source_total,
                    aggregate_total,
                ),
            }
        )

    return results


def reconcile_groups_vs_gabungan(
    header,
    groups,
    gabungan_rows,
    sum_columns,
):
    results = []

    for column in sum_columns:
        index = find_header_index(
            header,
            column,
        )

        group_total = sum(
            to_number(
                cell(
                    row,
                    index,
                )
            )
            for row in groups.values()
        )

        gabungan_total = sum(
            to_number(
                cell(
                    row,
                    index,
                )
            )
            for row in gabungan_rows
        )

        results.append(
            {
                "column": column,
                "groups": group_total,
                "gabungan": gabungan_total,
                "ok": numbers_equal(
                    group_total,
                    gabungan_total,
                ),
            }
        )

    return results


def metric_totals(
    header,
    rows,
    columns,
):
    totals = {}

    for column in columns:
        index = find_header_index(
            header,
            column,
        )

        totals[
            column
        ] = sum(
            to_number(
                cell(
                    row,
                    index,
                )
            )
            for row in rows
        )

    return totals


# =============================================================================
# Incremental monthly storage
# =============================================================================
# =============================================================================
# Incremental monthly storage
# =============================================================================

LOAD_LOG_SHEET = "_LOAD_LOG"

LOAD_LOG_HEADERS = [
    "Pipeline",
    "Load ID",
    "Source Spreadsheet ID",
    "Source Sheet",
    "Source Fingerprint",
    "Action",
    "Source Rows",
    "Stored ALL Rows",
    "Final ALL Rows",
    "Final RANGE Rows",
    "Range Start",
    "Range End",
    "Status",
    "Loaded At",
    "Data Start",
    "Data End",
]

# Schema sebelum kolom "Data Start/Data End" ditambahkan (14 kolom).
LEGACY_LOAD_LOG_HEADERS = LOAD_LOG_HEADERS[:14]

MASTER_META_HEADERS = [
    "_Pipeline",
    "_Load ID",
    "_Source Spreadsheet ID",
    "_Source Fingerprint",
    "_Loaded At",
]

def table_fingerprint(header, rows):
    """Return a stable SHA-256 fingerprint for a source table."""
    digest = hashlib.sha256()

    def update_digest(row):
        payload = json.dumps(
            row,
            ensure_ascii=False,
            separators=(",", ":"),
            default=str,
        )
        digest.update(payload.encode("utf-8"))
        digest.update(b"\n")

    update_digest(header)

    for row in rows:
        update_digest(row)

    return digest.hexdigest()


def resolve_load_id(explicit_load_id, source_spreadsheet_id):
    """
    Resolve the persistent identity of one source pull.

    When LOAD_ID is blank, the source spreadsheet ID is used. If the same
    spreadsheet is overwritten for different pulls, configure a different
    LOAD_ID for each pull so historical contributions remain independent.
    """
    load_id = str(explicit_load_id or "").strip()

    if load_id:
        return load_id

    source_id = str(source_spreadsheet_id or "").strip()

    if not source_id:
        raise ValueError("Source Spreadsheet ID kosong.")

    return source_id


def get_spreadsheet_sheets(service, spreadsheet_id):
    """Return sheet metadata keyed by title."""
    response = execute_with_retry(
        service
        .spreadsheets()
        .get(
            spreadsheetId=spreadsheet_id,
            fields=(
                "sheets("
                "properties("
                "sheetId,"
                "title,"
                "gridProperties"
                ")"
                ")"
            ),
        )
    )

    result = {}

    for sheet in response.get("sheets", []):
        properties = sheet.get("properties", {})
        title = properties.get("title", "")

        if title:
            result[title] = properties

    return result


def read_table_if_exists(
    service,
    spreadsheet_id,
    sheet_title,
    source_range="A:AZ",
):
    """Read a table when the target sheet exists; otherwise return an empty table."""
    sheets = get_spreadsheet_sheets(
        service,
        spreadsheet_id,
    )

    if sheet_title not in sheets:
        return [], []

    response = execute_with_retry(
        service
        .spreadsheets()
        .values()
        .get(
            spreadsheetId=spreadsheet_id,
            range=f"'{sheet_title}'!{source_range}",
            valueRenderOption="UNFORMATTED_VALUE",
            dateTimeRenderOption="FORMATTED_STRING",
        )
    )

    values = response.get("values", [])

    if not values:
        return [], []

    return values[0], values[1:]


def validate_header_layout(actual_header, expected_header, label):
    """Require an existing persistent table to use the expected column layout."""
    if not actual_header:
        return

    actual = [normalize_header(value) for value in actual_header]
    expected = [normalize_header(value) for value in expected_header]

    if actual != expected:
        raise ValueError(
            f'Header {label} tidak sesuai schema pipeline saat ini. '
            "Periksa struktur tab sistem sebelum melanjutkan.\n"
            f"Expected: {expected_header}\n"
            f"Actual  : {actual_header}"
        )



def normalize_load_log_schema(header, rows):
    """Validate and return the current _LOAD_LOG schema."""
    if not header:
        return [], [], False

    if (
        [normalize_header(value) for value in header]
        == [normalize_header(value) for value in LEGACY_LOAD_LOG_HEADERS]
    ):
        width = len(LOAD_LOG_HEADERS)

        return (
            list(LOAD_LOG_HEADERS),
            [
                [cell(row, index) for index in range(width)]
                for row in rows
            ],
            True,  # perlu ditulis ulang sekali (migrasi schema)
        )

    validate_header_layout(
        header,
        LOAD_LOG_HEADERS,
        LOAD_LOG_SHEET,
    )

    return list(header), list(rows), False


def normalize_master_schema(
    header,
    rows,
    expected_header,
):
    """Validate and return the current persistent-master schema."""
    if not header:
        return [], [], False

    validate_header_layout(
        header,
        expected_header,
        "master",
    )

    return list(header), list(rows), False


def master_rows_for_load(
    master_header,
    master_rows,
    output_header,
    pipeline,
    load_id,
):
    """Return only the stored output payload for one Pipeline + Load ID."""
    if not master_header:
        return []

    validate_header_layout(
        master_header,
        build_master_header(output_header),
        "master",
    )

    indexes = master_load_indexes(master_header)
    target_pipeline = str(pipeline).strip().upper()
    target_load_id = str(load_id).strip()
    width = len(output_header)

    return [
        [
            cell(row, index)
            for index in range(width)
        ]
        for row in master_rows
        if (
            str(
                cell(
                    row,
                    indexes["pipeline"],
                )
                or ""
            ).strip().upper()
            == target_pipeline
            and str(
                cell(
                    row,
                    indexes["load_id"],
                )
                or ""
            ).strip()
            == target_load_id
        )
    ]


def output_payload_fingerprint(
    output_header,
    rows,
):
    """Fingerprint one processed load payload for deterministic master QC."""
    return table_fingerprint(
        output_header,
        rows,
    )


def latest_load_record(
    log_header,
    log_rows,
    pipeline,
    load_id,
):
    """Return the latest successful load record for one pipeline and load ID."""
    if not log_header or not log_rows:
        return None

    required = [
        "Pipeline",
        "Load ID",
        "Status",
    ]

    indexes = {
        name: find_header_index(log_header, name)
        for name in required
    }

    matches = []

    for row in log_rows:
        row_pipeline = str(
            cell(row, indexes["Pipeline"]) or ""
        ).strip().upper()

        row_load_id = str(
            cell(row, indexes["Load ID"]) or ""
        ).strip()

        row_status = str(
            cell(row, indexes["Status"]) or ""
        ).strip().upper()

        if (
            row_pipeline == str(pipeline).strip().upper()
            and row_load_id == str(load_id).strip()
            and row_status == "SUCCESS"
        ):
            matches.append(row)

    if not matches:
        return None

    latest = matches[-1]

    return {
        name: cell(
            latest,
            find_header_index(log_header, name),
        )
        for name in log_header
    }


def determine_load_action(
    latest_record,
    source_spreadsheet_id,
    source_fingerprint,
    range_start=None,
    range_end=None,
    force=False,
):
    """
    Classify the current source pull as ADD, SKIP, or REPLACE.

    Identity is based only on:
      - Load ID (selected by caller)
      - Source Spreadsheet ID
      - Source Fingerprint
      - Range Start / Range End

    The caller may later upgrade SKIP to REPAIR if the stored processed
    contribution does not match deterministic processing of the same raw source.
    """
    if latest_record is None:
        return "ADD", "load ID belum pernah dimuat"

    if force:
        return "REPLACE", "--force digunakan"

    same_source = (
        str(
            latest_record.get(
                "Source Spreadsheet ID",
                "",
            )
        ).strip()
        == str(
            source_spreadsheet_id
        ).strip()
    )

    same_fingerprint = (
        str(
            latest_record.get(
                "Source Fingerprint",
                "",
            )
        ).strip()
        == str(
            source_fingerprint
        ).strip()
    )

    same_range = True

    if (
        range_start is not None
        or range_end is not None
    ):
        same_range = (
            str(
                latest_record.get(
                    "Range Start",
                    "",
                )
            ).strip()
            == str(
                range_start
                or ""
            ).strip()
            and str(
                latest_record.get(
                    "Range End",
                    "",
                )
            ).strip()
            == str(
                range_end
                or ""
            ).strip()
        )

    if (
        same_source
        and same_fingerprint
        and same_range
    ):
        return (
            "SKIP",
            "source, fingerprint, dan range sama",
        )

    reasons = []

    if not same_source:
        reasons.append(
            "source berubah"
        )

    if not same_fingerprint:
        reasons.append(
            "isi source berubah"
        )

    if not same_range:
        reasons.append(
            "range berubah"
        )

    return (
        "REPLACE",
        ", ".join(reasons),
    )


def latest_successful_load_records(
    log_header,
    log_rows,
    pipeline,
):
    """Return the latest successful log record for each load ID."""
    if not log_header or not log_rows:
        return {}

    required = [
        "Pipeline",
        "Load ID",
        "Status",
    ]

    indexes = {
        name: find_header_index(log_header, name)
        for name in required
    }

    target_pipeline = str(pipeline).strip().upper()
    records = {}

    for row in log_rows:
        row_pipeline = str(
            cell(row, indexes["Pipeline"]) or ""
        ).strip().upper()
        row_load_id = str(
            cell(row, indexes["Load ID"]) or ""
        ).strip()
        row_status = str(
            cell(row, indexes["Status"]) or ""
        ).strip().upper()

        if (
            row_pipeline != target_pipeline
            or not row_load_id
            or row_status != "SUCCESS"
        ):
            continue

        records[row_load_id] = {
            name: cell(
                row,
                find_header_index(log_header, name),
            )
            for name in log_header
        }

    return records


def merge_date_ranges(ranges):
    """Merge overlapping or adjacent date ranges."""
    if not ranges:
        return []

    normalized = sorted(ranges, key=lambda item: (item[0], item[1]))
    merged = [[normalized[0][0], normalized[0][1]]]

    for start_date, end_date in normalized[1:]:
        current = merged[-1]

        if start_date <= current[1] + timedelta(days=1):
            if end_date > current[1]:
                current[1] = end_date
        else:
            merged.append([start_date, end_date])

    return [tuple(item) for item in merged]


def cumulative_load_ranges(
    log_header,
    log_rows,
    pipeline,
    current_load_id,
    current_range_start,
    current_range_end,
):
    """Build cumulative Date Live/Post date ranges from successful loads.

    Each load ID contributes one active range. Historical records for the same
    load ID are superseded by its latest successful record. The current run is
    applied as an in-memory override before its SUCCESS log row is written.
    """
    records = latest_successful_load_records(
        log_header,
        log_rows,
        pipeline,
    )

    records[str(current_load_id).strip()] = {
        "Range Start": current_range_start,
        "Range End": current_range_end,
    }

    ranges = []

    for load_id, record in records.items():
        start_text = str(record.get("Range Start", "") or "").strip()
        end_text = str(record.get("Range End", "") or "").strip()

        if not start_text or not end_text:
            raise ValueError(
                f"Load ID {load_id!r} tidak memiliki Range Start/End yang valid."
            )

        start_date = parse_config_date(
            start_text,
            f"Range Start untuk load {load_id}",
        )
        end_date = parse_config_date(
            end_text,
            f"Range End untuk load {load_id}",
        )

        if start_date > end_date:
            raise ValueError(
                f"Range load {load_id!r} tidak valid: "
                f"{start_text} > {end_text}."
            )

        ranges.append((start_date, end_date))

    return merge_date_ranges(ranges)


def build_master_header(output_header):
    """Append persistent source-load metadata columns to the master schema."""
    return list(output_header) + list(MASTER_META_HEADERS)


def master_load_indexes(master_header):
    return {
        "pipeline": find_header_index(master_header, "_Pipeline"),
        "load_id": find_header_index(master_header, "_Load ID"),
    }


def attach_master_metadata(
    output_rows,
    pipeline,
    load_id,
    source_spreadsheet_id,
    source_fingerprint,
    loaded_at,
):
    """Add source-load metadata to processed ALL rows before persistent storage."""
    metadata = [
        pipeline,
        load_id,
        source_spreadsheet_id,
        source_fingerprint,
        loaded_at,
    ]

    return [
        list(row) + metadata
        for row in output_rows
    ]


def prepare_master_upsert(
    existing_header,
    existing_rows,
    expected_header,
    new_rows,
    pipeline,
    load_id,
):
    """
    Prepare an idempotent master update.

    A new load ID is appended. If the same load ID already exists, its stored
    contribution is replaced before the monthly result is rebuilt.
    """
    validate_header_layout(
        existing_header,
        expected_header,
        "master",
    )

    if not existing_header:
        return list(new_rows), "REPLACE"

    indexes = master_load_indexes(existing_header)
    target_pipeline = str(pipeline).strip().upper()
    target_load_id = str(load_id).strip()

    kept_rows = []
    matched = 0

    for row in existing_rows:
        is_target = (
            str(cell(row, indexes["pipeline"]) or "").strip().upper()
            == target_pipeline
            and str(cell(row, indexes["load_id"]) or "").strip()
            == target_load_id
        )

        if is_target:
            matched += 1
        else:
            kept_rows.append(row)

    if matched == 0:
        return list(existing_rows) + list(new_rows), "APPEND"

    return kept_rows + list(new_rows), "REPLACE"


def strip_master_metadata(master_header, master_rows, output_header):
    """Return the output portion of persistent master rows."""
    validate_header_layout(
        master_header,
        build_master_header(output_header),
        "master",
    )

    width = len(output_header)

    return [
        [cell(row, index) for index in range(width)]
        for row in master_rows
    ]


def validate_group_field_consistency(
    header,
    rows,
    key_headers,
    field_header,
    max_examples=20,
):
    """
    Validate that one content key does not carry conflicting reporting dates.

    The reporting field is expected to be stable for a content ID across all
    stored source pulls. Blank values are ignored when a non-blank value exists.
    """
    key_indexes = [
        find_header_index(header, name)
        for name in key_headers
    ]
    field_index = find_header_index(header, field_header)

    values_by_key = defaultdict(set)

    for row in rows:
        key = tuple(
            str(cell(row, index) or "").strip()
            for index in key_indexes
        )
        value = str(cell(row, field_index) or "").strip()

        if all(key) and value:
            values_by_key[key].add(value)

    conflicts = [
        (key, sorted(values))
        for key, values in values_by_key.items()
        if len(values) > 1
    ]

    if conflicts:
        examples = "\n".join(
            f"  {key}: {values}"
            for key, values in conflicts[:max_examples]
        )

        extra = (
            f"\n  ... {len(conflicts) - max_examples:,} conflict lainnya."
            if len(conflicts) > max_examples
            else ""
        )

        raise ValueError(
            f'Nilai "{field_header}" tidak konsisten untuk content key yang sama:\n'
            + examples
            + extra
        )

    return len(values_by_key)


# =============================================================================
# Spreadsheet capacity
# =============================================================================

def spreadsheet_grid_info(service, spreadsheet_id):
    sheets = get_spreadsheet_sheets(
        service,
        spreadsheet_id,
    )

    info = {}
    total_cells = 0

    for title, properties in sheets.items():
        grid = properties.get("gridProperties", {})
        row_count = grid.get("rowCount", 0)
        column_count = grid.get("columnCount", 0)
        cells = row_count * column_count
        total_cells += cells

        info[title] = {
            "sheetId": properties["sheetId"],
            "rowCount": row_count,
            "columnCount": column_count,
            "cells": cells,
        }

    return info, total_cells


def validate_target_capacity(
    service,
    spreadsheet_id,
    target_dimensions,
    label="OUTPUT",
    max_cells=MAX_CELLS,
):
    """Validate projected capacity when selected sheets are resized or created."""
    sheet_info, existing_cells = spreadsheet_grid_info(
        service,
        spreadsheet_id,
    )

    replaced_cells = sum(
        sheet_info.get(title, {}).get("cells", 0)
        for title in target_dimensions
    )

    target_cells = sum(
        max(int(rows), 2) * max(int(columns), 1)
        for rows, columns in target_dimensions.values()
    )

    projected_cells = (
        existing_cells
        - replaced_cells
        + target_cells
    )

    print(
        f"[{label}] sel sekarang {existing_cells:,} "
        f"- grid target lama {replaced_cells:,} "
        f"+ grid target baru {target_cells:,} "
        f"= {projected_cells:,} "
        f"(batas {max_cells:,})."
    )

    if projected_cells > max_cells:
        raise SystemExit(
            f"[{label}] Projected cell melebihi {max_cells:,}. "
            "Tidak ada output yang ditulis."
        )

    return existing_cells, target_cells, projected_cells


# =============================================================================
# Incremental writer
# =============================================================================

WRITE_BATCH_RANGES = 4


def _sheet_id(service, spreadsheet_id, sheet_title):
    sheets = get_spreadsheet_sheets(
        service,
        spreadsheet_id,
    )

    properties = sheets.get(sheet_title)

    return None if properties is None else properties["sheetId"]


def ensure_sheet_dimensions(
    service,
    spreadsheet_id,
    sheet_title,
    row_count,
    column_count,
    frozen_row_count=1,
):
    """Create a sheet when needed and set its grid to the requested dimensions."""
    row_count = max(int(row_count), 2)
    column_count = max(int(column_count), 1)

    sheet_id = _sheet_id(
        service,
        spreadsheet_id,
        sheet_title,
    )

    if sheet_id is None:
        response = execute_with_retry(
            service
            .spreadsheets()
            .batchUpdate(
                spreadsheetId=spreadsheet_id,
                body={
                    "requests": [
                        {
                            "addSheet": {
                                "properties": {
                                    "title": sheet_title,
                                    "gridProperties": {
                                        "rowCount": row_count,
                                        "columnCount": column_count,
                                        "frozenRowCount": frozen_row_count,
                                    },
                                }
                            }
                        }
                    ]
                },
            )
        )

        return response["replies"][0]["addSheet"]["properties"]["sheetId"]

    execute_with_retry(
        service
        .spreadsheets()
        .batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={
                "requests": [
                    {
                        "updateSheetProperties": {
                            "properties": {
                                "sheetId": sheet_id,
                                "gridProperties": {
                                    "rowCount": row_count,
                                    "columnCount": column_count,
                                    "frozenRowCount": frozen_row_count,
                                },
                            },
                            "fields": (
                                "gridProperties.rowCount,"
                                "gridProperties.columnCount,"
                                "gridProperties.frozenRowCount"
                            ),
                        }
                    }
                ]
            },
        )
    )

    return sheet_id


def clear_sheet_values(service, spreadsheet_id, sheet_title):
    execute_with_retry(
        service
        .spreadsheets()
        .values()
        .clear(
            spreadsheetId=spreadsheet_id,
            range=f"'{sheet_title}'",
            body={},
        )
    )


def _write_value_ranges(
    service,
    spreadsheet_id,
    value_ranges,
    batch_range_count=WRITE_BATCH_RANGES,
):
    """Write several non-overlapping ranges per Sheets API request."""
    if not value_ranges:
        return

    for start in range(
        0,
        len(value_ranges),
        batch_range_count,
    ):
        batch = value_ranges[
            start:
            start + batch_range_count
        ]

        execute_with_retry(
            service
            .spreadsheets()
            .values()
            .batchUpdate(
                spreadsheetId=spreadsheet_id,
                body={
                    "valueInputOption": "RAW",
                    "data": batch,
                },
            )
        )


def write_rows_at(
    service,
    spreadsheet_id,
    sheet_title,
    rows,
    start_row,
    chunk_rows=WRITE_CHUNK_ROWS,
):
    """Write rows from a known starting row using batched range updates."""
    if not rows:
        return

    value_ranges = []

    for offset in range(0, len(rows), chunk_rows):
        chunk = rows[offset:offset + chunk_rows]
        row_number = start_row + offset

        value_ranges.append(
            {
                "range": f"'{sheet_title}'!A{row_number}",
                "values": chunk,
            }
        )

    _write_value_ranges(
        service,
        spreadsheet_id,
        value_ranges,
    )


def replace_table(
    service,
    spreadsheet_id,
    sheet_title,
    header,
    rows,
):
    """Replace one sheet table while preserving the sheet name."""
    ensure_sheet_dimensions(
        service,
        spreadsheet_id,
        sheet_title,
        len(rows) + 1,
        len(header),
    )

    clear_sheet_values(
        service,
        spreadsheet_id,
        sheet_title,
    )

    write_rows_at(
        service,
        spreadsheet_id,
        sheet_title,
        [list(header)] + list(rows),
        start_row=1,
    )


def append_table_rows(
    service,
    spreadsheet_id,
    sheet_title,
    header,
    existing_row_count,
    new_rows,
):
    """Append rows to an existing persistent table without rewriting prior data."""
    if not new_rows:
        return

    sheet_id = _sheet_id(
        service,
        spreadsheet_id,
        sheet_title,
    )

    if sheet_id is None or existing_row_count == 0:
        replace_table(
            service,
            spreadsheet_id,
            sheet_title,
            header,
            new_rows,
        )
        return

    ensure_sheet_dimensions(
        service,
        spreadsheet_id,
        sheet_title,
        existing_row_count + len(new_rows) + 1,
        len(header),
    )

    write_rows_at(
        service,
        spreadsheet_id,
        sheet_title,
        new_rows,
        start_row=existing_row_count + 2,
    )


def upsert_master_table(
    service,
    spreadsheet_id,
    sheet_title,
    expected_header,
    existing_header,
    existing_rows,
    merged_rows,
    source_rows,
    write_mode,
    force_replace=False,
):
    """Persist master using append when safe and replace when required."""
    if (
        write_mode == "APPEND"
        and existing_header
        and not force_replace
    ):
        append_table_rows(
            service,
            spreadsheet_id,
            sheet_title,
            expected_header,
            len(existing_rows),
            source_rows,
        )
    else:
        replace_table(
            service,
            spreadsheet_id,
            sheet_title,
            expected_header,
            merged_rows,
        )


def append_load_log(
    service,
    spreadsheet_id,
    row,
    existing_header,
    existing_rows,
    force_replace=False,
):
    """Append one successful run record; replace once when migrating legacy schema."""
    validate_header_layout(
        existing_header,
        LOAD_LOG_HEADERS,
        LOAD_LOG_SHEET,
    )

    if force_replace:
        replace_table(
            service,
            spreadsheet_id,
            LOAD_LOG_SHEET,
            LOAD_LOG_HEADERS,
            list(existing_rows) + [list(row)],
        )
        return

    append_table_rows(
        service,
        spreadsheet_id,
        LOAD_LOG_SHEET,
        LOAD_LOG_HEADERS,
        len(existing_rows),
        [row],
    )


# =============================================================================
# Monthly output rebuild
# =============================================================================

def build_combined_from_groups(
    groups,
    header,
    gmv_column,
):
    """Build a combined output using the same brand ordering as the legacy flow."""
    gmv_index = find_header_index(
        header,
        gmv_column,
    )

    brand_tabs = build_brand_tabs(
        groups,
        gmv_index,
        "",
    )

    return build_gabungan(brand_tabs)


def rebuild_monthly_output(
    master_header,
    master_rows,
    output_header,
    content_id_header,
    date_header,
    sum_columns,
    gmv_column,
    brand_header=OUTPUT_BRAND_HEADER,
    jenis_header="Jenis",
    consistent_fields=None,
):
    """
    Re-aggregate all stored source pulls to the monthly Brand + Content ID grain.

    RANGE filtering is intentionally not performed here. The caller must first
    complete this monthly aggregation and only then filter the resulting groups
    by Date Live or Post date.
    """
    if not output_header:
        raise ValueError("Output header kosong.")

    if normalize_header(output_header[-1]) != normalize_header(jenis_header):
        raise ValueError(
            f'Kolom terakhir output harus "{jenis_header}".'
        )

    stored_rows = strip_master_metadata(
        master_header,
        master_rows,
        output_header,
    )

    base_header = list(output_header[:-1])
    base_width = len(base_header)

    aggregation_rows = [
        [cell(row, index) for index in range(base_width)]
        for row in stored_rows
    ]

    for field_header in consistent_fields or []:
        validate_group_field_consistency(
            base_header,
            aggregation_rows,
            [brand_header, content_id_header],
            field_header,
        )

    (
        groups,
        accepted_rows,
        skipped_blank,
        invalid_rows,
    ) = aggregate_brand_content(
        base_header,
        aggregation_rows,
        content_id_header=content_id_header,
        date_header=date_header,
        sum_columns=sum_columns,
        brand_header=brand_header,
        row_validator=None,
    )

    if invalid_rows:
        raise ValueError(
            "Monthly rebuild menghasilkan row invalid yang tidak diharapkan."
        )

    combined = build_combined_from_groups(
        groups,
        base_header,
        gmv_column,
    )

    validate_brand_column(
        output_header,
        combined,
        brand_header,
    )

    checks = reconcile_sum_columns(
        base_header,
        accepted_rows,
        groups,
        sum_columns,
    )

    if not all(item["ok"] for item in checks):
        raise ValueError(
            "Monthly reconciliation gagal."
        )

    return {
        "groups": groups,
        "rows": combined,
        "accepted_rows": accepted_rows,
        "skipped_blank": skipped_blank,
        "checks": checks,
        "base_header": base_header,
    }
def clean_label(value):
    """Trim and collapse internal whitespace."""
    return re.sub(r"\s+", " ", str(value or "").strip())


def label_key(value):
    """Case-insensitive comparison key for brand / sub product labels."""
    return clean_label(value).casefold()


def canonical_label(label, registry, variants=None):
    """
    Return one stable spelling per case-insensitive label.

    The first spelling seen wins. Other spellings are recorded in `variants`.
    """
    label = clean_label(label)
    key = label.casefold()
    first = registry.setdefault(key, label)

    if variants is not None and first != label:
        variants[key].add(label)

    return first


# =============================================================================
# BW Product ID -> Brand + Sub Product mapping
# =============================================================================

def detect_bw_mapping_header(
    service,
    spreadsheet_id,
    sheet_name=None,
    pid_header=BW_MAPPING_PRODUCT_ID_HEADER,
    brand_header=BW_MAPPING_BRAND_HEADER,
    sub_product_header=BW_MAPPING_SUB_PRODUCT_HEADER,
):
    """
    Detect the mapping sheet/header containing:
        Product ID | Brand | Sub Product

    When sheet_name=None, every sheet in the mapping spreadsheet is scanned.
    """
    sheets = get_spreadsheet_sheets(
        service,
        spreadsheet_id,
    )

    if not sheets:
        raise ValueError(
            "Spreadsheet mapping BW tidak memiliki sheet."
        )

    if sheet_name:
        if sheet_name not in sheets:
            raise ValueError(
                f'MAPPING_SHEET "{sheet_name}" tidak ditemukan. '
                f"Available: {sorted(sheets)}"
            )
        titles = [sheet_name]
    else:
        titles = list(sheets)

    targets = {
        "product_id": normalize_header(pid_header),
        "brand": normalize_header(brand_header),
        "sub_product": normalize_header(sub_product_header),
    }

    for title in titles:
        response = execute_with_retry(
            service
            .spreadsheets()
            .values()
            .get(
                spreadsheetId=spreadsheet_id,
                range=f"'{title}'!A1:AZ20",
                valueRenderOption="FORMATTED_VALUE",
            )
        )

        for row_number, row in enumerate(
            response.get("values", []),
            start=1,
        ):
            normalized = [
                normalize_header(value)
                for value in row
            ]

            found = {}

            for key, target in targets.items():
                found[key] = next(
                    (
                        index
                        for index, value
                        in enumerate(normalized)
                        if value == target
                    ),
                    None,
                )

            if all(
                found[key] is not None
                for key in (
                    "product_id",
                    "brand",
                    "sub_product",
                )
            ):
                return {
                    "sheet_title": title,
                    "header_row": row_number,
                    "product_id_index": found["product_id"],
                    "brand_index": found["brand"],
                    "sub_product_index": found["sub_product"],
                }

    raise ValueError(
        "Header mapping BW tidak ditemukan. "
        f'Harus ada "{pid_header}", "{brand_header}", '
        f'dan "{sub_product_header}".'
    )


def build_bw_mapping_from_rows(
    rows,
    product_id_index,
    brand_index,
    sub_product_index,
):
    """
    Build Product ID -> Brand + Sub Product from raw table rows.

    Shared by the Google Sheets loader and local-file runners. Brand and Sub
    Product spellings that differ only by case / spacing are merged into one
    label. A duplicate Product ID with a genuinely different Brand + Sub Product
    raises MASTER BW MAPPING CONFLICT.
    """
    mapping = {}
    conflicts = defaultdict(set)

    blank_product_id = 0
    blank_brand = 0
    blank_sub_product = 0
    duplicate_same = 0

    brand_labels = {}
    sub_labels = {}
    brand_variants = defaultdict(set)
    sub_variants = defaultdict(set)

    for row in rows:
        product_id = normalize_id(
            cell(
                row,
                product_id_index,
            )
        )

        brand = str(
            cell(
                row,
                brand_index,
            )
            or ""
        ).strip()

        sub_product = str(
            cell(
                row,
                sub_product_index,
            )
            or ""
        ).strip()

        if not product_id:
            blank_product_id += 1
            continue

        if not brand:
            blank_brand += 1
            brand = MISSING_BRAND

        if not sub_product:
            blank_sub_product += 1
            sub_product = MISSING_SUB_PRODUCT

        # "Vaseline" / "VASELINE" / "vaseline " -> satu label (huruf pertama yang muncul)
        brand = canonical_label(brand, brand_labels, brand_variants)
        sub_product = canonical_label(sub_product, sub_labels, sub_variants)

        mapped_value = (
            brand,
            sub_product,
        )

        if product_id in mapping:
            existing = (
                mapping[product_id]["brand"],
                mapping[product_id]["sub_product"],
            )

            if existing == mapped_value:
                duplicate_same += 1
                continue

            conflicts[product_id].update(
                {
                    existing,
                    mapped_value,
                }
            )
            continue

        mapping[product_id] = {
            "brand": brand,
            "sub_product": sub_product,
        }

    if conflicts:
        examples = []

        for product_id, values in list(
            conflicts.items()
        )[:20]:
            shown = sorted(
                f"{brand} | {sub_product}"
                for brand, sub_product in values
            )
            examples.append(
                f"{product_id}: {shown}"
            )

        raise ValueError(
            "MASTER BW MAPPING CONFLICT:\n  "
            + "\n  ".join(examples)
        )

    return mapping, {
        "mapping_count": len(mapping),
        "blank_product_id": blank_product_id,
        "blank_brand": blank_brand,
        "blank_sub_product": blank_sub_product,
        "duplicate_same": duplicate_same,
        "brand_case_variants": {
            brand_labels[key]: sorted(values)
            for key, values in brand_variants.items()
        },
        "sub_product_case_variants": {
            sub_labels[key]: sorted(values)
            for key, values in sub_variants.items()
        },
    }


def load_product_brand_subproduct_mapping(
    service,
    spreadsheet_id,
    sheet_name=None,
    pid_header=BW_MAPPING_PRODUCT_ID_HEADER,
    brand_header=BW_MAPPING_BRAND_HEADER,
    sub_product_header=BW_MAPPING_SUB_PRODUCT_HEADER,
):
    """
    Load Product ID -> Brand + Sub Product.

    Duplicate Product ID is allowed only when Brand + Sub Product is identical.
    A conflicting duplicate immediately raises an error.
    """
    layout = detect_bw_mapping_header(
        service,
        spreadsheet_id,
        sheet_name=sheet_name,
        pid_header=pid_header,
        brand_header=brand_header,
        sub_product_header=sub_product_header,
    )

    last_column = column_letter(
        max(
            layout["product_id_index"],
            layout["brand_index"],
            layout["sub_product_index"],
        )
    )

    response = execute_with_retry(
        service
        .spreadsheets()
        .values()
        .get(
            spreadsheetId=spreadsheet_id,
            range=(
                f"'{layout['sheet_title']}'!"
                f"A{layout['header_row'] + 1}:{last_column}"
            ),
            valueRenderOption="FORMATTED_VALUE",
        )
    )

    mapping, stats = build_bw_mapping_from_rows(
        response.get("values", []),
        layout["product_id_index"],
        layout["brand_index"],
        layout["sub_product_index"],
    )

    stats["sheet_title"] = layout["sheet_title"]
    stats["header_row"] = layout["header_row"]

    return mapping, stats

def add_brand_subproduct_from_product_id(
    header,
    rows,
    mapping,
    product_id_header=SOURCE_PRODUCT_ID_HEADER,
    brand_header=OUTPUT_BRAND_HEADER,
    sub_product_header=OUTPUT_SUB_PRODUCT_HEADER,
    missing_brand=MISSING_BRAND,
    missing_sub_product=MISSING_SUB_PRODUCT,
):
    """
    BW classification is mapping-only:

        source Product ID
            -> mapping Product ID
            -> Brand
            -> Sub Product

    Product name is not used to derive Brand/Sub Product.
    """
    product_id_index = find_header_index(
        header,
        product_id_header,
    )

    output_header = list(header)

    try:
        brand_index = find_header_index(
            output_header,
            brand_header,
        )
        output_header[brand_index] = brand_header
    except ValueError:
        brand_index = len(output_header)
        output_header.append(brand_header)

    try:
        sub_product_index = find_header_index(
            output_header,
            sub_product_header,
        )
        output_header[sub_product_index] = sub_product_header
    except ValueError:
        sub_product_index = len(output_header)
        output_header.append(sub_product_header)

    output_rows = []
    status = Counter()
    missing_product_ids = Counter()
    mapped_brand_counts = Counter()
    mapped_sub_product_counts = Counter()

    for row in rows:
        output_row = [
            cell(row, index)
            for index in range(len(output_header))
        ]

        product_id = normalize_id(
            cell(
                row,
                product_id_index,
            )
        )

        mapped = (
            mapping.get(product_id)
            if product_id
            else None
        )

        if mapped is None:
            brand = missing_brand
            sub_product = missing_sub_product

            if product_id:
                status["PRODUCT_ID_NOT_FOUND"] += 1
                missing_product_ids[product_id] += 1
            else:
                status["BLANK_PRODUCT_ID"] += 1
                missing_product_ids["<BLANK_PRODUCT_ID>"] += 1

        else:
            brand = str(
                mapped.get("brand", "")
                or missing_brand
            ).strip()

            sub_product = str(
                mapped.get("sub_product", "")
                or missing_sub_product
            ).strip()

            status["MAPPED"] += 1

            if brand == missing_brand:
                status["MAPPED_BUT_BRAND_MISSING"] += 1

            if sub_product == missing_sub_product:
                status["MAPPED_BUT_SUB_PRODUCT_MISSING"] += 1

        output_row[brand_index] = brand
        output_row[sub_product_index] = sub_product

        mapped_brand_counts[brand] += 1
        mapped_sub_product_counts[sub_product] += 1

        output_rows.append(output_row)

    return (
        output_header,
        output_rows,
        status,
        missing_product_ids,
        mapped_brand_counts,
        mapped_sub_product_counts,
    )




# =============================================================================
# BW aggregation: Brand + Content ID (FIRST PID representative)
# =============================================================================

def make_tap_row_validator(date_header="Date", id_headers=()):
    """
    Row validator for raw TAP exports.

    Rejects the "Summary" total row and rows whose ID columns hold "-" / "--"
    / blank, so they are never aggregated as if they were a real content.
    """
    cache = {}

    def validator(row, row_number, header):
        if "date" not in cache:
            cache["date"] = find_header_index(header, date_header)
            cache["ids"] = [
                (name, find_header_index(header, name))
                for name in id_headers
            ]

        if str(cell(row, cache["date"]) or "").strip().lower() == "summary":
            return False, "baris Summary"

        for name, index in cache["ids"]:
            value = str(cell(row, index) or "").strip()

            if value in {"", "-", "--"}:
                return False, f"{name} kosong ({value!r})"

        return True, ""

    return validator


def dedupe_source_rows(header, rows, key_headers):
    """
    Drop rows that are exact repeats of an earlier row with the same natural key.

    TAP rows are daily: one row per Date + content + product (+ campaign /
    creator). The same natural key with identical values is an export repeat and
    must not be summed twice. The same key with DIFFERENT values is kept and
    reported in `conflicts` so the caller can warn.

    Returns (kept_rows, dropped_count, conflicts).
    """
    key_indexes = [
        find_header_index(header, name)
        for name in key_headers
    ]
    width = len(header)

    seen = {}
    kept = []
    dropped = 0
    conflicts = []

    for row in rows:
        key = tuple(
            str(cell(row, index) if cell(row, index) is not None else "").strip()
            for index in key_indexes
        )

        if not any(key):
            kept.append(row)
            continue

        signature = tuple(
            str(cell(row, index) if cell(row, index) is not None else "").strip()
            for index in range(width)
        )

        if key not in seen:
            seen[key] = signature
            kept.append(row)
        elif seen[key] == signature:
            dropped += 1
        else:
            conflicts.append(key)
            kept.append(row)

    return kept, dropped, conflicts


def source_date_window(header, rows, date_header="Date"):
    """Return (min, max) ISO dates found in the raw pull's daily Date column."""
    index = find_header_index(header, date_header)
    values = []

    for row in rows:
        text = str(cell(row, index) or "").strip()

        if not text or text.lower() == "summary" or text in {"-", "--"}:
            continue

        values.append(_canonical_date_text(text)[:10])

    if not values:
        return "", ""

    return min(values), max(values)


def find_overlapping_loads(
    log_header,
    log_rows,
    pipeline,
    current_load_id,
    data_start,
    data_end,
):
    """
    Find OTHER load IDs whose pulled Date window overlaps the current pull.

    TAP rows are daily, so a date that appears in two different loads is counted
    twice when the monthly master is re-aggregated. Loads without a recorded
    window (legacy log rows) are skipped.
    """
    if not data_start or not data_end:
        return []

    records = latest_successful_load_records(
        log_header,
        log_rows,
        pipeline,
    )

    overlaps = []

    for load_id, record in records.items():
        if load_id == str(current_load_id).strip():
            continue

        other_start = str(record.get("Data Start", "") or "").strip()
        other_end = str(record.get("Data End", "") or "").strip()

        if not other_start or not other_end:
            continue

        if data_start <= other_end and other_start <= data_end:
            overlaps.append(
                {
                    "load_id": load_id,
                    "data_start": other_start,
                    "data_end": other_end,
                }
            )

    return overlaps


def enforce_no_overlapping_loads(
    log_header,
    log_rows,
    pipeline,
    current_load_id,
    data_start,
    data_end,
    strict=True,
):
    """Raise (strict) or print a warning when another load covers the same dates."""
    overlaps = find_overlapping_loads(
        log_header,
        log_rows,
        pipeline,
        current_load_id,
        data_start,
        data_end,
    )

    if not overlaps:
        return []

    detail = "\n".join(
        f"  load {item['load_id']!r}: {item['data_start']} s/d {item['data_end']}"
        for item in overlaps
    )

    message = (
        f"Pull {data_start} s/d {data_end} (load {current_load_id!r}) "
        "tumpang tindih tanggal dengan load lain sehingga baris harian yang sama "
        "akan terhitung dua kali:\n"
        + detail
        + "\nGunakan Load ID yang sama (REPLACE) atau pull tanpa tanggal yang bertumpuk."
    )

    if strict:
        raise ValueError(message)

    print("PERINGATAN: " + message)
    return overlaps


def aggregation_policy_report(
    header,
    key_columns,
    sum_columns,
    per_date_sum_columns=None,
    max_columns=None,
    ratio_columns=None,
):
    """List how every output column is aggregated, so nothing is silent."""
    explicit = {
        normalize_header(name): "key"
        for name in key_columns
    }

    for name in sum_columns:
        explicit[normalize_header(name)] = "sum"

    for name in per_date_sum_columns or []:
        explicit[normalize_header(name)] = "per_date_sum"

    for name in max_columns or []:
        explicit[normalize_header(name)] = "max"

    for name in (ratio_columns or {}):
        explicit[normalize_header(name)] = "ratio"

    report = defaultdict(list)

    for name in header:
        report[explicit.get(normalize_header(name), "first_row")].append(name)

    return dict(report)


def aggregate_bw_brand_content(
    header,
    rows,
    content_id_header,
    date_header,
    sum_columns,
    brand_header=OUTPUT_BRAND_HEADER,
    product_id_header=SOURCE_PRODUCT_ID_HEADER,
    sub_product_header=OUTPUT_SUB_PRODUCT_HEADER,
    row_validator=None,
    per_date_sum_columns=None,
    daily_header=None,
    max_columns=None,
    ratio_columns=None,
):
    """
    Aggregate TAP BW rows to exactly one row per (Brand, Content ID).

    Rules:
    - Product ID = FIRST accepted Product ID in the Brand + Content ID group.
    - Sub Product = Sub Product from that same FIRST PID row.
    - sum_columns: additive metrics, summed across every source row.
    - per_date_sum_columns: metrics repeated on every product row of the same
      content and day (e.g. views / likes). Per day the value is taken once
      (max), then summed over days. Needs daily_header (the raw "Date" column).
    - max_columns: metrics where the largest value is wanted (e.g. followers).
    - ratio_columns: {column: (numerator, denominator, multiplier)} recomputed
      from the final sums, never summed (e.g. AOV = GMV / Orders).
    - Date (date_header) keeps the earliest value, compared as normalized dates.
    - Everything else keeps the first accepted row (see aggregation_policy_report).
    - Brand labels are matched case-insensitively.
    """
    per_date_sum_columns = list(per_date_sum_columns or [])
    max_columns = list(max_columns or [])
    ratio_columns = dict(ratio_columns or {})

    if per_date_sum_columns and not daily_header:
        raise ValueError(
            "per_date_sum_columns membutuhkan daily_header (kolom tanggal harian, mis. 'Date')."
        )

    overlap = (
        {normalize_header(c) for c in sum_columns}
        & {normalize_header(c) for c in per_date_sum_columns}
        | {normalize_header(c) for c in sum_columns}
        & {normalize_header(c) for c in max_columns}
        | {normalize_header(c) for c in sum_columns}
        & {normalize_header(c) for c in ratio_columns}
    )

    if overlap:
        raise ValueError(
            f"Kolom tidak boleh punya dua aturan agregasi: {sorted(overlap)}"
        )

    content_id_index = find_header_index(header, content_id_header)
    brand_index = find_header_index(header, brand_header)
    product_id_index = find_header_index(header, product_id_header)
    sub_product_index = find_header_index(header, sub_product_header)
    date_index = find_header_index(header, date_header)
    daily_index = (
        find_header_index(header, daily_header)
        if daily_header
        else None
    )

    sum_indexes = [find_header_index(header, c) for c in sum_columns]
    per_date_indexes = [find_header_index(header, c) for c in per_date_sum_columns]
    max_indexes = [find_header_index(header, c) for c in max_columns]
    ratio_indexes = {
        find_header_index(header, name): (
            find_header_index(header, numerator),
            find_header_index(header, denominator),
            multiplier,
        )
        for name, (numerator, denominator, multiplier) in ratio_columns.items()
    }

    groups = {}
    daily_slots = {}
    brand_labels = {}
    accepted_rows = []
    skipped_blank = 0
    invalid_rows = []

    for row_number, row in enumerate(rows, start=2):
        content_id = normalize_id(cell(row, content_id_index))
        brand = canonical_label(cell(row, brand_index), brand_labels)

        if not content_id or not brand:
            skipped_blank += 1
            continue

        if row_validator is not None:
            is_valid, reason = row_validator(row, row_number, header)
            if not is_valid:
                invalid_rows.append((row_number, reason))
                continue

        accepted_rows.append(row)
        key = (brand, content_id)

        if key not in groups:
            base = [cell(row, index) for index in range(len(header))]
            base[brand_index] = brand
            base[content_id_index] = content_id
            base[product_id_index] = normalize_id(cell(row, product_id_index))
            base[sub_product_index] = (
                str(cell(row, sub_product_index) or "").strip()
                or MISSING_SUB_PRODUCT
            )
            for index in sum_indexes + per_date_indexes + max_indexes:
                base[index] = 0
            for index in ratio_indexes:
                base[index] = 0
            groups[key] = base

        base = groups[key]

        for index in sum_indexes:
            base[index] += to_number(cell(row, index))

        if per_date_indexes:
            day = _canonical_date_text(cell(row, daily_index))[:10]
            slots = daily_slots.setdefault(key, {})

            for index in per_date_indexes:
                value = to_number(cell(row, index))
                slot_key = (index, day)

                if slot_key not in slots or value > slots[slot_key]:
                    slots[slot_key] = value

        for index in max_indexes:
            value = to_number(cell(row, index))
            if value > base[index]:
                base[index] = value

        _update_earliest_date(base, date_index, cell(row, date_index))

    for key, slots in daily_slots.items():
        base = groups[key]
        for index in per_date_indexes:
            base[index] = sum(
                value
                for (slot_index, _day), value in slots.items()
                if slot_index == index
            )

    for base in groups.values():
        for index, (numerator, denominator, multiplier) in ratio_indexes.items():
            divisor = to_number(base[denominator])
            base[index] = (
                to_number(base[numerator]) / divisor * multiplier
                if divisor
                else 0
            )

    return groups, accepted_rows, skipped_blank, invalid_rows


def validate_unique_brand_content(
    header,
    rows,
    content_id_header,
    brand_header=OUTPUT_BRAND_HEADER,
):
    """Raise when output contains duplicate Brand + Content ID rows (brand case-insensitive)."""
    content_id_index = find_header_index(header, content_id_header)
    brand_index = find_header_index(header, brand_header)
    seen = set()
    duplicates = []

    for row_number, row in enumerate(rows, start=2):
        content_id = normalize_id(cell(row, content_id_index))
        brand = clean_label(cell(row, brand_index))
        if not content_id or not brand:
            continue
        key = (brand.casefold(), content_id)
        if key in seen:
            duplicates.append((row_number, brand, content_id))
        else:
            seen.add(key)

    if duplicates:
        shown = "\n".join(
            f"  row {row_number}: {brand} | {content_id}"
            for row_number, brand, content_id in duplicates[:20]
        )
        extra = (
            f"\n  ... {len(duplicates) - 20:,} duplicate lainnya."
            if len(duplicates) > 20 else ""
        )
        raise ValueError(
            "Output BW masih memiliki duplicate Brand + Content ID:\n"
            + shown + extra
        )

    return len(seen)


# =============================================================================
# BW aggregation: Brand + Sub Product + Content ID
# =============================================================================

def aggregate_brand_subproduct_content(
    header,
    rows,
    content_id_header,
    date_header,
    sum_columns,
    brand_header=OUTPUT_BRAND_HEADER,
    sub_product_header=OUTPUT_SUB_PRODUCT_HEADER,
    row_validator=None,
):
    """
    Aggregate source rows to:
        (Brand, Sub Product, Content ID)

    Additive metrics are summed.
    Date keeps the earliest value.
    Other non-additive fields keep the first accepted row.
    """
    content_id_index = find_header_index(
        header,
        content_id_header,
    )

    brand_index = find_header_index(
        header,
        brand_header,
    )

    sub_product_index = find_header_index(
        header,
        sub_product_header,
    )

    date_index = find_header_index(
        header,
        date_header,
    )

    sum_indexes = [
        find_header_index(
            header,
            column,
        )
        for column in sum_columns
    ]

    groups = {}
    accepted_rows = []
    skipped_blank = 0
    invalid_rows = []

    for row_number, row in enumerate(
        rows,
        start=2,
    ):
        content_id = normalize_id(
            cell(
                row,
                content_id_index,
            )
        )

        brand = str(
            cell(
                row,
                brand_index,
            )
            or ""
        ).strip()

        sub_product = str(
            cell(
                row,
                sub_product_index,
            )
            or ""
        ).strip()

        if (
            not content_id
            or not brand
            or not sub_product
        ):
            skipped_blank += 1
            continue

        if row_validator is not None:
            is_valid, reason = row_validator(
                row,
                row_number,
                header,
            )

            if not is_valid:
                invalid_rows.append(
                    (
                        row_number,
                        reason,
                    )
                )
                continue

        accepted_rows.append(row)

        key = (
            brand,
            sub_product,
            content_id,
        )

        if key not in groups:
            base = [
                cell(
                    row,
                    index,
                )
                for index in range(
                    len(header)
                )
            ]

            base[brand_index] = brand
            base[sub_product_index] = sub_product
            base[content_id_index] = content_id

            for index in sum_indexes:
                base[index] = 0

            groups[key] = base

        base = groups[key]

        for index in sum_indexes:
            base[index] += to_number(
                cell(
                    row,
                    index,
                )
            )

        _update_earliest_date(base, date_index, cell(row, date_index))

    return (
        groups,
        accepted_rows,
        skipped_blank,
        invalid_rows,
    )


def build_bw_brand_tabs(
    groups,
    i_gmv,
    tab_suffix,
):
    """
    Build one tab per Brand.

    Sub Product remains in the table only as representative metadata from the FIRST PID.
    """
    tabs = {}

    for key, row in groups.items():
        brand = key[0]

        jenis = (
            "GMV"
            if to_number(
                cell(
                    row,
                    i_gmv,
                )
            ) > 0
            else "NON GMV"
        )

        title = f"{brand}{tab_suffix}"

        tabs.setdefault(
            title,
            [],
        ).append(
            list(row)
            + [jenis]
        )

    for title in tabs:
        tabs[title].sort(
            key=lambda row: to_number(
                cell(
                    row,
                    i_gmv,
                )
            ),
            reverse=True,
        )

    return dict(
        sorted(
            tabs.items()
        )
    )


def build_bw_combined_from_groups(
    groups,
    header,
    gmv_column,
):
    gmv_index = find_header_index(
        header,
        gmv_column,
    )

    brand_tabs = build_bw_brand_tabs(
        groups,
        gmv_index,
        "",
    )

    return build_gabungan(
        brand_tabs
    )


def validate_bw_classification_columns(
    header,
    rows,
    brand_header=OUTPUT_BRAND_HEADER,
    sub_product_header=OUTPUT_SUB_PRODUCT_HEADER,
):
    brand_index = find_header_index(
        header,
        brand_header,
    )

    sub_product_index = find_header_index(
        header,
        sub_product_header,
    )

    blank_brand = 0
    blank_sub_product = 0

    for row in rows:
        if not str(
            cell(
                row,
                brand_index,
            )
            or ""
        ).strip():
            blank_brand += 1

        if not str(
            cell(
                row,
                sub_product_index,
            )
            or ""
        ).strip():
            blank_sub_product += 1

    if blank_brand or blank_sub_product:
        raise ValueError(
            "Output BW memiliki classification kosong: "
            f"Brand={blank_brand:,}, "
            f"Sub Product={blank_sub_product:,}."
        )

    return (
        brand_index,
        sub_product_index,
    )


def rebuild_monthly_output_bw(
    master_header,
    master_rows,
    output_header,
    content_id_header,
    date_header,
    sum_columns,
    gmv_column,
    brand_header=OUTPUT_BRAND_HEADER,
    sub_product_header=OUTPUT_SUB_PRODUCT_HEADER,
    jenis_header="Jenis",
    consistent_fields=None,
    max_columns=None,
    ratio_columns=None,
):
    """
    Re-aggregate all stored source pulls to exactly:
        Brand + Content ID

    Stored rows are already one row per load, so per-date metrics (views, likes)
    are plain additive here: pass them inside sum_columns at this stage.

    Product ID and Sub Product are representative values from the FIRST row
    encountered for that Brand + Content ID. Additive metrics are summed.
    RANGE filtering is intentionally performed only after this rebuild.
    """
    if not output_header:
        raise ValueError(
            "Output header kosong."
        )

    if (
        normalize_header(output_header[-1])
        != normalize_header(jenis_header)
    ):
        raise ValueError(
            f'Kolom terakhir output harus "{jenis_header}".'
        )

    stored_rows = strip_master_metadata(
        master_header,
        master_rows,
        output_header,
    )

    base_header = list(
        output_header[:-1]
    )
    base_width = len(base_header)

    aggregation_rows = [
        [
            cell(row, index)
            for index in range(base_width)
        ]
        for row in stored_rows
    ]

    for field_header in consistent_fields or []:
        validate_group_field_consistency(
            base_header,
            aggregation_rows,
            [
                brand_header,
                content_id_header,
            ],
            field_header,
        )

    (
        groups,
        accepted_rows,
        skipped_blank,
        invalid_rows,
    ) = aggregate_bw_brand_content(
        base_header,
        aggregation_rows,
        content_id_header=content_id_header,
        date_header=date_header,
        sum_columns=sum_columns,
        brand_header=brand_header,
        product_id_header=SOURCE_PRODUCT_ID_HEADER,
        sub_product_header=sub_product_header,
        row_validator=None,
        max_columns=max_columns,
        ratio_columns=ratio_columns,
    )

    if invalid_rows:
        raise ValueError(
            "Monthly BW rebuild menghasilkan "
            "row invalid yang tidak diharapkan."
        )

    combined = build_bw_combined_from_groups(
        groups,
        base_header,
        gmv_column,
    )

    validate_bw_classification_columns(
        output_header,
        combined,
        brand_header=brand_header,
        sub_product_header=sub_product_header,
    )

    validate_unique_brand_content(
        output_header,
        combined,
        content_id_header=content_id_header,
        brand_header=brand_header,
    )

    checks = reconcile_sum_columns(
        base_header,
        accepted_rows,
        groups,
        sum_columns,
    )

    if not all(
        item["ok"]
        for item in checks
    ):
        raise ValueError(
            "Monthly BW reconciliation gagal."
        )

    return {
        "groups": groups,
        "rows": combined,
        "accepted_rows": accepted_rows,
        "skipped_blank": skipped_blank,
        "checks": checks,
        "base_header": base_header,
    }

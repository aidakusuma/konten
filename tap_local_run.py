"""
Jalankan pipeline TAP (Short Video / Livestream) dari FILE EXCEL LOKAL.

Input : export TAP "Custom report" (.xlsx) + file mapping BW (Product ID | Brand | Sub Product)
Output: .xlsx berisi
          - satu sheet per Brand  -> hanya konten yang tanggalnya masuk RANGE yang dipilih
          - sheet GABUNGAN        -> semua brand, SEMUA range
          - sheet QC              -> rekonsiliasi dengan baris Summary + kebijakan agregasi

Contoh:
  python tap_local_run.py --type sv --raw TAP_VID_18-24_OCT_2026.xlsx --mapping mapping_bw.csv \
         --start 2026-10-20 --end 2026-10-22 --out hasil_sv.xlsx
  python tap_local_run.py --type ls --raw TAP_LS_18-24_OCT_2026.xlsx --mapping mapping_bw.csv \
         --start 2026-10-20 --end 2026-10-22 --out hasil_ls.xlsx

Tambahkan --self-check untuk membuktikan bahwa memecah data ke dua load (per tanggal)
lalu merge lewat rebuild_monthly_output_bw menghasilkan angka yang sama dengan sekali proses.

Syarat: tap_utils.py ada di folder yang sama.   pip install openpyxl pandas
"""

from __future__ import annotations

import argparse
import csv
import sys
import time
from collections import Counter
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parent))
import tap_utils as u  # noqa: E402

EST = "Estimated affiliate partner commission "
ACT = "Actual affiliate partner commission"
ITEMS = "Creator-attributed items sold"

CONFIG = {
    "sv": {
        "content": "Video ID",
        "date_src": "Post time",          # kolom mentah
        "date_col": "Post date",          # kolom turunan untuk filter RANGE
        "gmv": "Affiliate video-attributed GMV",
        "sum": [
            "Affiliate video-attributed GMV",
            "Creator video-attributed orders",
            "Affiliate video orders",
            EST, ACT, ITEMS,
        ],
        "per_date": ["Video views", "Video likes"],
        "max": ["Creator follower count"],
        "name": "Video name",
        "types": {"Comparison date": u.TYPE_TEXT, "Campaign duration": u.TYPE_TEXT,
                  "Post time": u.TYPE_TEXT},
        "suffix": " SV",
    },
    "ls": {
        "content": "Livestream room ID",
        "date_src": "LIVE time info",
        "date_col": "Date Live",
        "gmv": "Creator LIVE-attributed GMV",
        "sum": [
            "Creator LIVE-attributed GMV",
            "Creator LIVE-attributed orders",
            "Affiliate LIVE orders",
            EST, ACT, ITEMS,
        ],
        "per_date": ["LIVE views", "LIVE likes"],
        "max": ["Creator follower count"],
        "name": "Livestream name",
        "types": {"Comparison date": u.TYPE_TEXT, "Campaign duration": u.TYPE_TEXT,
                  "LIVE time info": u.TYPE_TEXT},
        "suffix": " LS",
    },
}


# --------------------------------------------------------------------------- input
def read_raw(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb["Custom report"] if "Custom report" in wb.sheetnames else wb.worksheets[0]
    ws.reset_dimensions()   # export TAP sering salah menulis dimensi sheet
    it = ws.iter_rows(values_only=True)
    header = [str(v).strip() if v is not None else "" for v in next(it)]
    while header and not header[-1]:
        header.pop()
    rows = [list(r)[: len(header)] for r in it]
    wb.close()
    return header, rows


def read_table_file(path):
    path = Path(path)
    if path.suffix.lower() in (".csv", ".txt"):
        with open(path, newline="", encoding="utf-8-sig") as f:
            return [list(r) for r in csv.reader(f)]
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    out = []
    for ws in wb.worksheets:
        ws.reset_dimensions()
        out = [list(r) for r in ws.iter_rows(values_only=True)]
        if out:
            break
    return out


def load_mapping(path):
    table = read_table_file(path)
    want = {k: u.normalize_header(k) for k in ("Product ID", "Brand", "Sub Product")}
    for i, row in enumerate(table[:20]):
        norm = [u.normalize_header(v) for v in row]
        if all(w in norm for w in want.values()):
            idx = [norm.index(w) for w in want.values()]
            mapping, stats = u.build_bw_mapping_from_rows(table[i + 1:], *idx)
            return mapping, stats
    raise ValueError("Header mapping tidak ditemukan (Product ID | Brand | Sub Product).")


# --------------------------------------------------------------------------- core
def prepare(cfg, header, rows, mapping):
    """Raw rows -> (header, rows) bersih, berklasifikasi brand, + info QC."""
    qc = {}
    validator = u.make_tap_row_validator("Date", [cfg["content"], "Product ID"])

    summary = None
    valid = []
    rejected = Counter()
    for n, row in enumerate(rows, start=2):
        if str(u.cell(row, 0) or "").strip().lower() == "summary":
            summary = dict(zip(header, row))
        ok, reason = validator(row, n, header)
        if ok:
            valid.append(row)
        else:
            rejected[reason] += 1
    qc["rows_raw"] = len(rows)
    qc["rows_rejected"] = dict(rejected)

    key = ["Date", cfg["content"], "Product ID", "Campaign ID", "Creator name"]
    valid, dropped, conflicts = u.dedupe_source_rows(header, valid, key)
    qc["dup_rows_dropped"] = dropped
    qc["dup_key_conflicts"] = len(conflicts)

    # sisakan kolom yang dipakai saja (jauh lebih cepat)
    keep = ["Date", "Campaign ID", "Campaign name", "Creator name", "Product ID", "Product name",
            cfg["content"], cfg["date_src"]] + cfg["sum"] + cfg["per_date"] + cfg["max"]
    if cfg["name"] not in keep:
        keep.append(cfg["name"])
    seen, keep_u = set(), []
    for k in keep:
        if u.normalize_header(k) not in seen:
            seen.add(u.normalize_header(k))
            keep_u.append(k)
    idx = [u.find_header_index(header, k) for k in keep_u]
    h1 = [header[i] for i in idx]
    r1 = [[u.cell(r, i) for i in idx] for r in valid]

    types = dict(cfg["types"])
    r2, col_types, _summary, _errors = u.coerce_rows_by_schema(h1, r1, column_types=types, strict=True)

    # kolom tanggal turunan (untuk filter RANGE)
    si = u.find_header_index(h1, cfg["date_src"])
    h2 = list(h1) + [cfg["date_col"]]
    r3 = [row + [u.normalize_date_text(str(row[si])[:10])] for row in r2]

    h4, r4, status, missing, _bc, _sc = u.add_brand_subproduct_from_product_id(h2, r3, mapping)
    qc["classification"] = dict(status)
    qc["missing_product_ids"] = len(missing)
    return h4, r4, summary, qc


def aggregate(cfg, h, rows, sum_cols=None, per_date=None):
    return u.aggregate_bw_brand_content(
        h, rows,
        content_id_header=cfg["content"],
        date_header=cfg["date_col"],
        sum_columns=sum_cols or cfg["sum"],
        per_date_sum_columns=per_date if per_date is not None else cfg["per_date"],
        daily_header="Date",
        max_columns=cfg["max"],
    )


def output_columns(cfg):
    return [cfg["date_col"], cfg["content"], cfg["name"], "Creator name", "Creator follower count",
            "Brand", "Sub Product", "Product ID", "Product name", "Campaign name"] \
        + cfg["sum"] + cfg["per_date"]


def to_output(cfg, h, groups, gmv_col):
    """-> (header_out, rows_all_with_jenis, tabs_for_range_fn)"""
    cols = output_columns(cfg)
    idx = [u.find_header_index(h, c) for c in cols]
    header_out = [h[i] for i in idx] + ["Jenis"]
    gi = u.find_header_index(h, gmv_col)

    def project(rows_with_jenis):
        return [[r[i] for i in idx] + [r[-1]] for r in rows_with_jenis]

    return header_out, gi, project


def write_xlsx(path, sheets):
    wb = openpyxl.Workbook(write_only=True)
    used = set()
    for title, header, rows in sheets:
        t = "".join(c for c in title if c not in '[]:*?/\\')[:31] or "Sheet"
        base, k = t, 2
        while t.lower() in used:
            t = f"{base[:28]}~{k}"
            k += 1
        used.add(t.lower())
        ws = wb.create_sheet(t)
        ws.append(header)
        for r in rows:
            ws.append(r)
    wb.save(path)


def run(args):
    cfg = CONFIG[args.type]
    t0 = time.time()
    mapping, mstats = load_mapping(args.mapping)
    print(f"Mapping: {mstats['mapping_count']:,} Product ID; "
          f"variasi huruf brand digabung: {mstats['brand_case_variants'] or '-'}")

    header, rows = read_raw(args.raw)
    print(f"Raw: {len(rows):,} baris ({time.time() - t0:.0f}s)")
    h, r, summary, qc = prepare(cfg, header, rows, mapping)
    print(f"Siap diagregasi: {len(r):,} baris | ditolak: {qc['rows_rejected']} | "
          f"duplikat persis dibuang: {qc['dup_rows_dropped']} | klasifikasi: {qc['classification']}")

    groups, accepted, skipped_blank, invalid = aggregate(cfg, h, r)
    checks = u.reconcile_sum_columns(h, accepted, groups, cfg["sum"])
    assert all(c["ok"] for c in checks), "rekonsiliasi agregasi gagal"
    print(f"Unik Brand+{cfg['content']}: {len(groups):,}")

    # rekonsiliasi dengan baris Summary dari TAP
    qc_rows = []
    if summary:
        for col in cfg["sum"]:
            raw_total = u.to_number(summary.get(col, 0))
            mine = next(c["aggregate"] for c in checks if c["column"] == col)
            qc_rows.append([col, raw_total, mine, "OK" if u.numbers_equal(raw_total, mine) else "BEDA"])
    for row in qc_rows:
        print("  Summary vs hasil:", row)

    header_out, gi, project = to_output(cfg, h, groups, cfg["gmv"])
    rs, re_ = args.start, args.end
    in_range, invalid_range = u.filter_groups_by_date_range(h, groups, cfg["date_col"], rs, re_)
    tabs = u.build_bw_brand_tabs(in_range, gi, cfg["suffix"])
    combined = u.build_bw_combined_from_groups(groups, h, cfg["gmv"])

    sheets = []
    for title, trows in tabs.items():
        sheets.append((title, header_out, project(trows)))
    sheets.append(("GABUNGAN (semua range)", header_out, project(combined)))

    policy = u.aggregation_policy_report(h, [cfg["content"], "Brand"], cfg["sum"], cfg["per_date"], cfg["max"])
    qc_sheet = [["Item", "Nilai"]]
    qc_sheet += [["Range dipilih", f"{rs} s/d {re_} (kolom {cfg['date_col']})"],
                 ["Baris raw", qc["rows_raw"]], ["Ditolak", str(qc["rows_rejected"])],
                 ["Duplikat persis dibuang", qc["dup_rows_dropped"]],
                 ["Konten unik (semua range)", len(groups)], ["Konten di range", len(in_range)],
                 ["Product ID belum ada di mapping", qc["missing_product_ids"]], []]
    qc_sheet += [["Kolom", "Summary TAP", "Hasil agregasi", "Status"]] + qc_rows + [[]]
    qc_sheet += [["Aturan agregasi", "Kolom"]] + [[k, ", ".join(v)] for k, v in policy.items()]
    sheets.append(("QC", qc_sheet[0], qc_sheet[1:]))
    write_xlsx(args.out, sheets)
    print(f"Tersimpan: {args.out}  ({time.time() - t0:.0f}s)")
    print({t: len(x) for t, x in tabs.items()}, "| gabungan:", len(combined))

    if args.self_check:
        self_check(cfg, h, r, groups, header_out)


def self_check(cfg, h, rows, single_groups, header_out):
    """Pecah per tanggal -> dua load -> master -> rebuild; harus sama dengan sekali proses."""
    di = u.find_header_index(h, "Date")
    dates = sorted({str(row[di]) for row in rows})
    cut = dates[len(dates) // 2]
    parts = [[x for x in rows if str(x[di]) < cut], [x for x in rows if str(x[di]) >= cut]]
    print(f"Self-check: load A = tanggal < {cut} ({len(parts[0]):,} baris), load B = >= {cut} ({len(parts[1]):,} baris)")

    out_h = None
    master = []
    for n, part in enumerate(parts, start=1):
        g, *_ = aggregate(cfg, h, part)
        hd, gi, project = to_output(cfg, h, g, cfg["gmv"])
        out_h = hd
        comb = u.build_bw_combined_from_groups(g, h, cfg["gmv"])
        # master menyimpan baris output penuh (kolom output + Jenis) per load
        stored = [list(x) for x in project(comb)]
        master += u.attach_master_metadata(stored, "TAP" + cfg["suffix"].strip(), f"LOAD{n}", "src", "fp", "now")
    mh = u.build_master_header(out_h)
    sum_all = cfg["sum"] + cfg["per_date"]  # di tahap rebuild views/likes sudah per-load -> aditif
    res = u.rebuild_monthly_output_bw(
        mh, master, out_h, cfg["content"], cfg["date_col"], sum_all, cfg["gmv"],
        max_columns=cfg["max"],
    )
    ok = True
    for col in sum_all:
        a = sum(u.to_number(u.cell(v, u.find_header_index(h, col))) for v in single_groups.values())
        b = sum(u.to_number(u.cell(v, u.find_header_index(res["base_header"], col))) for v in res["groups"].values())
        same = u.numbers_equal(a, b)
        ok &= same
        print(f"  {col}: sekali proses={a:,.0f} | merge dua load={b:,.0f} | {'OK' if same else 'BEDA'}")
    print(f"  jumlah konten: {len(single_groups):,} vs {len(res['groups']):,}")
    print("SELF-CHECK", "LULUS" if ok and len(single_groups) == len(res["groups"]) else "GAGAL")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--type", choices=["sv", "ls"], required=True)
    ap.add_argument("--raw", required=True)
    ap.add_argument("--mapping", required=True)
    ap.add_argument("--start", required=True, help="YYYY-MM-DD (Post date / Date Live)")
    ap.add_argument("--end", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--self-check", action="store_true")
    run(ap.parse_args())

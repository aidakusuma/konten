"""
Helper untuk deploy di Streamlit Community Cloud.

Di cloud, credentials.json TIDAK di-upload ke GitHub (rahasia). Isinya ditaruh di
Streamlit Secrets (App settings -> Secrets) dengan bagian [gcp_service_account].
Fungsi di bawah menulis ulang isi secret itu menjadi credentials.json di folder app
saat app start, sehingga semua kode yang membaca BASE_DIR / "credentials.json"
tetap jalan tanpa perubahan.

Di lokal (credentials.json sudah ada), fungsi ini tidak melakukan apa-apa.
"""

from __future__ import annotations

import json
from pathlib import Path


def ensure_credentials_file(base_dir: Path) -> None:
    target = base_dir / "credentials.json"
    if target.exists():
        return
    try:
        import streamlit as st
        if "gcp_service_account" not in st.secrets:
            return
        info = dict(st.secrets["gcp_service_account"])
    except Exception:  # noqa: BLE001  (tidak ada secrets.toml -> lewati)
        return
    # private_key kadang tersimpan dengan "\n" literal -> jadikan newline asli
    if isinstance(info.get("private_key"), str):
        info["private_key"] = info["private_key"].replace("\\n", "\n")
    try:
        target.write_text(json.dumps(info, indent=2), encoding="utf-8")
    except OSError:
        pass

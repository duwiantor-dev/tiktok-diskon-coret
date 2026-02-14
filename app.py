import io
import re
import zipfile
from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple

import streamlit as st
from openpyxl import load_workbook, Workbook
from openpyxl.cell.cell import MergedCell


# =========================
# FIXED ROWS sesuai request
# =========================
# INPUT (gambar bawah / kiri): header row 3, data mulai row 6
INPUT_HEADER_ROW = 3
INPUT_DATA_START_ROW = 6

# OUTPUT (gambar atas / kanan): header row 1, data mulai row 2
OUTPUT_HEADER_ROW = 1
OUTPUT_DATA_START_ROW = 2

# Pricelist header tetap row 2
PRICELIST_HEADER_ROW_FIXED = 2

# =========================
# INPUT column mapping (1-based)
# =========================
IN_COL_ID_PRODUK = 1   # A
IN_COL_ID_SKU = 4      # D
IN_COL_SKU_PENJUAL = 8 # H pada template lama? (tapi request bilang E)
# NOTE: sesuai request user: SKU penjual ada di kolom E
IN_COL_SKU_PENJUAL = 5 # E
IN_COL_HARGA = 6       # F
IN_COL_STOK = 7        # G

# =========================
# OUTPUT column mapping (1-based)
# =========================
OUT_COL_PRODUCT_ID = 1  # A
OUT_COL_SKU_ID = 2      # B
OUT_COL_HARGA = 3       # C
OUT_COL_STOK = 4        # D

# =========================
# Pricelist config
# =========================
PL_HEADER_SKU_CANDIDATES = ["KODEBARANG", "KODE BARANG", "SKU", "SKU NO", "SKU_NO", "KODEBARANG "]
PL_PRICE_COL_TIKTOK = "M3"
PL_PRICE_COL_SHOPEE = "M4"  # ada di file, tapi kita pakai M3

# Addon mapping config
ADDON_CODE_CANDIDATES = ["addon_code", "ADDON_CODE", "Addon Code", "Kode", "KODE", "KODE ADDON", "KODE_ADDON"]
ADDON_PRICE_CANDIDATES = ["harga", "HARGA", "Price", "PRICE", "Harga"]

# Heuristik kecil -> x1000
SMALL_TO_THOUSAND_THRESHOLD = 1_000_000
AUTO_MULTIPLIER_FOR_SMALL = 1000

SKU_SPLIT_PLUS = re.compile(r"\+")


# =========================
# Utils
# =========================
def normalize_text(x) -> str:
    if x is None:
        return ""
    return str(x).strip()


def normalize_addon_code(x) -> str:
    return normalize_text(x).upper()


def parse_platform_sku(full_sku: str) -> Tuple[str, List[str]]:
    if full_sku is None:
        return "", []
    s = str(full_sku).strip()
    if not s:
        return "", []
    parts = SKU_SPLIT_PLUS.split(s)
    base = parts[0].strip()
    addons = [p.strip() for p in parts[1:] if p and p.strip()]
    return base, addons


def parse_number_like_id(x) -> str:
    if x is None:
        return ""
    if isinstance(x, int):
        return str(x)
    if isinstance(x, float):
        if x.is_integer():
            return str(int(x))
        return str(x)
    return str(x).strip()


def parse_price_cell(val) -> Optional[int]:
    if val is None:
        return None

    if isinstance(val, (int, float)):
        try:
            if isinstance(val, float) and val.is_integer():
                return int(val)
            return int(round(float(val)))
        except Exception:
            return None

    s = str(val).strip()
    if not s:
        return None

    s = s.replace("Rp", "").replace("rp", "").replace(" ", "")

    if "." in s and "," in s:
        s = s.replace(".", "").replace(",", ".")
    elif "." in s and "," not in s:
        s = s.replace(".", "")
    elif "," in s and "." not in s:
        s = s.replace(",", "")

    try:
        f = float(s)
        if f.is_integer():
            return int(f)
        return int(round(f))
    except Exception:
        return None


def apply_multiplier_if_needed(x: int) -> int:
    if x is None:
        return 0
    if x < SMALL_TO_THOUSAND_THRESHOLD:
        return x * AUTO_MULTIPLIER_FOR_SMALL
    return x


def safe_set_cell_value(ws, row: int, col: int, value):
    cell = ws.cell(row=row, column=col)
    if isinstance(cell, MergedCell):
        coord = cell.coordinate
        for merged in ws.merged_cells.ranges:
            if coord in merged:
                ws.cell(row=merged.min_row, column=merged.min_col).value = value
                return
        return
    cell.value = value


# =========================
# Pricelist loader (header row 2)
# =========================
def find_header_row_and_cols_pricelist(ws) -> Tuple[int, int, int, int]:
    r = PRICELIST_HEADER_ROW_FIXED
    candidates = [c.strip().lower() for c in PL_HEADER_SKU_CANDIDATES]
    target_m3 = PL_PRICE_COL_TIKTOK.lower()
    target_m4 = PL_PRICE_COL_SHOPEE.lower()

    row_vals = []
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=r, column=c).value
        row_vals.append("" if v is None else str(v).strip())

    lower_to_col = {}
    for idx, v in enumerate(row_vals, start=1):
        lv = v.strip().lower()
        if lv and lv not in lower_to_col:
            lower_to_col[lv] = idx

    sku_col = None
    for cand in candidates:
        if cand in lower_to_col:
            sku_col = lower_to_col[cand]
            break

    if sku_col is None or target_m3 not in lower_to_col or target_m4 not in lower_to_col:
        raise ValueError(
            f"Header Pricelist row {PRICELIST_HEADER_ROW_FIXED} tidak sesuai. "
            f"Pastikan ada kolom KODEBARANG (atau setara) dan kolom M3 & M4."
        )

    return r, sku_col, lower_to_col[target_m3], lower_to_col[target_m4]


def load_pricelist_map(pl_bytes: bytes) -> Dict[str, Dict[str, int]]:
    wb = load_workbook(io.BytesIO(pl_bytes), data_only=True)
    ws = wb.active

    header_row, sku_col, m3_col, m4_col = find_header_row_and_cols_pricelist(ws)

    m: Dict[str, Dict[str, int]] = {}
    for r in range(header_row + 1, ws.max_row + 1):
        sku_val = ws.cell(row=r, column=sku_col).value
        sku = normalize_text(sku_val)
        if not sku:
            continue

        m3_raw = parse_price_cell(ws.cell(row=r, column=m3_col).value)
        m4_raw = parse_price_cell(ws.cell(row=r, column=m4_col).value)

        if m3_raw is None and m4_raw is None:
            continue

        m3 = apply_multiplier_if_needed(m3_raw) if m3_raw is not None else None
        m4 = apply_multiplier_if_needed(m4_raw) if m4_raw is not None else None

        m[sku] = {}
        if m3 is not None:
            m[sku]["M3"] = int(m3)
        if m4 is not None:
            m[sku]["M4"] = int(m4)

    return m


# =========================
# Addon loader
# =========================
def load_addon_map(addon_bytes: bytes) -> Dict[str, int]:
    wb = load_workbook(io.BytesIO(addon_bytes), data_only=True)
    ws = wb.active

    header_row = None
    code_col = None
    price_col = None

    code_cands = [c.strip().lower() for c in ADDON_CODE_CANDIDATES]
    price_cands = [c.strip().lower() for c in ADDON_PRICE_CANDIDATES]

    for r in range(1, 30):
        row_vals = []
        for c in range(1, ws.max_column + 1):
            v = ws.cell(row=r, column=c).value
            row_vals.append("" if v is None else str(v).strip())

        lower_to_col = {}
        for idx, v in enumerate(row_vals, start=1):
            lv = v.strip().lower()
            if not lv:
                continue
            if lv not in lower_to_col:
                lower_to_col[lv] = idx

        found_code = None
        for cc in code_cands:
            if cc in lower_to_col:
                found_code = lower_to_col[cc]
                break

        found_price = None
        for pc in price_cands:
            if pc in lower_to_col:
                found_price = lower_to_col[pc]
                break

        if found_code and found_price:
            header_row = r
            code_col = found_code
            price_col = found_price
            break

    if header_row is None or code_col is None or price_col is None:
        raise ValueError("Header Addon Mapping tidak ketemu. Pastikan ada kolom addon_code & harga (atau setara).")

    m: Dict[str, int] = {}
    for r in range(header_row + 1, ws.max_row + 1):
        code = normalize_addon_code(ws.cell(row=r, column=code_col).value)
        if not code:
            continue

        price_raw = parse_price_cell(ws.cell(row=r, column=price_col).value)
        if price_raw is None:
            continue

        price = apply_multiplier_if_needed(int(price_raw))
        m[code] = int(price)

    return m


# =========================
# Pricing compute (SELALU M3)
# =========================
def compute_new_price_from_sku_penjual(
    sku_penjual: str,
    pl_map: Dict[str, Dict[str, int]],
    addon_map: Dict[str, int],
    discount_rp: int,
) -> Tuple[Optional[int], str]:
    """
    Harga selalu pakai M3:
      final = base(M3) + sum(addon) - discount
    Jika base tidak ada / addon tidak ada -> return None (skip)
    """
    base_sku, addons = parse_platform_sku(sku_penjual)
    if not base_sku:
        return None, "SKU penjual kosong"

    pl = pl_map.get(base_sku)
    if not pl:
        return None, "Base SKU tidak ada di Pricelist"

    price_key = "M3"
    base_price = pl.get(price_key)
    if base_price is None:
        return None, f"Harga {price_key} kosong di Pricelist"

    addon_total = 0
    for a in addons:
        code = normalize_addon_code(a)
        if not code:
            continue
        if code not in addon_map:
            return None, f"Addon '{code}' tidak ada di file Addon Mapping"
        addon_total += int(addon_map[code])

    final_price = int(base_price) + int(addon_total) - int(discount_rp)
    if final_price < 0:
        final_price = 0

    return final_price, f"{price_key} + addon - diskon"


# =========================
# Output writer
# =========================
def write_output_rows_on_template(
    out_ws,
    rows: List[Tuple[str, str, int, int]],
    start_row: int,
):
    """
    rows: list of (product_id, sku_id, harga, stok)
    Ditulis ke template output mulai start_row.
    """
    r = start_row
    for product_id, sku_id, harga, stok in rows:
        out_ws.cell(row=r, column=OUT_COL_PRODUCT_ID).value = product_id
        out_ws.cell(row=r, column=OUT_COL_SKU_ID).value = sku_id
        out_ws.cell(row=r, column=OUT_COL_HARGA).value = int(harga)
        out_ws.cell(row=r, column=OUT_COL_STOK).value = int(stok)
        r += 1


def workbook_to_bytes(wb) -> bytes:
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


@dataclass
class RowIssue:
    file: str
    excel_row: int
    product_id: str
    sku_id: str
    sku_penjual: str
    reason: str


def make_issues_workbook(issues: List[RowIssue]) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "issues_report"
    ws.append(["file", "row", "product_id", "sku_id", "sku_penjual", "reason"])
    for x in issues:
        ws.append([x.file, x.excel_row, x.product_id, x.sku_id, x.sku_penjual, x.reason])
    return workbook_to_bytes(wb)


# =========================
# UI
# =========================
st.set_page_config(page_title="Product Discount (Input kiri -> Output kanan)", layout="wide")
st.title("Product Discount (Input template TikTok -> Output template Discount)")

c1, c2, c3, c4 = st.columns(4)
with c1:
    input_files = st.file_uploader(
        "Upload INPUT (template kiri) - bisa banyak",
        type=["xlsx"],
        accept_multiple_files=True,
    )
with c2:
    output_template_file = st.file_uploader("Upload TEMPLATE OUTPUT (gambar kanan)", type=["xlsx"])
with c3:
    pl_file = st.file_uploader("Upload Pricelist (header row 2)", type=["xlsx"])
with c4:
    addon_file = st.file_uploader("Upload Addon Mapping", type=["xlsx"])

st.divider()
discount_rp = st.number_input("Diskon (Rp) - mengurangi harga final", min_value=0, value=0, step=1000)

process = st.button("Proses")

if process:
    if not input_files or output_template_file is None or pl_file is None or addon_file is None:
        st.error("Wajib upload: INPUT (minimal 1), TEMPLATE OUTPUT, Pricelist, dan Addon Mapping.")
        st.stop()

    try:
        pl_map = load_pricelist_map(pl_file.getvalue())
    except Exception as e:
        st.error(f"Gagal baca Pricelist: {e}")
        st.stop()

    try:
        addon_map = load_addon_map(addon_file.getvalue())
    except Exception as e:
        st.error(f"Gagal baca Addon Mapping: {e}")
        st.stop()

    output_files: List[Tuple[str, bytes]] = []
    issues: List[RowIssue] = []

    for f in input_files:
        filename = f.name
        in_wb = load_workbook(io.BytesIO(f.getvalue()))
        in_ws = in_wb.active

        # Load output template fresh for setiap file input
        out_wb = load_workbook(io.BytesIO(output_template_file.getvalue()))
        out_ws = out_wb.active

        out_rows: List[Tuple[str, str, int, int]] = []

        for r in range(INPUT_DATA_START_ROW, in_ws.max_row + 1):
            product_id = parse_number_like_id(in_ws.cell(row=r, column=IN_COL_ID_PRODUK).value)
            sku_id = parse_number_like_id(in_ws.cell(row=r, column=IN_COL_ID_SKU).value)
            sku_penjual = normalize_text(in_ws.cell(row=r, column=IN_COL_SKU_PENJUAL).value)

            if not product_id and not sku_id and not sku_penjual:
                continue

            stok_raw = parse_price_cell(in_ws.cell(row=r, column=IN_COL_STOK).value)
            stok = int(stok_raw) if stok_raw is not None else 0

            # Hitung harga dari Pricelist+Addon (M3) - diskon
            new_price, reason = compute_new_price_from_sku_penjual(
                sku_penjual=sku_penjual,
                pl_map=pl_map,
                addon_map=addon_map,
                discount_rp=int(discount_rp),
            )

            if new_price is None:
                issues.append(RowIssue(
                    file=filename,
                    excel_row=r,
                    product_id=product_id,
                    sku_id=sku_id,
                    sku_penjual=sku_penjual,
                    reason=reason,
                ))
                continue

            # ✅ Output hanya data ini:
            # - product_id (A input -> A output)
            # - sku_id (D input -> B output)
            # - harga (hasil compute -> C output)
            # - stok (G input -> D output)
            out_rows.append((product_id, sku_id, int(new_price), int(stok)))

        if not out_rows:
            # tidak ada baris yang bisa digenerate (semua gagal mapping)
            continue

        # Tulis ke template output mulai row 2
        write_output_rows_on_template(out_ws, out_rows, start_row=OUTPUT_DATA_START_ROW)

        out_name = filename.replace(".xlsx", "_OUTPUT_DISCOUNT_M3.xlsx")
        output_files.append((out_name, workbook_to_bytes(out_wb)))

    st.subheader("Hasil")
    if not output_files:
        st.warning("Tidak ada file output yang dihasilkan (mungkin semua baris gagal mapping base/addon).")
    elif len(output_files) == 1:
        name, data = output_files[0]
        st.download_button(
            "Download Output (XLSX)",
            data=data,
            file_name=name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        zbuf = io.BytesIO()
        with zipfile.ZipFile(zbuf, "w", zipfile.ZIP_DEFLATED) as zf:
            for name, data in output_files:
                zf.writestr(name, data)

            if issues:
                zf.writestr("issues_report.xlsx", make_issues_workbook(issues))

        st.download_button(
            "Download Semua Output (ZIP)",
            data=zbuf.getvalue(),
            file_name="outputs_discount_m3.zip",
            mime="application/zip",
        )

    if issues:
        st.divider()
        st.subheader("Issues (baris yang gagal diproses)")
        st.caption("Baris ini tidak masuk ke output karena base SKU / addon tidak ketemu, atau harga M3 kosong.")

        import pandas as pd
        df_issues = pd.DataFrame([{
            "file": x.file,
            "row": x.excel_row,
            "product_id": x.product_id,
            "sku_id": x.sku_id,
            "sku_penjual": x.sku_penjual,
            "reason": x.reason,
        } for x in issues])
        st.dataframe(df_issues.head(300), use_container_width=True)

        rep_bytes = make_issues_workbook(issues)
        st.download_button(
            "Download Issues Report (XLSX)",
            data=rep_bytes,
            file_name="issues_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

import io
import re
from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st
from openpyxl import Workbook, load_workbook
from openpyxl.cell.cell import MergedCell


# =========================
# FIXED TEMPLATE SPECS
# =========================
INPUT_HEADER_ROW = 3
INPUT_DATA_START_ROW = 6

OUTPUT_HEADER_ROW = 1
OUTPUT_DATA_START_ROW = 2

# =========================
# INPUT HEADERS (row 3)
# =========================
IN_HDR_PRODUCT_ID_CANDS = ["ID Produk", "ID PRODUK", "Product ID", "Product_Id", "Product_id"]
IN_HDR_ID_SKU_CANDS = ["ID SKU", "ID_SKU", "Sku Id", "SKU ID", "SKU_Id", "SKU_id"]
IN_HDR_PRICE_CANDS = ["Harga Ritel (Mata Uang Lokal)", "Harga Ritel", "Harga", "PRICE"]
IN_HDR_STOCK_CANDS = ["Kuantitas", "Qty", "QTY", "Stock", "Stok"]
IN_HDR_SKU_PENJUAL_CANDS = ["SKU Penjual", "SKU_PENJUAL", "SKU PENJUAL", "Seller SKU", "SKU Seller"]

# =========================
# OUTPUT HEADERS (row 1) - MUST match template image atas
# =========================
OUT_COL_A = "Product_id (wajib)"
OUT_COL_B = "SKU_id (wajib)"
OUT_COL_C = "Harga Penawaran (wajib)"
OUT_COL_D = "Total Stok Promosi (optional)\n1. Total Stok Promosi≤ Stok\n2. Jika tidak diisi artinya tidak terbatas"
OUT_COL_E = "Batas Pembelian (optional)\n1. 1 ≤ Batas pembelian≤99\n2. Jika tidak diisi artinya tidak terbatas"

# =========================
# PRICELIST / ADDON
# =========================
PRICELIST_HEADER_ROW_FIXED = 2
PL_HEADER_SKU_CANDIDATES = ["KODEBARANG", "KODE BARANG", "SKU", "SKU NO", "SKU_NO", "KODEBARANG "]
PL_PRICE_COL_M3 = "M3"  # harga selalu M3

ADDON_CODE_CANDIDATES = ["addon_code", "ADDON_CODE", "Addon Code", "Kode", "KODE", "KODE ADDON", "KODE_ADDON"]
ADDON_PRICE_CANDIDATES = ["harga", "HARGA", "Price", "PRICE", "Harga"]

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


def lower_map_headers(ws, header_row: int) -> Dict[str, int]:
    row_vals = []
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=header_row, column=c).value
        row_vals.append("" if v is None else str(v).strip())

    m = {}
    for idx, v in enumerate(row_vals, start=1):
        lv = v.strip().lower()
        if lv and lv not in m:
            m[lv] = idx
    return m


def find_col_by_candidates(ws, header_row: int, candidates: List[str]) -> Optional[int]:
    m = lower_map_headers(ws, header_row)
    for cand in candidates:
        key = cand.strip().lower()
        if key in m:
            return m[key]
    return None


# =========================
# Pricelist + Addon loader
# =========================
def find_header_row_and_cols_pricelist(ws) -> Tuple[int, int, int]:
    r = PRICELIST_HEADER_ROW_FIXED
    m = lower_map_headers(ws, r)

    sku_col = None
    for cand in [c.strip().lower() for c in PL_HEADER_SKU_CANDIDATES]:
        if cand in m:
            sku_col = m[cand]
            break

    m3_key = PL_PRICE_COL_M3.lower()
    if sku_col is None or m3_key not in m:
        raise ValueError(
            f"Header Pricelist row {PRICELIST_HEADER_ROW_FIXED} tidak sesuai. "
            f"Pastikan ada kolom KODEBARANG (atau setara) dan kolom M3."
        )

    return r, sku_col, m[m3_key]


def load_pricelist_map(pl_bytes: bytes) -> Dict[str, int]:
    wb = load_workbook(io.BytesIO(pl_bytes), data_only=True)
    ws = wb.active

    header_row, sku_col, m3_col = find_header_row_and_cols_pricelist(ws)

    out: Dict[str, int] = {}
    for r in range(header_row + 1, ws.max_row + 1):
        sku = normalize_text(ws.cell(row=r, column=sku_col).value)
        if not sku:
            continue
        m3_raw = parse_price_cell(ws.cell(row=r, column=m3_col).value)
        if m3_raw is None:
            continue
        out[sku] = int(apply_multiplier_if_needed(int(m3_raw)))
    return out


def load_addon_map(addon_bytes: bytes) -> Dict[str, int]:
    wb = load_workbook(io.BytesIO(addon_bytes), data_only=True)
    ws = wb.active

    header_row = None
    code_col = None
    price_col = None

    code_cands = [c.strip().lower() for c in ADDON_CODE_CANDIDATES]
    price_cands = [c.strip().lower() for c in ADDON_PRICE_CANDIDATES]

    for r in range(1, 30):
        m = lower_map_headers(ws, r)

        found_code = None
        for cc in code_cands:
            if cc in m:
                found_code = m[cc]
                break

        found_price = None
        for pc in price_cands:
            if pc in m:
                found_price = m[pc]
                break

        if found_code and found_price:
            header_row = r
            code_col = found_code
            price_col = found_price
            break

    if header_row is None or code_col is None or price_col is None:
        raise ValueError("Header Addon Mapping tidak ketemu. Pastikan ada kolom addon_code & harga (atau setara).")

    out: Dict[str, int] = {}
    for r in range(header_row + 1, ws.max_row + 1):
        code = normalize_addon_code(ws.cell(row=r, column=code_col).value)
        if not code:
            continue
        price_raw = parse_price_cell(ws.cell(row=r, column=price_col).value)
        if price_raw is None:
            continue
        out[code] = int(apply_multiplier_if_needed(int(price_raw)))
    return out


# =========================
# Pricing
# =========================
def compute_new_price_for_row(
    sku_full: str,
    pl_map_m3: Dict[str, int],
    addon_map: Dict[str, int],
    discount_rp: int,
) -> Tuple[Optional[int], str]:
    base_sku, addons = parse_platform_sku(sku_full)
    if not base_sku:
        return None, "SKU Penjual kosong"

    if base_sku not in pl_map_m3:
        return None, "Base SKU tidak ada di Pricelist"

    base_price = int(pl_map_m3[base_sku])

    addon_total = 0
    for a in addons:
        code = normalize_addon_code(a)
        if not code:
            continue
        if code not in addon_map:
            return None, f"Addon '{code}' tidak ada di file Addon Mapping"
        addon_total += int(addon_map[code])

    final_price = base_price + addon_total - int(discount_rp)
    if final_price < 0:
        final_price = 0

    return int(final_price), "M3 + addon - diskon"


# =========================
# Output workbook builder
# =========================
def build_output_workbook(rows: List[Dict[str, object]]) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    # Header row 1 (template atas)
    ws.cell(row=OUTPUT_HEADER_ROW, column=1).value = OUT_COL_A
    ws.cell(row=OUTPUT_HEADER_ROW, column=2).value = OUT_COL_B
    ws.cell(row=OUTPUT_HEADER_ROW, column=3).value = OUT_COL_C
    ws.cell(row=OUTPUT_HEADER_ROW, column=4).value = OUT_COL_D
    ws.cell(row=OUTPUT_HEADER_ROW, column=5).value = OUT_COL_E

    # Data start row 2
    r = OUTPUT_DATA_START_ROW
    for it in rows:
        ws.cell(row=r, column=1).value = it.get("product_id", "")
        ws.cell(row=r, column=2).value = it.get("id_sku", "")
        ws.cell(row=r, column=3).value = it.get("harga", "")
        ws.cell(row=r, column=4).value = it.get("stok", "")
        # kolom E dibiarkan kosong sesuai permintaan
        r += 1

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


# =========================
# Streamlit UI
# =========================
st.set_page_config(page_title="Generate Product Discount (Harga selalu M3)", layout="wide")
st.title("Generate Product Discount (Harga selalu M3)")

c1, c2, c3 = st.columns(3)
with c1:
    input_file = st.file_uploader("Upload File Input (Template bawah)", type=["xlsx"])
with c2:
    pl_file = st.file_uploader("Upload Pricelist", type=["xlsx"])
with c3:
    addon_file = st.file_uploader("Upload Addon Mapping", type=["xlsx"])

st.divider()
discount_rp = st.number_input("Diskon (Rp) - mengurangi harga final", min_value=0, value=0, step=1000)
process = st.button("Proses")

if process:
    if input_file is None or pl_file is None or addon_file is None:
        st.error("Wajib upload: Input file, Pricelist, dan Addon Mapping.")
        st.stop()

    # Load maps
    try:
        pl_map_m3 = load_pricelist_map(pl_file.getvalue())
    except Exception as e:
        st.error(f"Gagal baca Pricelist: {e}")
        st.stop()

    try:
        addon_map = load_addon_map(addon_file.getvalue())
    except Exception as e:
        st.error(f"Gagal baca Addon Mapping: {e}")
        st.stop()

    # Load input workbook
    wb_in = load_workbook(io.BytesIO(input_file.getvalue()), data_only=True)
    ws_in = wb_in.active

    # Find required columns in input header row 3
    col_product_id = find_col_by_candidates(ws_in, INPUT_HEADER_ROW, IN_HDR_PRODUCT_ID_CANDS)
    col_id_sku = find_col_by_candidates(ws_in, INPUT_HEADER_ROW, IN_HDR_ID_SKU_CANDS)
    col_price = find_col_by_candidates(ws_in, INPUT_HEADER_ROW, IN_HDR_PRICE_CANDS)
    col_stock = find_col_by_candidates(ws_in, INPUT_HEADER_ROW, IN_HDR_STOCK_CANDS)
    col_sku_penjual = find_col_by_candidates(ws_in, INPUT_HEADER_ROW, IN_HDR_SKU_PENJUAL_CANDS)

    missing = []
    if col_product_id is None:
        missing.append("ID Produk")
    if col_id_sku is None:
        missing.append("ID SKU")
    if col_price is None:
        missing.append("Harga Ritel (Mata Uang Lokal)")
    if col_stock is None:
        missing.append("Kuantitas")
    if col_sku_penjual is None:
        missing.append("SKU Penjual")

    if missing:
        st.error(
            "Header input (row 3) tidak ketemu untuk kolom: "
            + ", ".join(missing)
            + ". Pastikan sesuai template."
        )
        st.stop()

    output_rows: List[Dict[str, object]] = []
    issues: List[Dict[str, object]] = []

    for r in range(INPUT_DATA_START_ROW, ws_in.max_row + 1):
        product_id = parse_number_like_id(ws_in.cell(row=r, column=col_product_id).value)
        id_sku = parse_number_like_id(ws_in.cell(row=r, column=col_id_sku).value)

        # harga input (kolom F) tidak dipakai sebagai final, tapi bisa dipakai sebagai referensi issue
        old_price_raw = parse_price_cell(ws_in.cell(row=r, column=col_price).value)
        old_price = int(old_price_raw) if old_price_raw is not None else 0

        stok_raw = ws_in.cell(row=r, column=col_stock).value
        stok = parse_price_cell(stok_raw)  # qty biasanya angka bulat
        stok = int(stok) if stok is not None else ""

        sku_penjual = parse_number_like_id(ws_in.cell(row=r, column=col_sku_penjual).value)

        # skip baris kosong total
        if not product_id and not id_sku and not sku_penjual:
            continue

        new_price, reason = compute_new_price_for_row(
            sku_full=sku_penjual,
            pl_map_m3=pl_map_m3,
            addon_map=addon_map,
            discount_rp=int(discount_rp),
        )

        if new_price is None:
            issues.append({
                "row": r,
                "product_id": product_id,
                "id_sku": id_sku,
                "sku_penjual": sku_penjual,
                "old_price": old_price,
                "reason": reason,
            })
            continue

        output_rows.append({
            "product_id": product_id,
            "id_sku": id_sku,
            "harga": int(new_price),
            "stok": stok,
        })

    # Preview hasil output (sesuai kolom output)
    st.subheader("Hasil Output (Preview) — sesuai template gambar atas")
    if not output_rows:
        st.warning("Tidak ada baris valid untuk di-generate (cek SKU Penjual / Pricelist / Addon).")
    else:
        df_out = pd.DataFrame(output_rows, columns=["product_id", "id_sku", "harga", "stok"])
        df_out.columns = ["Product_id", "SKU_id", "Harga Penawaran", "Total Stok Promosi"]
        st.dataframe(df_out, use_container_width=True, height=420)

        out_xlsx = build_output_workbook(output_rows)
        st.download_button(
            "Download Output XLSX (template gambar atas)",
            data=out_xlsx,
            file_name="product_discount_output_M3.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    # Issues
    if issues:
        st.divider()
        st.subheader("Issues (baris yang gagal dihitung)")
        df_issues = pd.DataFrame(issues)
        st.dataframe(df_issues, use_container_width=True, height=260)

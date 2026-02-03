import os
import re
import math
import datetime
from typing import List, Dict, Optional, Tuple
from copy import copy as _copy

import numpy as np
import pandas as pd
import openpyxl
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Border, Side
from openpyxl.comments import Comment
from openpyxl.styles import PatternFill

from io import BytesIO
from PIL import Image as PILImage
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string
from openpyxl.utils.units import pixels_to_EMU
from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker
from openpyxl.drawing.xdr import XDRPositiveSize2D


# =========================
# CONFIG
# =========================
EXPRESS_A_XLSX = "Express A.xlsx"
EXPRESS_G_XLSX = "Express G.xlsx"
VENDOR_INFO_XLSX = "รายงานข้อมูลผู้จำหน่าย.xlsx"
CATALOG_XLSX = "ข้อมูลสินค้า.xlsx"
TEMPLATE_PO_XLSX = "ตัวอย่างใบสั่งซื้อต่างประเทศ.xlsx"

PO_OUTPUT_FOLDER = "output_PO"
EXPRESS_FILES = [
    (EXPRESS_A_XLSX, "ASIA"),
    (EXPRESS_G_XLSX, "GREEN"),
]
EXPRESS_SHEET = 0  # first sheet

# Template assumptions
TEMPLATE_SHEET_NAME = "page1"
HEADER_ROW = 8
ITEM_START_ROW = 9
TEMPLATE_ITEM_ROW = 9
TEMPLATE_TOTAL_START_ROW = 14
TEMPLATE_TOTAL_END_ROW = 16

# Image behavior
IMAGE_WIDTH_BOOST = 1.20
IMAGE_PADDING_PX = 2

# Totals block fixed locations (based on your file)
TOTAL_VALUE_COL = "N"
TOTAL_FORBIDDEN_COL = "O"


# =========================
# THAI DATE UTILITIES
# =========================
THAI_MONTHS = {
    "ม.ค.": 1, "ก.พ.": 2, "มี.ค.": 3, "เม.ย.": 4,
    "พ.ค.": 5, "มิ.ย.": 6, "ก.ค.": 7, "ส.ค.": 8,
    "ก.ย.": 9, "ต.ค.": 10, "พ.ย.": 11, "ธ.ค.": 12
}

def thai_to_date(day: int, thai_month: str, thai_year: int) -> datetime.date:
    year = thai_year - 543
    thai_month = re.sub(r"\s+", "", thai_month)
    month = THAI_MONTHS.get(thai_month)
    if month is None:
        raise ValueError(f"Unknown Thai month token: {repr(thai_month)}")
    return datetime.date(year, month, day)

def last_day_of_month(year: int, month: int) -> int:
    if month == 12:
        next_first = datetime.date(year + 1, 1, 1)
    else:
        next_first = datetime.date(year, month + 1, 1)
    return (next_first - datetime.timedelta(days=1)).day

def add_months(dt: datetime.date, n: int) -> datetime.date:
    month = dt.month - 1 + n
    year = dt.year + month // 12
    month = month % 12 + 1
    day = min(dt.day, last_day_of_month(year, month))
    return datetime.date(year, month, day)

def calc_days_and_months(d1: datetime.date, d2: datetime.date) -> Tuple[int, int]:
    if d2 < d1:
        return 0, 0
    days = (d2 - d1).days
    months = 0
    cur = d1
    while True:
        next_m = add_months(cur, 1)
        if next_m <= d2:
            months += 1
            cur = next_m
        else:
            break
    leftover_days = (d2 - cur).days
    if leftover_days >= 15:
        months += 1
    return days, months

def parse_date_range_from_header(df_raw: pd.DataFrame) -> Dict[str, object]:
    pattern = r"วันที่จาก\s+(\d+)\s+(\S+)\s+(\d+)\s+ถึง\s+(\d+)\s+(\S+)\s+(\d+)"
    col0 = df_raw.iloc[:, 0].astype(str)
    for val in col0:
        if "วันที่จาก" not in val:
            continue
        text = str(val).replace("\xa0", " ")
        text = re.sub(r"\s+", " ", text).strip()
        m = re.search(pattern, text)
        if not m:
            continue
        d1, m1, y1, d2, m2, y2 = m.groups()
        start_date = thai_to_date(int(d1), m1, int(y1))
        end_date = thai_to_date(int(d2), m2, int(y2))
        days, months = calc_days_and_months(start_date, end_date)
        print(">>> พบช่วงวันที่ใน header:", text)
        print("    start =", start_date, "end =", end_date,
              "days =", days, "months(กฎ 15 วัน) =", months)
        return {
            "raw_line": text.strip(),
            "start_date": start_date,
            "end_date": end_date,
            "days": days,
            "months": months,
        }
    print(">>> ไม่พบช่วงวันที่ ('วันที่จาก ... ถึง ...') ใน header")
    return {"raw_line": "", "start_date": None, "end_date": None, "days": 0, "months": 1}


# =========================
# GENERAL UTILITIES
# =========================
def round_half_up(x: float) -> int:
    return int(math.floor(x + 0.5))

def norm_text(v) -> str:
    if v is None:
        return ""
    return str(v).replace("\xa0", " ").strip()

def load_vendor_map(path: str) -> dict:
    if not os.path.exists(path):
        return {}
    df = pd.read_excel(path, header=None)
    vendor_map = {}
    for _, r in df.iterrows():
        code = str(r.iloc[0]).strip() if not pd.isna(r.iloc[0]) else ""
        if not code:
            continue
        name = "" if pd.isna(r.iloc[1]) else str(r.iloc[1]).strip()
        addr = "" if pd.isna(r.iloc[2]) else str(r.iloc[2]).strip()
        vendor_map[code] = {"name": name, "address": addr}
    return vendor_map


# =========================
# MERGED-LINE HELPERS (Express parsing)
# =========================
def clean_cell(val) -> str:
    if pd.isna(val):
        return ""
    s = str(val)
    s = s.replace("\xa0", " ")
    s = re.sub(r"\s+", " ", s)
    return s.strip()

def fix_split_numbers_in_line(line: str) -> str:
    tokens = line.split()
    i = 0
    while i < len(tokens) - 1:
        cur = tokens[i]
        nxt = tokens[i + 1]
        if (re.fullmatch(r"\d{1,2}", cur) and re.fullmatch(r"\d{3,}(?:\.\d+)?\"?", nxt)):
            tokens[i] = cur + nxt
            del tokens[i + 1]
        else:
            i += 1
    return " ".join(tokens)

def row_to_merged_line(row: pd.Series) -> str:
    cells = [clean_cell(v) for v in row.tolist()]
    cells = [c for c in cells if c]
    raw_line = " ".join(cells)
    return fix_split_numbers_in_line(raw_line)

def looks_like_buyer_code(token: str) -> bool:
    return bool(re.fullmatch(r"[0-9A-Za-z]{5}", token))

def is_header_or_separator(line: str) -> bool:
    s = line.strip()
    if not s:
        return True
    if "BUYER" in s:
        return True
    if "วันที่จาก" in s:
        return True
    if re.search(r"-{5,}", s):
        return True
    return False


# =========================
# Product field split
# =========================
def split_product_field(s: str) -> Tuple[str, str]:
    if not isinstance(s, str):
        return "", ""
    s = s.strip()
    if not s:
        return "", ""

    parts = s.split(maxsplit=1)
    if len(parts) < 2:
        return "", ""
    rest = parts[1].strip()
    if not rest:
        return "", ""

    m_th = re.search(r'[\u0E00-\u0E7F]', rest)

    if m_th:
        th_pos = m_th.start()
        if th_pos == 0:
            return "", rest

        pre_th = rest[:th_pos].strip()
        tail_th = rest[th_pos:].strip()
        tokens = pre_th.split()
        if not tokens:
            return "", tail_th

        code_idx = 0
        for i, t in enumerate(tokens):
            if re.match(r"^(No[A-Za-z0-9\-]+|[A-Z]-\d+)", t):
                code_idx = i
                break

        base_code = tokens[code_idx]
        extra_tokens = tokens[code_idx + 1:]

        tag = ""
        if extra_tokens:
            t = extra_tokens[0]
            if re.fullmatch(r"[A-Za-z]", t) or re.fullmatch(r"\([A-Za-z]\)", t):
                tag = t
                extra_tokens = extra_tokens[1:]

        code_raw = f"{base_code} {tag}".strip() if tag else base_code
        code_raw = re.sub(r'^[Nn][Oo](?=[A-Za-z0-9])', '', code_raw)

        desc_pre = " ".join(extra_tokens).strip()
        desc = (desc_pre + " " + tail_th).strip() if desc_pre else tail_th
    else:
        tokens = rest.split()
        if not tokens:
            return "", ""
        if re.search(r'[\u0E00-\u0E7F]', tokens[0]):
            return "", rest
        code_raw = tokens[0]
        desc = " ".join(tokens[1:]).strip()

    code = re.sub(r'^[Nn][Oo](?=[A-Za-z0-9])', '', code_raw)
    return code, desc


# =========================
# Parse numeric tail (FIXED)
# =========================
def is_num_token(t: str) -> bool:
    """
    True ONLY for pure numeric tokens.
    Reject tokens that contain letters or dash to avoid grabbing BM-150, 01-24-8501, etc.
    """
    s = t.replace(",", "").replace('"', "").strip()
    if re.search(r"[A-Za-z\-]", s):
        return False
    return bool(re.fullmatch(r"-?\d+(\.\d+)?", s))

def parse_numeric_tail(tokens: List[str]) -> Tuple[Optional[float], Optional[float], Optional[float], Optional[float]]:
    """
    Extract numeric tail values from a merged line.

    Key fix:
    - Ignore integers like '150' that appear in description.
    - Start parsing numbers only from the FIRST token that looks like money: r'^\d+(\,\d{3})*(\.\d{2})$'
      (i.e. has exactly 2 decimal places)
    """
    def clean_num_text(t: str) -> str:
        return t.replace(",", "").replace('"', "").strip()

    money_pat = re.compile(r"^-?\d{1,3}(?:,\d{3})*(?:\.\d{2})$|^-?\d+(?:\.\d{2})$")

    # 1) Find first money-like token position (e.g. 559.00). Ignore earlier integers (e.g. 150).
    start_i = None
    for i, t in enumerate(tokens):
        tt = clean_num_text(t)
        if money_pat.fullmatch(tt):
            start_i = i
            break

    if start_i is None:
        return None, None, None, None

    tail_tokens = tokens[start_i:]

    # 2) Collect numeric tokens from that point onward
    num_texts: List[str] = []
    num_vals: List[Optional[float]] = []
    for t in tail_tokens:
        tt = clean_num_text(t)
        if re.fullmatch(r"-?\d+(\.\d+)?", tt):  # numeric
            num_texts.append(tt)
            try:
                num_vals.append(float(tt))
            except ValueError:
                num_vals.append(None)

    n = len(num_texts)
    if n < 3:
        return None, None, None, None

    # Keep only last 7 numeric tokens if too long
    if n > 7:
        num_texts = num_texts[-7:]
        num_vals = num_vals[-7:]
        n = len(num_texts)

    # 3) Detect yuan (same rule as before) but only if last token looks like money (2dp)
    has_yuan = False
    if n >= 4:
        last_text = num_texts[-1]
        last_val = num_vals[-1]
        if last_val is not None and (".") in last_text and 0 <= last_val < 1000:
            has_yuan = True

    def to_float(text: str) -> Optional[float]:
        try:
            return float(text)
        except ValueError:
            return None

    sale_val = None
    stock_val = None
    on_order_val = None
    yuan_val = None

    # 4) Your existing mapping logic (unchanged) — now works better because 150 is removed
    if has_yuan:
        yuan_val = to_float(num_texts[-1])
        if n >= 7:
            sale_val = to_float(num_texts[-5])
            stock_val = to_float(num_texts[-4])
            on_order_val = to_float(num_texts[-2])
        else:
            sale_val = None
            stock_val = to_float(num_texts[-4])
            on_order_val = to_float(num_texts[-2])
    else:
        if n >= 6:
            sale_val = to_float(num_texts[-4])
            stock_val = to_float(num_texts[-3])
            on_order_val = to_float(num_texts[-1])
        else:
            sale_val = None
            stock_val = to_float(num_texts[-3])
            on_order_val = to_float(num_texts[-1])

    return sale_val, stock_val, on_order_val, yuan_val


def parse_line_to_fields_merged(line: str) -> Optional[Dict[str, object]]:
    m = re.match(r"\s*([0-9A-Za-z]{5})\b(.*)", line)
    if not m:
        return None
    buyer = m.group(1)
    rest = m.group(2).strip()
    tokens = rest.split()
    if not tokens:
        return None

    barcode = ""
    idx = 0
    if idx < len(tokens) and re.fullmatch(r"\d+", tokens[idx]):
        barcode = tokens[idx]
        idx += 1

    product_tokens = tokens[idx:]
    while product_tokens and product_tokens[0] in {".", "}"}:
        product_tokens = product_tokens[1:]

    if len(product_tokens) < 2:
        return None

    sale, stock, on_order, yuan = parse_numeric_tail(tokens)

    if stock is None or on_order is None:
        return None
    if sale is None:
        sale = 0.0

    end_idx = len(product_tokens)
    while end_idx > 0 and is_num_token(product_tokens[end_idx - 1]):
        end_idx -= 1
    product_str = " ".join(product_tokens[:end_idx]).strip()

    # Uncomment if you need debug:
    # if yuan is None:
    #     nums_dbg = [t for t in tokens if is_num_token(t)]
    #     print("⚠️ yuan missing | buyer=", buyer, "| nums_tail=", nums_dbg[-7:], "| line=", line[:160])

    return {
        "buyer": buyer,
        "barcode": barcode,
        "สินค้า": product_str,
        "ยอดขาย": sale,
        "สินค้าคงเหลือ": stock,
        "ON_ORDER": on_order if on_order is not None else 0.0,
        "หยวน": yuan,
    }


# =========================
# Parse express file -> DF
# =========================
def parse_express_file(path: str, source_label: str) -> Tuple[pd.DataFrame, Dict[str, object]]:
    print(f">>> Parsing Express file: {path} ({source_label})")

    df_raw = pd.read_excel(path, sheet_name=EXPRESS_SHEET, header=None, dtype=str)
    date_info = parse_date_range_from_header(df_raw)

    rows = []
    for idx, row in df_raw.iterrows():
        merged = row_to_merged_line(row)
        if not merged:
            continue
        if is_header_or_separator(merged):
            continue

        first_token = merged.split(maxsplit=1)[0]
        if not looks_like_buyer_code(first_token):
            continue

        fields = parse_line_to_fields_merged(merged)
        if fields is None:
            continue

        fields["source"] = source_label
        fields["src_row"] = int(idx) + 1
        fields["src_file"] = os.path.basename(path)

        rows.append(fields)

    df = pd.DataFrame(rows)
    if not df.empty:
        df[["รหัสสินค้า", "รายละเอียดสินค้า"]] = df["สินค้า"].apply(lambda x: pd.Series(split_product_field(x)))
        df.drop(columns=["สินค้า"], inplace=True)

    return df, date_info


# =========================
# Combine ASIA + GREEN
# =========================
def combine_asia_green(df_asia: pd.DataFrame, df_green: pd.DataFrame,
                       months: int, min_factor: int = 4, max_factor: int = 7) -> pd.DataFrame:

    def agg_df(df: pd.DataFrame, label: str) -> pd.DataFrame:
        key_cols = ["buyer", "รหัสสินค้า"]

        if df.empty:
            return pd.DataFrame(columns=key_cols + [
                f"ยอดขาย_{label}", f"STOCK_{label}", f"ON_ORDER_{label}", f"หยวน_{label}",
                "barcode", "รายละเอียดสินค้า"
            ])

        g = df.groupby(key_cols, as_index=False).agg({
            "ยอดขาย": "sum",
            "สินค้าคงเหลือ": "sum",
            "ON_ORDER": "sum",
            "หยวน": "first",
            "barcode": "first",
            "รายละเอียดสินค้า": "first",
        })
        g = g.rename(columns={
            "ยอดขาย": f"ยอดขาย_{label}",
            "สินค้าคงเหลือ": f"STOCK_{label}",
            "ON_ORDER": f"ON_ORDER_{label}",
            "หยวน": f"หยวน_{label}",
        })
        return g

    g_asia = agg_df(df_asia, "ASIA")
    g_green = agg_df(df_green, "GREEN")

    combined = pd.merge(g_asia, g_green, on=["buyer", "รหัสสินค้า"], how="outer", suffixes=("", "_dup"))

    def coalesce(a, b):
        return a if pd.notna(a) and a != "" else b

    combined["barcode"] = [coalesce(a, b) for a, b in zip(
        combined.get("barcode_x", [None]*len(combined)),
        combined.get("barcode_y", [None]*len(combined))
    )]
    combined["รายละเอียดสินค้า"] = [coalesce(a, b) for a, b in zip(
        combined.get("รายละเอียดสินค้า_x", [None]*len(combined)),
        combined.get("รายละเอียดสินค้า_y", [None]*len(combined))
    )]
    for col in ["barcode_x", "barcode_y", "รายละเอียดสินค้า_x", "รายละเอียดสินค้า_y"]:
        if col in combined.columns:
            combined.drop(columns=[col], inplace=True)

    for col in ["ยอดขาย_ASIA", "STOCK_ASIA", "ON_ORDER_ASIA",
                "ยอดขาย_GREEN", "STOCK_GREEN", "ON_ORDER_GREEN"]:
        if col not in combined.columns:
            combined[col] = 0.0
        else:
            combined[col] = combined[col].fillna(0.0)

    combined["ยอดขาย_TOTAL"] = combined["ยอดขาย_ASIA"] + combined["ยอดขาย_GREEN"]
    combined["ON_ORDER_TOTAL"] = combined["ON_ORDER_ASIA"] + combined["ON_ORDER_GREEN"]

    def pick_yuan(row):
        yG = row.get("หยวน_GREEN", np.nan)
        yA = row.get("หยวน_ASIA", np.nan)
        if pd.notna(yG):
            return float(yG)
        if pd.notna(yA):
            return float(yA)
        return np.nan

    combined["หยวน"] = combined.apply(pick_yuan, axis=1)

    if months <= 0:
        months = 1

    combined["USE_MONTH"] = combined["ยอดขาย_TOTAL"].apply(lambda v: round_half_up(v / months) if v > 0 else 0)

    combined["TOTAL_QTY_NUM"] = combined["STOCK_ASIA"] + combined["STOCK_GREEN"] + combined["ON_ORDER_TOTAL"]
    combined["MIN_NUM"] = combined["USE_MONTH"] * int(min_factor)
    combined["MAX_NUM"] = combined["USE_MONTH"] * int(max_factor)

    mask = combined["TOTAL_QTY_NUM"] < combined["MIN_NUM"]

    filtered = combined[mask].reset_index(drop=True)
    print(f">>> Items after TOTAL_QTY < MIN filter: {len(filtered)}")

    return filtered


# =========================
# Catalog map + images
# =========================
def build_catalog_map(catalog_path: str, vendor_code: str) -> dict:

    wb = openpyxl.load_workbook(catalog_path)

    # ===== pick vendor sheet (tab name = vendor code) =====
    want = str(vendor_code).strip().upper()
    norm_map = {str(name).strip().upper(): name for name in wb.sheetnames}

    if want in norm_map:
        ws = wb[norm_map[want]]
    else:
        ws = None
        for k_norm, original in norm_map.items():
            if want in k_norm or k_norm in want:
                ws = wb[original]
                break
        if ws is None:
            raise RuntimeError(
                f"Catalog workbook has no sheet for vendor '{want}'. "
                f"Available sheets: {wb.sheetnames}"
            )

    # ===== FIXED COLUMN MAPPING (by Excel column) =====
    COL_ITEM_NO   = 1  # A
    COL_PICTURE   = 2  # B
    COL_DESC      = 3  # C
    COL_BRAND     = 4  # D
    COL_MATERIAL  = 5  # E
    COL_WEIGHT    = 6  # F
    COL_QTY_CT    = 7  # G

    HEADER_ROW = 1  # data starts below this

    # ===== read images (optional) =====
    img_at = {}
    for img in ws._images:
        try:
            r = img.anchor._from.row + 1
            c = img.anchor._from.col + 1
            img_at[(r, c)] = img._data()
        except Exception:
            pass

    # ===== build catalog map =====
    catalog = {}

    for r in range(HEADER_ROW + 1, ws.max_row + 1):
        item_no = ws.cell(r, COL_ITEM_NO).value
        if not item_no:
            continue

        item_no = str(item_no).strip()

        catalog[item_no] = {
            "goods_desc": ws.cell(r, COL_DESC).value,
            "brand": ws.cell(r, COL_BRAND).value,
            "material": ws.cell(r, COL_MATERIAL).value,
            "weight": ws.cell(r, COL_WEIGHT).value,
            "qty_per_carton": ws.cell(r, COL_QTY_CT).value,
            "img_bytes": img_at.get((r, COL_PICTURE)),
        }

    print(f">>> Catalog loaded from sheet '{ws.title}': {len(catalog)} items")
    return catalog


# =========================
# Excel helpers
# =========================
def copy_column_widths(src_ws, dst_ws):
    for col_letter, dim in src_ws.column_dimensions.items():
        if dim.width is not None:
            dst_ws.column_dimensions[col_letter].width = dim.width

def force_bottom_border(ws, row: int, start_col: int, end_col: int):
    thin = Side(style="thin")
    for c in range(start_col, end_col + 1):
        cell = ws.cell(row, c)
        b = _copy(cell.border) if cell.border else Border()
        b.bottom = thin
        cell.border = b

def copy_row_style(ws, src_row: int, dst_row: int, max_col: int):
    for c in range(1, max_col + 1):
        src = ws.cell(src_row, c)
        dst = ws.cell(dst_row, c)
        if src.has_style:
            dst._style = src._style

def copy_row_height(ws, src_row: int, dst_row: int):
    ws.row_dimensions[dst_row].height = ws.row_dimensions[src_row].height

def copy_template_rows(src_ws, dst_ws, src_start, src_end, dst_start, max_col):
    row_offset = dst_start - src_start

    for r in range(src_start, src_end + 1):
        for c in range(1, max_col + 1):
            src = src_ws.cell(r, c)
            dst = dst_ws.cell(r + row_offset, c)
            dst.value = src.value
            if src.has_style:
                dst._style = src._style

    for r in range(src_start, src_end + 1):
        dst_ws.row_dimensions[r + row_offset].height = src_ws.row_dimensions[r].height

    for merged in src_ws.merged_cells.ranges:
        if merged.min_row >= src_start and merged.max_row <= src_end and merged.max_col <= max_col:
            dst_ws.merge_cells(
                start_row=merged.min_row + row_offset,
                start_column=merged.min_col,
                end_row=merged.max_row + row_offset,
                end_column=merged.max_col,
            )

def add_signature_under_last_item(ws, last_item_row: int,
                                 sig_path: str = "footer_signatures.png",
                                 gap_rows: int = 4,
                                 anchor_col: str = "A"):
    if not os.path.exists(sig_path):
        print(f"⚠️ Signature file not found: {sig_path}")
        return
    sig = XLImage(sig_path)
    sig.width = 1000
    sig.height = 750
    sig_row = last_item_row + gap_rows
    ws.add_image(sig, f"{anchor_col}{sig_row}")


# =========================
# IMAGE placement
# =========================
def _excel_colwidth_to_pixels(width):
    if width is None:
        width = 8.43
    return int(width * 7 + 5)

def _excel_rowheight_to_pixels(height_pts):
    if height_pts is None:
        height_pts = 15
    return int(height_pts * 96 / 72)

def _get_cell_rect_pixels(ws, col_letter, row_num):
    col_w = _excel_colwidth_to_pixels(ws.column_dimensions[col_letter].width)
    row_h = _excel_rowheight_to_pixels(ws.row_dimensions[row_num].height)

    for mr in ws.merged_cells.ranges:
        if mr.min_col <= column_index_from_string(col_letter) <= mr.max_col and mr.min_row <= row_num <= mr.max_row:
            total_w = 0
            for c in range(mr.min_col, mr.max_col + 1):
                letter = get_column_letter(c)
                total_w += _excel_colwidth_to_pixels(ws.column_dimensions[letter].width)
            total_h = 0
            for r in range(mr.min_row, mr.max_row + 1):
                total_h += _excel_rowheight_to_pixels(ws.row_dimensions[r].height)
            return total_w, total_h

    return col_w, row_h

def add_image_to_cell(ws, cell_addr: str, img_bytes: bytes,
                      width_boost: float = IMAGE_WIDTH_BOOST,
                      padding_px: int = IMAGE_PADDING_PX):
    if not img_bytes:
        return

    col_letter, row_num = coordinate_from_string(cell_addr)
    row_num = int(row_num)

    cell_w_px, cell_h_px = _get_cell_rect_pixels(ws, col_letter, row_num)

    max_w = max(1, int((cell_w_px - padding_px) * width_boost))
    max_w = min(max_w, cell_w_px - padding_px)
    max_h = max(1, cell_h_px - padding_px)

    pil = PILImage.open(BytesIO(img_bytes)).convert("RGBA")
    w, h = pil.size

    scale = min(max_w / w, max_h / h, 1.0)
    new_w = max(1, int(w * scale))
    new_h = max(1, int(h * scale))
    pil = pil.resize((new_w, new_h))

    bio = BytesIO()
    pil.save(bio, format="PNG")
    bio.seek(0)

    img = XLImage(bio)
    img.width = new_w
    img.height = new_h

    if not hasattr(ws, "_img_buffers"):
        ws._img_buffers = []
    ws._img_buffers.append(bio)

    col_idx0 = column_index_from_string(col_letter) - 1
    row_idx0 = row_num - 1

    x_off_px = max(0, int((cell_w_px - new_w) / 2))
    y_off_px = max(0, int((cell_h_px - new_h) / 2))

    marker = AnchorMarker(
        col=col_idx0, colOff=pixels_to_EMU(x_off_px),
        row=row_idx0, rowOff=pixels_to_EMU(y_off_px)
    )

    img.anchor = OneCellAnchor(
        _from=marker,
        ext=XDRPositiveSize2D(pixels_to_EMU(new_w), pixels_to_EMU(new_h))
    )
    ws.add_image(img)


# =========================
# PO column mapping
# =========================
def norm_header(x) -> str:
    return str(x).replace("\n", " ").replace("\xa0", " ").strip() if x is not None else ""

def get_po_col_map(ws, header_row=HEADER_ROW):
    col_map = {}
    for c in range(1, ws.max_column + 1):
        v = norm_header(ws.cell(header_row, c).value)
        if v:
            col_map[v] = c
    return col_map


# =========================
# TOTAL block
# =========================
def paste_total_block_and_fix(
    ws,
    template_ws,
    total_block_start: int,
    po_last_col: int,
    item_start_row: int,
    last_item_row: int,
    col_amt: str,
    rate_thb_per_cny: float,
):
    copy_template_rows(
        template_ws, ws,
        TEMPLATE_TOTAL_START_ROW, TEMPLATE_TOTAL_END_ROW,
        total_block_start,
        max_col=po_last_col
    )

    r1 = total_block_start
    r_mid = r1 + 1
    r2 = r1 + 2

    sum_pat = re.compile(
        rf'(SUM\()(\$?[A-Z]{{1,3}}\$?){item_start_row}:(\$?[A-Z]{{1,3}}\$?)13(\))',
        flags=re.IGNORECASE
    )
    for row in ws.iter_rows(min_row=r1, max_row=r2, min_col=1, max_col=po_last_col):
        for cell in row:
            v = cell.value
            if isinstance(v, str) and v.startswith("=") and "SUM" in v.upper():
                cell.value = sum_pat.sub(
                    rf'\g<1>\g<2>{item_start_row}:\g<3>{last_item_row}\g<4>',
                    v
                )

    ws[f"O{r1}"].value = None
    ws[f"O{r2}"].value = None

    ws[f"N{r1}"].value = f"=SUM({col_amt}{item_start_row}:{col_amt}{last_item_row})"
    ws[f"N{r_mid}"].value = rate_thb_per_cny
    ws[f"N{r2}"].value = f"=N{r_mid}*N{r1}"


def find_label_cell(ws, label: str, max_row: int = 60, max_col: int = 30):
    target = norm_text(label)
    for r in range(1, min(max_row, ws.max_row) + 1):
        for c in range(1, min(max_col, ws.max_column) + 1):
            if norm_text(ws.cell(r, c).value) == target:
                return r, c
    return None


# =========================
# Generate PO
# =========================
def generate_po_from_combined(
    combined_df: pd.DataFrame,
    vendor_code: str,
    po_date: Optional[datetime.date] = None,
    rate_thb_per_cny: float = 6.0,
    template_path: str = TEMPLATE_PO_XLSX,
    catalog_path: str = CATALOG_XLSX,
    vendor_info_path: str = VENDOR_INFO_XLSX,
    min_factor: int = 4,
    max_factor: int = 7,
) -> str:

    if po_date is None:
        po_date = datetime.date.today()

    os.makedirs(PO_OUTPUT_FOLDER, exist_ok=True)
    output_path = os.path.join(PO_OUTPUT_FOLDER, f"PO_{vendor_code}.xlsx")

    vendor_key = str(vendor_code).strip().upper()

    vendor_map_raw = load_vendor_map(vendor_info_path)
    vendor_map = {str(k).strip().upper(): v for k, v in vendor_map_raw.items()}

    supplier_name = vendor_map.get(vendor_key, {}).get("name", "")
    supplier_addr = vendor_map.get(vendor_key, {}).get("address", "")

    catalog_map = {}
    if os.path.exists(catalog_path):
        catalog_map = build_catalog_map(catalog_path, vendor_code=vendor_code)

    else:
        print(f"⚠️ Catalog not found: {catalog_path}")

    wb = openpyxl.load_workbook(template_path)
    template_ws = wb[TEMPLATE_SHEET_NAME]
    ws = wb.copy_worksheet(template_ws)
    ws.title = "PO"

    copy_column_widths(template_ws, ws)
    ws.column_dimensions["B"].width = (ws.column_dimensions["B"].width or 18) + 6

    po_cols = get_po_col_map(ws, header_row=HEADER_ROW)

    def _find_col_key(possible_keys):
        for k in possible_keys:
            if k in po_cols:
                return k
        return None

    min_key_old = _find_col_key(["MIN*4", "MIN * 4", "MIN×4", "MIN x4", "MINX4"])
    max_key_old = _find_col_key(["MAX*7", "MAX * 7", "MAX×7", "MAX x7", "MAXX7"])

    if not min_key_old or not max_key_old:
        raise RuntimeError(f"Cannot find MIN/MAX header in template.")

    min_col_idx = po_cols[min_key_old]
    max_col_idx = po_cols[max_key_old]

    ws.cell(HEADER_ROW, min_col_idx).value = f"MIN*{int(min_factor)}"
    ws.cell(HEADER_ROW, max_col_idx).value = f"MAX*{int(max_factor)}"

    col_min = get_column_letter(min_col_idx)
    col_max = get_column_letter(max_col_idx)

    PO_LAST_COL = max(po_cols.values())

    col_cart = get_column_letter(po_cols["CARTONS"])
    col_tot_order = get_column_letter(po_cols["TOTAL QTY (ORDER)"])
    col_amt = get_column_letter(po_cols["AMOUNT (CNY)"])
    col_thb = get_column_letter(po_cols["THB"])

    ws["H6"] = vendor_code
    ws["H6"].font = Font(color="FF0000", bold=True, size=18)

    pos = find_label_cell(ws, "DATE", max_row=20, max_col=30)
    if pos:
        r, c = pos
        ws.cell(r, c + 1).value = po_date

    pos = find_label_cell(ws, "SUPPLIER", max_row=20, max_col=30)
    if pos and supplier_name:
        r, c = pos
        ws.cell(r, c + 1).value = supplier_name

    pos = find_label_cell(ws, "ADDRESS", max_row=25, max_col=30)
    if pos and supplier_addr:
        r, c = pos
        ws.cell(r, c + 1).value = supplier_addr

    if os.path.exists("logo.png"):
        logo_img = XLImage("logo.png")
        logo_img.width = 210
        logo_img.height = 110
        ws.add_image(logo_img, "A1")

    combined_df = (
        combined_df.sort_values(by=["รหัสสินค้า", "รายละเอียดสินค้า"], ascending=[True, True])
        .reset_index(drop=True)
    )

    current_row = ITEM_START_ROW

    for _, row in combined_df.iterrows():
        line = current_row
        current_row += 1

        copy_row_style(ws, TEMPLATE_ITEM_ROW, line, PO_LAST_COL)
        copy_row_height(ws, TEMPLATE_ITEM_ROW, line)

        buyer_item = str(row["รหัสสินค้า"]).strip()
        cat = catalog_map.get(buyer_item, {})

        qty_per_carton = cat.get("qty_per_carton", "")
        try:
            qty_per_carton_num = float(qty_per_carton) if qty_per_carton not in [None, ""] else 0.0
        except:
            qty_per_carton_num = 0.0

        use_month = int(row["USE_MONTH"]) if not pd.isna(row["USE_MONTH"]) else 0
        total_qty_num = float(row["TOTAL_QTY_NUM"])
        max_num = float(row["MAX_NUM"])
        units_to_order = max_num - total_qty_num
        if units_to_order < 0:
            units_to_order = 0

        yuan = row["หยวน"] if not pd.isna(row["หยวน"]) else None
        try:
            yuan_num = float(yuan) if yuan is not None else None
        except:
            yuan_num = None

        ws.cell(line, po_cols["BUYER ITEM NO."]).value = buyer_item

        if cat.get("img_bytes"):
            add_image_to_cell(ws, f"B{line}", cat["img_bytes"])

        ws.cell(line, po_cols["GOODS DESCRIPTION"]).value = cat.get("goods_desc", row.get("รายละเอียดสินค้า", ""))
        ws.cell(line, po_cols["BRAND"]).value = cat.get("brand", "")
        ws.cell(line, po_cols["MATERIAL"]).value = cat.get("material", "")
        ws.cell(line, po_cols["Weight"]).value = cat.get("weight", "")
        ws.cell(line, po_cols["QTY PER CARTON"]).value = qty_per_carton_num if qty_per_carton_num > 0 else None

        ws.cell(line, po_cols["STOCK GREEN"]).value = float(row["STOCK_GREEN"])
        ws.cell(line, po_cols["STOCK ASIA"]).value = float(row["STOCK_ASIA"])
        ws.cell(line, po_cols["ON ORDER"]).value = float(row["ON_ORDER_TOTAL"])
        ws.cell(line, po_cols["USE MONTH"]).value = use_month

        col_use = get_column_letter(po_cols["USE MONTH"])
        col_sg = get_column_letter(po_cols["STOCK GREEN"])
        col_sa = get_column_letter(po_cols["STOCK ASIA"])
        col_on = get_column_letter(po_cols["ON ORDER"])
        col_tq = get_column_letter(po_cols["TOTAL QTY"])
        col_zan = get_column_letter(po_cols["จน./USE MONTH"])
        col_remain0 = get_column_letter(po_cols["คงเหลือ (จน./USE MONTH เดิม)"])
        col_qpc = get_column_letter(po_cols["QTY PER CARTON"])
        col_green = get_column_letter(po_cols["GREEN"])
        col_asia = get_column_letter(po_cols["ASIA"])
        col_fob = get_column_letter(po_cols["FOB PRICE (CNY)"])

        ws[f"{col_min}{line}"] = f"={col_use}{line}*{int(min_factor)}"
        ws[f"{col_max}{line}"] = f"={col_use}{line}*{int(max_factor)}"

        ws[f"{col_tq}{line}"] = f"={col_sg}{line}+{col_sa}{line}+{col_on}{line}"
        ws[f"{col_zan}{line}"] = f"=ROUND(({col_tq}{line}+{col_tot_order}{line})/{col_use}{line},0)"
        ws[f"{col_remain0}{line}"] = f"=ROUND({col_tq}{line}/{col_use}{line},0)"

        ws[f"{col_cart}{line}"] = f"=ROUND(({col_max}{line}-{col_tq}{line})/{col_qpc}{line},0)"
        ws[f"{col_green}{line}"] = f"={col_cart}{line}*{col_qpc}{line}"
        ws[f"{col_asia}{line}"] = 0
        ws[f"{col_tot_order}{line}"] = f"={col_green}{line}+{col_asia}{line}"

        if yuan_num is not None:
            ws[f"{col_fob}{line}"] = yuan_num
            ws[f"{col_thb}{line}"] = yuan_num * rate_thb_per_cny
        else:
            ws[f"{col_fob}{line}"] = None
            ws[f"{col_thb}{line}"] = None

        ws[f"{col_amt}{line}"] = f"={col_fob}{line}*{col_tot_order}{line}"

    if len(combined_df) > 0:
        last_item_row = ITEM_START_ROW + len(combined_df) - 1
        force_bottom_border(ws, last_item_row, 1, PO_LAST_COL)

        total_block_start = last_item_row + 1
        paste_total_block_and_fix(
            ws=ws,
            template_ws=template_ws,
            total_block_start=total_block_start,
            po_last_col=PO_LAST_COL,
            item_start_row=ITEM_START_ROW,
            last_item_row=last_item_row,
            col_amt=col_amt,
            rate_thb_per_cny=rate_thb_per_cny,
        )

        total_block_end = total_block_start + 2
        add_signature_under_last_item(
            ws,
            last_item_row=total_block_end,
            sig_path="footer_signatures.png",
            gap_rows=4,
            anchor_col="A",
        )

    wb.remove(template_ws)
    wb.save(output_path)
    print(f"✔ PO created: {output_path}")
    return output_path

HIGHLIGHT_BELOW_MIN = PatternFill(fill_type="solid", start_color="FFF2CC", end_color="FFF2CC")  # light yellow

def generate_po_streamlit(
    express_asia_path: str,
    express_green_path: str,
    catalog_path: str,
    vendor_info_path: str,
    template_path: str,
    vendor_code: str,
    po_date,
    rate_thb_per_cny: float,
    min_factor: int = 4,
    max_factor: int = 7,
) -> dict:
    """
    Returns dict:
      {
        "po_filtered": <path or None>,
        "po_all_items": <path always>,
        "count_all": int,
        "count_filtered": int
      }
    """
    vendor_code = str(vendor_code).strip().upper()

    # 1) parse each file
    df_asia, date_info_asia = parse_express_file(express_asia_path, "ASIA")
    df_green, date_info_green = parse_express_file(express_green_path, "GREEN")

    for df in (df_asia, df_green):
        if not df.empty and "buyer" in df.columns:
            df["buyer"] = df["buyer"].astype(str).str.replace("\u00A0", " ", regex=False).str.strip().str.upper()

    # 2) choose months
    months = 1
    if date_info_asia and date_info_asia.get("months", 0) > 0:
        months = int(date_info_asia["months"])
    elif date_info_green and date_info_green.get("months", 0) > 0:
        months = int(date_info_green["months"])

    # 3) Build combined_all (no filter) using same logic
    def _agg_df(df: pd.DataFrame, label: str) -> pd.DataFrame:
        key_cols = ["buyer", "รหัสสินค้า"]
        if df.empty:
            return pd.DataFrame(columns=key_cols + [
                f"ยอดขาย_{label}", f"STOCK_{label}", f"ON_ORDER_{label}", f"หยวน_{label}",
                "barcode", "รายละเอียดสินค้า"
            ])
        g = df.groupby(key_cols, as_index=False).agg({
            "ยอดขาย": "sum",
            "สินค้าคงเหลือ": "sum",
            "ON_ORDER": "sum",
            "หยวน": "first",
            "barcode": "first",
            "รายละเอียดสินค้า": "first",
        })
        return g.rename(columns={
            "ยอดขาย": f"ยอดขาย_{label}",
            "สินค้าคงเหลือ": f"STOCK_{label}",
            "ON_ORDER": f"ON_ORDER_{label}",
            "หยวน": f"หยวน_{label}",
        })

    g_asia = _agg_df(df_asia, "ASIA")
    g_green = _agg_df(df_green, "GREEN")

    combined_all = pd.merge(g_asia, g_green, on=["buyer", "รหัสสินค้า"], how="outer", suffixes=("", "_dup"))

    def _coalesce(a, b):
        return a if pd.notna(a) and a != "" else b

    combined_all["barcode"] = [
        _coalesce(a, b) for a, b in zip(
            combined_all.get("barcode_x", [None] * len(combined_all)),
            combined_all.get("barcode_y", [None] * len(combined_all)),
        )
    ]
    combined_all["รายละเอียดสินค้า"] = [
        _coalesce(a, b) for a, b in zip(
            combined_all.get("รายละเอียดสินค้า_x", [None] * len(combined_all)),
            combined_all.get("รายละเอียดสินค้า_y", [None] * len(combined_all)),
        )
    ]
    for col in ["barcode_x", "barcode_y", "รายละเอียดสินค้า_x", "รายละเอียดสินค้า_y"]:
        if col in combined_all.columns:
            combined_all.drop(columns=[col], inplace=True)

    for col in ["ยอดขาย_ASIA", "STOCK_ASIA", "ON_ORDER_ASIA",
                "ยอดขาย_GREEN", "STOCK_GREEN", "ON_ORDER_GREEN"]:
        if col not in combined_all.columns:
            combined_all[col] = 0.0
        else:
            combined_all[col] = combined_all[col].fillna(0.0)

    combined_all["ยอดขาย_TOTAL"] = combined_all["ยอดขาย_ASIA"] + combined_all["ยอดขาย_GREEN"]
    combined_all["ON_ORDER_TOTAL"] = combined_all["ON_ORDER_ASIA"] + combined_all["ON_ORDER_GREEN"]

    def _pick_yuan(row):
        yG = row.get("หยวน_GREEN", np.nan)
        yA = row.get("หยวน_ASIA", np.nan)
        if pd.notna(yG):
            return float(yG)
        if pd.notna(yA):
            return float(yA)
        return np.nan

    combined_all["หยวน"] = combined_all.apply(_pick_yuan, axis=1)

    if months <= 0:
        months = 1

    combined_all["USE_MONTH"] = combined_all["ยอดขาย_TOTAL"].apply(
        lambda v: round_half_up(v / months) if v > 0 else 0
    )
    combined_all["TOTAL_QTY_NUM"] = combined_all["STOCK_ASIA"] + combined_all["STOCK_GREEN"] + combined_all["ON_ORDER_TOTAL"]
    combined_all["MIN_NUM"] = combined_all["USE_MONTH"] * int(min_factor)
    combined_all["MAX_NUM"] = combined_all["USE_MONTH"] * int(max_factor)

    # vendor rows ALL (no filter)
    vendor_rows_all = combined_all[combined_all["buyer"] == vendor_code].copy()
    if vendor_rows_all.empty:
        buyers = sorted(set(combined_all["buyer"].dropna().astype(str).tolist()))
        preview = ", ".join(buyers[:30])
        raise ValueError(f"Vendor '{vendor_code}' not found. Parsed buyers (first 30): {preview}")

    # Export ALL items file (always)
    path_all = export_vendor_all_items_excel(vendor_rows_all, vendor_code=vendor_code)

    # vendor rows FILTERED (below MIN)
    vendor_rows_filtered = vendor_rows_all[vendor_rows_all["TOTAL_QTY_NUM"] < vendor_rows_all["MIN_NUM"]].copy()

    path_filtered = None
    if not vendor_rows_filtered.empty:
        path_filtered = generate_po_from_combined(
            combined_df=vendor_rows_filtered,
            vendor_code=vendor_code,
            po_date=po_date,
            rate_thb_per_cny=float(rate_thb_per_cny),
            template_path=template_path,
            catalog_path=catalog_path,
            vendor_info_path=vendor_info_path,
            min_factor=int(min_factor),
            max_factor=int(max_factor),
        )

    return {
        "po_filtered": path_filtered,
        "po_all_items": path_all,
        "count_all": int(len(vendor_rows_all)),
        "count_filtered": int(len(vendor_rows_filtered)),
    }

def export_vendor_all_items_excel(
    vendor_rows_all: pd.DataFrame,
    vendor_code: str,
    out_folder: str = PO_OUTPUT_FOLDER,
) -> str:
    """
    Export ALL items for the vendor (no MIN filter) into Excel.
    Highlight rows where TOTAL_QTY_NUM < MIN_NUM.
    """
    os.makedirs(out_folder, exist_ok=True)
    vendor_code = str(vendor_code).strip().upper()
    out_path = os.path.join(out_folder, f"PO_{vendor_code}_ALL_ITEMS.xlsx")

    cols_wanted = [
        "buyer", "รหัสสินค้า", "รายละเอียดสินค้า", "barcode",
        "ยอดขาย_ASIA", "STOCK_ASIA", "ON_ORDER_ASIA", "หยวน_ASIA",
        "ยอดขาย_GREEN", "STOCK_GREEN", "ON_ORDER_GREEN", "หยวน_GREEN",
        "ยอดขาย_TOTAL", "ON_ORDER_TOTAL",
        "USE_MONTH", "TOTAL_QTY_NUM", "MIN_NUM", "MAX_NUM", "หยวน",
    ]
    cols = [c for c in cols_wanted if c in vendor_rows_all.columns]

    df_out = vendor_rows_all.copy()

    # Sort like PO
    if "รหัสสินค้า" in df_out.columns and "รายละเอียดสินค้า" in df_out.columns:
        df_out = df_out.sort_values(["รหัสสินค้า", "รายละเอียดสินค้า"], ascending=[True, True])

    # Write with pandas first
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        df_out[cols].to_excel(writer, sheet_name="all_items", index=False)

    # Re-open and highlight rows below MIN
    wb = openpyxl.load_workbook(out_path)
    ws = wb["all_items"]

    # Find the column indexes of TOTAL_QTY_NUM and MIN_NUM from header row 1
    header = {str(ws.cell(1, c).value).strip(): c for c in range(1, ws.max_column + 1)}

    if "TOTAL_QTY_NUM" not in header or "MIN_NUM" not in header:
        wb.save(out_path)
        return out_path

    col_total = header["TOTAL_QTY_NUM"]
    col_min = header["MIN_NUM"]

    # Apply fill for rows where total < min
    for r in range(2, ws.max_row + 1):
        total_v = ws.cell(r, col_total).value
        min_v = ws.cell(r, col_min).value

        try:
            total_f = float(total_v) if total_v is not None else None
            min_f = float(min_v) if min_v is not None else None
        except Exception:
            continue

        if total_f is not None and min_f is not None and total_f < min_f:
            for c in range(1, ws.max_column + 1):
                ws.cell(r, c).fill = HIGHLIGHT_BELOW_MIN

    wb.save(out_path)
    return out_path

# =========================
# MAIN
# =========================
def main():
    df_asia = pd.DataFrame()
    df_green = pd.DataFrame()
    months_global = None

    for path, label in EXPRESS_FILES:
        if not os.path.exists(path):
            print(f"⚠️ File not found, skipping: {path}")
            continue

        df_raw = pd.read_excel(path, sheet_name=EXPRESS_SHEET, header=None, dtype=str)
        date_info = parse_date_range_from_header(df_raw)
        months = date_info["months"] if date_info["months"] > 0 else 1

        if months_global is None:
            months_global = months
        elif months_global != months:
            print(f"⚠️ Months differ between files, using {months_global}")

        df_tmp, _ = parse_express_file(path, label)
        if label == "ASIA":
            df_asia = df_tmp
        else:
            df_green = df_tmp

    if months_global is None:
        months_global = 1

    print(f">>> Months used for USE_MONTH: {months_global}")
    combined = combine_asia_green(df_asia, df_green, months=months_global)

    vendor = input("กรุณาระบุรหัส Vendor เช่น A0029: ").strip()
    if not vendor:
        print("ไม่ได้ระบุ Vendor code, stop.")
        return

    date_str = input("ระบุวันที่ PO (รูปแบบ YYYY-MM-DD, ว่าง = วันนี้): ").strip()
    if date_str:
        try:
            po_date = datetime.datetime.strptime(date_str, "%Y-%m-%d").date()
        except ValueError:
            print("รูปแบบวันที่ไม่ถูกต้อง ใช้วันที่วันนี้แทน")
            po_date = datetime.date.today()
    else:
        po_date = datetime.date.today()

    vendor_combined = combined[combined["buyer"] == vendor].copy()
    if vendor_combined.empty:
        print(f"⚠️ No rows after filter for vendor {vendor}")
        return

    generate_po_from_combined(
        combined_df=vendor_combined,
        vendor_code=vendor,
        po_date=po_date,
        rate_thb_per_cny=6.0,
    )

if __name__ == "__main__":
    main()
import os
import re
import math
import unicodedata
import datetime
from typing import List, Dict, Optional, Tuple
from copy import copy as _copy
from io import BytesIO

import numpy as np
import pandas as pd
import openpyxl
from PIL import Image as PILImage

from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Border, Side, PatternFill, Alignment
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string
from openpyxl.utils.units import pixels_to_EMU
from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, AnchorMarker
from openpyxl.drawing.xdr import XDRPositiveSize2D


# =========================
# CONFIG (template/layout)
# =========================
EXPRESS_SHEET = 0

TEMPLATE_SHEET_NAME = "page1"
HEADER_ROW = 8
ITEM_START_ROW = 9
TEMPLATE_ITEM_ROW = 9
TEMPLATE_TOTAL_START_ROW = 14
TEMPLATE_TOTAL_END_ROW = 16

PO_OUTPUT_FOLDER = "output_PO"

IMAGE_WIDTH_BOOST = 1.20
IMAGE_PADDING_PX = 2

HIGHLIGHT_BELOW_MIN = PatternFill(fill_type="solid", start_color="FFF2CC", end_color="FFF2CC")
HIGHLIGHT_MISSING_CATALOG = PatternFill(fill_type="solid", fgColor="FFFF00")
BARCODE_MISMATCH_NOTE = "barcode ทั้ง2ไฟล์ไม่ตรงกัน"


# =========================
# THAI DATE
# =========================
THAI_MONTHS = {
    "ม.ค.": 1, "ก.พ.": 2, "มี.ค.": 3, "เม.ย.": 4,
    "พ.ค.": 5, "มิ.ย.": 6, "ก.ค.": 7, "ส.ค.": 8,
    "ก.ย.": 9, "ต.ค.": 10, "พ.ย.": 11, "ธ.ค.": 12
}


def thai_to_date(day: int, thai_month: str, thai_year: int) -> datetime.date:
    """Convert Thai BE date tokens into Gregorian date."""
    year = thai_year - 543
    thai_month = re.sub(r"\s+", "", thai_month)
    month = THAI_MONTHS.get(thai_month)
    if month is None:
        raise ValueError(f"Unknown Thai month token: {repr(thai_month)}")
    return datetime.date(year, month, day)


def last_day_of_month(year: int, month: int) -> int:
    """Return last day number of a given month."""
    if month == 12:
        next_first = datetime.date(year + 1, 1, 1)
    else:
        next_first = datetime.date(year, month + 1, 1)
    return (next_first - datetime.timedelta(days=1)).day


def add_months(dt: datetime.date, n: int) -> datetime.date:
    """Add N months to a date while keeping a valid day."""
    month = dt.month - 1 + n
    year = dt.year + month // 12
    month = month % 12 + 1
    day = min(dt.day, last_day_of_month(year, month))
    return datetime.date(year, month, day)


def calc_days_and_months(d1: datetime.date, d2: datetime.date) -> Tuple[int, int]:
    """
    Compute (days, months) where leftover >= 15 days counts as an extra month.
    """
    if d2 < d1:
        return 0, 0
    days = (d2 - d1).days
    months = 0
    cur = d1
    while True:
        nxt = add_months(cur, 1)
        if nxt <= d2:
            months += 1
            cur = nxt
        else:
            break
    leftover = (d2 - cur).days
    if leftover >= 15:
        months += 1
    return days, months


def parse_date_range_from_header(df_raw: pd.DataFrame) -> Dict[str, object]:
    """
    Extract Thai date range text like:
      'วันที่จาก 1 ม.ค. 2569 ถึง 31 ม.ค. 2569'
    and convert to months by the 15-day rule.
    """
    pattern = r"วันที่จาก\s+(\d+)\s+(\S+)\s+(\d+)\s+ถึง\s+(\d+)\s+(\S+)\s+(\d+)"
    col0 = df_raw.iloc[:, 0].astype(str)
    for val in col0:
        if "วันที่จาก" not in str(val):
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
        return {"raw_line": text, "start_date": start_date, "end_date": end_date, "days": days, "months": months}
    return {"raw_line": "", "start_date": None, "end_date": None, "days": 0, "months": 1}


# =========================
# SMALL UTILS
# =========================
def round_half_up(x: float) -> int:
    """Round half up (0.5 -> 1)."""
    return int(math.floor(x + 0.5))


def norm_text(v) -> str:
    """Normalize cell text."""
    if v is None:
        return ""
    return str(v).replace("\xa0", " ").strip()


def clean_cell(val) -> str:
    """Clean a raw cell value into single-space text."""
    if pd.isna(val):
        return ""
    s = str(val).replace("\xa0", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s


def fix_split_numbers_in_line(line: str) -> str:
    """
    Fix patterns where a number is split into two tokens (rare Excel formatting artifact).
    Example: '1' '234.00' -> '1234.00'
    """
    tokens = line.split()
    i = 0
    while i < len(tokens) - 1:
        cur = tokens[i]
        nxt = tokens[i + 1]
        if re.fullmatch(r"\d{1,2}", cur) and re.fullmatch(r"\d{3,}(?:\.\d+)?\"?", nxt):
            tokens[i] = cur + nxt
            del tokens[i + 1]
        else:
            i += 1
    return " ".join(tokens)


def row_to_merged_line(row: pd.Series) -> str:
    """
    Merge all non-empty cells for robust parsing of buyer/product text.
    (We do NOT rely on merging for yuan; yuan is found by scanning row cells.)
    """
    cells = [clean_cell(v) for v in row.tolist()]
    cells = [c for c in cells if c]
    raw = " ".join(cells).strip()
    return fix_split_numbers_in_line(raw)


def looks_like_buyer_code(token: str) -> bool:
    """Buyer code is 5 alphanumeric characters."""
    return bool(re.fullmatch(r"[0-9A-Za-z]{5}", token or ""))


def is_header_or_separator(line: str) -> bool:
    """Skip header-like rows."""
    s = (line or "").strip()
    if not s:
        return True
    if "BUYER" in s:
        return True
    if "วันที่จาก" in s:
        return True
    if re.search(r"-{5,}", s):
        return True
    return False

def build_dense_chunks_no_space(row: pd.Series, max_scan_cols: int = 60) -> List[Tuple[int, int, str]]:
    """
    Returns list of (start_idx, end_idx, chunk_text)
    chunk_text is concatenation of adjacent non-empty cells WITHOUT spaces.
    A chunk ends when we hit an empty cell.
    """
    chunks = []
    cur = []
    start = None

    def is_empty(v) -> bool:
        if v is None:
            return True
        s = str(v).replace("\xa0", " ").strip()
        return s == "" or s.lower() == "nan"

    n = min(len(row), max_scan_cols)

    for i in range(n):
        v = row.iloc[i]
        if is_empty(v):
            if cur:
                chunks.append((start, i - 1, "".join(cur)))
                cur = []
                start = None
            continue

        s = str(v).replace("\xa0", " ").strip()
        if start is None:
            start = i
        cur.append(s)

    if cur:
        chunks.append((start, n - 1, "".join(cur)))

    return chunks

# =========================
# PRODUCT SPLIT
# =========================
_DASH_TRANS = str.maketrans({
    "‐": "-", "-": "-", "‒": "-", "–": "-", "—": "-", "−": "-",
})


def split_product_field(s: str) -> Tuple[str, str]:
    """
    Split merged product string into (product_code, description).

    Supports:
      - doc token like 01-15-0730-D3-5 (ignore)
      - codes like NR123-60, DGS-2318, MC401-18, BM-150, NRW, etc
      - NoBM-150 => BM-150
    """
    if not isinstance(s, str):
        return "", ""
    s = s.strip()
    if not s:
        return "", ""

    s = s.translate(_DASH_TRANS)
    parts = s.split(maxsplit=1)
    if len(parts) < 2:
        return "", ""
    rest = parts[1].strip()
    if not rest:
        return "", ""

    def is_doc_token(t: str) -> bool:
        t = t.strip().translate(_DASH_TRANS)
        return bool(re.fullmatch(r"\d{2}-\d{2}-\d{3,6}(?:-[A-Za-z0-9]+)*", t))

    def is_product_code(t: str) -> bool:
        t = t.strip().translate(_DASH_TRANS)
        t2 = re.sub(r'^[Nn][Oo](?=[A-Za-z0-9])', '', t).strip()
        if not t2:
            return False
        t2 = re.sub(r"\([^()]*\)$", "", t2)
        if re.fullmatch(r"[A-Za-z]{1,6}[A-Za-z0-9]*-[A-Za-z0-9]+", t2):
            return True
        if re.fullmatch(r"[A-Za-z]{1,6}\d+[A-Za-z0-9]*", t2):
            return True
        if re.fullmatch(r"[A-Za-z]{2,5}", t2):  # e.g. NRW
            return True
        return False

    # A parenthetical suffix can contain Thai words and spaces. Capture the
    # full code before searching for the first Thai character in the line.
    rest_tokens = rest.split()
    while rest_tokens and is_doc_token(rest_tokens[0]):
        rest_tokens = rest_tokens[1:]
    code_with_suffix = re.match(
        r"^(?P<code>(?:[Nn][Oo])?[A-Za-z]{1,6}[A-Za-z0-9]*(?:-[A-Za-z0-9]+)*\([^()]*\)(?:-?[A-Za-z0-9]+)*)(?P<description>.*)$",
        " ".join(rest_tokens),
    )
    if code_with_suffix:
        code = re.sub(r"^[Nn][Oo](?=[A-Za-z0-9])", "", code_with_suffix.group("code"))
        return code, code_with_suffix.group("description").strip()
    if rest_tokens and is_product_code(rest_tokens[0]):
        code = re.sub(r"^[Nn][Oo](?=[A-Za-z0-9])", "", rest_tokens[0]).strip()
        return code, " ".join(rest_tokens[1:]).strip()

    m_th = re.search(r"[\u0E00-\u0E7F]", rest)
    if m_th:
        th_pos = m_th.start()
        pre_th = rest[:th_pos].strip()
        tail_th = rest[th_pos:].strip()
        tokens = pre_th.split()
        if not tokens:
            return "", tail_th

        while tokens and is_doc_token(tokens[0]):
            tokens = tokens[1:]
        if not tokens:
            return "", tail_th

        code_idx = None
        for i, t in enumerate(tokens):
            if is_product_code(t):
                code_idx = i
                break
        if code_idx is None:
            code_idx = 0

        code_raw = tokens[code_idx].translate(_DASH_TRANS)
        code = re.sub(r'^[Nn][Oo](?=[A-Za-z0-9])', '', code_raw).strip()

        desc_pre = " ".join(tokens[code_idx + 1:]).strip()
        desc = (desc_pre + " " + tail_th).strip() if desc_pre else tail_th
        return code, desc

    tokens = rest.split()
    while tokens and is_doc_token(tokens[0]):
        tokens = tokens[1:]
    if not tokens:
        return "", ""

    code_raw = tokens[0].translate(_DASH_TRANS)
    code = re.sub(r'^[Nn][Oo](?=[A-Za-z0-9])', '', code_raw).strip()
    desc = " ".join(tokens[1:]).strip()
    return code, desc


# =========================
# MONEY BLOCK + 5/6 NUMBERS
# =========================
_MONEY_2DP_RE = re.compile(r"-?\d+(?:,\d{3})*\.\d{2}")
_YUAN_RE = re.compile(r"^[Yy]?\s*(-?\d+(?:\.\d+)?)\s*$")


def _money_tokens_2dp_in_text(s: str) -> List[float]:
    out: List[float] = []
    if not s:
        return out
    for raw in _MONEY_2DP_RE.findall(s):
        try:
            out.append(float(raw.replace(",", "")))
        except Exception:
            pass
    return out


def extract_money_2dp_numbers_from_row(row: pd.Series, scan_cols: int = 25) -> List[float]:
    """
    Extract 2dp money numbers from the first `scan_cols` columns (left area).
    Keeps the order left->right.
    """
    out: List[float] = []
    for i in range(min(scan_cols, len(row))):
        v = row.iloc[i]
        if v is None:
            continue
        s = str(v).replace("\xa0", " ").strip()
        if not s:
            continue
        out.extend(_money_tokens_2dp_in_text(s))
    return out

def extract_5_6_block_from_money_chunk(row: pd.Series) -> Optional[List[float]]:
    """
    Extract 5/6 numbers from the detected dense chunk (not from whole row).
    """
    found = find_money_block_chunk(row, min_tokens=3, max_scan_cols=60)
    if not found:
        return None

    _, _, txt = found
    nums = _money_tokens_2dp_in_text(txt)

    if len(nums) >= 6:
        return nums[-6:]
    if len(nums) >= 5:
        return nums[-5:]
    return None


def find_money_block_chunk(row: pd.Series, min_tokens: int = 3, max_scan_cols: int = 60):
    """
    Find the chunk that contains >= min_tokens money 2dp tokens.
    Returns (start_idx, end_idx, chunk_text) or None.
    """
    chunks = build_dense_chunks_no_space(row, max_scan_cols=max_scan_cols)
    best = None

    for (s, e, txt) in chunks:
        cnt = len(_MONEY_2DP_RE.findall(txt))
        if cnt >= min_tokens:
            # pick the first matching chunk (or you can pick the one with most tokens)
            return (s, e, txt)

    return None


def parse_yuan_value(v) -> Optional[float]:
    """
    Accept:
      - 7.98
      - Y7.98 / Y 7.98
    Reject:
      - % values
      - junk text
      - 0 / 0.00
    """
    if v is None:
        return None
    s = str(v).replace("\xa0", " ").strip()
    if not s:
        return None
    if "%" in s:
        return None

    m = _YUAN_RE.match(s)
    if not m:
        return None
    try:
        val = float(m.group(1))
    except Exception:
        return None
    if abs(val) < 1e-12:
        return None
    return val


def extract_yuan_after_money_block(row: pd.Series, lookahead_cells: int = 20) -> Optional[float]:
    """
    Same logic: find money area, then scan next 15-20 cells for yuan.
    """
    found = find_money_block_chunk(row, min_tokens=3, max_scan_cols=60)
    if not found:
        return None

    _, end_idx, _ = found  # <-- anchor to the END of the dense chunk
    start = end_idx + 1
    end = min(len(row), start + int(lookahead_cells))

    for j in range(start, end):
        v = row.iloc[j]
        if v is None:
            continue
        s = str(v).replace("\xa0", " ").strip()
        if not s:
            continue
        y = parse_yuan_value(s)
        if y is not None:
            return y

    return None
    """
    FINAL RULE YOU CONFIRMED:
      - locate money-block cell (contains many 2dp numbers)
      - scan next 15-20 cells to the right
      - the first valid numeric cell there is yuan
      - if none -> None
    """
    found = find_money_block_chunk(row, min_tokens=3, max_scan_cols=60)
    if not found:
        return None

    _, end_idx, _ = found  # <-- anchor to the END of the dense chunk
    start = end_idx + 1
    end = min(len(row), start + int(lookahead_cells))

    for j in range(start, end):
        v = row.iloc[j]
        if v is None:
            continue
        s = str(v).replace("\xa0", " ").strip()
        if not s:
            continue
        y = parse_yuan_value(s)
        if y is not None:
            return y

    return None



_REPORT_VALUES_RE = re.compile(
    r"(?<!\S)-?\d+(?:,\d{3})*\.\d{2}(?:\s*-?\d+(?:,\d{3})*\.\d{2}){2,}"
)


def strip_report_values(text: str) -> str:
    """Keep product wording and remove the trailing Express amount block."""
    value = str(text or "").replace("\xa0", " ")
    match = _REPORT_VALUES_RE.search(value)
    if match:
        value = value[:match.start()]
    return re.sub(r"\s+", " ", value).strip()


# =========================
# PARSE ONE LINE -> FIELDS
# =========================
def parse_line_to_fields(row: pd.Series, merged_line: str) -> Optional[Dict[str, object]]:
    """
    Parse an Express row into:
      buyer, barcode(optional), สินค้า(product_str), ยอดขาย, สินค้าคงเหลือ, ON_ORDER, หยวน
    The 5/6-number block is extracted from row (2dp tokens).
    Yuan is extracted by money-block anchor + lookahead scan.
    """
    m = re.match(r"\s*([0-9A-Za-z]{5})\b(.*)", merged_line or "")
    if not m:
        return None

    buyer = m.group(1).strip().upper()
    rest = m.group(2).strip()
    if not rest:
        return None

    tokens = rest.split()
    if not tokens:
        return None

    barcode = ""
    idx = 0
    if idx < len(tokens) and re.fullmatch(r"\d+|\d{8,}\.", tokens[idx] or ""):
        barcode = normalize_barcode(tokens[idx])
        idx += 1

    product_str = strip_report_values(" ".join(tokens[idx:]))
    if not product_str:
        return None

    block = extract_5_6_block_from_money_chunk(row)

    # fallback (if no dense money cell found)
    if not block:
        nums = extract_money_2dp_numbers_from_row(row, scan_cols=25)
        if len(nums) < 5:
            return None
        block = nums[-6:] if len(nums) >= 6 else nums[-5:]

    # your old mapping:
    # if 6 numbers: sale=block[2], stock=block[3], on_order=block[5]
    # if 5 numbers: sale=0, stock=block[2], on_order=block[4]
    if len(block) == 6:
        sale = float(block[2])
        stock = float(block[3])
        on_order = float(block[5])
    else:
        sale = 0.0
        stock = float(block[2])
        on_order = float(block[4])

    yuan_val = extract_yuan_after_money_block(row, lookahead_cells=20)

    return {
        "buyer": buyer,
        "barcode": barcode,
        "สินค้า": product_str,
        "ยอดขาย": sale,
        "สินค้าคงเหลือ": stock,
        "ON_ORDER": on_order,
        "หยวน": yuan_val,
    }


def parse_express_file(path: str, source_label: str) -> Tuple[pd.DataFrame, Dict[str, object]]:
    """
    Parse Express export into DataFrame:
      buyer, barcode, รหัสสินค้า, รายละเอียดสินค้า, ยอดขาย, สินค้าคงเหลือ, ON_ORDER, หยวน
    """
    df_raw = pd.read_excel(path, sheet_name=EXPRESS_SHEET, header=None, dtype=str)
    date_info = parse_date_range_from_header(df_raw)

    rows = []
    for idx, row in df_raw.iterrows():
        merged = row_to_merged_line(row)
        if not merged:
            continue
        if is_header_or_separator(merged):
            continue

        first_token = merged.split(maxsplit=1)[0] if merged.split() else ""
        if not looks_like_buyer_code(first_token):
            continue

        fields = parse_line_to_fields(row, merged)
        if fields is None:
            continue

        fields["source"] = source_label
        fields["src_row"] = int(idx) + 1
        fields["src_file"] = os.path.basename(path)
        rows.append(fields)

    df = pd.DataFrame(rows)
    if not df.empty:
        df["buyer"] = df["buyer"].astype(str).str.replace("\u00A0", " ", regex=False).str.strip().str.upper()
        df[["รหัสสินค้า", "รายละเอียดสินค้า"]] = df["สินค้า"].apply(lambda x: pd.Series(split_product_field(x)))
        df.drop(columns=["สินค้า"], inplace=True)

    return df, date_info


# =========================
# COMBINE + AGG
# =========================
def normalize_item_code(value) -> str:
    """Treat punctuation variants as one code; retain parenthetical variants."""
    if value is None or pd.isna(value):
        return ""
    text = unicodedata.normalize("NFKC", str(value)).translate(_DASH_TRANS)
    return re.sub(r"[\s-]+", "", text).casefold()


def normalize_product_description(value) -> str:
    """Remove Express source tags without changing product variant wording."""
    if value is None or pd.isna(value):
        return ""
    text = unicodedata.normalize("NFKC", str(value)).translate(_DASH_TRANS)
    text = re.sub(r"\s+", " ", text).strip()
    text = re.sub(
        r"(?:[\s/,-]+(?:IR|MR|FN|VN|OEM))+(?:[\s/,-]*)$",
        "", text, flags=re.IGNORECASE,
    )
    text = re.sub(r"\s*-\s*", "-", text).strip().rstrip("/").strip()
    return text.casefold()


_COLOR_NAMES = {
    "น้ำเงิน": "blue", "น้ำตาล": "brown", "เขียว": "green", "แดง": "red",
    "ขาว": "white", "ดำ": "black", "เทา": "gray", "ฟ้า": "lightblue",
    "ชมพู": "pink", "เหลือง": "yellow", "ส้ม": "orange", "ม่วง": "purple",
    "เงิน": "silver", "ทอง": "gold", "งา": "ivory",
    "ครีมงาช้าง": "ivory", "งาช้าง": "ivory", "ครีม": "ivory",
    "light blue": "lightblue", "dark blue": "darkblue", "lightblue": "lightblue",
    "red": "red", "blue": "blue", "green": "green", "white": "white",
    "black": "black", "grey": "gray", "gray": "gray", "brown": "brown",
    "pink": "pink", "yellow": "yellow", "orange": "orange", "purple": "purple",
    "silver": "silver", "gold": "gold", "ivory": "ivory", "cream": "ivory",
}
# Match dictionaries use the same normalization as descriptions. Thai sara am
# (ำ) decomposes under NFKC, including in ดำ, น้ำเงิน and น้ำตาล.
for _name, _canonical in list(_COLOR_NAMES.items()):
    if not _name.isascii():
        for _shade, _suffix in (("เข้ม", "dark"), ("อ่อน", "light"), ("พาสเทล", "pastel")):
            _COLOR_NAMES[_name + _shade] = _canonical + "-" + _suffix
_COLOR_NAMES = {unicodedata.normalize("NFKC", name): color for name, color in _COLOR_NAMES.items()}
_COLOR_PATTERN = re.compile(
    "|".join(
        (r"(?<![a-z])" + re.escape(name) + r"(?![a-z])") if name.isascii()
        else (r"(?<![ก-๙])งา(?![ก-๙])|(?<=สี)งา" if name == "งา" else re.escape(name))
        for name in sorted(_COLOR_NAMES, key=len, reverse=True)
    )
)
_STYLE_NAMES = {
    "ด้ามสั้น": "short-handle", "ด้ามยาว": "long-handle",
    "คอสั้น": "short-neck", "คอยาว": "long-neck",
    "รุ่นใหญ่": "large", "รุ่นเล็ก": "small",
    "ขนาดใหญ่": "large", "ขนาดเล็ก": "small",
    "วงรี": "oval", "กลม": "round", "เหลี่ยม": "square",
    "รุ่นมาตรฐาน": "standard", "standard": "standard",
    "ดีลัก": "deluxe", "deluxe": "deluxe", "block": "block",
    "ซาติน": "satin", "ซาต": "satin", "satin": "satin",
    "เงา": "gloss", "glossy": "gloss", "ด้าน": "matte", "matte": "matte",
}
_STYLE_NAMES = {unicodedata.normalize("NFKC", name): style for name, style in _STYLE_NAMES.items()}
_NUMBER = r"(?:\d+\s*[- ]\s*\d+\s*/\s*\d+|\d+\s*/\s*\d+|\d+(?:\.\d+)?)"
_SIZE_PATTERN = re.compile(
    r"#?\s*" + _NUMBER + r"(?:\s*[x×*]\s*" + _NUMBER + r"){0,2}"
    r"(?:\s*(?:mm\.?|cm\.?|inch(?:es)?|in\b|มม\.?|ซม\.?|นิ้ว|เมตร|ฟุต|[\"″]))?"
)


def product_variant_signature(description, item_code="") -> tuple:
    """Code is the legacy identity; retain explicit colors, sizes and styles.

    Trailing name/report wording does not create another item. Variant details
    may occur on either side of a space, so truncating at the first space would
    erase real colors and dimensions. The original display text is untouched.
    """
    text = normalize_product_description(description)
    parts = text.split(maxsplit=1)
    if parts and normalize_item_code(parts[0]) == normalize_item_code(item_code):
        text = parts[1] if len(parts) > 1 else ""
    colors = {_COLOR_NAMES[match.group()] for match in _COLOR_PATTERN.finditer(text)}
    styles = {canonical for word, canonical in _STYLE_NAMES.items() if word in text}
    # Some Express names end halfway through a color. Infer a color only when
    # its prefix identifies one choice; retain uncertain prefixes as variants.
    color_tokens = re.findall(r"สี\s*-?\s*([ก-๙]+)|(?<![ก-๙])(น้ํา[ก-๙]+)", text)
    for explicit, standalone in color_tokens:
        token = explicit or standalone
        if any(token.startswith(name) for name in _COLOR_NAMES):
            continue
        if any(token.startswith(word) for word in _STYLE_NAMES):
            continue
        simple = re.sub(r"[\u0e31\u0e34-\u0e3a\u0e47-\u0e4e]", "", token)
        options = {
            color for name, color in _COLOR_NAMES.items()
            if not name.isascii() and re.sub(
                r"[\u0e31\u0e34-\u0e3a\u0e47-\u0e4e]", "", name
            ).startswith(simple)
        }
        # Shade suffixes do not turn a unique base color into many options.
        options = {color.split("-")[0] for color in options}
        colors.add(next(iter(options)) if len(options) == 1 else "partial:" + token)
    for model in re.findall(r"(?:รุ่น|model)\s*([a-z][a-z0-9-]*)", text):
        styles.add("model:" + model.rstrip("-"))
    # A whole product and a separately sold component are real variants,
    # rather than incomplete descriptions of each other.
    if unicodedata.normalize("NFKC", "ฝักบัวชำระ") in text:
        head_only = "เฉพาะหัว" in text or "หัวอย่างเดียว" in text
        styles.add("head-only" if head_only else "complete")
    if "อ่างล้างหน้า" in text:
        styles.add("with-legs" if "พร้อมขา" in text else "without-legs")
    color_codes = {"r": "red", "g": "green", "b": "blue", "lb": "lightblue"}
    for parenthetical in re.findall(r"\(([^()]*)\)", text):
        if parenthetical in color_codes:
            # LB identifies the light-blue option even when its Thai name says
            # น้ำเงิน rather than ฟ้า. Keep other conflicting color details.
            if parenthetical == "lb":
                colors.discard("blue")
            colors.add(color_codes[parenthetical])
        elif not _COLOR_PATTERN.search(parenthetical) and parenthetical:
            styles.add(parenthetical)
    sizes = set()
    for match in _SIZE_PATTERN.finditer(text):
        size = match.group().strip().replace("×", "x").replace("*", "x")
        size = re.sub(r"(?<=\d)[ -]+(?=\d+\s*/)", "+", size)
        size = re.sub(r"\s+", "", size)
        size = re.sub(r'(?:inch(?:es)?|นิ้ว|["″])$', "in", size)
        size = re.sub(r"(?:มม|mm)\.?$", "mm", size)
        size = re.sub(r"(?:ซม|cm)\.?$", "cm", size)
        size = re.sub(r"\d+\.\d+", lambda number: (
            str(int(number.group().split(".")[0])) + "." + number.group().split(".")[1].rstrip("0")
        ).rstrip("."), size)
        sizes.add(size)
    return tuple(sorted(colors)), tuple(sorted(sizes)), tuple(sorted(styles))


def variants_compatible(left: tuple, right: tuple) -> bool:
    """A missing detail may match one variant; explicit differences may not."""
    return all(not a or not b or a == b for a, b in zip(left, right))


def catalog_variant_counts(df: pd.DataFrame, barcode_col: str) -> tuple[dict, dict, dict]:
    """Count every supplier variant under normalized code and description keys."""
    by_code = {}
    by_code_description = {}
    barcode_descriptions = {}
    for _, row in df.iterrows():
        code = normalize_item_code(row.get("รหัสสินค้า"))
        description = normalize_product_description(row.get("รายละเอียดสินค้า"))
        barcode = normalize_barcode(row.get(barcode_col))
        by_code[code] = by_code.get(code, 0) + 1
        code_description = (code, description)
        by_code_description[code_description] = by_code_description.get(code_description, 0) + 1
        barcode_descriptions.setdefault((code, barcode), set()).add(description)
    return (
        by_code,
        by_code_description,
        {key: len(values) for key, values in barcode_descriptions.items()},
    )


def normalize_variant_count_keys(counts: dict, key_kind: str) -> dict:
    """Accept caller-provided counts using either raw or normalized item codes."""
    normalized = {}
    for key, count in counts.items():
        if key_kind == "code":
            new_key = normalize_item_code(key)
        elif key_kind == "description":
            new_key = (normalize_item_code(key[0]), normalize_product_description(key[1]))
        else:
            new_key = (normalize_item_code(key[0]), normalize_barcode(key[1]))
        normalized[new_key] = normalized.get(new_key, 0) + int(count)
    return normalized


def _agg_one(df: pd.DataFrame, label: str, variant_context: Optional[dict] = None) -> pd.DataFrame:
    """Sum barcode variants, with legacy code/variant fallback when blank."""
    columns = [
        "buyer", "รหัสสินค้า", "_norm_code", "barcode", "_norm_desc", "_variant_key",
        f"รายละเอียดสินค้า_{label}", f"ยอดขาย_{label}", f"STOCK_{label}",
        f"ON_ORDER_{label}", f"หยวน_{label}",
    ]
    if df.empty:
        return pd.DataFrame(columns=columns)

    source = df.copy()
    for col in ("buyer", "รหัสสินค้า", "รายละเอียดสินค้า"):
        source[col] = source[col].fillna("").astype(str).str.strip()
    source["buyer"] = source["buyer"].str.upper()
    source["barcode"] = source["barcode"].map(normalize_barcode)
    source["_barcode_originally_present"] = source["barcode"].ne("")
    source["_norm_code"] = source["รหัสสินค้า"].map(normalize_item_code)
    source["_norm_desc"] = source["รายละเอียดสินค้า"].map(normalize_product_description)

    source["_variant_key"] = [
        product_variant_signature(description, code)
        for description, code in zip(source["รายละเอียดสินค้า"], source["รหัสสินค้า"])
    ]
    # Without an item code there is no legacy code identity. Preserve the full
    # normalized name so unrelated uncoded products cannot collapse together.
    for index in source.index[source["_norm_code"].eq("")]:
        signature = source.at[index, "_variant_key"]
        name = source.at[index, "_norm_desc"] or f"unknown-row-{index}"
        source.at[index, "_variant_key"] = signature[:2] + (
            signature[2] + ("description:" + name,),
        )
    variants_by_code = variant_context if variant_context is not None else source.groupby(
        ["buyer", "_norm_code"], sort=False
    )["_variant_key"].agg(lambda values: set(values)).to_dict()
    # Complete truncated descriptions only when one explicit variant fits.
    # A plain/unknown row stays separate when several colors or sizes fit.
    for index in source.index[source["barcode"].eq("")]:
        key = (source.at[index, "buyer"], source.at[index, "_norm_code"])
        variant = source.at[index, "_variant_key"]
        if not key[1]:
            continue
        options = [
            other for other in variants_by_code.get(key, {variant})
            if all(not a or a == b for a, b in zip(variant, other))
        ]
        fullest = [
            other for other in options
            if not any(
                other != candidate and all(not a or a == b for a, b in zip(other, candidate))
                for candidate in options
            )
        ]
        if len(fullest) == 1:
            source.at[index, "_variant_key"] = fullest[0]

    known_barcodes = source[source["barcode"].ne("")].groupby(
        ["buyer", "_norm_code", "_variant_key"], sort=False, dropna=False
    )["barcode"].agg(lambda values: set(values)).to_dict()
    for index in source.index[source["barcode"].eq("")]:
        key = (source.at[index, "buyer"], source.at[index, "_norm_code"],
               source.at[index, "_variant_key"])
        options = known_barcodes.get(key, set())
        if len(options) == 1:
            source.at[index, "barcode"] = next(iter(options))

    source["_group_desc"] = [
        variant if not barcode else ()
        for variant, barcode in zip(source["_variant_key"], source["barcode"])
    ]
    source["หยวน"] = pd.to_numeric(source["หยวน"], errors="coerce")
    for col in ("ยอดขาย", "สินค้าคงเหลือ", "ON_ORDER"):
        source[col] = pd.to_numeric(source[col], errors="coerce").fillna(0.0)
    source["_active"] = source[["ยอดขาย", "สินค้าคงเหลือ", "ON_ORDER"]].ne(0).any(axis=1)
    # Active wording wins; among equally active rows, use the one that supplied
    # the barcode before a blank row later attached to it.
    source = source.sort_values(
        ["_active", "_barcode_originally_present"],
        ascending=[False, False],
        kind="stable",
    )

    keys = ["buyer", "_norm_code", "barcode", "_group_desc"]
    source["_selected_price"] = np.nan
    for group_key, group in source.groupby(keys, sort=False, dropna=False):
        active_prices = group.loc[group["_active"], "หยวน"].dropna().unique()
        all_prices = group["หยวน"].dropna().unique()
        # Missing-barcode groups retain the legacy first available price.
        # Populated barcodes still identify one variant with one active price.
        if len(active_prices) > 1 and group_key[2]:
            buyer, _, barcode, _ = group_key
            code = group["รหัสสินค้า"].iloc[0]
            raise ValueError(
                f"Conflicting {label} prices for supplier {buyer}, item {code}, "
                f"barcode {barcode or '(blank)'}. Correct the source prices before generating the PO."
            )
        if len(active_prices) >= 1:
            selected_price = active_prices[0]
        elif len(all_prices) == 1 or (len(all_prices) and not group_key[2]):
            selected_price = all_prices[0]
        else:
            selected_price = np.nan
        source.loc[group.index, "_selected_price"] = selected_price

    grouped = source.groupby(keys, as_index=False, dropna=False, sort=False).agg({
        "รหัสสินค้า": "first",
        "_norm_desc": "first",
        "_variant_key": "first",
        "รายละเอียดสินค้า": "first",
        "ยอดขาย": "sum",
        "สินค้าคงเหลือ": "sum",
        "ON_ORDER": "sum",
        "_selected_price": "first",
    })
    grouped.rename(columns={
        "รายละเอียดสินค้า": f"รายละเอียดสินค้า_{label}",
        "ยอดขาย": f"ยอดขาย_{label}",
        "สินค้าคงเหลือ": f"STOCK_{label}",
        "ON_ORDER": f"ON_ORDER_{label}",
        "_selected_price": f"หยวน_{label}",
    }, inplace=True)
    return grouped[columns]


def build_combined_all(
    df_asia: pd.DataFrame,
    df_green: pd.DataFrame,
    months: int,
    min_factor: int,
    max_factor: int,
) -> pd.DataFrame:
    """Match barcodes first, then one-to-one code variants; GREEN takes priority."""
    # Use both files when deciding whether missing variant details are unique.
    # Otherwise a colorless ASIA row could be assigned red before GREEN's white
    # option is seen.
    variant_context = {}
    for frame in (df_asia, df_green):
        for _, row in frame.iterrows():
            code = normalize_item_code(row.get("รหัสสินค้า"))
            if not code:
                continue
            buyer = "" if pd.isna(row["buyer"]) else str(row["buyer"]).strip().upper()
            key = (buyer, code)
            variant_context.setdefault(key, set()).add(product_variant_signature(
                row.get("รายละเอียดสินค้า"), row.get("รหัสสินค้า")
            ))
    # Uncoded rows already use full names and do not infer variant details.
    asia = _agg_one(df_asia, "ASIA", variant_context).to_dict("records")
    green = _agg_one(df_green, "GREEN", variant_context).to_dict("records")
    matches = {}
    used_green = set()

    # A matching barcode on the same normalized code is the strongest identity.
    for ai, a in enumerate(asia):
        if not a["barcode"]:
            continue
        candidates = [
            gi for gi, g in enumerate(green)
            if gi not in used_green
            and a["buyer"] == g["buyer"]
            and a["_norm_code"] == g["_norm_code"]
            and a["barcode"] == g["barcode"]
        ]
        if len(candidates) == 1:
            matches[ai] = candidates[0]
            used_green.add(candidates[0])

    # If Express codes differ, a vendor-wide barcode can still identify the
    # same variant when it appears exactly once in each source.
    for ai, a in enumerate(asia):
        if ai in matches or not a["barcode"]:
            continue
        same_asia_barcode = [
            item for item in asia if item["buyer"] == a["buyer"] and item["barcode"] == a["barcode"]
        ]
        candidates = [
            gi for gi, g in enumerate(green)
            if gi not in used_green
            and g["buyer"] == a["buyer"]
            and g["barcode"] == a["barcode"]
        ]
        same_green_barcode = [
            item for item in green if item["buyer"] == a["buyer"] and item["barcode"] == a["barcode"]
        ]
        if len(same_asia_barcode) == len(same_green_barcode) == len(candidates) == 1:
            matches[ai] = candidates[0]
            used_green.add(candidates[0])

    # Exact variant matches precede partial/truncated matches. The same product
    # may have different barcodes in the two reports; retain one-to-one variant
    # matching and record the mismatch instead of producing duplicate PO lines.
    for exact_only in (True, False):
        for ai, a in enumerate(asia):
            if ai in matches:
                continue

            def eligible(other, candidate):
                return (
                    other["buyer"] == candidate["buyer"]
                    and other["_norm_code"] == candidate["_norm_code"]
                    and (other["_variant_key"] == candidate["_variant_key"] if exact_only
                         else variants_compatible(other["_variant_key"], candidate["_variant_key"]))
                )

            # Count all original choices, including already matched rows, so
            # processing order cannot make an ambiguous missing color unique.
            candidates = [gi for gi, g in enumerate(green) if eligible(a, g)]
            if len(candidates) != 1 or candidates[0] in used_green:
                continue
            gi = candidates[0]
            reverse = [other_ai for other_ai, other in enumerate(asia) if eligible(other, green[gi])]
            if len(reverse) == 1:
                matches[ai] = gi
                used_green.add(gi)

    records = []
    pairs = [(a, green[matches[ai]] if ai in matches else None) for ai, a in enumerate(asia)]
    pairs.extend((None, g) for gi, g in enumerate(green) if gi not in used_green)
    for a, g in pairs:
        chosen = g if g is not None else a
        asia_barcode = a["barcode"] if a is not None else ""
        green_barcode = g["barcode"] if g is not None else ""
        source_barcode = green_barcode or asia_barcode
        barcode_mismatch = bool(asia_barcode and green_barcode and asia_barcode != green_barcode)
        green_price = g.get("หยวน_GREEN", np.nan) if g is not None else np.nan
        asia_price = a.get("หยวน_ASIA", np.nan) if a is not None else np.nan
        record = {
            "buyer": chosen["buyer"],
            "รหัสสินค้า": chosen["รหัสสินค้า"],
            "รายละเอียดสินค้า": (
                g["รายละเอียดสินค้า_GREEN"] if g is not None
                else a["รายละเอียดสินค้า_ASIA"]
            ),
            "barcode": source_barcode,
            "barcode_ASIA": asia_barcode,
            "barcode_GREEN": green_barcode,
            "barcode_mismatch": barcode_mismatch,
            "หมายเหตุ": BARCODE_MISMATCH_NOTE if barcode_mismatch else "",
            "catalog_match_barcode": source_barcode,
            "หยวน_ASIA": asia_price,
            "หยวน_GREEN": green_price,
            "หยวน": green_price if pd.notna(green_price) else asia_price,
        }
        for label, item in (("ASIA", a), ("GREEN", g)):
            for col in (f"ยอดขาย_{label}", f"STOCK_{label}", f"ON_ORDER_{label}"):
                record[col] = item[col] if item is not None else 0.0
        records.append(record)

    combined = pd.DataFrame(records, columns=[
        "buyer", "รหัสสินค้า", "รายละเอียดสินค้า", "barcode", "barcode_ASIA", "barcode_GREEN",
        "barcode_mismatch", "หมายเหตุ", "catalog_match_barcode",
        "ยอดขาย_ASIA", "STOCK_ASIA", "ON_ORDER_ASIA", "หยวน_ASIA",
        "ยอดขาย_GREEN", "STOCK_GREEN", "ON_ORDER_GREEN", "หยวน_GREEN", "หยวน",
    ])
    for col in ("ยอดขาย_ASIA", "STOCK_ASIA", "ON_ORDER_ASIA",
                "ยอดขาย_GREEN", "STOCK_GREEN", "ON_ORDER_GREEN"):
        combined[col] = pd.to_numeric(combined[col], errors="coerce").fillna(0.0)

    combined["ยอดขาย_TOTAL"] = combined["ยอดขาย_ASIA"] + combined["ยอดขาย_GREEN"]
    combined["ON_ORDER_TOTAL"] = combined["ON_ORDER_ASIA"] + combined["ON_ORDER_GREEN"]
    combined["USE_MONTH"] = combined["ยอดขาย_TOTAL"].apply(
        lambda value: round_half_up(value / max(months, 1)) if value > 0 else 0
    )
    combined["TOTAL_QTY_NUM"] = (
        combined["STOCK_ASIA"] + combined["STOCK_GREEN"] + combined["ON_ORDER_TOTAL"]
    )
    combined["MIN_NUM"] = combined["USE_MONTH"] * int(min_factor)
    combined["MAX_NUM"] = combined["USE_MONTH"] * int(max_factor)
    return combined


# =========================
# VENDOR INFO
# =========================
def load_vendor_map(path: str) -> dict:
    """Read vendor info file: column 0=code, 1=name, 2=address."""
    if not os.path.exists(path):
        return {}
    df = pd.read_excel(path, header=None)
    out = {}
    for _, r in df.iterrows():
        code = str(r.iloc[0]).strip() if not pd.isna(r.iloc[0]) else ""
        if not code:
            continue
        out[str(code).strip().upper()] = {
            "name": "" if pd.isna(r.iloc[1]) else str(r.iloc[1]).strip(),
            "address": "" if pd.isna(r.iloc[2]) else str(r.iloc[2]).strip(),
        }
    return out


# =========================
# CATALOG (multi-sheet per vendor)
# =========================
def normalize_barcode(value, number_format: str = "") -> str:
    """Read an identifier as text, including simple Excel zero-padded cells."""
    if value is None or isinstance(value, bool):
        return ""
    if isinstance(value, (int, float)):
        if not math.isfinite(value):
            return ""
        barcode = str(int(value)) if float(value).is_integer() else str(value)
    else:
        barcode = str(value)
    barcode = re.sub(r"\s+", "", barcode)
    trailing_period = re.fullmatch(r"(\d{8,})\.", barcode)
    if trailing_period:
        barcode = trailing_period.group(1)
    fmt = str(number_format or "").strip()
    if barcode.isdigit() and re.fullmatch(r"0+", fmt):
        barcode = barcode.zfill(len(fmt))
    return barcode


def source_barcode_mismatch(row) -> bool:
    """Warn only when both source identifiers exist and differ."""
    asia = normalize_barcode(row.get("barcode_ASIA"))
    green = normalize_barcode(row.get("barcode_GREEN"))
    return bool(asia and green and asia != green)


def merge_notes(*values) -> str:
    """Preserve each warning on its own line without repeating it."""
    lines = []
    for value in values:
        if value is None or pd.isna(value):
            continue
        for line in str(value).splitlines():
            line = line.strip()
            if line and line not in lines:
                lines.append(line)
    return "\n".join(lines)


class CatalogMap(dict):
    """Catalog rows by item code, with a vendor-sheet-wide barcode index."""

    def __init__(self):
        super().__init__()
        self.by_barcode = {}
        self.by_norm_code = {}


def build_catalog_map(catalog_path: str, vendor_code: str) -> dict:
    """
    Read catalog workbook where each vendor has its own sheet.
    Column mapping by Excel position:
      A=item no, B=picture, C=desc, D=brand, E=material, F=weight,
      G=qty/carton, H=unit price (unused), I=barcode
    """
    wb = openpyxl.load_workbook(catalog_path)
    want = str(vendor_code).strip().upper()
    norm_map = {str(n).strip().upper(): n for n in wb.sheetnames}

    if want in norm_map:
        ws = wb[norm_map[want]]
    else:
        ws = None
        for k_norm, original in norm_map.items():
            if want in k_norm or k_norm in want:
                ws = wb[original]
                break
        if ws is None:
            wb.close()
            return CatalogMap()

    COL_ITEM_NO = 1
    COL_PIC = 2
    COL_DESC = 3
    COL_BRAND = 4
    COL_MAT = 5
    COL_WEIGHT = 6
    COL_QTYCT = 7
    COL_BARCODE = 9
    HEADER_ROW_LOCAL = 1

    img_at = {}
    for img in ws._images:
        try:
            r = img.anchor._from.row + 1
            c = img.anchor._from.col + 1
            img_at[(r, c)] = img._data()
        except Exception:
            pass

    catalog = CatalogMap()
    for r in range(HEADER_ROW_LOCAL + 1, ws.max_row + 1):
        item_no = ws.cell(r, COL_ITEM_NO).value
        item_no = str(item_no).strip() if item_no is not None else ""
        barcode = normalize_barcode(
            ws.cell(r, COL_BARCODE).value,
            ws.cell(r, COL_BARCODE).number_format,
        )
        if not item_no and not barcode:
            continue
        entry = {
            "goods_desc": ws.cell(r, COL_DESC).value,
            "brand": ws.cell(r, COL_BRAND).value,
            "material": ws.cell(r, COL_MAT).value,
            "weight": ws.cell(r, COL_WEIGHT).value,
            "qty_per_carton": ws.cell(r, COL_QTYCT).value,
            "barcode": barcode,
            "item_code": item_no,
            "img_bytes": img_at.get((r, COL_PIC)),
        }
        if item_no:
            catalog.setdefault(item_no, []).append(entry)
            catalog.by_norm_code.setdefault(normalize_item_code(item_no), []).append(entry)
        if barcode:
            catalog.by_barcode.setdefault(barcode, []).append(entry)
    return catalog


def resolve_catalog_variant(
    catalog_map: dict,
    item_code: str,
    description: str,
    variant_count: int,
    description_variant_count: int = 1,
    barcode: str = "",
    barcode_description_count: int = 1,
) -> dict:
    """Prefer barcodes; missing barcodes use the March item-code fallback.

    A description/variant match improves legacy selection. If none is available,
    the last eligible worksheet row wins, as in the pre-barcode catalog map.
    Source variant counts no longer block a code fallback.
    """
    normalized_code = normalize_item_code(item_code)
    normalized_index = getattr(catalog_map, "by_norm_code", None)
    if normalized_index is not None:
        entries = normalized_index.get(normalized_code, [])
    else:
        entries = [
            entry for code, code_entries in catalog_map.items()
            if normalize_item_code(code) == normalized_code
            for entry in code_entries
        ]
    source_barcode = normalize_barcode(barcode)
    barcode_index = getattr(catalog_map, "by_barcode", None)
    barcode_matches = (
        barcode_index.get(source_barcode, []) if barcode_index is not None else [
            entry for code_entries in catalog_map.values()
            for entry in code_entries
            if source_barcode and normalize_barcode(entry.get("barcode")) == source_barcode
        ]
    ) if source_barcode else []
    if len(barcode_matches) == 1:
        return barcode_matches[0].copy()
    if not barcode_matches:
        # A blank catalog barcode is eligible for the old lookup even during
        # a partial update. Two different populated barcodes remain distinct.
        eligible = (
            [entry for entry in entries if not normalize_barcode(entry.get("barcode"))]
            if source_barcode else entries
        )
        if not eligible:
            if source_barcode and entries:
                raise ValueError(
                    f"No catalog BARCODE match for item {item_code}, barcode {source_barcode}. "
                    "QTY PER CARTON cannot be verified from another variant. "
                    "Add this barcode in catalog column I or leave the matching item-code "
                    "row's barcode blank to use the legacy lookup."
                )
            return {}
        normalized_desc = normalize_product_description(description)
        exact = [
            entry for entry in eligible if normalized_desc and
            normalize_product_description(entry.get("goods_desc")) == normalized_desc
        ]
        if exact:
            return exact[-1].copy()
        signature = product_variant_signature(description, item_code)
        variants = [
            entry for entry in eligible if any(signature) and
            product_variant_signature(entry.get("goods_desc"), item_code) == signature
        ]
        return (variants or eligible)[-1].copy()

    # Duplicate populated barcodes keep only fields common to all matching rows.
    candidates = barcode_matches
    resolved = {}
    for field in ("brand", "material", "weight", "qty_per_carton"):
        values = [entry.get(field) for entry in candidates]
        normalized_values = []
        for value in values:
            if value is None or str(value).strip() == "":
                normalized_values.append(("blank", ""))
            elif field == "qty_per_carton":
                try:
                    normalized_values.append(("number", float(value)))
                except (TypeError, ValueError):
                    normalized_values.append(("text", str(value).strip()))
            else:
                normalized_values.append(("text", re.sub(r"\s+", " ", str(value)).strip().casefold()))
        if len(set(normalized_values)) == 1:
            resolved[field] = "" if normalized_values[0][0] == "blank" else values[0]
        elif field == "qty_per_carton":
            detail = f", barcode {source_barcode}" if source_barcode else ""
            raise ValueError(
                f"Ambiguous QTY PER CARTON for item {item_code}{detail} ({description}). "
                "Add a unique BARCODE in catalog column I or correct the carton quantities."
            )
        else:
            resolved[field] = ""

    resolved["img_bytes"] = None
    return resolved


# =========================
# EXCEL IMAGE + STYLE HELPERS
# =========================
def _add_png(ws, png_path: str, anchor_cell: str, width_px: Optional[int] = None):
    """
    Add an image (png/jpg) into a sheet at anchor_cell, optionally resizing by width.
    Keeps aspect ratio. Safe: no undefined variables, no special buffering needed for file-path images.
    """
    if not png_path:
        return
    if not os.path.exists(png_path):
        return

    img = XLImage(png_path)

    # Resize by width (keep aspect ratio)
    if width_px is not None:
        try:
            w0 = float(img.width or 0)
            h0 = float(img.height or 0)
            if w0 > 0 and h0 > 0:
                scale = float(width_px) / w0
                img.width = int(round(w0 * scale))
                img.height = int(round(h0 * scale))
        except Exception:
            # If anything weird happens, just keep original size
            pass

    img.anchor = anchor_cell
    ws.add_image(img)


def add_logo_and_footer(ws, base_dir: str, footer_row: int,
                        logo_cell: str = "A1",
                        logo_width_px: int = 260,
                        footer_width_px: int = 650):
    """
    Add logo at fixed top location + footer below totals.
    Expects:
      base_dir/logo.png
      base_dir/footer_signatures.png
    """
    logo_path = os.path.join(base_dir, "logo.png")
    footer_path = os.path.join(base_dir, "footer_signatures.png")

    _add_png(ws, logo_path, anchor_cell=logo_cell, width_px=logo_width_px)
    _add_png(ws, footer_path, anchor_cell=f"A{int(footer_row)}", width_px=footer_width_px)

def _excel_colwidth_to_pixels(width):
    """Convert Excel col width to pixels."""
    if width is None:
        width = 8.43
    return int(width * 7 + 5)


def _excel_rowheight_to_pixels(height_pts):
    """Convert Excel row height to pixels."""
    if height_pts is None:
        height_pts = 15
    return int(height_pts * 96 / 72)


def _get_cell_rect_pixels(ws, col_letter, row_num):
    """Return (w_px, h_px) including merged cell extents."""
    col_w = _excel_colwidth_to_pixels(ws.column_dimensions[col_letter].width)
    row_h = _excel_rowheight_to_pixels(ws.row_dimensions[row_num].height)

    for mr in ws.merged_cells.ranges:
        if mr.min_col <= column_index_from_string(col_letter) <= mr.max_col and mr.min_row <= row_num <= mr.max_row:
            total_w = 0
            for c in range(mr.min_col, mr.max_col + 1):
                letter = get_column_letter(c)
                total_w += _excel_colwidth_to_pixels(ws.column_dimensions[letter].width)
            total_h = 0
            for rr in range(mr.min_row, mr.max_row + 1):
                total_h += _excel_rowheight_to_pixels(ws.row_dimensions[rr].height)
            return total_w, total_h

    return col_w, row_h


def add_image_to_cell(ws, cell_addr: str, img_bytes: bytes,
                      width_boost: float = IMAGE_WIDTH_BOOST,
                      padding_px: int = IMAGE_PADDING_PX):
    """Place an image into an Excel cell with center alignment."""
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


def norm_header(x) -> str:
    """Normalize header text."""
    return str(x).replace("\n", " ").replace("\xa0", " ").strip() if x is not None else ""


def get_po_col_map(ws, header_row=HEADER_ROW):
    """Create header->col index mapping from template."""
    col_map = {}
    for c in range(1, ws.max_column + 1):
        v = norm_header(ws.cell(header_row, c).value)
        if v:
            col_map[v] = c
    return col_map


def copy_column_widths(src_ws, dst_ws):
    """Copy column widths."""
    for col_letter, dim in src_ws.column_dimensions.items():
        if dim.width is not None:
            dst_ws.column_dimensions[col_letter].width = dim.width


def copy_row_style(ws, src_row: int, dst_row: int, max_col: int):
    """Copy style from src_row to dst_row."""
    for c in range(1, max_col + 1):
        src = ws.cell(src_row, c)
        dst = ws.cell(dst_row, c)
        if src.has_style:
            dst._style = src._style


def copy_row_height(ws, src_row: int, dst_row: int):
    """Copy row height."""
    ws.row_dimensions[dst_row].height = ws.row_dimensions[src_row].height


def force_bottom_border(ws, row: int, start_col: int, end_col: int):
    """Force thin bottom border on a row range."""
    thin = Side(style="thin")
    for c in range(start_col, end_col + 1):
        cell = ws.cell(row, c)
        b = _copy(cell.border) if cell.border else Border()
        b.bottom = thin
        cell.border = b


def copy_template_rows(src_ws, dst_ws, src_start, src_end, dst_start, max_col):
    """Copy a row block including merges/styles."""
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
    """
    Paste total section and update SUM formulas to match new last_item_row.
    """
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
    """Find an exact label cell."""
    target = norm_text(label)
    for r in range(1, min(max_row, ws.max_row) + 1):
        for c in range(1, min(max_col, ws.max_column) + 1):
            if norm_text(ws.cell(r, c).value) == target:
                return r, c
    return None


def add_catalog_notes_column(ws):
    """Insert the notes column in the PO table, preserving the supplier header."""
    # The template table starts at row 8. Its formulas refer only to table
    # columns D onward, so moving that block translates every affected reference.
    last_col = max(get_po_col_map(ws, header_row=HEADER_ROW).values())
    last_row = ws.max_row
    ws.move_range(f"D{HEADER_ROW}:{get_column_letter(last_col)}{last_row}", cols=1, translate=True)
    # Expand grouped dimensions (for example I:J) before shifting widths.
    original_dimensions = list(ws.column_dimensions.items())
    table_dimensions = {}
    for col in range(4, last_col + 1):
        dimension = next((dim for key, dim in original_dimensions
                          if (dim.min or column_index_from_string(key)) <= col
                          <= (dim.max or column_index_from_string(key))), None)
        if dimension is None:
            dimension = ws.column_dimensions[get_column_letter(col)]
        table_dimensions[col] = _copy(dimension)
    for col in range(last_col, 3, -1):
        new_letter = get_column_letter(col + 1)
        dimension = table_dimensions[col]
        dimension.index = new_letter
        dimension.min = dimension.max = col + 1
        ws.column_dimensions[new_letter] = dimension
    ws.column_dimensions["D"].width = 48
    for row in range(HEADER_ROW, last_row + 1):
        ws.cell(row, 4)._style = _copy(ws.cell(row, 3)._style)
        ws.cell(row, 4).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ws.cell(HEADER_ROW, 4).value = "หมายเหตุ"


def catalog_missing_fields(catalog_entry):
    """Return catalog metadata that is unavailable; barcode is optional."""
    fields = (
        ("qty_per_carton", "QTY PER CARTON", "carton"),
        ("goods_desc", "GOODS DESCRIPTION", "รายละเอียดสินค้า"),
        ("img_bytes", "GOODS PICTURE", "รูปสินค้า"),
        ("brand", "BRAND", "ยี่ห้อ"),
        ("material", "MATERIAL", "วัสดุ"),
        ("weight", "Weight", "น้ำหนัก"),
    )
    missing = []
    for key, header, label in fields:
        value = catalog_entry.get(key)
        if key == "qty_per_carton":
            try:
                valid = not isinstance(value, bool) and math.isfinite(float(value)) and float(value) > 0
            except (TypeError, ValueError):
                valid = False
            if not valid:
                missing.append((header, label))
        elif value is None or (isinstance(value, str) and not value.strip()) or (
            isinstance(value, (int, float)) and not math.isfinite(value)
        ) or (key == "img_bytes" and not value):
            missing.append((header, label))
    return missing


# =========================
# PO GENERATION
# =========================
def generate_po_from_combined(
    combined_df: pd.DataFrame,
    vendor_code: str,
    po_date: Optional[datetime.date],
    rate_thb_per_cny: float,
    template_path: str,
    catalog_path: str,
    vendor_info_path: str,
    min_factor: int,
    max_factor: int,
    variant_counts_by_code: Optional[Dict[str, int]] = None,
    variant_counts_by_code_description: Optional[Dict[Tuple[str, str], int]] = None,
    barcode_description_counts: Optional[Dict[Tuple[str, str], int]] = None,
    catalog_filename: Optional[str] = None,
) -> str:

    if po_date is None:
        po_date = datetime.date.today()

    os.makedirs(PO_OUTPUT_FOLDER, exist_ok=True)

    vendor_key = str(vendor_code).strip().upper()
    output_path = os.path.join(PO_OUTPUT_FOLDER, f"PO_{vendor_key}_BELOW_MIN.xlsx")

    vendor_map = load_vendor_map(vendor_info_path)
    supplier_name = vendor_map.get(vendor_key, {}).get("name", "")
    supplier_addr = vendor_map.get(vendor_key, {}).get("address", "")

    catalog_name = catalog_filename or os.path.basename(catalog_path)
    catalog_map = {}
    if os.path.exists(catalog_path):
        catalog_map = build_catalog_map(catalog_path, vendor_code=vendor_key)

    wb = openpyxl.load_workbook(template_path)
    template_ws = wb[TEMPLATE_SHEET_NAME]
    ws = wb.copy_worksheet(template_ws)
    ws.title = "PO"

    # locate total block
    pos_total = find_label_cell(ws, "TOTAL AMOUNT CNY", max_row=200, max_col=60)
    if not pos_total:
        raise RuntimeError("Cannot find TOTAL AMOUNT CNY in template.")

    BASE_TOTAL_ROW = pos_total[0]

    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

    # add logo
    _add_png(ws, os.path.join(BASE_DIR, "logo.png"), anchor_cell="A1", width_px=260)

    copy_column_widths(template_ws, ws)

    add_catalog_notes_column(ws)
    # Keep the preferred available source barcode at the end after adding notes.
    ws["Y8"].value = "BARCODE"
    ws["Y8"]._style = _copy(ws["X8"]._style)
    ws["Y9"]._style = _copy(ws["X9"]._style)
    ws.column_dimensions["Y"].width = 20

    po_cols = get_po_col_map(ws, header_row=HEADER_ROW)

    def find_one(keys):
        for k in keys:
            if k in po_cols:
                return k
        return None

    min_key_old = find_one(["MIN*4", "MIN * 4", "MIN×4", "MIN x4", "MINX4"])
    max_key_old = find_one(["MAX*7", "MAX * 7", "MAX×7", "MAX x7", "MAXX7"])

    if not min_key_old or not max_key_old:
        raise RuntimeError("Cannot find MIN/MAX header in template.")

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

    # Header fields
    ws["H6"] = vendor_key
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

    match_barcode_col = (
        "catalog_match_barcode"
        if "catalog_match_barcode" in combined_df.columns else "barcode"
    )
    fallback_counts = catalog_variant_counts(combined_df, match_barcode_col)
    variant_counts_by_code = normalize_variant_count_keys(
        variant_counts_by_code if variant_counts_by_code is not None else fallback_counts[0],
        "code",
    )
    variant_counts_by_code_description = normalize_variant_count_keys(
        variant_counts_by_code_description
        if variant_counts_by_code_description is not None else fallback_counts[1],
        "description",
    )
    barcode_description_counts = normalize_variant_count_keys(
        barcode_description_counts if barcode_description_counts is not None else fallback_counts[2],
        "barcode",
    )

    combined_df = combined_df.sort_values(
        by=["รหัสสินค้า", "รายละเอียดสินค้า", "barcode"],
        ascending=[True, True, True]
    ).reset_index(drop=True)

    # ensure totals section is not overwritten
    needed_item_rows = len(combined_df)
    available_item_rows = BASE_TOTAL_ROW - ITEM_START_ROW

    extra = max(0, needed_item_rows - available_item_rows)

    if extra > 0:
        ws.insert_rows(BASE_TOTAL_ROW, amount=extra)
        BASE_TOTAL_ROW += extra

    current_row = ITEM_START_ROW
    # Snapshot before filling the first item so its warnings never leak to later rows.
    item_styles = [_copy(ws.cell(TEMPLATE_ITEM_ROW, col)._style) for col in range(1, PO_LAST_COL + 1)]
    item_height = ws.row_dimensions[TEMPLATE_ITEM_ROW].height or 15
    incomplete_cartons = False
    has_catalog_warnings = False

    for _, row in combined_df.iterrows():

        line = current_row
        current_row += 1

        for col, style in enumerate(item_styles, start=1):
            ws.cell(line, col)._style = _copy(style)
        ws.row_dimensions[line].height = item_height

        buyer_item = str(row["รหัสสินค้า"]).strip()
        source_desc = str(row.get("รายละเอียดสินค้า") or "").strip()
        po_barcode = normalize_barcode(row.get("barcode", ""))
        source_barcode = normalize_barcode(row.get(match_barcode_col, ""))
        count_code = normalize_item_code(buyer_item)
        count_description = normalize_product_description(source_desc)
        cat = resolve_catalog_variant(
            catalog_map,
            buyer_item,
            source_desc,
            variant_counts_by_code.get(count_code, 1),
            variant_counts_by_code_description.get((count_code, count_description), 1),
            barcode=source_barcode,
            barcode_description_count=barcode_description_counts.get(
                (count_code, source_barcode), 1
            ),
        )

        missing_fields = catalog_missing_fields(cat)
        missing_carton = any(header == "QTY PER CARTON" for header, _ in missing_fields)
        incomplete_cartons = incomplete_cartons or missing_carton
        qty_per_carton_num = None if missing_carton else float(cat["qty_per_carton"])
        if not cat:
            notes = f'ไม่มีรายละเอียดสินค้าตัวนี้ อัปเดต "{catalog_name}"'
            for col in range(1, PO_LAST_COL + 1):
                ws.cell(line, col).fill = HIGHLIGHT_MISSING_CATALOG
        else:
            notes = "\n".join(f'ไม่มี "{label}" ใน "{catalog_name}"' for _, label in missing_fields)
            for header, _ in missing_fields:
                ws.cell(line, po_cols[header]).fill = HIGHLIGHT_MISSING_CATALOG
        barcode_mismatch = source_barcode_mismatch(row)
        notes = merge_notes(row.get("หมายเหตุ"),
                            BARCODE_MISMATCH_NOTE if barcode_mismatch else "", notes)
        note_cell = ws.cell(line, po_cols["หมายเหตุ"])
        note_cell.value = notes or None
        if notes:
            has_catalog_warnings = True
            note_cell.fill = HIGHLIGHT_MISSING_CATALOG
            # Include the template font size when allowing space for wrapped Thai.
            font_size = note_cell.font.sz or 11
            chars_per_line = max(12, int(ws.column_dimensions["D"].width * 7 / (font_size * 0.6)))
            wrapped_lines = sum(max(1, math.ceil(len(part) / chars_per_line)) for part in notes.split("\n"))
            ws.row_dimensions[line].height = min(409, max(item_height, font_size * 1.5 * wrapped_lines + 12))

        use_month = int(row["USE_MONTH"]) if not pd.isna(row["USE_MONTH"]) else 0

        yuan = row["หยวน"] if not pd.isna(row["หยวน"]) else None
        yuan_num = float(yuan) if yuan is not None else None

        ws.cell(line, po_cols["BUYER ITEM NO."]).value = buyer_item

        if cat.get("img_bytes"):
            add_image_to_cell(ws, f"B{line}", cat["img_bytes"])

        ws.cell(line, po_cols["GOODS DESCRIPTION"]).value = source_desc
        barcode_cell = ws.cell(line, po_cols["BARCODE"])
        barcode_cell.value = po_barcode
        barcode_cell.number_format = "@"
        if barcode_mismatch:
            barcode_cell.fill = HIGHLIGHT_MISSING_CATALOG
        ws.cell(line, po_cols["BRAND"]).value = cat.get("brand", "")
        ws.cell(line, po_cols["MATERIAL"]).value = cat.get("material", "")
        ws.cell(line, po_cols["Weight"]).value = cat.get("weight", "")
        for header, _ in missing_fields:
            if header not in ("GOODS DESCRIPTION", "GOODS PICTURE"):
                ws.cell(line, po_cols[header]).value = None

        ws.cell(line, po_cols["QTY PER CARTON"]).value = qty_per_carton_num

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
            ws[f"{col_thb}{line}"] = yuan_num * float(rate_thb_per_cny)
        else:
            ws[f"{col_fob}{line}"] = None
            ws[f"{col_thb}{line}"] = None

        ws[f"{col_amt}{line}"] = f"={col_fob}{line}*{col_tot_order}{line}"

        if missing_carton:
            # Blank dependencies until a positive carton size is supplied in Excel.
            # These formulas resume calculating if the user fills the yellow cell.
            for col in (col_cart, col_green, col_asia, col_tot_order, col_zan, col_amt):
                cell = ws[f"{col}{line}"]
                expression = str(cell.value).lstrip("=")
                cell.value = f'=IF(IFERROR(AND(ISNUMBER({col_qpc}{line}),{col_qpc}{line}>0),FALSE),{expression},"")'
                cell.fill = HIGHLIGHT_MISSING_CATALOG

    if len(combined_df) > 0:

        last_item_row = ITEM_START_ROW + len(combined_df) - 1
        if has_catalog_warnings:
            verified_label = find_label_cell(ws, "เอกสารชุดนี้ได้ผ่านการตรวจสอบความถูกต้องจากผู้จัดทำแล้ว 100%",
                                            max_row=ws.max_row, max_col=PO_LAST_COL)
            if verified_label:
                ws.cell(*verified_label).value = "โปรดอัปเดตข้อมูลที่ไฮไลต์สีเหลืองตามหมายเหตุ"
        force_bottom_border(ws, last_item_row, 1, PO_LAST_COL)

        # openpyxl moves the template's total rows when items are inserted,
        # but it does not update the formulas inside those rows.
        for col in (col_cart, col_tot_order, col_amt):
            item_range = f"{col}{ITEM_START_ROW}:{col}{last_item_row}"
            ws[f"{col}{BASE_TOTAL_ROW}"] = (
                f'=IF(COUNT({item_range})={len(combined_df)},SUM({item_range}),"")'
                if incomplete_cartons else f"=SUM({item_range})"
            )
            if incomplete_cartons:
                ws[f"{col}{BASE_TOTAL_ROW}"].fill = HIGHLIGHT_MISSING_CATALOG
        ws[f"{col_amt}{BASE_TOTAL_ROW + 1}"] = float(rate_thb_per_cny)
        ws[f"{col_amt}{BASE_TOTAL_ROW + 2}"] = (
            f'=IF({col_amt}{BASE_TOTAL_ROW}="","",{col_amt}{BASE_TOTAL_ROW}*{col_amt}{BASE_TOTAL_ROW + 1})'
            if incomplete_cartons else f"={col_amt}{BASE_TOTAL_ROW}*{col_amt}{BASE_TOTAL_ROW + 1}"
        )
        if incomplete_cartons:
            ws[f"{col_amt}{BASE_TOTAL_ROW + 2}"].fill = HIGHLIGHT_MISSING_CATALOG

        pos_thb = find_label_cell(ws, "TOTAL AMOUNT THB", max_row=400, max_col=60)
        if not pos_thb:
            raise RuntimeError("Cannot find TOTAL AMOUNT THB in sheet.")

        thb_row = pos_thb[0]
        exc_row = thb_row - 1
        ws.row_dimensions[exc_row].height = 40

        footer_row = thb_row + 2

        _add_png(
            ws,
            os.path.join(BASE_DIR, "footer_signatures.png"),
            anchor_cell=f"A{footer_row}",
            width_px=1200,
        )

    wb.remove(template_ws)
    ws.print_area = f"A1:{get_column_letter(PO_LAST_COL)}{ws.max_row}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    wb.save(output_path)

    return output_path

def export_vendor_all_items_excel(vendor_rows_all: pd.DataFrame, vendor_code: str, out_folder: str = PO_OUTPUT_FOLDER) -> str:
    """Export all vendor items, keeping barcode warnings beside product names."""
    os.makedirs(out_folder, exist_ok=True)
    vendor_code = str(vendor_code).strip().upper()
    out_path = os.path.join(out_folder, f"PO_{vendor_code}_ALL_ITEMS.xlsx")

    df_out = vendor_rows_all.copy()
    df_out["หมายเหตุ"] = [
        merge_notes(row.get("หมายเหตุ"), BARCODE_MISMATCH_NOTE if source_barcode_mismatch(row) else "")
        for _, row in df_out.iterrows()
    ]
    cols_wanted = [
        "buyer", "รหัสสินค้า", "รายละเอียดสินค้า", "หมายเหตุ", "barcode", "barcode_ASIA", "barcode_GREEN",
        "ยอดขาย_ASIA", "STOCK_ASIA", "ON_ORDER_ASIA", "หยวน_ASIA",
        "ยอดขาย_GREEN", "STOCK_GREEN", "ON_ORDER_GREEN", "หยวน_GREEN",
        "ยอดขาย_TOTAL", "ON_ORDER_TOTAL",
        "USE_MONTH", "TOTAL_QTY_NUM", "MIN_NUM", "MAX_NUM", "หยวน",
    ]
    cols = [c for c in cols_wanted if c in df_out.columns]
    if "รหัสสินค้า" in df_out.columns and "รายละเอียดสินค้า" in df_out.columns:
        df_out = df_out.sort_values(["รหัสสินค้า", "รายละเอียดสินค้า"], kind="stable")

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        df_out[cols].to_excel(writer, sheet_name="all_items", index=False)

    wb = openpyxl.load_workbook(out_path)
    ws = wb["all_items"]
    header = {str(ws.cell(1, c).value).strip(): c for c in range(1, ws.max_column + 1)}
    for name in ("barcode", "barcode_ASIA", "barcode_GREEN"):
        if name not in header:
            continue
        barcode_col = header[name]
        ws.column_dimensions[get_column_letter(barcode_col)].width = 20
        for r in range(2, ws.max_row + 1):
            cell = ws.cell(r, barcode_col)
            if cell.value is not None:
                cell.value = str(cell.value)
            cell.number_format = "@"

    if "TOTAL_QTY_NUM" in header and "MIN_NUM" in header:
        col_total = header["TOTAL_QTY_NUM"]
        col_min = header["MIN_NUM"]
        for r in range(2, ws.max_row + 1):
            tv = ws.cell(r, col_total).value
            mv = ws.cell(r, col_min).value
            try:
                t = float(tv) if tv is not None else None
                m = float(mv) if mv is not None else None
            except (TypeError, ValueError):
                continue
            if t is not None and m is not None and t < m:
                for c in range(1, ws.max_column + 1):
                    ws.cell(r, c).fill = HIGHLIGHT_BELOW_MIN

    note_col = header["หมายเหตุ"]
    ws.column_dimensions[get_column_letter(note_col)].width = 40
    for r, (_, row) in enumerate(df_out.iterrows(), start=2):
        note_cell = ws.cell(r, note_col)
        note_cell.alignment = Alignment(vertical="top", wrap_text=True)
        if source_barcode_mismatch(row):
            note_cell.fill = HIGHLIGHT_MISSING_CATALOG
            if "barcode" in header:
                ws.cell(r, header["barcode"]).fill = HIGHLIGHT_MISSING_CATALOG
            ws.row_dimensions[r].height = 32

    wb.save(out_path)
    return out_path


def generate_po_streamlit(
    express_asia_path: str,
    express_green_path: str,
    catalog_path: str,
    vendor_info_path: str,
    template_path: str,
    vendor_code: str,
    po_date,
    rate_thb_per_cny: float,
    min_factor: int,
    max_factor: int,
    catalog_filename: Optional[str] = None,
) -> dict:
    """
    Streamlit entry:
      - parse files
      - build combined_all (no vendor filter)
      - export ALL items file for vendor
      - build filtered rows below MIN and create PO if any
    """
    vendor_code = str(vendor_code).strip().upper()

    # allow missing express files
    if express_asia_path and os.path.exists(express_asia_path):
        df_asia, info_asia = parse_express_file(express_asia_path, "ASIA")
    else:
        df_asia, info_asia = pd.DataFrame(), {"months": 0}

    if express_green_path and os.path.exists(express_green_path):
        df_green, info_green = parse_express_file(express_green_path, "GREEN")
    else:
        df_green, info_green = pd.DataFrame(), {"months": 0}

    months = 1
    if info_asia and info_asia.get("months", 0) > 0:
        months = int(info_asia["months"])
    elif info_green and info_green.get("months", 0) > 0:
        months = int(info_green["months"])

    combined_all = build_combined_all(df_asia, df_green, months=months, min_factor=min_factor, max_factor=max_factor)

    vendor_rows_all = combined_all[combined_all["buyer"] == vendor_code].copy()
    if vendor_rows_all.empty:
        buyers = sorted(set(combined_all["buyer"].dropna().astype(str).tolist()))
        preview = ", ".join(buyers[:40])
        raise ValueError(f"Vendor '{vendor_code}' not found. Parsed buyers (first 40): {preview}")

    path_all = export_vendor_all_items_excel(vendor_rows_all, vendor_code=vendor_code)

    vendor_rows_filtered = vendor_rows_all[vendor_rows_all["TOTAL_QTY_NUM"] < vendor_rows_all["MIN_NUM"]].copy()

    path_filtered = None
    if not vendor_rows_filtered.empty:
        count_by_code, count_by_description, count_by_barcode_description = (
            catalog_variant_counts(vendor_rows_all, "catalog_match_barcode")
        )
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
            variant_counts_by_code=count_by_code,
            variant_counts_by_code_description=count_by_description,
            barcode_description_counts=count_by_barcode_description,
            catalog_filename=catalog_filename,
        )

    return {
        "po_filtered": path_filtered,
        "po_all_items": path_all,
        "count_all": int(len(vendor_rows_all)),
        "count_filtered": int(len(vendor_rows_filtered)),
    }

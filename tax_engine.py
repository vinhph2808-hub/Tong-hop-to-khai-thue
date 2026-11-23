# -*- coding: utf-8 -*-
"""
tax_engine.py - Bộ máy trích xuất tờ khai thuế & xuất Excel độc lập đa nền tảng.
Chạy thuần túy bằng Python (OpenPyXL, PyPDF, pdfplumber, xml.etree, RapidOCR).
Tương thích 100% trên Windows, Linux, macOS và Streamlit Cloud!
Hỗ trợ ưu tiên định dạng: XML > PDF > Hình ảnh (PNG, JPG, BMP...) chuyển sang PDF/Scan.
"""

from __future__ import annotations

import datetime
import os
import re
import shutil
import stat
import subprocess
import tempfile
import unicodedata
import xml.etree.ElementTree as ET
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Any

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import pandas as pd
from pypdf import PdfReader
import pdfplumber

# Các hằng số định danh loại thuế
TAX_VAT = "GTGT"
TAX_PIT = "TNCN 05/KK"
TAX_CIT = "TNDN 03/TNDN"
TAX_FCT = "NTNN 01/NTNN"

# Các chỉ tiêu mẫu biểu
E343_INDICATORS = (
    "26", "32a", "29", "30", "32", "33", "31",
    "24", "24a", "25", "38", "37", "42", "43", "22",
)
VAT_INDICATORS = E343_INDICATORS

PIT_INDICATORS = (
    "21", "22", "27", "28", "30", "31", "32",
    "33", "34", "35", "36", "29",
)

CIT_INDICATORS = (
    "A1", "B1", "B8", "B13", "C3", "C4", "C9", "C10", "E",
)

FCT_INDICATORS = ("09", "13")

IMAGE_SUFFIXES = {".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".webp"}
SOURCE_SUFFIXES = {".pdf", ".xml", *IMAGE_SUFFIXES}
ARCHIVE_SUFFIXES = {".zip", ".rar"}


@dataclass
class Declaration:
    path: Path
    month: int
    year: int
    revision: int
    values: dict[str, int]
    missing: list[str]
    modified: float
    tax_type: str = TAX_VAT
    period_type: str = "month"
    note: str = ""
    source_format: str = "PDF"  # XML, PDF, IMAGE, SCAN_PDF
    priority_rank: int = 2      # XML: 3, PDF: 2, IMAGE/SCAN: 1


def fold(text: str) -> str:
    text = unicodedata.normalize("NFD", text or "")
    text = "".join(ch for ch in text if unicodedata.category(ch) != "Mn")
    text = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]+", " ", text)
    return text.replace("đ", "d").replace("Đ", "D").lower()


def clean_pdf_text(text: str) -> str:
    text = re.sub(r"\(cid:\d+\)", " ", text)
    text = text.replace("\u00a0", " ").replace("\x00", " ")
    text = re.sub(r"[\x01-\x08\x0b\x0c\x0e-\x1f]+", " ", text)
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\r\n?", "\n", text)
    return text


def parse_amount(raw: str) -> int:
    raw = raw.strip()
    negative = raw.startswith("(") and raw.endswith(")")
    repeated = re.fullmatch(r"((?:\d{1,3})(?:[.,]\d{3})+)\1+", raw)
    if repeated:
        raw = repeated.group(1)
    raw = re.sub(r"[^0-9,.-]", "", raw)
    if not raw:
        raise ValueError("empty amount")
    digits = re.sub(r"\D", "", raw)
    if not digits:
        raise ValueError(f"invalid amount: {raw}")
    value = int(digits)
    return -value if negative or raw.startswith("-") else value


def extract_period(text: str, filename: str) -> tuple[int, int] | None:
    source = fold(text + "\n" + filename)

    # 1. Khai báo theo Tháng (Ưu tiên cao nhất)
    month_patterns = [
        r"ky\s*tinh\s*thue.{0,50}?thang\s*0?([1-9]|1[0-2])\s*(?:nam|/)?\s*(20\d{2})",
        r"ky\s*tinh\s*thue.{0,50}?t0?([1-9]|1[0-2])\s*[-_./\s]*(20\d{2})",
        r"thang\s*0?([1-9]|1[0-2])\s*(?:nam|/)\s*(20\d{2})",
        r"ky\s*tinh\s*thue\s*:?\s*0?([1-9]|1[0-2])/(20\d{2})",
        r"thang\s*0?([1-9]|1[0-2])\s*nam\s*(20\d{2})",
    ]
    for pattern in month_patterns:
        match = re.search(pattern, source, re.DOTALL)
        if match:
            month = int(match.group(1))
            year = int(match.group(2))
            if 1 <= month <= 12:
                return month, year

    # 2. Khai báo theo Quý
    quarter_patterns = [
        r"(?:ky\s*tinh\s*thue.{0,50}?)?quy\s*([1-4])\s*(?:nam|/)?\s*(20\d{2})",
        r"(?:ky\s*tinh\s*thue.{0,50}?)?quy\s*([iI]{1,3}|IV)\s*(?:nam|/)?\s*(20\d{2})",
        r"(?:ky\s*tinh\s*thue.{0,50}?)?quy\s*([1-4])\s*ky\s*tinh\s*tu\s*thang\s*\d{1,2}/\s*(20\d{2})",
        r"(?:ky\s*tinh\s*thue.{0,50}?)?q([1-4])\s*[-_./\s]*(20\d{2})",
    ]
    for pattern in quarter_patterns:
        match = re.search(pattern, source, re.DOTALL)
        if match:
            q_str = match.group(1).upper()
            q_map = {"1": 1, "2": 2, "3": 3, "4": 4, "I": 1, "II": 2, "III": 3, "IV": 4}
            quarter = q_map.get(q_str, 1)
            year = int(match.group(2))
            return quarter * 3, year

    # 3. Quét kỳ theo tên tệp (Filename)
    fn_source = fold(filename)
    fn_match = re.search(r"(?:thang|t)[-_.\s]*0?([1-9]|1[0-2])[-_.\s]+(20\d{2})", fn_source)
    if not fn_match:
        fn_match = re.search(r"t0?([1-9]|1[0-2])[-_.](20\d{2})", fn_source)
    if not fn_match:
        fn_match = re.search(r"[-_. ]0?([1-9]|1[0-2])\.(20\d{2})", fn_source)
    if fn_match:
        month = int(fn_match.group(1))
        year = int(fn_match.group(2))
        if 1 <= month <= 12:
            return month, year

    fn_quarter = re.search(r"(?:quy|q)[-_.\s]*([1-4])[-_.\s]+(20\d{2})", fn_source)
    if fn_quarter:
        quarter = int(fn_quarter.group(1))
        year = int(fn_quarter.group(2))
        return quarter * 3, year

    # 4. Khai báo theo Năm (Chỉ xét khi không có Tháng và không có Quý)
    annual_patterns = [
        r"ky\s*tinh\s*thue\s*:?\s*nam\s*(20\d{2})",
        r"nam\s*(20\d{2})\s*(?:tu\s*\d{1,2}/\d{1,2}/\d{4}\s*den|\(tu|tu\s*01/01/)",
        r"ky\s*tinh\s*thue.{0,50}?nam\s*(20\d{2})",
    ]
    for pattern in annual_patterns:
        match = re.search(pattern, source, re.DOTALL)
        if match:
            year = int(match.group(1))
            return 12, year

    return None


def extract_revision(text: str, filename: str) -> int:
    source = fold(text + "\n" + filename)
    match = re.search(r"bo\s*sung\s*lan\s*(?:thu\s*)?(\d+)", source)
    if match:
        return int(match.group(1))
    fn_match = re.search(r"(?:lan\s*|l)(\d+)", fold(filename))
    if fn_match and "lan dau" not in source:
        try:
            return int(fn_match.group(1))
        except ValueError:
            pass
    return 0


def detect_tax_type(text: str) -> str | None:
    source = fold(text[:12000])
    if "01/gtgt" in source or "gia tri gia tang" in source:
        return TAX_VAT
    if "05/kk-tncn" in source or "02/kk-tncn" in source or ("05/kk-" in source and "thu nhap ca nhan" in source) or "thu nhap ca nhan" in source:
        return TAX_PIT
    if "01/ntnn" in source or "nha thau" in source:
        return TAX_FCT
    if "03/tndn" in source or "03-tndn" in source or "03_tndn" in source or ("quyet toan" in source and "thu nhap doanh nghiep" in source) or "thu nhap doanh nghiep" in source:
        return TAX_CIT
    # Mặc định kiểm tra chỉ tiêu [43] hoặc [22]
    if "[43]" in source or "[22]" in source or "[25]" in source:
        return TAX_VAT
    return None


def extract_indicators(text: str) -> tuple[dict[str, int], list[str]]:
    results: dict[str, int] = {}
    lines = clean_pdf_text(text).splitlines()
    marker = re.compile(r"\[\s*(24a|24|25|26|29|30|31|32a|32|33|37|38|42|43|22)\s*\]", re.I)
    amount_at_start = re.compile(
        r"^\s*(?:(?:vnd|vnđ|đ|đồng|:)\s*)*"
        r"(\(\s*(?:\d{1,3}(?:[.,]\d{3})+|[1-9]\d*|0)\s*\)|"
        r"-?\s*(?:\d{1,3}(?:[.,]\d{3})+|[1-9]\d*|0))",
        re.I,
    )
    for line in lines:
        matches = list(marker.finditer(line))
        for index, match in enumerate(matches):
            code = match.group(1).lower()
            if code in results:
                continue
            end = matches[index + 1].start() if index + 1 < len(matches) else len(line)
            next_bracket = line.find("[", match.end())
            if next_bracket != -1 and (index + 1 >= len(matches) or next_bracket < matches[index + 1].start()):
                end = next_bracket
            after = line[match.end():end]
            value_match = amount_at_start.match(after)
            if not value_match:
                continue
            try:
                results[code] = parse_amount(value_match.group(1))
            except ValueError:
                continue

    if "22" not in results:
        results["22"] = 0

    missing = [code for code in VAT_INDICATORS if code not in results]
    return results, missing


def extract_pit_indicator_29(text: str) -> int | None:
    lines = clean_pdf_text(text).splitlines()
    for i, line in enumerate(lines):
        for m in re.finditer(r"\[\s*29\s*\](?!\s*=)", line):
            after = line[m.end():]
            amt_match = re.search(r"(?:vn[đd]|dong|[:\-])?\s*([(-]?\s*\d{1,3}(?:[.,]\d{3})+|-?\s*\d+)\b", after, re.I)
            if amt_match:
                try:
                    return parse_amount(amt_match.group(1))
                except Exception:
                    pass
        if re.search(r"\[\s*29\s*\](?!\s*=)\s*$", line):
            if i + 1 < len(lines):
                next_line = lines[i + 1]
                amt_match = re.search(r"(?:vn[đd]|dong|[:\-])?\s*([(-]?\s*\d{1,3}(?:[.,]\d{3})+|-?\s*\d+)\b", next_line, re.I)
                if amt_match:
                    try:
                        return parse_amount(amt_match.group(1))
                    except Exception:
                        pass

    match = re.search(r"\[\s*29\s*\](?!\s*=)[^0-9\n]{0,30}?([(-]?\s*\d{1,3}(?:[.,]\d{3})+|-?\s*\d+)", text, re.I)
    if match:
        try:
            return parse_amount(match.group(1))
        except Exception:
            pass

    return None


def extract_cit_indicators(text: str) -> tuple[dict[str, int], list[str]]:
    results: dict[str, int] = {}
    lines = clean_pdf_text(text).splitlines()
    for code in CIT_INDICATORS:
        pattern = re.compile(
            rf"(?:\[\s*{code}\s*\]|\b{code}\b)\s+([(-]?\s*\d[\d.,]*\)?)(?:\s*(?:vnd|vnđ|đ|đồng))?\s*$",
            re.I,
        )
        for line in lines:
            line_clean = line.strip()
            match = pattern.search(line_clean)
            if match:
                try:
                    results[code] = parse_amount(match.group(1))
                    break
                except ValueError:
                    continue
    for code in CIT_INDICATORS:
        if code in results:
            continue
        pattern = re.compile(
            rf"(?:\[\s*{code}\s*\]|\b{code}\b)\s+([(-]?\s*\d[\d.,]*\)?)\b",
            re.I,
        )
        for line in lines:
            line_clean = line.strip()
            if f"={code}" in line_clean or f"+{code}" in line_clean or f"-{code}" in line_clean:
                continue
            match = pattern.search(line_clean)
            if match:
                try:
                    results[code] = parse_amount(match.group(1))
                    break
                except ValueError:
                    continue
    missing = [code for code in CIT_INDICATORS if code not in results]
    return results, missing


def _fct_header_code(text: str) -> str | None:
    compact = fold(text).replace(" ", "")
    exact = {"(7)": "07", "(8)": "08", "(10)": "10", "(11)": "11", "(12)": "12"}
    if compact in exact:
        return exact[compact]
    if compact.startswith("(9=7x8"):
        return "09"
    if compact.startswith("[13="):
        return "13"
    if compact.startswith("(14)"):
        return "14"
    return None


def extract_fct_indicators_from_word_pages(
    word_pages: list[list[dict]],
) -> tuple[dict[str, int], list[str]]:
    column_centers: dict[str, float] | None = None
    for words in word_pages:
        markers = []
        for word in words:
            code = _fct_header_code(str(word.get("text", "")))
            if code:
                markers.append((code, word))
        marker_09 = next((word for code, word in markers if code == "09"), None)
        if marker_09 is not None:
            marker_top = float(marker_09["top"])
            nearby: dict[str, dict] = {}
            for code, word in markers:
                if abs(float(word["top"]) - marker_top) <= 12:
                    current = nearby.get(code)
                    if current is None or abs(float(word["top"]) - marker_top) < abs(float(current["top"]) - marker_top):
                        nearby[code] = word
            if all(code in nearby for code in ("08", "09", "10", "12", "13", "14")):
                column_centers = {
                    code: (float(word["x0"]) + float(word["x1"])) / 2
                    for code, word in nearby.items()
                }

        if not column_centers:
            continue
        ordered = sorted(column_centers.items(), key=lambda item: item[1])
        positions = {code: index for index, (code, _) in enumerate(ordered)}
        for total_word in words:
            if fold(str(total_word.get("text", ""))).strip(" :") != "tong":
                continue
            row_top = float(total_word["top"])
            row_words = [word for word in words if abs(float(word["top"]) - row_top) <= 3]
            if not any(fold(str(word.get("text", ""))).strip(" :") == "cong" for word in row_words):
                continue
            numeric_words = []
            for word in row_words:
                raw = str(word.get("text", "")).strip()
                if not re.fullmatch(r"\(?\s*-?\d[\d.,]*\s*\)?", raw):
                    continue
                try:
                    value = parse_amount(raw)
                except ValueError:
                    continue
                center = (float(word["x0"]) + float(word["x1"])) / 2
                numeric_words.append((center, value))
            results: dict[str, int] = {}
            for target in ("09", "13"):
                index = positions[target]
                center = ordered[index][1]
                left = (ordered[index - 1][1] + center) / 2 if index > 0 else center - 30
                right = (center + ordered[index + 1][1]) / 2 if index + 1 < len(ordered) else center + 30
                candidates = [item for item in numeric_words if left <= item[0] < right]
                results[target] = min(candidates, key=lambda item: abs(item[0] - center))[1] if candidates else 0
            return results, []
    return {}, ["09", "13"]


def extract_fct_indicators_from_pdf(path: Path) -> tuple[dict[str, int], list[str]]:
    word_pages: list[list[dict]] = []
    try:
        with pdfplumber.open(path) as pdf:
            for page in pdf.pages:
                deduped = page.dedupe_chars(tolerance=1)
                word_pages.append(deduped.extract_words(
                    x_tolerance=2,
                    y_tolerance=3,
                    keep_blank_chars=False,
                    use_text_flow=False,
                ))
    except Exception:
        return {}, ["09", "13"]
    return extract_fct_indicators_from_word_pages(word_pages)


def extract_fct_indicators_layout(layout_text: str) -> tuple[dict[str, int], list[str]]:
    amount_pattern = re.compile(r"(?<!\d)(?:\(\s*\d[\d.,]*\s*\)|-?\d[\d.,]*)(?!\d)")
    marker_patterns = {
        "07": re.compile(r"\(\s*7\s*\)", re.I),
        "08": re.compile(r"\(\s*8\s*\)", re.I),
        "09": re.compile(r"\(\s*9\s*(?:=\s*7\s*[x×*]\s*8\s*)?\)", re.I),
        "10": re.compile(r"\(\s*10\s*\)", re.I),
        "11": re.compile(r"\(\s*11\s*\)", re.I),
        "12": re.compile(r"\(\s*12\s*\)", re.I),
        "13": re.compile(r"\[\s*13\s*=\s*.*?(?=\s{2,}\(\s*14\s*\)|$)", re.I),
        "14": re.compile(r"\(\s*14\s*(?:=\s*\(?\s*9\s*\)?\s*\+\s*\(?\s*13\s*\)?\s*)?\)", re.I),
    }
    layout_text = clean_pdf_text(layout_text)
    column_centers: dict[str, float] | None = None
    total_rows: list[str] = []
    for line in layout_text.splitlines():
        centers: dict[str, float] = {}
        for code, pattern in marker_patterns.items():
            match = pattern.search(line)
            if match:
                centers[code] = (match.start() + match.end()) / 2
        if all(code in centers for code in ("07", "09", "10", "13", "14")):
            column_centers = centers
        if "tong cong" in fold(line):
            total_rows.append(line)
            if column_centers:
                ordered = sorted(column_centers.items(), key=lambda item: item[1])
                left_edge = ordered[0][1] - ((ordered[1][1] - ordered[0][1]) / 2 if len(ordered) > 1 else 20)
                right_edge = ordered[-1][1] + ((ordered[-1][1] - ordered[-2][1]) / 2 if len(ordered) > 1 else 20)
                assigned: dict[str, int] = {}
                for match in amount_pattern.finditer(line):
                    token_center = (match.start() + match.end()) / 2
                    if not left_edge <= token_center <= right_edge:
                        continue
                    code = min(ordered, key=lambda item: abs(item[1] - token_center))[0]
                    try:
                        assigned[code] = parse_amount(match.group(0))
                    except ValueError:
                        continue
                return {"09": assigned.get("09", 0), "13": assigned.get("13", 0)}, []

    for line in total_rows:
        tokens = amount_pattern.findall(line)
        if len(tokens) < 7:
            continue
        try:
            values = [parse_amount(token) for token in tokens[-7:]]
        except ValueError:
            continue
        return {"09": values[2], "13": values[5]}, []
    return {}, ["09", "13"]


def convert_image_to_pdf(image_path: Path, output_pdf_path: Path) -> Path:
    """Chuyển đổi file hình ảnh (PNG, JPG, BMP...) sang định dạng chuẩn PDF."""
    from PIL import Image
    with Image.open(str(image_path)) as img:
        rgb = img.convert("RGB")
        rgb.save(str(output_pdf_path), "PDF", resolution=150.0)
    return output_pdf_path


def read_scan_declaration(path: Path) -> Declaration:
    """Đọc tờ khai từ bản scan/ảnh qua OCR (RapidOCR)."""
    try:
        from tax_scan_ocr import scan_pages, ocr_pages_to_text, printed_pit_amounts
    except Exception as exc:
        raise ValueError(f"Thiếu bộ xử lý OCR scan/ảnh: {exc}")

    pages = scan_pages(path)
    full_text = ocr_pages_to_text(pages)
    if not full_text.strip():
        raise ValueError("OCR: Không nhận diện được chữ hoặc số liệu từ bản scan/ảnh")

    tax_type = detect_tax_type(full_text)
    if tax_type is None:
        source = fold(full_text[:8000])
        if "06/tncn" in source:
            raise ValueError("Bỏ qua mẫu 06/TNCN: chỉ tiêu 29 không phải số thuế khấu trừ")
        raise ValueError("OCR: Không nhận diện được mẫu tờ khai 01/GTGT, 05/KK-TNCN, 01/NTNN hoặc 03/TNDN")

    period = extract_period(full_text, path.name)
    if not period:
        raise ValueError("OCR: Không xác định được tháng/quý/năm kê khai từ bản scan/ảnh")

    period_source = fold(full_text + "\n" + path.name)
    has_month = bool(re.search(r"thang\s*0?([1-9]|1[0-2])", period_source)) or bool(re.search(r"\bt0?([1-9]|1[0-2])\b", period_source))
    is_quarterly = not has_month and bool(re.search(r"(?:ky\s*tinh\s*thue.{0,50}?)?quy\s*([1-4]|iv|iii|ii|i)", period_source, re.DOTALL))
    is_annual = not has_month and not is_quarterly and ((tax_type == TAX_CIT) or bool(re.search(r"ky\s*tinh\s*thue\s*:?\s*nam\s*20\d{2}", period_source, re.DOTALL)))
    revision = extract_revision(full_text, path.name)
    period_type = "year" if is_annual else ("quarter" if is_quarterly else "month")

    mtime = path.stat().st_mtime if path.exists() else 0.0

    if tax_type == TAX_VAT:
        values, missing = extract_indicators(full_text)
        note = "Đọc từ bản scan/ảnh OCR (01/GTGT)."
        return Declaration(
            path=path, month=period[0], year=period[1], revision=revision,
            values=values, missing=missing, modified=mtime, tax_type=TAX_VAT,
            period_type=period_type, note=note,
            source_format="IMAGE" if path.suffix.lower() in IMAGE_SUFFIXES else "SCAN_PDF",
            priority_rank=1,
        )
    elif tax_type == TAX_PIT:
        try:
            amounts = printed_pit_amounts(pages)
            pit_val = amounts["29"]
            note = "OCR số in trong bảng; đối chiếu [29] = [30] + [31]."
        except Exception:
            pit_val = extract_pit_indicator_29(full_text)
            note = "Đọc từ bản scan/ảnh OCR (05/KK-TNCN)."
        values = {"29": pit_val} if pit_val is not None else {}
        missing = [] if pit_val is not None else ["29"]
        return Declaration(
            path=path, month=period[0], year=period[1], revision=revision,
            values=values, missing=missing, modified=mtime, tax_type=TAX_PIT,
            period_type=period_type, note=note,
            source_format="IMAGE" if path.suffix.lower() in IMAGE_SUFFIXES else "SCAN_PDF",
            priority_rank=1,
        )
    elif tax_type == TAX_CIT:
        values, missing = extract_cit_indicators(full_text)
        note = "Đọc từ bản scan/ảnh OCR (03/TNDN)."
        return Declaration(
            path=path, month=period[0], year=period[1], revision=revision,
            values=values, missing=missing, modified=mtime, tax_type=TAX_CIT,
            period_type=period_type, note=note,
            source_format="IMAGE" if path.suffix.lower() in IMAGE_SUFFIXES else "SCAN_PDF",
            priority_rank=1,
        )
    else:  # TAX_FCT
        values, missing = extract_fct_indicators_layout(full_text)
        note = "Đọc từ bản scan/ảnh OCR (01/NTNN)."
        return Declaration(
            path=path, month=period[0], year=period[1], revision=revision,
            values=values, missing=missing, modified=mtime, tax_type=TAX_FCT,
            period_type=period_type, note=note,
            source_format="IMAGE" if path.suffix.lower() in IMAGE_SUFFIXES else "SCAN_PDF",
            priority_rank=1,
        )


def read_pdf_declaration(path: Path) -> Declaration:
    """Đọc tờ khai từ file PDF kỹ thuật số (hoặc tự động chuyển fallback sang OCR nếu là bản scan)."""
    reader = PdfReader(str(path))
    if reader.is_encrypted:
        try:
            reader.decrypt("")
        except Exception as exc:
            raise ValueError(f"File PDF '{path.name}' bị khóa mật khẩu") from exc

    text = ""
    try:
        with pdfplumber.open(path) as pdf:
            pages = [page.extract_text() or "" for page in pdf.pages]
            text = clean_pdf_text("\n".join(pages))
    except Exception:
        text = ""

    if len(text.strip()) < 40:
        text = clean_pdf_text("\n".join((page.extract_text() or "") for page in reader.pages))

    # Nếu văn bản quá ít (< 40 ký tự) -> Đây là file PDF scan/ảnh chụp -> Dùng bộ OCR
    if len(text.strip()) < 40:
        return read_scan_declaration(path)

    tax_type = detect_tax_type(text)
    if tax_type is None:
        source = fold(text[:8000])
        if "06/tncn" in source:
            raise ValueError("Bỏ qua mẫu 06/TNCN: chỉ tiêu 29 không phải số thuế khấu trừ")
        # Thử fallback qua OCR phòng trường hợp font chữ bị mã hóa CID
        try:
            return read_scan_declaration(path)
        except Exception:
            raise ValueError(f"Không nhận diện được loại tờ khai thuế hợp lệ trong file '{path.name}'")

    period_source = fold(text + "\n" + path.name)
    has_month = bool(re.search(r"thang\s*0?([1-9]|1[0-2])", period_source)) or bool(re.search(r"\bt0?([1-9]|1[0-2])\b", period_source))
    is_quarterly = not has_month and bool(re.search(r"(?:ky\s*tinh\s*thue.{0,50}?)?quy\s*([1-4]|iv|iii|ii|i)", period_source, re.DOTALL))
    is_annual = not has_month and not is_quarterly and ((tax_type == TAX_CIT) or bool(re.search(r"ky\s*tinh\s*thue\s*:?\s*nam\s*20\d{2}", period_source, re.DOTALL)))
    period = extract_period(text, path.name)
    if not period:
        raise ValueError(f"Không xác định được tháng/quý/năm kê khai trong file '{path.name}'")

    period_type = "year" if is_annual else ("quarter" if is_quarterly else "month")

    if tax_type == TAX_VAT:
        values, missing = extract_indicators(text)
    elif tax_type == TAX_PIT:
        layout_text = "\n".join((page.extract_text(extraction_mode="layout") or "") for page in reader.pages)
        pit_value = extract_pit_indicator_29(layout_text)
        if pit_value is None:
            pit_value = extract_pit_indicator_29(text)
        if pit_value is None:
            try:
                scan_decl = read_scan_declaration(path)
                if "29" in scan_decl.values and scan_decl.values["29"] is not None:
                    pit_value = scan_decl.values["29"]
            except Exception:
                pass
        values = {"29": pit_value} if pit_value is not None else {}
        missing = [] if pit_value is not None else ["29"]
    elif tax_type == TAX_CIT:
        values, missing = extract_cit_indicators(text)
    else:  # TAX_FCT
        values, missing = extract_fct_indicators_from_pdf(path)
        if missing:
            layout_text = "\n".join((page.extract_text(extraction_mode="layout") or "") for page in reader.pages)
            values, missing = extract_fct_indicators_layout(layout_text)

    mtime = path.stat().st_mtime if path.exists() else 0.0
    return Declaration(
        path=path,
        month=period[0],
        year=period[1],
        revision=extract_revision(text, path.name),
        values=values,
        missing=missing,
        modified=mtime,
        tax_type=tax_type,
        period_type=period_type,
        note="",
        source_format="PDF",
        priority_rank=2,
    )


def xml_local_name(tag: str) -> str:
    return fold(tag.rsplit("}", 1)[-1]).replace("_", "").replace("-", "")


def parse_xml_period_value(raw_period: str, is_quarterly: bool) -> tuple[int, int] | None:
    raw_period = raw_period.strip()
    match = re.fullmatch(r"\s*(?:nam\s*|y)?(20\d{2})\s*", fold(raw_period))
    if match:
        return 12, int(match.group(1))
    match = re.fullmatch(r"\s*(\d{1,2})\s*[/.-]\s*(\d{1,2})\s*[/.-]\s*(20\d{2})\s*", raw_period)
    if match:
        month, year = int(match.group(2)), int(match.group(3))
        if 1 <= month <= 12:
            return month, year
    match = re.fullmatch(r"\s*(20\d{2})\s*[/.-]\s*(\d{1,2})\s*[/.-]\s*(\d{1,2})\s*", raw_period)
    if match:
        month, year = int(match.group(2)), int(match.group(1))
        if 1 <= month <= 12:
            return month, year
    match = re.fullmatch(r"\s*(0?[1-9]|1[0-2])\s*/\s*(20\d{2})\s*", raw_period)
    if match:
        number, year = int(match.group(1)), int(match.group(2))
        return (number * 3, year) if is_quarterly and 1 <= number <= 4 else (number, year)
    return extract_period(raw_period, "")


def read_xml_declaration(path: Path) -> Declaration:
    """Đọc tờ khai từ file XML chuẩn nộp Thuế điện tử eTax (Tổng cục Thuế)."""
    try:
        root = ET.parse(path).getroot()
    except ET.ParseError as exc:
        raise ValueError(f"File XML '{path.name}' bị lỗi cấu trúc") from exc

    fields: list[tuple[str, str]] = []
    for element in root.iter():
        value = (element.text or "").strip()
        if value:
            fields.append((xml_local_name(element.tag), value))

    searchable = "\n".join(value for _, value in fields) + "\n" + path.name
    tax_type = detect_tax_type(searchable)
    if tax_type is None:
        raise ValueError(f"Không phải XML tờ khai 01/GTGT, 05/KK-TNCN, 01/NTNN hoặc 03/TNDN: '{path.name}'")

    # Bóc tách các thẻ con định danh theo dạng <ChiTieu><ma>...</ma><giaTri>...</giaTri></ChiTieu>
    structured_values: dict[str, str] = {}
    for parent in root.iter():
        children = list(parent)
        if len(children) >= 2:
            code_child = None
            val_child = None
            for ch in children:
                ch_tag = xml_local_name(ch.tag)
                if ch_tag in {"ma", "maso", "mact", "machitieu", "id", "code"}:
                    code_child = (ch.text or "").strip()
                elif ch_tag in {"giatri", "sotien", "tien", "value", "amount", "so"}:
                    val_child = (ch.text or "").strip()
            if code_child and val_child:
                norm_code = xml_local_name(code_child)
                structured_values[norm_code] = val_child

    quarter_markers = [value for name, value in fields if name in {"kieuky", "kykkhai", "kykekhai", "kytinhthue"}]
    quarter_source = fold(" ".join(quarter_markers) + " " + path.name)
    is_quarterly = bool(re.search(r"(?:quy\s*|\bq\s*)[1-4]", quarter_source)) or any(
        fold(value).strip() in {"q", "quy"} for value in quarter_markers
    )
    is_annual = (tax_type == TAX_CIT)
    period = None
    if tax_type == TAX_CIT:
        for name, raw_value in fields:
            if name in {"namkkhai", "namkekhai", "namtinhthue", "nam", "namqt", "namqtt", "kykkhai", "kykekhai", "kytinhthue", "kykkhaitungay", "kykekhaitungay"}:
                match = re.search(r"20\d{2}", raw_value)
                if match:
                    period = (12, int(match.group(0)))
                    break
    if period is None:
        primary_period_names = {"kykkhai", "kykekhai", "kytinhthue"}
        secondary_period_names = {"kykkhaitungay", "kykekhaitungay", "tungay"}
        period_values = [value for name, value in fields if name in primary_period_names]
        if not period_values:
            period_values = [value for name, value in fields if name in secondary_period_names]
        for raw_period in period_values:
            period = parse_xml_period_value(raw_period, is_quarterly)
            if period:
                break
    if period is None:
        period = extract_period(searchable if tax_type == TAX_CIT else "", path.name)
    if not period:
        raise ValueError(f"Không xác định được tháng/quý/năm kê khai trong file XML '{path.name}'")

    requested = (
        (*VAT_INDICATORS, "22")
        if tax_type == TAX_VAT
        else (("29",) if tax_type == TAX_PIT else (CIT_INDICATORS if tax_type == TAX_CIT else ("09", "13")))
    )
    values: dict[str, int] = {}
    for code in requested:
        normalized_code = code.lower()
        compact_code = normalized_code.lstrip("0") or "0"
        accepted_names = {
            normalized_code,
            compact_code,
            f"ct{normalized_code}",
            f"ct{compact_code}",
            f"chitieu{normalized_code}",
            f"chitieu{compact_code}",
            f"chitieutokhai{normalized_code}",
            f"chitieutokhai{compact_code}",
            f"ctkkhai{normalized_code}",
            f"ctkkhai{compact_code}",
        }
        if tax_type == TAX_VAT and normalized_code == "24":
            preferred_names = {"ct24", "chitieu24", "chitieutokhai24", "ctkkhai24"}
            candidate_fields = [item for item in fields if item[0] in preferred_names]
            candidate_fields += [item for item in fields if item[0] in accepted_names]
        elif tax_type == TAX_FCT and normalized_code in {"09", "13"}:
            preferred_names = {
                f"tongthuegtgtphatsinh{compact_code}",
                f"tongthuetndnphatsinh{compact_code}",
                f"tongchitieu{compact_code}",
            }
            candidate_fields = [item for item in fields if item[0] in preferred_names]
            candidate_fields += [item for item in fields if item[0] in accepted_names]
        else:
            candidate_fields = [item for item in fields if item[0] in accepted_names]
            for candidate_key in accepted_names:
                if candidate_key in structured_values:
                    candidate_fields.append((candidate_key, structured_values[candidate_key]))
        for _, raw_value in candidate_fields:
            try:
                values[code] = parse_amount(raw_value)
                break
            except ValueError:
                continue

    if tax_type == TAX_CIT:
        found_any = any(code in values for code in CIT_INDICATORS)
        if found_any:
            for code in CIT_INDICATORS:
                if code not in values:
                    values[code] = 0

    if tax_type == TAX_VAT:
        if "22" not in values:
            values["22"] = 0

    required_codes = VAT_INDICATORS if tax_type == TAX_VAT else requested
    missing = [code for code in required_codes if code not in values]
    revision = extract_revision(searchable, path.name)
    for name, raw_value in fields:
        if name in {"solan", "lanbosung", "lanbs"} and raw_value.strip().isdigit():
            revision = max(revision, int(raw_value.strip()))

    mtime = path.stat().st_mtime if path.exists() else 0.0
    return Declaration(
        path=path,
        month=period[0],
        year=period[1],
        revision=revision,
        values=values,
        missing=missing,
        modified=mtime,
        tax_type=tax_type,
        period_type="year" if (tax_type == TAX_CIT or is_annual) else ("quarter" if is_quarterly else "month"),
        note="Chuẩn XML eTax Tổng cục Thuế (Độ chính xác cao nhất).",
        source_format="XML",
        priority_rank=3,
    )


def read_declaration_auto(path: Path) -> Declaration:
    """Tự động nhận diện định dạng tệp và bóc tách theo đúng ưu tiên."""
    ext = path.suffix.lower()
    if ext == ".xml":
        return read_xml_declaration(path)
    elif ext in IMAGE_SUFFIXES:
        return read_scan_declaration(path)
    elif ext == ".pdf":
        return read_pdf_declaration(path)
    else:
        raise ValueError(f"Định dạng tệp không được hỗ trợ: {path.name}")


def extract_archive(archive_path: Path, extract_dir: Path) -> list[Path]:
    """Giải nén an toàn file ZIP hoặc RAR, quét đệ quy các file hợp lệ."""
    extracted: list[Path] = []
    suffix = archive_path.suffix.lower()

    if suffix == ".zip":
        with zipfile.ZipFile(str(archive_path), "r") as zf:
            for info in zf.infolist():
                norm = info.filename.replace("\\", "/")
                parts = Path(norm).parts
                if not norm or Path(norm).is_absolute() or ".." in parts:
                    continue
                target = (extract_dir / Path(*parts)).resolve()
                if info.is_dir():
                    target.mkdir(parents=True, exist_ok=True)
                    continue
                target.parent.mkdir(parents=True, exist_ok=True)
                with zf.open(info, "r") as src, open(target, "wb") as dst:
                    shutil.copyfileobj(src, dst)
    elif suffix == ".rar":
        unrar = shutil.which("unrar") or shutil.which("UnRAR.exe") or "C:\\Program Files\\WinRAR\\UnRAR.exe"
        if unrar and os.path.exists(unrar):
            subprocess.run([unrar, "x", "-y", str(archive_path), str(extract_dir) + os.sep], check=False)
        else:
            try:
                import rarfile
                with rarfile.RarFile(str(archive_path), "r") as rf:
                    rf.extractall(str(extract_dir))
            except Exception as exc:
                raise ValueError(f"Không thể giải nén file RAR: {exc}")

    for root, _, files in os.walk(str(extract_dir)):
        for f in files:
            p = Path(root) / f
            if p.suffix.lower() in SOURCE_SUFFIXES:
                extracted.append(p)
    return extracted


def select_latest_declarations(
    sources: list[Path],
    logger: Callable[[str, str], None] | None = None,
    period_mode: str = "all",
) -> tuple[list[Declaration], list[tuple[str, str]], list[dict]]:
    """
    Đọc tất cả các tệp và chọn lọc tờ khai tối ưu theo nguyên tắc:
    1. Ưu tiên định dạng: XML (hạng 3) > PDF (hạng 2) > Hình ảnh/Scan (hạng 1).
    2. Ưu tiên lần nộp: Bổ sung lần cao hơn > Lần đầu.
    3. Thời gian cập nhật mới hơn.
    Trả về: (selected_declarations, errors, audit_records)
    """
    if logger is None:
        logger = lambda m, t="info": None

    declarations: list[Declaration] = []
    errors: list[tuple[str, str]] = []
    audit_records: list[dict] = []

    for index, path in enumerate(sources, 1):
        try:
            decl = read_declaration_auto(path)
            # Lọc theo kỳ
            if period_mode != "all" and decl.period_type in {"month", "quarter"} and decl.period_type != period_mode:
                audit_records.append({
                    "Tệp": path.name,
                    "Định dạng": decl.source_format,
                    "Loại tờ khai": decl.tax_type,
                    "Kỳ thuế": f"{'Quý' if decl.period_type == 'quarter' else 'Tháng'} {decl.month // 3 if decl.period_type == 'quarter' else decl.month:02d}/{decl.year}",
                    "Lần nộp": "Lần đầu" if decl.revision == 0 else f"Bổ sung L{decl.revision}",
                    "Trạng thái": f"Bỏ qua (Do lọc {period_mode})",
                    "Ghi chú": decl.note,
                })
                logger(f"Bỏ qua {path.name}: không đúng bộ lọc kỳ ({period_mode})", "info")
                continue

            declarations.append(decl)
            logger(f"✔ Đã đọc: {path.name} [{decl.source_format}] ({decl.tax_type} - Tháng/Quý {decl.month}/{decl.year} - Lần {decl.revision})", "success")
        except Exception as exc:
            err_msg = str(exc)
            errors.append((path.name, err_msg))
            audit_records.append({
                "Tệp": path.name,
                "Định dạng": path.suffix.upper().replace(".", ""),
                "Loại tờ khai": "Chưa xác định",
                "Kỳ thuế": "-",
                "Lần nộp": "-",
                "Trạng thái": "❌ Lỗi đọc",
                "Ghi chú": err_msg,
            })
            logger(f"⚠ Lỗi {path.name}: {err_msg}", "warning")

    # Áp dụng bộ lọc chọn lọc tối ưu
    latest: dict[tuple[str, int, int], Declaration] = {}
    fct_declarations: list[Declaration] = []

    for item in declarations:
        if item.tax_type == TAX_FCT:
            fct_declarations.append(item)
            continue

        key = (item.tax_type, item.year, item.month)
        current = latest.get(key)

        # Trọng số so sánh: (Lần bổ sung, Ưu tiên định dạng XML > PDF > Scan, Thời gian sửa đổi, Tên file)
        score_new = (item.revision, item.priority_rank, item.modified, item.path.name)
        score_current = (current.revision, current.priority_rank, current.modified, current.path.name) if current else (-1, -1, -1, "")

        if current is None or score_new > score_current:
            latest[key] = item

    selected = sorted([*latest.values(), *fct_declarations], key=lambda x: (x.tax_type, x.year, x.month, x.revision))
    selected_ids = {id(d) for d in selected}

    for decl in declarations:
        period_str = f"Quý {decl.month // 3}" if decl.period_type == "quarter" else f"Tháng {decl.month:02d}"
        if decl.period_type == "year":
            period_str = "Cả năm"
        rev_str = "Lần đầu" if decl.revision == 0 else f"Bổ sung L{decl.revision}"

        is_chosen = id(decl) in selected_ids
        audit_records.append({
            "Tệp": decl.path.name,
            "Định dạng": decl.source_format,
            "Loại tờ khai": decl.tax_type,
            "Kỳ thuế": f"{period_str}/{decl.year}",
            "Lần nộp": rev_str,
            "Trạng thái": "✅ ĐƯỢC CHỌN" if is_chosen else "Thay thế bởi bản mới/XML",
            "Ghi chú": decl.note or ("Đã chọn vào bảng tính" if is_chosen else "Đã có bản ưu tiên hơn"),
        })

    return selected, errors, audit_records


def write_to_excel_openpyxl(
    template_path: Path,
    output_path: Path,
    declarations: list[Declaration],
    parse_errors: list[tuple[str, str]],
    logger: Callable[[str, str], None] | None = None,
) -> Path:
    """
    Ghi kết quả vào file Excel dựa trên mẫu MẪU.xlsx bằng thư viện OpenPyXL.
    100% độc lập, không cần Microsoft Excel hay Win32 COM, chạy mượt trên mọi hệ điều hành!
    """
    if logger is None:
        logger = lambda m, t="info": None

    if not template_path.exists():
        raise FileNotFoundError(f"Không tìm thấy tệp mẫu: {template_path}")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb = openpyxl.load_workbook(str(template_path))

    number_format = "#,##0;(-#,##0);0"

    # --- 1. SHEET E343 (THUẾ GTGT) ---
    if "E343" in wb.sheetnames:
        ws_vat = wb["E343"]
        header_rows = []
        for r in range(1, ws_vat.max_row + 1):
            if fold(str(ws_vat.cell(r, 1).value or "")).strip() == "chi tieu":
                header_rows.append(r)

        vat_targets: dict[tuple[int, str], tuple[int, int]] = {}
        for h_row in header_rows:
            code_cols: dict[str, int] = {}
            for c in range(2, ws_vat.max_column + 1):
                val = str(ws_vat.cell(h_row, c).value or "").strip().lower().replace("[", "").replace("]", "").replace(" ", "")
                if val in E343_INDICATORS:
                    code_cols[val] = c
            if not code_cols:
                continue
            for r in range(h_row + 1, min(ws_vat.max_row + 1, h_row + 20)):
                m_val = str(ws_vat.cell(r, 1).value or "").strip()
                try:
                    m = int(float(m_val))
                except ValueError:
                    continue
                if 1 <= m <= 12:
                    for code, col_idx in code_cols.items():
                        vat_targets[(m, code)] = (r, col_idx)

        # Ghi giá trị 0 mặc định cho tất cả ô dữ liệu VAT
        for (m, code), (r, c) in vat_targets.items():
            cell = ws_vat.cell(r, c)
            cell.value = 0
            cell.number_format = number_format

        # Ghi số liệu từ các tờ khai GTGT đã bóc tách
        for decl in declarations:
            if decl.tax_type == TAX_VAT:
                for code, val in decl.values.items():
                    target = vat_targets.get((decl.month, code))
                    if target is not None:
                        r, c = target
                        cell = ws_vat.cell(r, c)
                        cell.value = val
                        cell.number_format = number_format

    # --- 2. SHEET E350 (THUẾ TNCN) ---
    if "E350" in wb.sheetnames:
        ws_pit = wb["E350"]
        decl_col = None
        for r in range(1, 10):
            for c in range(1, ws_pit.max_column + 1):
                v = fold(str(ws_pit.cell(r, c).value or "")).strip()
                if "to khai" in v or "phat sinh" in v:
                    decl_col = c
                    break
            if decl_col:
                break
        if decl_col is None:
            decl_col = 2

        pit_rows: dict[int, int] = {}
        for r in range(1, ws_pit.max_row + 1):
            v = str(ws_pit.cell(r, 1).value or "").strip()
            try:
                m = int(float(v))
                if 1 <= m <= 12:
                    pit_rows[m] = r
                    cell = ws_pit.cell(r, decl_col)
                    cell.value = 0
                    cell.number_format = number_format
            except ValueError:
                pass

        for decl in declarations:
            if decl.tax_type == TAX_PIT and decl.month in pit_rows:
                r = pit_rows[decl.month]
                val = decl.values.get("29", 0)
                cell = ws_pit.cell(r, decl_col)
                cell.value = val
                cell.number_format = number_format

    # --- 3. SHEET E370 (THUẾ NHÀ THẦU NTNN) ---
    if "E370" in wb.sheetnames:
        ws_fct = wb["E370"]
        cols_fct: dict[str, int] = {}
        for r in range(1, 20):
            for c in range(1, ws_fct.max_column + 1):
                v = fold(str(ws_fct.cell(r, c).value or "")).replace(" ", "")
                if "9=7x8" in v:
                    cols_fct["09"] = c
                if "13=" in v and "10x11" in v:
                    cols_fct["13"] = c

        fct_rows: dict[int, int] = {}
        for r in range(1, ws_fct.max_row + 1):
            v = str(ws_fct.cell(r, 1).value or "").strip()
            try:
                m = int(float(v))
                if 1 <= m <= 12:
                    fct_rows[m] = r
                    for c in cols_fct.values():
                        cell = ws_fct.cell(r, c)
                        cell.value = 0
                        cell.number_format = number_format
            except ValueError:
                pass

        for decl in declarations:
            if decl.tax_type == TAX_FCT and decl.month in fct_rows:
                r = fct_rows[decl.month]
                for code in ("09", "13"):
                    if code in cols_fct:
                        cell = ws_fct.cell(r, cols_fct[code])
                        existing = cell.value or 0
                        cell.value = existing + decl.values.get(code, 0)
                        cell.number_format = number_format

    # --- 4. SHEET E382 (QUYẾT TOÁN TNDN) ---
    if "E382" in wb.sheetnames:
        ws_cit = wb["E382"]
        cit_targets: dict[str, tuple[int, int]] = {}
        for r in range(1, ws_cit.max_row + 1):
            for c in range(1, ws_cit.max_column + 1):
                v = str(ws_cit.cell(r, c).value or "").strip().upper().replace("[", "").replace("]", "")
                if v in CIT_INDICATORS:
                    cit_targets[v] = (r, c + 1)
                    cell = ws_cit.cell(r, c + 1)
                    cell.value = 0
                    cell.number_format = number_format

        for decl in declarations:
            if decl.tax_type == TAX_CIT:
                for code, val in decl.values.items():
                    code_up = code.upper()
                    if code_up in cit_targets:
                        r, c = cit_targets[code_up]
                        cell = ws_cit.cell(r, c)
                        cell.value = val
                        cell.number_format = number_format

    # --- 5. TẠO HOẶC CẬP NHẬT SHEET NHẬT KÝ NHẬP ---
    log_sheet_name = "NHẬT KÝ NHẬP PDF-XML"
    if log_sheet_name in wb.sheetnames:
        ws_log = wb[log_sheet_name]
        wb.remove(ws_log)
    ws_log = wb.create_sheet(title=log_sheet_name, index=0)

    header_font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_align = Alignment(horizontal="left", vertical="center")
    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )

    headers = ["Năm", "Tháng/Kỳ", "Loại tờ khai", "Lần nộp", "Định dạng", "Tên tệp tờ khai", "Chỉ tiêu thiếu", "Ghi chú"]
    for col_idx, h in enumerate(headers, 1):
        cell = ws_log.cell(1, col_idx, h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center_align
    ws_log.row_dimensions[1].height = 26

    row_idx = 2
    for decl in declarations:
        period_str = f"Quý {decl.month // 3}" if decl.period_type == "quarter" else f"Tháng {decl.month:02d}"
        if decl.period_type == "year":
            period_str = "Cả năm"
        rev_str = "Lần đầu" if decl.revision == 0 else f"Bổ sung lần {decl.revision}"

        ws_log.cell(row_idx, 1, decl.year).alignment = center_align
        ws_log.cell(row_idx, 2, period_str).alignment = center_align
        ws_log.cell(row_idx, 3, decl.tax_type).alignment = center_align
        ws_log.cell(row_idx, 4, rev_str).alignment = center_align
        ws_log.cell(row_idx, 5, decl.source_format).alignment = center_align
        ws_log.cell(row_idx, 6, decl.path.name).alignment = left_align
        ws_log.cell(row_idx, 7, ", ".join(decl.missing) if decl.missing else "Đầy đủ").alignment = center_align
        ws_log.cell(row_idx, 8, decl.note or "Thành công").alignment = left_align

        for c in range(1, 9):
            ws_log.cell(row_idx, c).border = thin_border
            ws_log.cell(row_idx, c).font = Font(name="Segoe UI", size=9.5)
        ws_log.row_dimensions[row_idx].height = 20
        row_idx += 1

    col_widths = [8, 14, 16, 16, 12, 45, 20, 30]
    for i, w in enumerate(col_widths, 1):
        col_letter = openpyxl.utils.get_column_letter(i)
        ws_log.column_dimensions[col_letter].width = w

    wb.save(str(output_path))
    logger(f"✔ Đã ghi thành công kết quả ra file Excel: '{output_path.name}'", "success")
    return output_path


VAT_COLUMN_LABELS = {
    "22": "[22] Thuế GTGT còn được KT kỳ trước chuyển sang",
    "24": "[24] Giá trị HHDV mua vào",
    "24a": "[24a] Giá trị HHDV mua vào nhập khẩu",
    "25": "[25] Thuế GTGT HHDV mua vào được khấu trừ",
    "26": "[26] HHDV bán ra không chịu thuế GTGT",
    "29": "[29] HHDV bán ra chịu thuế suất 0%",
    "30": "[30] HHDV bán ra chịu thuế suất 5%",
    "31": "[31] Thuế GTGT HHDV bán ra chịu thuế suất 5%",
    "32": "[32] HHDV bán ra chịu thuế suất 10%",
    "32a": "[32a] HHDV bán ra không tính thuế",
    "33": "[33] Thuế GTGT HHDV bán ra chịu thuế suất 10%",
    "37": "[37] Điều chỉnh giảm thuế GTGT còn được KT kỳ trước",
    "38": "[38] Điều chỉnh tăng thuế GTGT còn được KT kỳ trước",
    "42": "[42] Thuế GTGT đề nghị hoàn",
    "43": "[43] Thuế GTGT còn được KT chuyển kỳ sau",
}

CIT_LABELS = {
    "A1": "Tổng lợi nhuận kế toán trước thuế TNDN",
    "B1": "Điều chỉnh tăng tổng lợi nhuận trước thuế TNDN",
    "B8": "Chi phí không được trừ khi xác định thu nhập chịu thuế",
    "B13": "Các khoản điều chỉnh làm giảm lợi nhuận trước thuế",
    "C3": "Thu nhập tính thuế TNDN",
    "C4": "Thu nhập tính thuế áp dụng thuế suất 20%",
    "C9": "Thuế TNDN từ hoạt động SXKD tính theo thuế suất phổ thông",
    "C10": "Thuế TNDN được miễn, giảm theo quy định",
    "E": "Số thuế TNDN phải nộp trong kỳ",
}


def generate_summary_dataframes(declarations: list[Declaration]) -> dict[str, pd.DataFrame]:
    """Tạo các DataFrame thống kê trực quan kèm đầy đủ tên nội dung chỉ tiêu để hiển thị trên giao diện Web."""
    dfs: dict[str, pd.DataFrame] = {}

    # 1. Bảng GTGT (01/GTGT)
    vat_order = ("22", "24", "24a", "25", "26", "29", "30", "31", "32", "32a", "33", "37", "38", "42", "43")
    vat_records = []
    for m in range(1, 13):
        row = {"Tháng": f"Tháng {m:02d}"}
        matched = [d for d in declarations if d.tax_type == TAX_VAT and d.month == m]
        if matched:
            d = matched[0]
            for code in vat_order:
                col_name = VAT_COLUMN_LABELS.get(code, f"[{code}]")
                row[col_name] = d.values.get(code, 0)
        else:
            for code in vat_order:
                col_name = VAT_COLUMN_LABELS.get(code, f"[{code}]")
                row[col_name] = 0
        vat_records.append(row)
    dfs["VAT"] = pd.DataFrame(vat_records)

    # 2. Bảng TNCN (05/KK-TNCN)
    pit_records = []
    for m in range(1, 13):
        matched = [d for d in declarations if d.tax_type == TAX_PIT and d.month == m]
        val = matched[0].values.get("29", 0) if matched else 0
        pit_records.append({
            "Tháng": f"Tháng {m:02d}",
            "[29] Tổng số thuế TNCN đã khấu trừ trong kỳ": val,
        })
    dfs["PIT"] = pd.DataFrame(pit_records)

    # 3. Bảng TNDN (03/TNDN)
    cit_records = []
    matched_cit = [d for d in declarations if d.tax_type == TAX_CIT]
    if matched_cit:
        d = matched_cit[0]
        for code in CIT_INDICATORS:
            cit_records.append({
                "Mã chỉ tiêu": f"[{code}]",
                "Nội dung chỉ tiêu": CIT_LABELS.get(code, ""),
                "Số tiền (VNĐ)": d.values.get(code, 0),
            })
    else:
        for code in CIT_INDICATORS:
            cit_records.append({
                "Mã chỉ tiêu": f"[{code}]",
                "Nội dung chỉ tiêu": CIT_LABELS.get(code, ""),
                "Số tiền (VNĐ)": 0,
            })
    dfs["CIT"] = pd.DataFrame(cit_records)

    # 4. Bảng NTNN (01/NTNN)
    fct_records = []
    for m in range(1, 13):
        matched = [d for d in declarations if d.tax_type == TAX_FCT and d.month == m]
        c09 = sum(d.values.get("09", 0) for d in matched)
        c13 = sum(d.values.get("13", 0) for d in matched)
        fct_records.append({
            "Tháng": f"Tháng {m:02d}",
            "[09] Thuế GTGT nhà thầu phải nộp": c09,
            "[13] Thuế TNDN nhà thầu phải nộp": c13,
            "Tổng cộng số thuế nhà thầu phải nộp ([09] + [13])": c09 + c13,
        })
    dfs["FCT"] = pd.DataFrame(fct_records)

    return dfs

# -*- coding: utf-8 -*-
"""
tax_engine.py - Bộ máy trích xuất tờ khai thuế & xuất Excel độc lập đa nền tảng.
Chạy thuần túy bằng Python (OpenPyXL, PyPDF, pdfplumber, xml.etree).
Tương thích 100% trên cả Windows, Linux, macOS và Streamlit Cloud!
"""

from __future__ import annotations

import datetime
import os
import re
import shutil
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
TAX_VAT = "01/GTGT"
TAX_PIT = "05/KK-TNCN"
TAX_PIT_02 = "02/KK-TNCN"
TAX_CIT = "03/TNDN"
TAX_FCT = "01/NTNN"

# Các chỉ tiêu mẫu biểu
E343_INDICATORS = (
    "26", "32a", "29", "30", "32", "33", "31",
    "24", "24a", "25", "38", "37", "42", "43", "22"
)
VAT_INDICATORS = E343_INDICATORS

PIT_INDICATORS = (
    "21", "22", "27", "28", "30", "31", "32",
    "33", "34", "35", "36", "29"
)

CIT_INDICATORS = (
    "A1", "B1", "B8", "B13", "C3", "C4", "C9", "C10", "E"
)

FCT_INDICATORS = ("09", "13")

IMAGE_SUFFIXES = {".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".webp"}
ARCHIVE_SUFFIXES = {".zip", ".rar"}


@dataclass
class Declaration:
    path: Path
    year: int
    month: int
    revision: int
    values: dict[str, int]
    missing: list[str]
    note: str = ""
    period_type: str = "month"
    tax_type: str = TAX_VAT
    modified: float = 0.0


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


def extract_period(text: str, filename: str) -> tuple[int, int, str] | None:
    source = fold(text + "\n" + filename)
    annual_patterns = [
        r"ky\s*tinh\s*thue.{0,100}?nam\s*(20\d{2})",
        r"nam\s*(20\d{2})\s*(?:tu\s*\d{1,2}/\d{1,2}/\d{4}\s*den|\(tu|tu\s*01/01/)",
    ]
    for pattern in annual_patterns:
        match = re.search(pattern, source)
        if match:
            year = int(match.group(1))
            return 12, year, "year"

    quarter_patterns = [
        r"quy\s*([1-4])\s*(?:nam|/)?\s*(20\d{2})",
        r"quy\s*([iI]{1,3}|IV)\s*(?:nam|/)?\s*(20\d{2})",
    ]
    for pattern in quarter_patterns:
        match = re.search(pattern, source)
        if match:
            q_str = match.group(1).upper()
            q_map = {"1": 1, "2": 2, "3": 3, "4": 4, "I": 1, "II": 2, "III": 3, "IV": 4}
            quarter = q_map.get(q_str, 1)
            year = int(match.group(2))
            return quarter * 3, year, "quarter"

    month_patterns = [
        r"thang\s*0?(\d{1,2})\s*(?:nam|/)?\s*(20\d{2})",
        r"ky\s*tinh\s*thue.*?0?(\d{1,2})/(20\d{2})",
    ]
    for pattern in month_patterns:
        match = re.search(pattern, source)
        if match:
            month = int(match.group(1))
            year = int(match.group(2))
            if 1 <= month <= 12:
                return month, year, "month"

    fn_match = re.search(r"t0?(\d{1,2})[-_. ]*(20\d{2})", source)
    if fn_match:
        month = int(fn_match.group(1))
        year = int(fn_match.group(2))
        if 1 <= month <= 12:
            return month, year, "month"

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


def detect_tax_type(text: str, filename: str) -> str:
    source = fold(text + "\n" + filename)
    if "01/ntnn" in source or "nha thau" in source:
        return TAX_FCT
    if "03/tndn" in source or "thu nhap doanh nghiep" in source:
        return TAX_CIT
    if "05/kk-tncn" in source or "05/tncn" in source or "02/kk-tncn" in source or "thu nhap ca nhan" in source:
        return TAX_PIT
    return TAX_VAT


def extract_indicators_from_text(text: str, tax_type: str = TAX_VAT) -> tuple[dict[str, int], list[str]]:
    results: dict[str, int] = {}
    missing: list[str] = []

    if tax_type == TAX_VAT:
        target_codes = VAT_INDICATORS
    elif tax_type == TAX_PIT:
        target_codes = PIT_INDICATORS
    elif tax_type == TAX_CIT:
        target_codes = CIT_INDICATORS
    else:
        target_codes = FCT_INDICATORS

    normalized = clean_pdf_text(text)
    lines = normalized.splitlines()

    for code in target_codes:
        found = False
        patterns = [
            rf"\[{re.escape(code)}\]\s*[:=-]?\s*([0-9.,() -]+)",
            rf"\b{re.escape(code)}\b\s*[:=-]?\s*([0-9.,() -]+)",
        ]
        for pattern in patterns:
            for line in lines:
                m = re.search(pattern, line, re.IGNORECASE)
                if m:
                    try:
                        results[code] = parse_amount(m.group(1))
                        found = True
                        break
                    except ValueError:
                        pass
            if found:
                break
        if not found:
            # Gán mặc định bằng 0 nếu không tìm thấy
            results[code] = 0
            missing.append(code)

    return results, missing


def read_pdf_declaration(path: Path) -> Declaration:
    text_parts = []
    with pdfplumber.open(str(path)) as pdf:
        for page in pdf.pages:
            t = page.extract_text()
            if t:
                text_parts.append(t)

    full_text = "\n".join(text_parts)
    if not full_text.strip():
        # Fallback PyPDF
        reader = PdfReader(str(path))
        full_text = "\n".join(page.extract_text() or "" for page in reader.pages)

    if not full_text.strip():
        raise ValueError(f"Không thể đọc nội dung văn bản trong file PDF '{path.name}'.")

    tax_type = detect_tax_type(full_text, path.name)
    period = extract_period(full_text, path.name)
    if not period:
        raise ValueError(f"Không xác định được tháng/quý/năm kê khai trong file '{path.name}'.")

    month, year, period_type = period
    revision = extract_revision(full_text, path.name)
    values, missing = extract_indicators_from_text(full_text, tax_type)

    mtime = path.stat().st_mtime if path.exists() else 0.0
    return Declaration(
        path=path,
        year=year,
        month=month,
        revision=revision,
        values=values,
        missing=missing,
        period_type=period_type,
        tax_type=tax_type,
        modified=mtime,
    )


def read_xml_declaration(path: Path) -> Declaration:
    tree = ET.parse(str(path))
    root = tree.getroot()

    xml_text = " ".join(root.itertext())
    tax_type = detect_tax_type(xml_text, path.name)

    # Tìm kỳ tính thuế trong XML
    month = None
    year = None
    period_type = "month"

    for tag in ("kyKKhai", "kyTinhThue", "kyKTT"):
        elem = root.find(f".//{tag}")
        if elem is not None and elem.text:
            p = extract_period(elem.text, path.name)
            if p:
                month, year, period_type = p
                break

    if not month or not year:
        p = extract_period(xml_text, path.name)
        if p:
            month, year, period_type = p
        else:
            raise ValueError(f"Không xác định được kỳ kê khai trong file XML '{path.name}'.")

    # Lần nộp / bổ sung
    revision = 0
    for tag in ("soLan", "lanBS", "lanThayDoi"):
        elem = root.find(f".//{tag}")
        if elem is not None and elem.text:
            try:
                revision = int(elem.text.strip())
                break
            except ValueError:
                pass

    values: dict[str, int] = {}
    missing: list[str] = []

    if tax_type == TAX_VAT:
        target_codes = VAT_INDICATORS
    elif tax_type == TAX_PIT:
        target_codes = PIT_INDICATORS
    elif tax_type == TAX_CIT:
        target_codes = CIT_INDICATORS
    else:
        target_codes = FCT_INDICATORS

    # Quét tất cả các thẻ XML có thuộc tính hoặc tên khớp với mã chỉ tiêu
    for elem in root.iter():
        text_val = (elem.text or "").strip()
        tag_name = elem.tag
        # Check attributes
        for attr_k, attr_v in elem.attrib.items():
            attr_clean = attr_v.replace("[", "").replace("]", "").strip()
            if attr_clean in target_codes and text_val:
                try:
                    values[attr_clean] = parse_amount(text_val)
                except ValueError:
                    pass

        # Check tag name (ví dụ ct22, ct24, ct25, ct32a...)
        m = re.match(r"^ct(\w+)$", tag_name, re.I)
        if m and text_val:
            code = m.group(1).lower()
            if code in target_codes:
                try:
                    values[code] = parse_amount(text_val)
                except ValueError:
                    pass

    # Bổ sung các chỉ tiêu thiếu
    for code in target_codes:
        if code not in values:
            values[code] = 0
            missing.append(code)

    mtime = path.stat().st_mtime if path.exists() else 0.0
    return Declaration(
        path=path,
        year=year,
        month=month,
        revision=revision,
        values=values,
        missing=missing,
        period_type=period_type,
        tax_type=tax_type,
        modified=mtime,
    )


def extract_archive(archive_path: Path, extract_dir: Path) -> list[Path]:
    """Giải nén tệp zip hoặc rar và trả về danh sách các tệp trích xuất."""
    extracted: list[Path] = []
    suffix = archive_path.suffix.lower()

    if suffix == ".zip":
        with zipfile.ZipFile(str(archive_path), "r") as zf:
            zf.extractall(str(extract_dir))
    elif suffix == ".rar":
        try:
            import rarfile
            with rarfile.RarFile(str(archive_path), "r") as rf:
                rf.extractall(str(extract_dir))
        except Exception:
            # Fallback nếu trên máy tính có WinRAR/UnRAR
            unrar = shutil.which("unrar") or shutil.which("UnRAR.exe") or "C:\\Program Files\\WinRAR\\UnRAR.exe"
            if os.path.exists(unrar):
                import subprocess
                subprocess.run([unrar, "x", "-y", str(archive_path), str(extract_dir)], check=True)

    for root, _, files in os.walk(str(extract_dir)):
        for f in files:
            p = Path(root) / f
            if p.suffix.lower() in {".pdf", ".xml"}:
                extracted.append(p)
    return extracted


def read_declaration_auto(path: Path) -> Declaration:
    ext = path.suffix.lower()
    if ext == ".xml":
        return read_xml_declaration(path)
    elif ext == ".pdf":
        return read_pdf_declaration(path)
    else:
        raise ValueError(f"Định dạng tệp không được hỗ trợ: {path.name}")


def select_latest_declarations(
    sources: list[Path],
    logger: Callable[[str, str], None] | None = None,
    period_mode: str = "all"
) -> tuple[list[Declaration], list[tuple[str, str]]]:
    """Đọc và chọn lọc các tờ khai mới nhất cho từng kỳ tính thuế."""
    if logger is None:
        logger = lambda m, t="info": None

    parsed: list[Declaration] = []
    errors: list[tuple[str, str]] = []

    for src in sources:
        try:
            decl = read_declaration_auto(src)
            # Lọc theo chế độ kỳ
            if period_mode == "quarter" and decl.period_type != "quarter":
                continue
            if period_mode == "month" and decl.period_type != "month":
                continue
            parsed.append(decl)
            logger(f"✔ Đã đọc thành công: {src.name} ({decl.tax_type} - Tháng/Quý {decl.month}/{decl.year} - Lần {decl.revision})", "success")
        except Exception as exc:
            errors.append((src.name, str(exc)))
            logger(f"⚠ Bỏ qua {src.name}: {exc}", "warning")

    # Nhóm theo (tax_type, year, month) và chọn bản mới nhất (revision cao nhất hoặc modified time mới nhất)
    grouped: dict[tuple[str, int, int], Declaration] = {}
    for item in parsed:
        key = (item.tax_type, item.year, item.month)
        if key not in grouped:
            grouped[key] = item
        else:
            curr = grouped[key]
            if item.revision > curr.revision:
                grouped[key] = item
            elif item.revision == curr.revision and item.modified > curr.modified:
                grouped[key] = item

    latest = sorted(grouped.values(), key=lambda x: (x.tax_type, x.year, x.month))
    return latest, errors


def write_to_excel_openpyxl(
    template_path: Path,
    output_path: Path,
    declarations: list[Declaration],
    parse_errors: list[tuple[str, str]],
    logger: Callable[[str, str], None] | None = None
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

        # Ghi giá trị 0 mặc định cho tất cả 180 ô dữ liệu VAT
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
            decl_col = 2  # Cột B mặc định

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

    # Định dạng tiêu đề Nhật ký
    header_font = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_align = Alignment(horizontal="left", vertical="center")
    right_align = Alignment(horizontal="right", vertical="center")
    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )

    headers = ["Năm", "Tháng/Kỳ", "Loại tờ khai", "Lần nộp", "Tên tệp tờ khai", "Chỉ tiêu thiếu", "Ghi chú"]
    for col_idx, h in enumerate(headers, 1):
        cell = ws_log.cell(1, col_idx, h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center_align
    ws_log.row_dimensions[1].height = 26

    # Ghi từng dòng nhật ký
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
        ws_log.cell(row_idx, 5, decl.path.name).alignment = left_align
        ws_log.cell(row_idx, 6, ", ".join(decl.missing) if decl.missing else "Đầy đủ").alignment = center_align
        ws_log.cell(row_idx, 7, decl.note or "Thành công").alignment = left_align

        for c in range(1, 8):
            ws_log.cell(row_idx, c).border = thin_border
            ws_log.cell(row_idx, c).font = Font(name="Segoe UI", size=9.5)
        ws_log.row_dimensions[row_idx].height = 20
        row_idx += 1

    # Tự động căn chỉnh độ rộng cột
    col_widths = [8, 14, 16, 16, 45, 20, 25]
    for i, w in enumerate(col_widths, 1):
        col_letter = openpyxl.utils.get_column_letter(i)
        ws_log.column_dimensions[col_letter].width = w

    wb.save(str(output_path))
    logger(f"✔ Đã ghi thành công kết quả ra file Excel: '{output_path.name}'", "success")
    return output_path


def generate_summary_dataframes(declarations: list[Declaration]) -> dict[str, pd.DataFrame]:
    """Tạo các DataFrame thống kê trực quan để hiển thị trên giao diện Web."""
    dfs: dict[str, pd.DataFrame] = {}

    # 1. Bảng GTGT (01/GTGT)
    vat_records = []
    for m in range(1, 13):
        row = {"Tháng": f"Tháng {m:02d}"}
        # Tìm tờ khai tháng m
        matched = [d for d in declarations if d.tax_type == TAX_VAT and d.month == m]
        if matched:
            d = matched[0]
            for code in ("22", "24", "25", "26", "29", "30", "32", "32a", "33", "37", "38", "42", "43"):
                row[f"[{code}]"] = d.values.get(code, 0)
        else:
            for code in ("22", "24", "25", "26", "29", "30", "32", "32a", "33", "37", "38", "42", "43"):
                row[f"[{code}]"] = 0
        vat_records.append(row)
    dfs["VAT"] = pd.DataFrame(vat_records)

    # 2. Bảng TNCN (05/KK-TNCN)
    pit_records = []
    for m in range(1, 13):
        matched = [d for d in declarations if d.tax_type == TAX_PIT and d.month == m]
        val = matched[0].values.get("29", 0) if matched else 0
        pit_records.append({"Tháng": f"Tháng {m:02d}", "Thuế TNCN Phát Sinh [29]": val})
    dfs["PIT"] = pd.DataFrame(pit_records)

    # 3. Bảng TNDN (03/TNDN)
    cit_records = []
    matched_cit = [d for d in declarations if d.tax_type == TAX_CIT]
    if matched_cit:
        d = matched_cit[0]
        for code in CIT_INDICATORS:
            cit_records.append({"Chỉ tiêu": f"[{code}]", "Số tiền (VNĐ)": d.values.get(code, 0)})
    else:
        for code in CIT_INDICATORS:
            cit_records.append({"Chỉ tiêu": f"[{code}]", "Số tiền (VNĐ)": 0})
    dfs["CIT"] = pd.DataFrame(cit_records)

    # 4. Bảng NTNN (01/NTNN)
    fct_records = []
    for m in range(1, 13):
        matched = [d for d in declarations if d.tax_type == TAX_FCT and d.month == m]
        c09 = sum(d.values.get("09", 0) for d in matched)
        c13 = sum(d.values.get("13", 0) for d in matched)
        fct_records.append({"Tháng": f"Tháng {m:02d}", "Chỉ tiêu 09 (GTGT)": c09, "Chỉ tiêu 13 (TNDN)": c13, "Tổng cộng": c09 + c13})
    dfs["FCT"] = pd.DataFrame(fct_records)

    return dfs

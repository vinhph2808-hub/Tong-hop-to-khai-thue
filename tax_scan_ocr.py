"""Offline OCR for printed tax return documents (PDF scans, images, photos).

Supports 01/GTGT, 05/KK-TNCN, and 01/NTNN.
Reconstructs spatial text lines from bounding boxes and extracts numeric indicators.
"""
from __future__ import annotations

import re
import unicodedata
from pathlib import Path

_engine = None

IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".webp"}


def _normalize_text(s: str) -> str:
    n = unicodedata.normalize("NFD", s)
    return "".join(c for c in n if unicodedata.category(c) != "Mn").lower()


def _result_to_words(result) -> list[dict]:
    words = []
    for box, text, confidence in result or []:
        words.append({
            "text": text,
            "confidence": float(confidence),
            "left": min(p[0] for p in box),
            "right": max(p[0] for p in box),
            "top": min(p[1] for p in box),
            "bottom": max(p[1] for p in box),
        })
    return words


def _evaluate_orientation(words: list[dict], page_height: float) -> float:
    if not words:
        return -9999.0
    horiz_count = sum(1 for w in words if (w["right"] - w["left"]) >= (w["bottom"] - w["top"]))
    score = (horiz_count / len(words)) * 100.0

    header_kw = ("cong hoa", "viet nam", "to khai", "tong cuc", "chi cuc", "bo tai chinh", "ma so thue")

    for w in words:
        txt = _normalize_text(w["text"])
        mid_y = (w["top"] + w["bottom"]) / 2
        for h in header_kw:
            if h in txt:
                if mid_y < page_height * 0.4:
                    score += 50.0
                elif mid_y > page_height * 0.6:
                    score -= 50.0
        # Số trang ví dụ "1/2", "2/2", "Trang 1": luôn ở nửa dưới trang
        if re.search(r"\b\d+/\d+\b", txt):
            if mid_y > page_height * 0.6:
                score += 50.0
            elif mid_y < page_height * 0.3:
                score -= 50.0

    # Nếu có chỉ tiêu bảng và chữ ký thì bảng phải luôn nằm trên chữ ký
    table_indicators = [w for w in words if re.search(r"\[\s*(?:41|42|43|21|22|23|24|25)\s*\]", w["text"])]
    signatures = [w for w in words if any(f in _normalize_text(w["text"]) for f in ("nguoi nop thue", "nhan vien dai ly", "ky, ghi ro ho ten", "wang"))]
    if table_indicators and signatures:
        min_table_y = min(w["top"] for w in table_indicators)
        max_sig_y = max(w["bottom"] for w in signatures)
        if min_table_y < max_sig_y:
            score += 60.0
        else:
            score -= 60.0

    # Thứ tự chỉ tiêu: [41] nằm trên [43]
    y_41 = [w["top"] for w in words if "[41]" in w["text"]]
    y_43 = [w["top"] for w in words if "[43]" in w["text"]]
    if y_41 and y_43:
        if y_41[0] < y_43[0]:
            score += 50.0
        else:
            score -= 50.0

    return score


def _ocr_page_pixels(pixels, preferred_rotation: int | None = None) -> tuple[list[dict], int | None]:
    global _engine
    import numpy as np

    cleaned = pixels.copy()
    channels = cleaned.astype(np.int16)
    colored = ((channels.max(2) - channels.min(2)) > 25) & (channels.max(2) > 80)
    cleaned[colored] = 255

    # Nếu tài liệu có hướng xoay đã xác định từ trang trước, áp dụng ngay
    if preferred_rotation:
        rot_pixels = np.rot90(cleaned, preferred_rotation)
        result, _ = _engine(rot_pixels[:, :, ::-1])
        return _result_to_words(result), preferred_rotation

    result, _ = _engine(cleaned[:, :, ::-1])
    words = _result_to_words(result)
    best_k = 0

    if words:
        horiz_count = sum(1 for w in words if (w["right"] - w["left"]) >= (w["bottom"] - w["top"]))
        horiz_ratio = horiz_count / len(words)
        # Nếu chữ chủ yếu theo chiều dọc (xoay 90 độ ngang/dọc) hoặc số lượng từ ngang thấp,
        # tự động thử xoay trang về hướng đúng để các dòng văn bản không bị lệch cột.
        if horiz_ratio < 0.65:
            best_words = words
            best_score = _evaluate_orientation(words, cleaned.shape[0])
            for k in (1, 3):
                rot_pixels = np.rot90(cleaned, k)
                r_res, _ = _engine(rot_pixels[:, :, ::-1])
                r_words = _result_to_words(r_res)
                score = _evaluate_orientation(r_words, rot_pixels.shape[0])
                if score > best_score:
                    best_score = score
                    best_words = r_words
                    best_k = k
                if score >= 150:
                    best_k = k
                    break
            words = best_words

    return words, best_k


def scan_pages(path: Path):
    global _engine
    try:
        import numpy as np
        from PIL import Image
        import pypdfium2 as pdfium
        from rapidocr_onnxruntime import RapidOCR
    except ImportError as exc:
        raise ValueError("Thiếu bộ OCR. Vui lòng dùng đầy đủ bản tool có hỗ trợ scan.") from exc
    if _engine is None:
        _engine = RapidOCR()
    pages = []
    suffix = path.suffix.lower()

    # Xử lý tệp hình ảnh trực tiếp (PNG, JPG, JPEG, BMP, TIFF, WEBP)
    if suffix in IMAGE_EXTENSIONS:
        try:
            with Image.open(str(path)) as img:
                rgb_img = img.convert("RGB")
                pixels = np.array(rgb_img)
                words, _ = _ocr_page_pixels(pixels)
                pages.append(words)
        except Exception as exc:
            raise ValueError(f"Không thể mở hoặc đọc file ảnh: {exc}") from exc
        return pages

    # Xử lý tệp PDF scan
    doc_rotation = None
    with pdfium.PdfDocument(str(path)) as document:
        for page in document:
            try:
                bitmap = page.render(scale=3)
                try:
                    pixels = np.array(bitmap.to_pil().convert("RGB"))
                finally:
                    bitmap.close()
                words, page_rot = _ocr_page_pixels(pixels, preferred_rotation=doc_rotation)
                if doc_rotation is None and page_rot:
                    doc_rotation = page_rot
                pages.append(words)
            finally:
                page.close()
    return pages


def words_to_lines(words: list[dict]) -> str:
    """Ghép các từ OCR thành từng dòng văn bản theo tọa độ không gian."""
    if not words:
        return ""
    sorted_words = sorted(words, key=lambda w: (w["top"], w["left"]))
    lines = []
    current_line = []
    for w in sorted_words:
        if not current_line:
            current_line.append(w)
            continue
        line_mid_y = sum((item["top"] + item["bottom"]) / 2 for item in current_line) / len(current_line)
        line_h = max(item["bottom"] - item["top"] for item in current_line)
        w_mid_y = (w["top"] + w["bottom"]) / 2
        if abs(w_mid_y - line_mid_y) <= line_h * 0.65:
            current_line.append(w)
        else:
            current_line.sort(key=lambda item: item["left"])
            lines.append(" ".join(item["text"] for item in current_line))
            current_line = [w]
    if current_line:
        current_line.sort(key=lambda item: item["left"])
        lines.append(" ".join(item["text"] for item in current_line))
    return "\n".join(lines)


def ocr_pages_to_text(pages: list[list[dict]]) -> str:
    """Chuyển đổi toàn bộ các trang OCR thành văn bản có cấu trúc dòng hoàn chỉnh."""
    return "\n".join(words_to_lines(page) for page in pages)


def printed_pit_amounts(pages):
    amounts = {}
    amount_pattern = re.compile(r"(?:\d{1,3}(?:\.\d{3})+|\d{1,3}(?:,\d{3})+|\d+)")
    for words in pages:
        for code in ("29", "30", "31"):
            markers = [w for w in words if re.fullmatch(r"\[\s*" + code + r"\s*\]", w["text"])]
            for marker in markers:
                mid_y = (marker["top"] + marker["bottom"]) / 2
                tolerance = (marker["bottom"] - marker["top"]) * 0.65
                row = sorted((w for w in words
                              if w["left"] > marker["right"]
                              and abs((w["top"] + w["bottom"]) / 2 - mid_y) <= tolerance),
                             key=lambda w: w["left"])
                if not row or row[0]["text"].strip().upper() not in ("VND", "VNĐ"):
                    continue
                # More than one number on a row may be a handwritten correction.
                candidates = [w for w in row[1:] if amount_pattern.fullmatch(w["text"].strip())]
                if len(candidates) != 1 or min(marker["confidence"], row[0]["confidence"], candidates[0]["confidence"]) < 0.9:
                    raise ValueError(f"OCR: chưa đọc chắc chắn số in tại [{code}]; cần đối chiếu PDF.")
                if code in amounts:
                    raise ValueError(f"OCR: có nhiều dòng [{code}]; chưa thể chọn một tờ khai duy nhất.")
                amounts[code] = int(re.sub(r"[.,]", "", candidates[0]["text"].strip()))
    if set(amounts) != {"29", "30", "31"}:
        raise ValueError("OCR: thiếu số in tại [29], [30] hoặc [31]; không tự thay số chưa đọc được bằng 0.")
    if amounts["29"] != amounts["30"] + amounts["31"]:
        raise ValueError("OCR: số in [29] không khớp [30] + [31]; cần đối chiếu PDF.")
    return amounts

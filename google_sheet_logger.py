# -*- coding: utf-8 -*-
"""
google_sheet_logger.py - Module tự động ghi nhật ký các máy tính đăng nhập
và sử dụng hệ thống vào Google Sheets qua Google Apps Script Webhook.
Chạy thuần thư viện chuẩn (urllib, json, threading) - Hoàn toàn độc lập,
không phụ thuộc thư viện bên ngoài và không làm chậm ứng dụng.
"""

from __future__ import annotations

import datetime
import json
import os
import platform
import socket
import threading
import urllib.parse
import urllib.request
from pathlib import Path
from typing import Any

CONFIG_FILE = Path(__file__).resolve().parent / "sheet_config.json"


def get_configured_sheet_url() -> str:
    """Lấy đường dẫn Google Apps Script Webhook từ cấu hình."""
    # 1. Kiểm tra biến môi trường
    env_url = os.environ.get("GOOGLE_SHEET_URL", "").strip()
    if env_url:
        return env_url

    # 2. Kiểm tra Streamlit secrets (nếu đang chạy trên Streamlit)
    try:
        import streamlit as st
        if hasattr(st, "secrets") and "GOOGLE_SHEET_URL" in st.secrets:
            sec_url = str(st.secrets["GOOGLE_SHEET_URL"]).strip()
            if sec_url:
                return sec_url
    except Exception:
        pass

    # 3. Kiểm tra file cấu hình cục bộ sheet_config.json
    if CONFIG_FILE.exists():
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                return str(data.get("sheet_url", "")).strip()
        except Exception:
            pass

    return ""


def save_configured_sheet_url(url: str) -> None:
    """Lưu đường dẫn Google Apps Script Webhook vào file cấu hình."""
    url = url.strip()
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump({"sheet_url": url}, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def get_client_info_from_streamlit() -> dict[str, str]:
    """Trích xuất thông tin thiết bị và mạng của người dùng từ Streamlit."""
    info = {
        "ip": "127.0.0.1",
        "user_agent": "Trình duyệt Web",
        "hostname": "Khách vãng lai",
        "os_name": "Web Browser",
    }
    try:
        import streamlit as st
        if hasattr(st, "context") and hasattr(st.context, "headers"):
            headers = st.context.headers
            # Lấy địa chỉ IP thực của người dùng qua proxy / cloudflare
            ip = (
                headers.get("x-forwarded-for")
                or headers.get("x-real-ip")
                or headers.get("cf-connecting-ip")
                or headers.get("host")
                or "127.0.0.1"
            )
            # Nếu có nhiều IP phân cách dấu phẩy, lấy IP đầu tiên
            if "," in ip:
                ip = ip.split(",")[0].strip()
            info["ip"] = ip

            ua = headers.get("user-agent", "")
            if ua:
                info["user_agent"] = ua
                # Phân tích sơ bộ hệ điều hành từ User-Agent
                ua_lower = ua.lower()
                if "windows" in ua_lower:
                    info["os_name"] = "Windows PC"
                elif "macintosh" in ua_lower or "mac os" in ua_lower:
                    info["os_name"] = "Macbook / macOS"
                elif "iphone" in ua_lower:
                    info["os_name"] = "Apple iPhone"
                elif "android" in ua_lower:
                    info["os_name"] = "Android Mobile"
                elif "linux" in ua_lower:
                    info["os_name"] = "Linux"

                # Phân tích trình duyệt
                if "edg/" in ua_lower:
                    info["browser"] = "Microsoft Edge"
                elif "chrome/" in ua_lower:
                    info["browser"] = "Google Chrome"
                elif "safari/" in ua_lower:
                    info["browser"] = "Apple Safari"
                elif "firefox/" in ua_lower:
                    info["browser"] = "Mozilla Firefox"
                elif "coc_coc" in ua_lower:
                    info["browser"] = "Cốc Cốc"
    except Exception:
        pass
    return info


def send_log_payload(payload: dict[str, Any], sheet_url: str = "") -> tuple[bool, str]:
    """Gửi dữ liệu nhật ký trực tiếp đến Google Apps Script Webhook."""
    target_url = (sheet_url or get_configured_sheet_url()).strip()
    if not target_url or not target_url.startswith("http"):
        return False, "Chưa cấu hình URL Google Sheet Webhook"

    try:
        data_bytes = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        req = urllib.request.Request(
            target_url,
            data=data_bytes,
            headers={
                "Content-Type": "application/json; charset=utf-8",
                "User-Agent": "Antigravity-TaxTool/1.0",
            },
            method="POST",
        )
        with urllib.request.urlopen(req, timeout=8) as response:
            res_text = response.read().decode("utf-8", errors="replace")
            try:
                res_json = json.loads(res_text)
                if res_json.get("status") in {"SUCCESS", "OK"}:
                    return True, res_json.get("message", "Thành công")
                return True, res_text[:100]
            except Exception:
                return True, res_text[:100]
    except Exception as exc:
        return False, str(exc)


def log_device_async(
    action: str = "Đăng nhập hệ thống",
    platform_type: str = "Web Online (Streamlit)",
    device_id: str = "",
    client_name: str = "",
    ip_address: str = "",
    user_agent: str = "",
    file_count: int = 0,
    note: str = "",
    sheet_url: str = "",
) -> None:
    """Ghi nhật ký ngầm ở luồng nền (Background Thread), hoàn toàn không làm chậm trải nghiệm ứng dụng."""
    def worker():
        payload = {
            "timestamp": datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
            "action": action,
            "platform": platform_type,
            "device_id": device_id or f"WEB-{os.urandom(3).hex().upper()}",
            "client_name": client_name or "Người dùng Web",
            "ip": ip_address or "127.0.0.1",
            "user_agent": user_agent or "Web Client",
            "file_count": file_count,
            "note": note,
        }
        send_log_payload(payload, sheet_url=sheet_url)

    thread = threading.Thread(target=worker, daemon=True)
    thread.start()

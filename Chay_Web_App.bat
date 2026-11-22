@echo off
chcp 65001 >nul
title Hệ Thống Tổng Hợp Tờ Khai Thuế (Bản Web)
echo ======================================================================
echo    HỆ THỐNG XỬ LÝ & TỔNG HỢP SỐ LIỆU TỜ KHAI THUẾ - GIAO DIỆN WEB
echo ======================================================================
echo Đang khởi chạy máy chủ Web... Trình duyệt sẽ tự động mở!
echo Bạn cũng có thể mở thủ công tại địa chỉ: http://localhost:8502
echo ======================================================================
if exist ".venv\Scripts\streamlit.exe" (
    ".venv\Scripts\streamlit.exe" run web_app.py --server.port 8502
) else (
    "d:\Vinh.ph\AI\Codex\Thông báo BHXH\.venv\Scripts\streamlit.exe" run web_app.py --server.port 8502
)
pause

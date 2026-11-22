@echo off
chcp 65001 >nul
title ĐỒNG BỘ CẬP NHẬT LÊN WEBSITE (STREAMLIT CLOUD)
echo ======================================================================
echo       ĐỒNG BỘ TỰ ĐỘNG CẬP NHẬT LÊN WEBSITE (STREAMLIT CLOUD)
echo ======================================================================
echo.
set GIT_EXE="C:\Users\VINHPH\AppData\Local\GitHubDesktop\app-3.6.2\resources\app\git\cmd\git.exe"

echo 1. Đang ghi nhận các thay đổi trên máy tính...
%GIT_EXE% add .

echo 2. Đang tạo bản cập nhật...
%GIT_EXE% commit -m "Cap nhat phien ban moi tu may tinh"

echo 3. Đang đẩy lên GitHub và kích hoạt máy chủ Cloud cập nhật...
%GIT_EXE% push origin main

echo.
echo ======================================================================
echo  THÀNH CÔNG!
echo  Hệ thống Streamlit Cloud đang tự động cập nhật phiên bản mới.
echo  Bạn hãy vào lại link website sau 15 - 30 giây để kiểm tra!
echo ======================================================================
pause

@echo off
title DIEM DANH PRO (FIXED)
cls

echo ==============================================
echo        HE THONG DIEM DANH (VER 2.0)
echo ==============================================

echo [1] Kiem tra va cai dat thu vien...
:: Tìm file thuốc dlib (nếu có) để cài trước
if exist *.whl pip install *.whl

:: Cài các thư viện khác
pip install -r requirements.txt

echo.
echo [2] Dang mo App...
echo Hay dung dien thoai quet QR Code hien ra...

:: Dùng lệnh này để tránh lỗi 'not recognized'
python -m streamlit run app.py

pause
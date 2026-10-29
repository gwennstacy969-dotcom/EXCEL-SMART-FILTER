@echo off
title EXCEL SMART FILTER - Responsive App
cd /d "%~dp0"

echo ===============================================================
echo        MEMBUKA APLIKASI EXCEL SMART FILTER (DESKTOP & MOBILE)
echo ===============================================================
echo.

where node >nul 2>nul
if %errorlevel% equ 0 (
    echo [OK] Menjalankan server aplikasi lokal...
    node server.js
) else (
    echo [OK] Membuka aplikasi langsung di browser...
    start "" "%~dp0index.html"
)
pause


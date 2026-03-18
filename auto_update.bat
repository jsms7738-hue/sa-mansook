@echo off
chcp 65001 > nul
set BASE_DIR=c:\Users\yoonh\Desktop\AI\PBA 생산 결과
cd /d "%BASE_DIR%"

echo [%date% %time%] Starting Auto Update...

echo [1/3] Extracting Data...
python extract_data_pba.py
if %errorlevel% neq 0 (
    echo [ERROR] Data extraction failed.
    exit /b 1
)

echo [2/3] Building Dashboards...
python build_dashboard_pba.py
if %errorlevel% neq 0 (
    echo [ERROR] Dashboard build failed.
    exit /b 1
)

echo [3/3] Uploading to GitHub...
git add .
git commit -m "Auto-update: %date% %time%"
git push origin master
if %errorlevel% neq 0 (
    echo [WARNING] GitHub push failed. Manual check required.
)

echo [%date% %time%] Auto Update Completed.

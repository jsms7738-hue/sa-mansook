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

echo [2/4] Building PBA Dashboards...
python build_dashboard_pba.py
if %errorlevel% neq 0 (
    echo [ERROR] PBA Dashboard build failed.
    exit /b 1
)

echo [3/4] Updating ICT Dashboard...
pushd "c:\Users\yoonh\Desktop\AI\ICT데이타"
python extract_data.py
python build_dashboard.py
popd

echo [4/4] Uploading to GitHub...
git add .
git commit -m "Auto-update: %date% %time%"
git push origin master
if %errorlevel% neq 0 (
    echo [WARNING] GitHub push failed. Manual check required.
)

echo [%date% %time%] Auto Update Completed.

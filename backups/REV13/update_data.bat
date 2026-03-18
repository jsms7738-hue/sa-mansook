@echo off
chcp 65001 > nul
echo ============================================
echo  PBA 생산 결과 데이터 업데이트 시작
echo ============================================
echo.

cd /d "c:\Users\yoonh\Desktop\AI\PBA 생산 결과"

echo [1/2] Excel 데이터 읽는 중...
python extract_data_pba.py
if %errorlevel% neq 0 (
    echo.
    echo [오류] 데이터 추출에 실패했습니다.
    echo Python이 설치되어 있는지, 필요한 패키지가 설치되었는지 확인하세요.
    echo 필요 패키지: pip install pandas openpyxl
    pause
    exit /b 1
)

echo.
echo [2/2] 대시보드 HTML 파일 생성 중...
python build_dashboard_pba.py
if %errorlevel% neq 0 (
    echo.
    echo [오류] 대시보드 생성에 실패했습니다.
    pause
    exit /b 1
)

echo.
echo ============================================
echo  업데이트 완료! 브라우저에서 새로고침하세요.
echo ============================================
echo.
timeout /t 3

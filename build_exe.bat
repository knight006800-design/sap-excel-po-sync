@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo [1/3] 의존성 설치...
python -m pip install -r requirements.txt
if errorlevel 1 goto :err
echo [2/3] PyInstaller 빌드...
python -m PyInstaller --noconfirm --clean --windowed --name "구매오더수량동기화" --add-data "config.json;." main.py
if errorlevel 1 goto :err
echo [3/3] 배포 폴더 정리...
if not exist "배포" mkdir "배포"
xcopy /E /I /Y "dist\구매오더수량동기화" "배포\구매오더수량동기화\"
copy /Y "config.json" "배포\구매오더수량동기화\"
copy /Y "사용설명서.txt" "배포\구매오더수량동기화\"
copy /Y "README.md" "배포\구매오더수량동기화\" 2>nul
echo.
echo 완료: 배포\구매오더수량동기화\
echo zip으로 묶어서 다른 PC에 전달하세요. Tesseract OCR 설치도 필요합니다.
pause
exit /b 0
:err
echo 빌드 실패
pause
exit /b 1

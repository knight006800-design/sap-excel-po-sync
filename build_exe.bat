@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo [1/3] 의존성 설치...
py -3.7 -m pip install -r requirements.txt
if errorlevel 1 goto :err
echo [2/3] PyInstaller 빌드 (웅이전용.exe)...
py -3.7 -m PyInstaller --noconfirm --clean --windowed --name "웅이전용" --add-data "config.json;." main.py
if errorlevel 1 goto :err
echo [3/3] 배포 폴더 정리...
if not exist "배포" mkdir "배포"
if exist "배포\웅이전용" rmdir /s /q "배포\웅이전용"
xcopy /E /I /Y "dist\웅이전용" "배포\웅이전용\"
copy /Y "config.json" "배포\웅이전용\"
copy /Y "사용설명서.txt" "배포\웅이전용\"
powershell -NoProfile -Command "Compress-Archive -Path '배포\웅이전용\*' -DestinationPath 'woongyi-windows.zip' -Force"
echo.
echo 완료: 배포\웅이전용\웅이전용.exe
echo zip: woongyi-windows.zip
pause
exit /b 0
:err
echo 빌드 실패
pause
exit /b 1

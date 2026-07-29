@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo [0/3] 안내 그림(빨간 점) 생성...
py -3.7 make_guide_images.py
echo [1/3] 의존성 설치 (pywin32 포함)...
py -3.7 -m pip install -r requirements.txt
py -3.7 -m pip install --upgrade pywin32
if errorlevel 1 goto :err
echo [2/3] PyInstaller 단일 exe 빌드 (폴더 없음)...
py -3.7 -m PyInstaller --noconfirm --clean --windowed --onefile --name "웅이전용" --add-data "config.json;." --add-data "guide_images;guide_images" --hidden-import=win32api --hidden-import=win32con --hidden-import=win32gui --hidden-import=win32timezone --hidden-import=pythoncom --hidden-import=pywintypes --hidden-import=win32com --hidden-import=win32com.client --hidden-import=win32com.client.gencache --collect-all win32com --collect-binaries pywintypes --collect-binaries pythoncom main.py
if errorlevel 1 goto :err
echo [3/3] 배포 zip (exe를 폴더 없이)...
if not exist "배포" mkdir "배포"
copy /Y "dist\웅이전용.exe" "배포\웅이전용.exe"
copy /Y "사용설명서.txt" "배포\사용설명서.txt"
copy /Y "config.json" "배포\config.json"
powershell -NoProfile -Command "if (Test-Path 'woongyi-windows.zip') { Remove-Item 'woongyi-windows.zip' -Force -ErrorAction SilentlyContinue }; Compress-Archive -Path '배포\웅이전용.exe','배포\사용설명서.txt','배포\config.json' -DestinationPath 'woongyi-windows.zip' -Force"
echo.
echo 완료: 배포\웅이전용.exe
echo zip: woongyi-windows.zip
pause
exit /b 0
:err
echo 빌드 실패
pause
exit /b 1

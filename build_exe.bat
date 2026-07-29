@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo [0/4] 안내 그림 복사...
py -3.7 copy_guides.py
echo [1/4] 의존성 설치...
py -3.7 -m pip install -r requirements.txt
if errorlevel 1 goto :err
echo [2/4] PyInstaller 빌드 (웅이전용.exe)...
py -3.7 -m PyInstaller --noconfirm --clean --windowed --name "웅이전용" --add-data "config.json;." --add-data "guide_images;guide_images" main.py
if errorlevel 1 goto :err
echo [3/4] 배포 폴더 정리...
if not exist "배포" mkdir "배포"
if exist "배포\웅이전용" rmdir /s /q "배포\웅이전용"
xcopy /E /I /Y "dist\웅이전용" "배포\웅이전용\"
copy /Y "config.json" "배포\웅이전용\"
copy /Y "사용설명서.txt" "배포\웅이전용\"
if exist "guide_images" xcopy /E /I /Y "guide_images" "배포\웅이전용\guide_images\"
echo [4/4] zip 생성...
powershell -NoProfile -Command "Compress-Archive -Path '배포\웅이전용\*' -DestinationPath '웅이전용-windows.zip' -Force"
echo.
echo 완료: 배포\웅이전용\웅이전용.exe
echo zip: 웅이전용-windows.zip
pause
exit /b 0
:err
echo 빌드 실패
pause
exit /b 1

# SAP–엑셀 구매오더 수량 동기화 RPA

엑셀 **B열 자재코드**와 SAP 화면의 자재코드를 맞춰, 엑셀 **D열 수량**을 SAP **오더수량**에 자동 입력합니다.

- SAP에 없는 코드 → 엑셀 B열 음영
- 수량 일치 → 엑셀 D열 노란색
- 원격 디버깅 로그: `구동점검.txt`

## 필요 환경

- Windows 10/11, 듀얼모니터 (좌측 SAP, 1920×1080 권장)
- Microsoft Excel
- [Tesseract OCR](https://github.com/UB-Mannheim/tesseract/wiki)
- Python 3.7+ (소스 실행 시) 또는 빌드된 exe

## 사용 방법

1. SAP「구매오더 생성」을 좌측에 자재코드가 **전부 보이게** 띄움
2. `python main.py` 또는 `구매오더수량동기화.exe` 실행
3. 엑셀 파일 연결 → **SAP 화면 보정** → **동기화 실행**
4. 오류 시 `구동점검.txt` / `구동점검_캡처_*.png` 확인

상세: [사용설명서.txt](사용설명서.txt)

## 개발

```bash
pip install -r requirements.txt
python main.py
build_exe.bat
```

# 웅이전용 — SAP–엑셀 구매오더 수량 동기화

## 다른 PC (Python 없음)

1. [Releases](https://github.com/knight006800-design/sap-excel-po-sync/releases) 에서 **웅이전용-windows.zip** 다운로드  
2. 압축 해제 후 **`웅이전용.exe`** 실행  
3. **Microsoft Excel** + [Tesseract OCR](https://github.com/UB-Mannheim/tesseract/wiki) 만 있으면 됩니다

## 기능

- 엑셀 B열 자재코드 ↔ SAP 자재코드 매칭
- 엑셀 D열 수량 → SAP 오더수량 입력
- SAP에 없으면 엑셀 B열 음영 / 수량 같으면 D열 노란색
- 보정 시 **안내 그림** 표시
- 로그는 **`구동점검` 폴더**에 모음

## 개발

```bash
pip install -r requirements.txt
python copy_guides.py
python main.py
build_exe.bat
```

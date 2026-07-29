# SAP–엑셀 구매오더 수량 동기화 RPA

엑셀 **B열 자재코드**와 SAP 화면의 자재코드를 맞춰, 엑셀 **D열 수량**을 SAP **오더수량**에 자동 입력합니다.

## 다른 PC (Python 없음) — exe만 받기

1. [Releases](https://github.com/knight006800-design/sap-excel-po-sync/releases) 에서 **PurchaseOrderSync-windows.zip** 다운로드
2. 압축 해제 후 `PurchaseOrderSync.exe` 실행
3. 그 PC에 **Microsoft Excel** 과 [Tesseract OCR](https://github.com/UB-Mannheim/tesseract/wiki) 만 설치되어 있으면 됩니다 (Python 불필요)

## 기능

- SAP에 없는 코드 → 엑셀 B열 음영
- 수량 일치 → 엑셀 D열 노란색
- 원격 디버깅 로그: `구동점검.txt`

## 사용 방법

1. SAP「구매오더 생성」을 좌측에 자재코드가 **전부 보이게** 띄움
2. exe 실행 → 엑셀 파일 연결 → **SAP 화면 보정** → **동기화 실행**
3. 오류 시 `구동점검.txt` / `구동점검_캡처_*.png` 확인

상세: [사용설명서.txt](사용설명서.txt)

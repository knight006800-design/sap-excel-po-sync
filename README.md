# 웅이전용 — SAP–엑셀 구매오더 수량 비교

## 사용
1. 엑셀 연결
2. SAP 자재코드만 Ctrl+V
3. **비교·결과 생성**
4. 오류 없으면 **결과 수량** 복사 → SAP 오더수량에 Ctrl+V

## 규칙
- SAP 코드 = 엑셀 코드 → 엑셀 수량
- SAP에만 있음 → 엑셀 코드기입 오류/휴먼에러 (수정 후 재작업)
- 엑셀에만 있음 → 지정 파일로 추출 (원본 미수정)

## 다운로드
https://github.com/knight006800-design/sap-excel-po-sync/releases

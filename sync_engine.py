# -*- coding: utf-8 -*-
"""SAP 코드 ↔ 엑셀 비교 → 결과 수량 / 오류 안내 / 미매칭 추출."""
from __future__ import print_function

from difflib import SequenceMatcher

from excel_io import (
    ExcelWorkbook,
    format_qty_text,
    normalize_code,
    parse_pasted_sap_codes,
)


SIMILAR_THRESHOLD = 0.70


def find_similar_codes(
    code, candidates, threshold=SIMILAR_THRESHOLD, limit=1, exclude=None
):
    """유사 엑셀 코드 1건. SAP에 이미 있는 품번은 후보에서 제외."""
    if not code or not candidates:
        return []
    exclude_set = set(exclude or ())
    exclude_set.add(code)
    clen = len(code)
    # 길이 차이가 크면 1글자 휴먼에러 후보가 아님
    max_delta = 2 if clen <= 12 else max(2, clen // 5)
    best_ratio = 0.0
    best_code = None
    for c in candidates:
        if not c or c in exclude_set:
            continue
        if abs(len(c) - clen) > max_delta:
            continue
        sm = SequenceMatcher(None, code, c)
        if sm.quick_ratio() < 0.55:
            continue
        ratio = sm.ratio()
        if ratio < threshold:
            # 길이가 거의 같으면 1글자 오타 하한 완화
            if abs(len(c) - clen) <= 1 and ratio >= 0.60:
                pass
            else:
                continue
        if ratio > best_ratio:
            best_ratio = ratio
            best_code = c
            if best_ratio >= 0.99:
                break
    if best_code is None:
        return []
    return [(best_ratio, best_code)][:limit]


class SyncEngine(object):
    def __init__(self, config, logger, stop_flag=None):
        self.config = config
        self.logger = logger
        self.stop_flag = stop_flag  # 호환용 (UI 중지 제거)

    def run(self, excel_path, codes_text, unmatched_path=None):
        sap_codes = parse_pasted_sap_codes(codes_text)
        if not sap_codes:
            raise RuntimeError(
                u"붙여넣은 SAP 코드가 없습니다. SAP 자재번호 열을 복사해 코드 칸에 Ctrl+V 하세요."
            )

        n = len(sap_codes)
        self.logger.info(u"SAP 자재코드 {0}행 (순서 유지)".format(n))

        book = ExcelWorkbook(excel_path, logger=self.logger)
        book.open()
        try:
            excel_rows = book.load_rows()
            if not excel_rows:
                raise RuntimeError(u"엑셀에 자재코드(B열) 데이터가 없습니다.")

            excel_qty = {}
            for er in excel_rows:
                code = normalize_code(er["code"])
                if not code:
                    continue
                if code not in excel_qty:
                    excel_qty[code] = er["qty"]

            excel_codes = list(excel_qty.keys())
            sap_code_set = set(c for c in sap_codes if c)
            unmatched_excel = [
                er
                for er in excel_rows
                if normalize_code(er["code"]) not in sap_code_set
            ]

            result_rows = []
            matched = 0
            missing_in_excel = []
            no_qty = []
            # 유사도 계산은 오류 코드에만, 동일 코드는 캐시
            similar_cache = {}

            for code in sap_codes:
                if not code:
                    result_rows.append(
                        {
                            "code": u"",
                            "result_qty": None,
                            "result_qty_text": u"",
                            "status": u"",
                            "ok": False,
                            "suggestions": [],
                        }
                    )
                    continue

                if code not in excel_qty:
                    if code in similar_cache:
                        suggestions = similar_cache[code]
                    else:
                        suggestions = find_similar_codes(
                            code, excel_codes, exclude=sap_code_set
                        )
                        similar_cache[code] = suggestions
                    missing_in_excel.append({"code": code, "suggestions": suggestions})
                    if suggestions:
                        best = suggestions[0]
                        status = u"→ 유사품번 {0} ({1}%)".format(
                            best[1], int(round(best[0] * 100))
                        )
                    else:
                        status = u"엑셀 코드기입 오류/휴먼에러"
                    result_rows.append(
                        {
                            "code": code,
                            "result_qty": None,
                            "result_qty_text": u"",
                            "status": status,
                            "ok": False,
                            "suggestions": suggestions,
                        }
                    )
                    continue

                qty = excel_qty[code]
                if qty is None:
                    no_qty.append(code)
                    result_rows.append(
                        {
                            "code": code,
                            "result_qty": None,
                            "result_qty_text": u"",
                            "status": u"엑셀 수량 없음",
                            "ok": False,
                            "suggestions": [],
                        }
                    )
                    continue

                matched += 1
                result_rows.append(
                    {
                        "code": code,
                        "result_qty": qty,
                        "result_qty_text": format_qty_text(qty),
                        "status": u"엑셀수량",
                        "ok": True,
                        "suggestions": [],
                    }
                )

            seen_err = set()
            missing_unique = []
            for item in missing_in_excel:
                c = item["code"]
                if c not in seen_err:
                    seen_err.add(c)
                    missing_unique.append(item)

            self.logger.info(
                u"매칭(엑셀수량)={0} / SAP→엑셀없음={1} / 엑셀만있음(추출)={2}".format(
                    matched, len(missing_unique), len(unmatched_excel)
                )
            )
            # 개별 WARN은 상위 몇 건만 (파일 I/O 부담 감소)
            for item in missing_unique[:20]:
                c = item["code"]
                sug = item.get("suggestions") or []
                if sug:
                    self.logger.warn(
                        u"SAP만 있음 [{0}] — 유사: {1}({2}%)".format(
                            c, sug[0][1], int(round(sug[0][0] * 100))
                        )
                    )
                else:
                    self.logger.warn(u"SAP만 있음 [{0}]".format(c))
            if len(missing_unique) > 20:
                self.logger.warn(
                    u"… 외 {0}건 생략".format(len(missing_unique) - 20)
                )

            moved_path = None
            has_errors = bool(missing_unique) or bool(no_qty)
            # 오류가 없을 때만 발주(미매칭) 파일 생성
            if unmatched_excel and not has_errors:
                self.logger.info(u"SAP에 없는 엑셀 코드 → 새 파일 추출 (원본 유지)")
                if unmatched_path:
                    self.logger.info(u"저장 위치: {0}".format(unmatched_path))
                moved_path = book.move_unmatched_rows(
                    unmatched_excel, out_path=unmatched_path
                )
            elif unmatched_excel and has_errors:
                self.logger.warn(
                    u"오류 {0}건으로 발주 파일 추출을 생략합니다.".format(
                        len(missing_unique) + len(no_qty)
                    )
                )

            paste_qty = u"\n".join(r["result_qty_text"] for r in result_rows)

            summary = {
                "sap_rows": n,
                "excel_rows": len(excel_rows),
                "matched": matched,
                "missing_count": len(missing_unique),
                "missing_codes": [m["code"] for m in missing_unique],
                "missing_details": missing_unique,
                "no_qty_codes": no_qty,
                "unmatched": len(unmatched_excel),
                "moved_path": moved_path or u"",
                "has_errors": has_errors,
                "result_rows": result_rows,
                "paste_qty_text": paste_qty,
            }
            if has_errors:
                self.logger.warn(
                    u"엑셀 오류가 있어 SAP 붙여넣기 전에 수정 후 재작업이 필요합니다."
                )
            else:
                self.logger.info(
                    u"완료: 결과수량 {0}행. 결과 수량을 SAP에 붙여넣으세요.".format(n)
                )
            if moved_path:
                self.logger.info(u"미매칭 파일: {0}".format(moved_path))
            return summary
        finally:
            book.close()

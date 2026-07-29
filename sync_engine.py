# -*- coding: utf-8 -*-
"""SAP 붙여넣기 ↔ 엑셀 비교 → 결과 수량(순서 유지) 생성."""
from __future__ import print_function

from excel_io import (
    ExcelWorkbook,
    format_qty_text,
    normalize_code,
    parse_pasted_sap_codes,
    parse_pasted_sap_qtys,
)


def _qty_equal(a, b):
    if a is None and b is None:
        return True
    if a is None or b is None:
        return False
    try:
        return abs(float(a) - float(b)) < 1e-9
    except Exception:
        return str(a).strip() == str(b).strip()


class SyncEngine(object):
    def __init__(self, config, logger, stop_flag=None):
        self.config = config
        self.logger = logger
        self.stop_flag = stop_flag

    def _stopped(self):
        if self.stop_flag is not None:
            try:
                return bool(self.stop_flag())
            except Exception:
                return False
        return False

    def run(self, excel_path, codes_text, qtys_text, unmatched_path=None):
        sap_codes = parse_pasted_sap_codes(codes_text)
        sap_qtys = parse_pasted_sap_qtys(qtys_text)

        if not sap_codes:
            raise RuntimeError(
                u"붙여넣은 SAP 코드가 없습니다. SAP 자재번호 열을 복사해 코드 칸에 Ctrl+V 하세요."
            )
        if not sap_qtys:
            raise RuntimeError(
                u"붙여넣은 SAP 수량이 없습니다. SAP 오더수량 열을 복사해 수량 칸에 Ctrl+V 하세요."
            )

        n = max(len(sap_codes), len(sap_qtys))
        if len(sap_qtys) != len(sap_codes):
            self.logger.warn(
                u"코드 {0}개 / 수량 {1}개 — 행 수가 다릅니다. 빈 칸은 비워 둡니다.".format(
                    len(sap_codes), len(sap_qtys)
                )
            )
            while len(sap_codes) < n:
                sap_codes.append(u"")
            while len(sap_qtys) < n:
                sap_qtys.append(None)

        self.logger.info(
            u"SAP 붙여넣기: {0}행 (코드·수량 순서 유지)".format(n)
        )

        book = ExcelWorkbook(excel_path, logger=self.logger)
        book.open()
        try:
            excel_rows = book.load_rows()
            if not excel_rows:
                raise RuntimeError(u"엑셀에 자재코드(B열) 데이터가 없습니다.")

            # 코드 → 엑셀 수량 (동일 코드 여러 행이면 첫 값)
            excel_qty = {}
            excel_by_code = {}
            for er in excel_rows:
                code = normalize_code(er["code"])
                if not code:
                    continue
                if code not in excel_qty:
                    excel_qty[code] = er["qty"]
                    excel_by_code[code] = er

            sap_code_set = set(c for c in sap_codes if c)
            unmatched_excel = [
                er
                for er in excel_rows
                if normalize_code(er["code"]) not in sap_code_set
            ]

            result_rows = []
            changed = 0
            same = 0
            kept = 0

            for i in range(n):
                if self._stopped():
                    self.logger.warn(u"사용자 중지")
                    break
                code = sap_codes[i]
                sap_q = sap_qtys[i]
                if code and code in excel_qty and excel_qty[code] is not None:
                    out_q = excel_qty[code]
                    if _qty_equal(out_q, sap_q):
                        status = u"동일"
                        same += 1
                    else:
                        status = u"수정"
                        changed += 1
                else:
                    out_q = sap_q
                    status = u"유지"  # 엑셀에 없음 → SAP 값 그대로
                    kept += 1

                result_rows.append(
                    {
                        "code": code,
                        "sap_qty": sap_q,
                        "result_qty": out_q,
                        "status": status,
                        "sap_qty_text": format_qty_text(sap_q),
                        "result_qty_text": format_qty_text(out_q),
                    }
                )

            self.logger.info(
                u"비교: 수정={0} 동일={1} 유지(엑셀없음)={2} / 엑셀미매칭추출={3}".format(
                    changed, same, kept, len(unmatched_excel)
                )
            )

            moved_path = None
            if unmatched_excel and not self._stopped():
                self.logger.info(u"엑셀에만 있는 행 → 새 파일 추출 (원본 유지)")
                if unmatched_path:
                    self.logger.info(u"저장 위치: {0}".format(unmatched_path))
                moved_path = book.move_unmatched_rows(
                    unmatched_excel, out_path=unmatched_path
                )

            paste_qty = u"\n".join(r["result_qty_text"] for r in result_rows)
            detail_lines = []
            for i, r in enumerate(result_rows):
                detail_lines.append(
                    u"{0}\t{1}\t{2}\t{3}".format(
                        r["code"],
                        r["sap_qty_text"],
                        r["result_qty_text"],
                        r["status"],
                    )
                )

            summary = {
                "sap_rows": n,
                "excel_rows": len(excel_rows),
                "changed": changed,
                "same": same,
                "kept": kept,
                "unmatched": len(unmatched_excel),
                "moved_path": moved_path or u"",
                "result_rows": result_rows,
                "paste_qty_text": paste_qty,
                "detail_text": u"\n".join(detail_lines),
            }
            self.logger.info(
                u"완료: 결과수량 {0}행 생성 (순서 유지). SAP에 결과 수량만 붙여넣으세요.".format(
                    len(result_rows)
                )
            )
            if moved_path:
                self.logger.info(u"미매칭 파일: {0}".format(moved_path))
            return summary
        finally:
            book.close()

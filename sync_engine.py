# -*- coding: utf-8 -*-
"""붙여넣기 SAP 코드·수량 ↔ 엑셀 비교 후 수량 동기화."""
from __future__ import print_function

import time

from excel_io import (
    ExcelWorkbook,
    normalize_code,
    parse_pasted_sap_codes,
    parse_pasted_sap_qtys,
)
from sap_input import click_and_type_qty, esc_pressed
from settings_ui import sap_config_ready


def _qty_equal(a, b):
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
        if esc_pressed():
            return True
        if self.stop_flag is not None:
            try:
                return bool(self.stop_flag())
            except Exception:
                return False
        return False

    def run(self, excel_path, codes_text, qtys_text):
        if not sap_config_ready(self.config):
            raise RuntimeError(
                u"SAP 행/수량 위치가 없습니다. '행/수량 초기설정'에서 클릭지정 후 저장하세요."
            )

        sap = self.config.get("sap") or {}
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

        if len(sap_qtys) != len(sap_codes):
            self.logger.warn(
                u"코드 {0}개 / 수량 {1}개 — 행 수가 다릅니다. 앞쪽부터 맞춰 비교합니다.".format(
                    len(sap_codes), len(sap_qtys)
                )
            )

        self.logger.info(
            u"SAP 붙여넣기: 코드 {0}개, 수량 {1}개".format(len(sap_codes), len(sap_qtys))
        )
        for i, c in enumerate(sap_codes[:25]):
            q = sap_qtys[i] if i < len(sap_qtys) else None
            self.logger.info(u"  SAP[{0}] {1} / 수량={2}".format(i + 1, c, q))
        if len(sap_codes) > 25:
            self.logger.info(u"  ... 외 {0}개".format(len(sap_codes) - 25))

        sap_index = {}
        for i, code in enumerate(sap_codes):
            if code not in sap_index:
                sap_index[code] = i

        book = ExcelWorkbook(excel_path, logger=self.logger)
        book.open()
        try:
            excel_rows = book.load_rows()
            if not excel_rows:
                raise RuntimeError(u"엑셀에 자재코드(B열) 데이터가 없습니다.")

            matched = []
            unmatched = []
            for er in excel_rows:
                code = normalize_code(er["code"])
                if code in sap_index:
                    matched.append((er, sap_index[code]))
                else:
                    unmatched.append(er)

            self.logger.info(
                u"비교 결과: 매칭 {0}건 / 미매칭(엑셀추출) {1}건".format(
                    len(matched), len(unmatched)
                )
            )

            qty_x = int(sap["qty_center_x"])
            first_y = int(sap["first_row_y"])
            row_h = max(8, int(sap["row_height"]))
            delay_between = float(self.config.get("delay_between_rows", 0.15))

            updated = 0
            same = 0
            skipped = 0
            failed = 0

            for er, sap_i in matched:
                if self._stopped():
                    self.logger.warn(u"사용자 중지")
                    break
                if er["qty"] is None:
                    self.logger.warn(
                        u"엑셀 수량 없음 → 스킵: {0}".format(er["code_display"])
                    )
                    skipped += 1
                    continue

                sap_q = sap_qtys[sap_i] if sap_i < len(sap_qtys) else None
                if _qty_equal(er["qty"], sap_q):
                    self.logger.info(
                        u"수량동일 스킵: {0} (= {1})".format(er["code_display"], er["qty"])
                    )
                    same += 1
                    continue

                y = int(first_y + sap_i * row_h)
                self.logger.info(
                    u"수량수정 {0}: SAP {1} → 엑셀 {2} (행{3}, y={4})".format(
                        er["code_display"], sap_q, er["qty"], sap_i + 1, y
                    )
                )
                try:
                    ok = click_and_type_qty(
                        qty_x, y, er["qty"], self.config, logger=self.logger
                    )
                except Exception as e:
                    ok = False
                    self.logger.warn(u"입력 예외: {0}".format(e))
                if ok:
                    updated += 1
                else:
                    failed += 1
                    self.logger.warn(u"입력 실패: {0}".format(er["code_display"]))
                time.sleep(delay_between)

            moved_path = None
            if unmatched and not self._stopped():
                self.logger.info(u"미매칭 행을 새 엑셀로 자동 추출 중...")
                moved_path = book.move_unmatched_rows(unmatched)
                book.save()
            elif updated:
                book.save()

            summary = {
                "sap_codes": len(sap_codes),
                "excel_rows": len(excel_rows),
                "matched": len(matched),
                "unmatched": len(unmatched),
                "updated": updated,
                "same": same,
                "skipped": skipped,
                "failed": failed,
                "moved_path": moved_path or u"",
            }
            self.logger.info(
                u"완료: 매칭={matched}, 수정={updated}, 동일스킵={same}, "
                u"미매칭추출={unmatched}, 스킵={skipped}, 실패={failed}".format(
                    **summary
                )
            )
            if moved_path:
                self.logger.info(u"미매칭 파일: {0}".format(moved_path))
            return summary
        finally:
            book.close()

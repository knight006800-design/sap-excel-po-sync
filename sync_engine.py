# -*- coding: utf-8 -*-
"""붙여넣기 SAP 코드 ↔ 엑셀 매칭 후 수량 동기화."""
from __future__ import print_function

import time

from excel_io import ExcelWorkbook, normalize_code, parse_pasted_sap_codes
from sap_input import click_and_type_qty, esc_pressed
from settings_ui import sap_config_ready


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

    def run(self, excel_path, pasted_text):
        if not sap_config_ready(self.config):
            raise RuntimeError(
                u"SAP 행/수량 위치가 없습니다. '행/수량 초기설정'에서 클릭지정 후 저장하세요."
            )

        sap = self.config.get("sap") or {}
        sap_codes = parse_pasted_sap_codes(pasted_text)
        if not sap_codes:
            raise RuntimeError(
                u"붙여넣은 SAP 코드가 없습니다. SAP에서 자재번호를 복사한 뒤 "
                u"아래 칸에 Ctrl+V 하세요."
            )

        self.logger.info(u"붙여넣기 SAP 코드 {0}개".format(len(sap_codes)))
        for i, c in enumerate(sap_codes[:30]):
            self.logger.info(u"  SAP[{0}] {1}".format(i + 1, c))
        if len(sap_codes) > 30:
            self.logger.info(u"  ... 외 {0}개".format(len(sap_codes) - 30))

        # 코드 → SAP 행 인덱스(0부터, 화면 위→아래 = 붙여넣기 순서)
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
                u"매칭 {0}건 / 미매칭(이동) {1}건".format(len(matched), len(unmatched))
            )

            qty_x = int(sap["qty_center_x"])
            first_y = int(sap["first_row_y"])
            row_h = max(8, int(sap["row_height"]))
            delay_between = float(self.config.get("delay_between_rows", 0.15))

            updated = 0
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

                y = int(first_y + sap_i * row_h)
                self.logger.info(
                    u"수량입력 {0} → {1} (SAP행{2}, 클릭 y={3})".format(
                        er["code_display"], er["qty"], sap_i + 1, y
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
                self.logger.info(u"미매칭 행을 새 엑셀로 이동 중...")
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
                "skipped": skipped,
                "failed": failed,
                "moved_path": moved_path or u"",
            }
            self.logger.info(
                u"완료: 매칭={matched}, 수량수정={updated}, 미매칭이동={unmatched}, "
                u"스킵={skipped}, 실패={failed}".format(**summary)
            )
            if moved_path:
                self.logger.info(u"미매칭 파일: {0}".format(moved_path))
            return summary
        finally:
            book.close()

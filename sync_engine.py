# -*- coding: utf-8 -*-
"""구매오더 수량 동기화 — 메인 실행 엔진."""
from __future__ import print_function

import platform
import time

import pyautogui

from drive_check import DriveCheckLog
from excel_io import ExcelWorkbook, normalize_qty
from paths_util import load_config
from sap_input import click_and_type_qty, esc_pressed
from sap_vision import SapScreenReader


class SyncEngine(object):
    def __init__(self, gui_log=None, stop_flag=None):
        self.gui_log = gui_log
        self.stop_flag = stop_flag or (lambda: False)
        self.logger = None

    def _stopped(self):
        return self.stop_flag() or esc_pressed()

    def run(self, excel_path=None):
        self.logger = DriveCheckLog(gui_callback=self.gui_log)
        log = self.logger
        cfg = load_config()
        if excel_path:
            cfg["excel_path"] = excel_path

        try:
            log.step(u"환경정보")
            log.info(u"OS: {0}".format(platform.platform()))
            try:
                size = pyautogui.size()
                log.info(u"화면크기: {0}x{1}".format(size[0], size[1]))
            except Exception as e:
                log.warn(u"화면크기 확인 실패: {0}".format(e))
            log.info(u"엑셀경로: {0}".format(cfg.get("excel_path")))
            log.info(u"SAP설정: {0}".format(cfg.get("sap")))

            path = cfg.get("excel_path") or ""
            if not path:
                raise ValueError(u"엑셀 파일이 연결되지 않았습니다.")

            log.step(u"엑셀로드")
            wb = ExcelWorkbook(path, logger=log)
            wb.open()
            rows = wb.load_rows()
            if not rows:
                raise ValueError(u"엑셀에 자재코드(B열) 데이터가 없습니다.")

            if self._stopped():
                log.warn(u"사용자 중지")
                wb.close()
                log.finish(u"중지됨")
                return

            log.step(u"SAP화면OCR")
            for sec in (3, 2, 1):
                log.info(u"{0}초 후 SAP 화면 캡처… (ESC=중지)".format(sec))
                time.sleep(1)
                if self._stopped():
                    log.warn(u"사용자 중지")
                    wb.close()
                    log.finish(u"중지됨")
                    return

            reader = SapScreenReader(cfg, logger=log)
            reader.scan()

            stats = {"match": 0, "updated": 0, "missing": 0, "skip": 0, "error": 0}
            log.step(u"행처리시작", u"엑셀 {0}행".format(len(rows)))

            for item in rows:
                if self._stopped():
                    log.warn(u"사용자 중지 (처리 중)")
                    break

                code = item["code"]
                excel_qty = item["qty"]
                excel_row = item["row"]
                sap = reader.find(code)

                if sap is None:
                    wb.shade_missing(excel_row)
                    stats["missing"] += 1
                    log.info(u"미존재 → B음영: {0} (엑셀행{1})".format(code, excel_row))
                    continue

                sap_qty = sap.get("qty")
                eq = normalize_qty(excel_qty)
                sq = normalize_qty(sap_qty)

                if eq is not None and sq is not None and eq == sq:
                    wb.shade_qty_match(excel_row)
                    stats["match"] += 1
                    log.info(
                        u"수량일치 → D노랑: {0} excel={1} sap={2}".format(code, eq, sq)
                    )
                    continue

                if eq is None:
                    stats["skip"] += 1
                    log.warn(u"엑셀 수량 없음 스킵: {0}".format(code))
                    continue

                try:
                    click_and_type_qty(sap["qty_x"], sap["qty_y"], eq, cfg, logger=log)
                    stats["updated"] += 1
                    log.info(
                        u"수량수정: {0} sap={1} → excel={2} @({3},{4})".format(
                            code, sq, eq, sap["qty_x"], sap["qty_y"]
                        )
                    )
                    time.sleep(float(cfg.get("delay_between_rows", 0.15)))
                except Exception as e:
                    stats["error"] += 1
                    log.error(u"입력 실패 {0}: {1}".format(code, e))
                    log.exception(u"입력")

            log.step(u"엑셀저장")
            wb.save()
            wb.close()

            summary = (
                u"일치={match} 수정={updated} 미존재={missing} 스킵={skip} 오류={error}".format(
                    **stats
                )
            )
            log.finish(summary)
            return stats
        except Exception:
            if self.logger:
                self.logger.exception(u"실행전체")
                try:
                    shot = pyautogui.screenshot()
                    self.logger.save_screenshot(shot, u"치명오류")
                except Exception:
                    pass
                self.logger.finish(u"오류로 종료")
            raise

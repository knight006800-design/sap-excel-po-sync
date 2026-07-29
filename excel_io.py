# -*- coding: utf-8 -*-
"""엑셀 .xls / .xlsx 읽기 및 B·D열 노란색 음영 (Excel COM)."""
from __future__ import print_function

import os
import re

# Excel Interior.Color = BGR. 노랑 = RGB(255,255,0) → 0x0000FFFF
COLOR_YELLOW = 0x00FFFF
# ColorIndex 6 = 표준 노랑 (xls에서 더 안정적)
XL_YELLOW = 6


def normalize_qty(value):
    if value is None:
        return None
    if isinstance(value, (int, float)):
        try:
            if float(value) == int(value):
                return int(value)
            return float(value)
        except Exception:
            return None
    text = str(value).strip()
    if not text:
        return None
    text = text.replace(",", "").replace(" ", "")
    text = re.sub(r"[^\d.\-]", "", text)
    if not text or text in (".", "-", "-."):
        return None
    try:
        num = float(text)
        if num == int(num):
            return int(num)
        return num
    except Exception:
        return None


def normalize_code(value):
    if value is None:
        return u""
    return str(value).strip().upper()


class ExcelWorkbook(object):
    def __init__(self, path, logger=None):
        self.path = os.path.abspath(path)
        self.logger = logger
        self._excel = None
        self._wb = None
        self._ws = None
        self.rows = []
        self._shade_count = 0
        self._com_inited = False

    def _log(self, msg):
        if self.logger:
            self.logger.info(msg)

    def _warn(self, msg):
        if self.logger:
            self.logger.warn(msg)

    def open(self):
        if not os.path.isfile(self.path):
            raise IOError(u"엑셀 파일이 없습니다: {0}".format(self.path))
        import pythoncom
        import win32com.client

        # 백그라운드 스레드에서 COM 사용 시 필수
        pythoncom.CoInitialize()
        self._com_inited = True

        self._excel = None
        try:
            self._excel = win32com.client.GetActiveObject("Excel.Application")
            self._log(u"기존 Excel 인스턴스에 연결")
        except Exception:
            self._excel = win32com.client.DispatchEx("Excel.Application")
            self._log(u"새 Excel 인스턴스 시작")

        self._excel.DisplayAlerts = False
        self._excel.Visible = True

        # 이미 열린 통합문서가 있으면 재사용
        self._wb = None
        target = self.path.lower()
        try:
            for i in range(1, self._excel.Workbooks.Count + 1):
                wb = self._excel.Workbooks(i)
                try:
                    full = str(wb.FullName).lower()
                except Exception:
                    continue
                if full == target:
                    self._wb = wb
                    self._log(u"이미 열린 엑셀에 연결: {0}".format(self.path))
                    break
        except Exception:
            pass

        if self._wb is None:
            self._wb = self._excel.Workbooks.Open(self.path, UpdateLinks=0, ReadOnly=False)
            self._log(u"엑셀 열림: {0}".format(self.path))

        self._ws = self._wb.Worksheets(1)

    def load_rows(self):
        self.rows = []
        used = self._ws.UsedRange
        max_row = used.Row + used.Rows.Count - 1
        for r in range(2, max_row + 1):
            code_raw = self._ws.Cells(r, 2).Value
            qty_raw = self._ws.Cells(r, 4).Value
            code = normalize_code(code_raw)
            if not code:
                continue
            qty = normalize_qty(qty_raw)
            self.rows.append(
                {
                    "row": r,
                    "code": code,
                    "code_display": str(code_raw).strip()
                    if code_raw is not None
                    else code,
                    "qty": qty,
                    "qty_raw": qty_raw,
                }
            )
        self._log(u"엑셀 데이터 행 수: {0}".format(len(self.rows)))
        return self.rows

    def _paint_yellow(self, cell):
        try:
            cell.Interior.Pattern = 1  # xlSolid
            cell.Interior.ColorIndex = XL_YELLOW
        except Exception:
            cell.Interior.Color = COLOR_YELLOW
        self._shade_count += 1

    def shade_missing(self, excel_row):
        """SAP에 없는 자재코드 — B열 노란색."""
        self._paint_yellow(self._ws.Cells(excel_row, 2))

    def shade_qty_match(self, excel_row):
        """수량 일치 — D열 노란색."""
        self._paint_yellow(self._ws.Cells(excel_row, 4))

    def save(self):
        try:
            self._wb.Save()
            self._log(u"엑셀 저장 완료 (음영 {0}칸)".format(self._shade_count))
        except Exception as e:
            # .xls 잠금 등 — 다른 이름으로 저장 시도
            self._warn(u"엑셀 저장 실패: {0}".format(e))
            alt = os.path.splitext(self.path)[0] + u"_결과.xls"
            try:
                self._wb.SaveAs(alt)
                self._log(u"다른 이름으로 저장: {0}".format(alt))
            except Exception as e2:
                raise IOError(u"엑셀 저장 실패: {0} / {1}".format(e, e2))

    def close(self):
        self._ws = None
        self._wb = None
        self._excel = None
        if self._com_inited:
            try:
                import pythoncom

                pythoncom.CoUninitialize()
            except Exception:
                pass
            self._com_inited = False

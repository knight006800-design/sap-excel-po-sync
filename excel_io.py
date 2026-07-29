# -*- coding: utf-8 -*-
"""엑셀 .xls / .xlsx 읽기 및 B·D열 음영 표시 (Excel COM)."""
from __future__ import print_function

import os
import re

COLOR_MISSING = 0xB4C8DC
COLOR_MATCH_YELLOW = 0x00FFFF


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

    def _log(self, msg):
        if self.logger:
            self.logger.info(msg)

    def open(self):
        if not os.path.isfile(self.path):
            raise IOError(u"엑셀 파일이 없습니다: {0}".format(self.path))
        import win32com.client

        self._excel = win32com.client.DispatchEx("Excel.Application")
        self._excel.Visible = False
        self._excel.DisplayAlerts = False
        self._wb = self._excel.Workbooks.Open(self.path)
        self._ws = self._wb.Worksheets(1)
        self._log(u"엑셀 열림: {0}".format(self.path))

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

    def shade_missing(self, excel_row):
        self._ws.Cells(excel_row, 2).Interior.Color = COLOR_MISSING

    def shade_qty_match(self, excel_row):
        self._ws.Cells(excel_row, 4).Interior.Color = COLOR_MATCH_YELLOW

    def save(self):
        self._wb.Save()
        self._log(u"엑셀 저장 완료")

    def close(self):
        try:
            if self._wb is not None:
                self._wb.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            if self._excel is not None:
                self._excel.Quit()
        except Exception:
            pass
        self._wb = None
        self._excel = None
        self._ws = None

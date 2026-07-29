# -*- coding: utf-8 -*-
"""엑셀 .xls / .xlsx 읽기·미매칭 행 이동 (Excel COM)."""
from __future__ import print_function

import os
import re
from datetime import datetime


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
    if not text or text in (u"-", u"."):
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
    text = str(value).strip().upper()
    text = re.sub(r"[^A-Z0-9\-]", "", text)
    return text


def parse_pasted_sap_codes(text):
    """Ctrl+V 붙여넣기 텍스트 → 코드 목록 (SAP 화면 위→아래 순서)."""
    if not text:
        return []
    codes = []
    seen = set()
    for line in str(text).replace("\r\n", "\n").replace("\r", "\n").split("\n"):
        line = line.strip()
        if not line or line.startswith("#"):
            continue
        # 탭/여러 칸이면 첫 토큰 우선, 없으면 코드처럼 보이는 토큰
        parts = re.split(r"[\t;|]+", line)
        candidate = u""
        for p in parts:
            c = normalize_code(p)
            if c and len(c) >= 4:
                candidate = c
                break
        if not candidate:
            candidate = normalize_code(line)
        if not candidate or len(candidate) < 4:
            continue
        if candidate in seen:
            continue
        seen.add(candidate)
        codes.append(candidate)
    return codes


class ExcelWorkbook(object):
    def __init__(self, path, logger=None):
        self.path = os.path.abspath(path)
        self.logger = logger
        self._excel = None
        self._wb = None
        self._ws = None
        self.rows = []
        self._com_inited = False
        self._max_col = 1

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
        self._max_col = max(1, used.Column + used.Columns.Count - 1)
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
        self._log(
            u"엑셀 데이터 행 수: {0} (사용열={1})".format(len(self.rows), self._max_col)
        )
        return self.rows

    def move_unmatched_rows(self, excel_rows, out_path=None):
        """
        미매칭 행 전체를 동일 양식(1행 서식 포함) 새 엑셀로 옮김.
        - 1행(헤더) EntireRow 복사
        - 해당 데이터 행 EntireRow 복사 후 원본에서 삭제
        """
        if not excel_rows:
            return None

        row_nums = sorted(set(int(r["row"]) for r in excel_rows))
        base, ext = os.path.splitext(self.path)
        if not ext:
            ext = u".xls"
        if not out_path:
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            out_path = u"{0}_미매칭_{1}{2}".format(base, stamp, ext)
        out_path = os.path.abspath(out_path)

        # 새 통합문서
        new_wb = self._excel.Workbooks.Add()
        new_ws = new_wb.Worksheets(1)

        # 1행 서식+값 복사
        self._ws.Rows(1).Copy()
        new_ws.Rows(1).PasteSpecial(Paste=-4104)  # xlPasteAll
        try:
            self._excel.CutCopyMode = False
        except Exception:
            pass

        dest = 2
        for r in row_nums:
            self._ws.Rows(r).Copy()
            new_ws.Rows(dest).PasteSpecial(Paste=-4104)
            dest += 1
        try:
            self._excel.CutCopyMode = False
        except Exception:
            pass

        # 열 너비 대략 맞춤
        try:
            for c in range(1, self._max_col + 1):
                new_ws.Columns(c).ColumnWidth = self._ws.Columns(c).ColumnWidth
        except Exception:
            pass

        # 원본에서 아래 행부터 삭제 (행 번호 밀림 방지)
        for r in sorted(row_nums, reverse=True):
            self._ws.Rows(r).Delete()

        try:
            if os.path.isfile(out_path):
                os.remove(out_path)
        except Exception:
            pass

        # .xls / .xlsx 저장
        # 56 = xlExcel8 (.xls), 51 = xlOpenXMLWorkbook (.xlsx)
        file_format = 56 if ext.lower() == u".xls" else 51
        try:
            new_wb.SaveAs(out_path, FileFormat=file_format)
        except Exception:
            # 확장자 강제 xlsx 재시도
            out_path = os.path.splitext(out_path)[0] + u".xlsx"
            new_wb.SaveAs(out_path, FileFormat=51)

        new_wb.Close(SaveChanges=False)
        self._log(u"미매칭 {0}행 이동 → {1}".format(len(row_nums), out_path))
        return out_path

    def save(self):
        try:
            self._wb.Save()
            self._log(u"엑셀 저장 완료")
        except Exception as e:
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

# -*- coding: utf-8 -*-
"""SAP 화면 캡처 및 OCR (자재코드·오더수량)."""
from __future__ import print_function

import os
import re
import sys

import pyautogui
from PIL import Image, ImageOps, ImageFilter, ImageEnhance

from excel_io import normalize_code, normalize_qty


# 자재코드처럼 보이는 패턴 (예: 88510-AA000, DPT42-MX510, RP-01-V1.0A)
CODE_RE = re.compile(
    r"[A-Z0-9][A-Z0-9._/\-]{2,30}",
    re.IGNORECASE,
)


def find_tesseract(cfg):
    candidates = []
    cmd = (cfg or {}).get("tesseract_cmd") or ""
    if cmd:
        candidates.append(cmd)
    if getattr(sys, "frozen", False):
        base = os.path.dirname(sys.executable)
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    candidates.extend(
        [
            os.path.join(base, "tesseract", "tesseract.exe"),
            r"C:\Program Files\Tesseract-OCR\tesseract.exe",
            r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
        ]
    )
    for p in candidates:
        if p and os.path.isfile(p):
            return p
    return None


def setup_tesseract(cfg, logger=None):
    import pytesseract

    path = find_tesseract(cfg)
    if path:
        pytesseract.pytesseract.tesseract_cmd = path
        if logger:
            logger.info(u"Tesseract: {0}".format(path))
        return True
    if logger:
        logger.error(
            u"Tesseract OCR를 찾을 수 없습니다. "
            u"https://github.com/UB-Mannheim/tesseract/wiki 에서 설치하거나 "
            u"프로그램 폴더에 tesseract\\tesseract.exe 를 두세요."
        )
    return False


def grab_region(left, top, right, bottom):
    width = max(1, int(right) - int(left))
    height = max(1, int(bottom) - int(top))
    return pyautogui.screenshot(region=(int(left), int(top), width, height))


def preprocess(img, scale=2.5):
    if scale and scale != 1:
        w, h = img.size
        resample = Image.LANCZOS if hasattr(Image, "LANCZOS") else Image.ANTIALIAS
        img = img.resize((max(1, int(w * scale)), max(1, int(h * scale))), resample)
    gray = ImageOps.grayscale(img)
    gray = ImageOps.autocontrast(gray)
    gray = ImageEnhance.Contrast(gray).enhance(1.6)
    gray = gray.filter(ImageFilter.SHARPEN)
    # 이진화로 SAP 글자 대비 강화
    try:
        gray = gray.point(lambda x: 255 if x > 160 else 0)
    except Exception:
        pass
    return gray


def ocr_raw(img, cfg, psm=6, digits_only=False):
    import pytesseract

    scale = float(cfg.get("ocr_scale", 2.8))
    proc = preprocess(img, scale=scale)
    config = "--psm {0}".format(psm)
    if digits_only:
        config += " -c tessedit_char_whitelist=0123456789,."
    try:
        text = pytesseract.image_to_string(proc, lang="eng", config=config)
    except Exception:
        text = pytesseract.image_to_string(proc, config=config)
    return (text or "").strip()


def extract_codes(text):
    """OCR 텍스트에서 자재코드 후보 추출."""
    if not text:
        return []
    # O/0 혼동 보정은 코드 단위로
    found = []
    for m in CODE_RE.finditer(text.upper().replace(" ", "")):
        token = m.group(0).strip("-._/")
        if len(token) < 4:
            continue
        # 순수 숫자만(날짜 등) 제외 — 하이픈/문자 없으면 스킵하되 5자리 이상은 허용
        if token.isdigit() and len(token) < 6:
            continue
        # 헤더성 단어 제외
        if token in (
            u"EA",
            u"CN7",
            u"LX3",
            u"STATUS",
            u"MATERIAL",
        ):
            continue
        found.append(normalize_code(token))
    return found


def clean_material_code(text):
    codes = extract_codes(text)
    if not codes:
        return u""
    # 가장 긴 후보
    return max(codes, key=len)


def ocr_cell_code(img, cfg):
    """셀 이미지를 여러 PSM으로 시도."""
    best = u""
    raws = []
    for psm in (7, 8, 6, 13):
        raw = ocr_raw(img, cfg, psm=psm, digits_only=False)
        raws.append(raw)
        code = clean_material_code(raw)
        if code and len(code) > len(best):
            best = code
        if best and len(best) >= 6:
            break
    return best, u" | ".join([r.replace("\n", " ") for r in raws if r])


def ocr_cell_qty(img, cfg):
    raw = ocr_raw(img, cfg, psm=7, digits_only=True)
    if not raw:
        raw = ocr_raw(img, cfg, psm=8, digits_only=False)
    return normalize_qty(raw), raw


class SapScreenReader(object):
    def __init__(self, cfg, logger=None):
        self.cfg = cfg
        self.logger = logger
        self.rows = []
        self.by_code = {}

    def _log(self, msg):
        if self.logger:
            self.logger.info(msg)

    def _warn(self, msg):
        if self.logger:
            self.logger.warn(msg)

    def scan(self):
        sap = self.cfg["sap"]
        left = int(sap["material_left"])
        top = int(sap["material_top"])
        right = int(sap["material_right"])
        bottom = int(sap["material_bottom"])
        qty_x = int(sap["qty_center_x"])
        first_y = int(sap["first_row_y"])
        row_h = max(8, int(sap["row_height"]))
        n_rows = int(sap.get("visible_rows") or 25)

        if right <= left or bottom <= top:
            raise RuntimeError(
                u"보정 좌표가 잘못되었습니다 (영역 크기 0). SAP 화면 보정을 다시 하세요."
            )

        if not setup_tesseract(self.cfg, self.logger):
            raise RuntimeError(u"Tesseract OCR 미설치")

        self._log(
            u"OCR영역 material=({0},{1})-({2},{3}) qty_x={4} first_y={5} row_h={6} n={7}".format(
                left, top, right, bottom, qty_x, first_y, row_h, n_rows
            )
        )

        full = grab_region(left, top, right, bottom)
        if self.logger:
            self.logger.save_screenshot(full, u"자재코드영역")
            # 전체 화면 캡처도 남겨 보정 위치 확인
            try:
                self.logger.save_screenshot(pyautogui.screenshot(), u"전체화면")
            except Exception:
                pass

        # 1차: 전체 열 OCR (디버그 + 보조)
        col_text = ocr_raw(full, self.cfg, psm=6)
        self._log(u"열전체OCR: {0}".format(col_text.replace("\n", " / ")[:300]))

        self.rows = []
        # first_y가 영역 밖이면 영역 상단+행높이/2로 보정
        if first_y < top or first_y > bottom:
            self._warn(
                u"첫 행 Y({0})가 자재코드 영역 밖입니다. 영역 기준으로 자동 조정합니다.".format(
                    first_y
                )
            )
            first_y = top + row_h // 2

        empty_streak = 0
        for i in range(max(n_rows, 40)):
            cy = first_y + i * row_h
            if cy > bottom + row_h:
                break
            cell_top = max(top, int(cy - row_h * 0.45))
            cell_bottom = min(bottom, int(cy + row_h * 0.45))
            if cell_bottom - cell_top < 4:
                continue

            mat_img = grab_region(left, cell_top, right, cell_bottom)
            code, raw = ocr_cell_code(mat_img, self.cfg)
            if not code or len(code) < 4:
                empty_streak += 1
                if empty_streak >= 3 and len(self.rows) > 0:
                    break
                continue
            empty_streak = 0

            q_left = qty_x - 55
            q_right = qty_x + 55
            qty_img = grab_region(q_left, cell_top, q_right, cell_bottom)
            qty, qty_raw = ocr_cell_qty(qty_img, self.cfg)

            item = {
                "index": i,
                "code": code,
                "qty": qty,
                "qty_raw": qty_raw,
                "qty_x": qty_x,
                "qty_y": cy,
                "ocr_code_raw": raw,
            }
            self.rows.append(item)
            self._log(
                u"SAP행{0}: code={1} qty={2} (qraw={3}) click=({4},{5})".format(
                    i, code, qty, qty_raw, qty_x, cy
                )
            )

        # 2차: 행 OCR 실패 시 열 전체에서 코드만 이라도 수집
        if not self.rows:
            self._warn(u"행단위 OCR 실패 → 열전체 텍스트에서 코드 재추출 시도")
            codes = extract_codes(col_text)
            # 중복 제거 순서 유지
            seen = set()
            uniq = []
            for c in codes:
                if c not in seen:
                    seen.add(c)
                    uniq.append(c)
            for i, code in enumerate(uniq):
                cy = first_y + i * row_h
                if cy > bottom + row_h:
                    cy = top + row_h // 2 + i * row_h
                self.rows.append(
                    {
                        "index": i,
                        "code": code,
                        "qty": None,
                        "qty_raw": u"",
                        "qty_x": qty_x,
                        "qty_y": cy,
                        "ocr_code_raw": u"(열전체)",
                    }
                )
                self._log(u"열추출 SAP행{0}: code={1} click=({2},{3})".format(i, code, qty_x, cy))

        self.by_code = {}
        for r in self.rows:
            if r["code"] not in self.by_code:
                self.by_code[r["code"]] = r

        self._log(u"SAP OCR 인식 행 수: {0}".format(len(self.rows)))
        if not self.rows:
            raise RuntimeError(
                u"SAP 화면에서 자재코드를 하나도 읽지 못했습니다.\n"
                u"1) SAP가 좌측 모니터에 보이는지\n"
                u"2) [SAP 화면 보정]을 다시 했는지\n"
                u"3) 구동점검 폴더의 캡처 그림에 자재코드가 들어있는지\n"
                u"확인 후 다시 실행하세요."
            )
        return self.rows

    def find(self, code):
        key = normalize_code(code)
        hit = self.by_code.get(key)
        if hit:
            return hit
        # 하이픈 제거 비교 (OCR이 하이픈을 빠뜨린 경우)
        key2 = key.replace("-", "").replace(".", "")
        for c, row in self.by_code.items():
            if c.replace("-", "").replace(".", "") == key2:
                return row
        return None

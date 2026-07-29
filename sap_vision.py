# -*- coding: utf-8 -*-
"""SAP 화면 캡처 및 OCR (자재코드·오더수량)."""
from __future__ import print_function

import os
import re
import sys

import pyautogui
from PIL import Image, ImageOps, ImageFilter

from excel_io import normalize_code, normalize_qty


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
    width = max(1, right - left)
    height = max(1, bottom - top)
    return pyautogui.screenshot(region=(left, top, width, height))


def preprocess(img, scale=2.5):
    if scale and scale != 1:
        w, h = img.size
        resample = Image.LANCZOS if hasattr(Image, "LANCZOS") else Image.ANTIALIAS
        img = img.resize((int(w * scale), int(h * scale)), resample)
    gray = ImageOps.grayscale(img)
    gray = ImageOps.autocontrast(gray)
    gray = gray.filter(ImageFilter.SHARPEN)
    return gray


def ocr_text(img, cfg, digits_only=False):
    import pytesseract

    scale = float(cfg.get("ocr_scale", 2.5))
    proc = preprocess(img, scale=scale)
    config = "--psm 7"
    if digits_only:
        config += " -c tessedit_char_whitelist=0123456789,."
    else:
        config += (
            " -c tessedit_char_whitelist="
            "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-._/"
        )
    try:
        text = pytesseract.image_to_string(proc, lang="eng", config=config)
    except Exception:
        text = pytesseract.image_to_string(proc, config=config)
    return (text or "").strip()


def clean_material_code(text):
    if not text:
        return u""
    parts = re.split(r"[\s\n\r]+", text)
    parts = [p.strip() for p in parts if p.strip()]
    if not parts:
        return u""
    best = max(parts, key=len)
    best = best.replace("O", "0") if re.match(r"^[\dO\-]+$", best) else best
    return normalize_code(best)


class SapScreenReader(object):
    def __init__(self, cfg, logger=None):
        self.cfg = cfg
        self.logger = logger
        self.rows = []
        self.by_code = {}

    def _log(self, msg):
        if self.logger:
            self.logger.info(msg)

    def scan(self):
        sap = self.cfg["sap"]
        left = int(sap["material_left"])
        top = int(sap["material_top"])
        right = int(sap["material_right"])
        bottom = int(sap["material_bottom"])
        qty_x = int(sap["qty_center_x"])
        first_y = int(sap["first_row_y"])
        row_h = int(sap["row_height"])
        n_rows = int(sap.get("visible_rows") or 25)

        if not setup_tesseract(self.cfg, self.logger):
            raise RuntimeError(u"Tesseract OCR 미설치")

        full = grab_region(left, top, right, bottom)
        if self.logger:
            self.logger.save_screenshot(full, u"자재코드영역")

        self.rows = []
        for i in range(n_rows):
            cy = first_y + i * row_h
            if cy < top or cy > bottom:
                break
            cell_top = max(top, int(cy - row_h * 0.45))
            cell_bottom = min(bottom, int(cy + row_h * 0.45))
            if cell_bottom <= cell_top:
                continue
            mat_img = grab_region(left, cell_top, right, cell_bottom)
            raw = ocr_text(mat_img, self.cfg, digits_only=False)
            code = clean_material_code(raw)
            if not code or len(code) < 3:
                continue

            q_left = qty_x - 45
            q_right = qty_x + 45
            qty_img = grab_region(q_left, cell_top, q_right, cell_bottom)
            qty_raw = ocr_text(qty_img, self.cfg, digits_only=True)
            qty = normalize_qty(qty_raw)

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
                u"SAP행{0}: code={1} (raw={2}) qty={3} (raw={4}) click=({5},{6})".format(
                    i, code, raw.replace("\n", " "), qty, qty_raw, qty_x, cy
                )
            )

        self.by_code = {}
        for r in self.rows:
            if r["code"] not in self.by_code:
                self.by_code[r["code"]] = r

        self._log(u"SAP OCR 인식 행 수: {0}".format(len(self.rows)))
        return self.rows

    def find(self, code):
        return self.by_code.get(normalize_code(code))

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
    found = []
    cleaned = text.upper().replace(" ", "")
    for m in CODE_RE.finditer(cleaned):
        token = m.group(0).strip("-._/")
        if len(token) < 4:
            continue
        if token.isdigit() and len(token) < 6:
            continue
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


def ocr_confusions(code):
    """OCR 혼동 문자 치환 후보 (O/0, I/1, S/5, B/8)."""
    if not code:
        return []
    base = normalize_code(code)
    variants = set([base, base.replace("-", "").replace(".", "")])
    pairs = [("O", "0"), ("0", "O"), ("I", "1"), ("1", "I"), ("S", "5"), ("5", "S"), ("B", "8"), ("8", "B")]
    for a, b in pairs:
        if a in base:
            variants.add(base.replace(a, b))
    return list(variants)


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

        # 1차: 전체 열 OCR (전체 기록)
        col_text = ocr_raw(full, self.cfg, psm=6)
        self._log(u"열전체OCR(전체): {0}".format(col_text.replace("\n", " / ")))
        if self.logger and hasattr(self.logger, "save_text"):
            self.logger.save_text(u"OCR_열전체", col_text)
        elif self.logger:
            self.logger.info(u"[OCR원문저장] 열전체 길이={0}".format(len(col_text)))

        self.rows = []
        if first_y < top or first_y > bottom:
            self._warn(
                u"첫 행 Y({0})가 자재코드 영역 밖입니다. 영역 기준으로 자동 조정합니다.".format(
                    first_y
                )
            )
            first_y = top + row_h // 2

        # 영역 끝까지 전부 스캔 (중간에 실패해도 중단하지 않음)
        max_i = max(n_rows, int((bottom - top) / float(row_h)) + 3)
        for i in range(max_i):
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
                self._log(
                    u"SAP행{0} 인식실패 y={1} raw=[{2}]".format(
                        i, cy, (raw or u"").replace("\n", " ")[:120]
                    )
                )
                continue

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
                u"SAP행{0}: code={1} qty={2} (qraw={3}) raw=[{4}] click=({5},{6})".format(
                    i,
                    code,
                    qty,
                    qty_raw,
                    (raw or u"").replace("\n", " ")[:80],
                    qty_x,
                    cy,
                )
            )

        # 열전체에서 뽑은 코드도 항상 병합 (행 OCR이 일부만 성공해도 보완)
        col_codes = extract_codes(col_text)
        existing = set([r["code"] for r in self.rows])
        for code in col_codes:
            if code in existing:
                continue
            # 클릭 Y는 대략 배치 (수량 수정 정확도는 행OCR 우선)
            cy = first_y + len(self.rows) * row_h
            self.rows.append(
                {
                    "index": len(self.rows),
                    "code": code,
                    "qty": None,
                    "qty_raw": u"",
                    "qty_x": qty_x,
                    "qty_y": min(cy, bottom - 2),
                    "ocr_code_raw": u"(열전체병합)",
                }
            )
            existing.add(code)
            self._log(u"열병합 추가코드: {0}".format(code))

        if not self.rows and col_codes:
            self._warn(u"행단위 OCR 실패 → 열전체 코드만 사용")
            for i, code in enumerate(col_codes):
                cy = first_y + i * row_h
                self.rows.append(
                    {
                        "index": i,
                        "code": code,
                        "qty": None,
                        "qty_raw": u"",
                        "qty_x": qty_x,
                        "qty_y": cy if cy <= bottom else top + row_h // 2 + i * row_h,
                        "ocr_code_raw": u"(열전체)",
                    }
                )
                self._log(
                    u"열추출 SAP행{0}: code={1} click=({2},{3})".format(
                        i, code, qty_x, cy
                    )
                )

        self.by_code = {}
        for r in self.rows:
            if r["code"] not in self.by_code:
                self.by_code[r["code"]] = r

        # OCR 결과 전부 파일로 저장
        dump_lines = [
            u"=== SAP OCR 결과 전체 ===",
            u"인식행수: {0}".format(len(self.rows)),
            u"열전체OCR:",
            col_text,
            u"",
            u"--- 행별 ---",
        ]
        for r in self.rows:
            dump_lines.append(
                u"idx={0} code={1} qty={2} qraw={3} raw={4} click=({5},{6})".format(
                    r["index"],
                    r["code"],
                    r["qty"],
                    r.get("qty_raw"),
                    r.get("ocr_code_raw"),
                    r["qty_x"],
                    r["qty_y"],
                )
            )
        dump = u"\n".join(dump_lines)
        if self.logger and hasattr(self.logger, "save_text"):
            path = self.logger.save_text(u"OCR결과", dump)
            self._log(u"OCR결과 파일: {0}".format(path))
        else:
            self._log(dump)

        self._log(u"SAP OCR 인식 행 수: {0}".format(len(self.rows)))
        self._log(
            u"SAP OCR 코드목록: {0}".format(
                u", ".join([r["code"] for r in self.rows])
            )
        )
        if not self.rows:
            raise RuntimeError(
                u"SAP 화면에서 자재코드를 하나도 읽지 못했습니다.\n"
                u"1) SAP가 주모니터에 보이는지\n"
                u"2) [SAP 화면 보정]을 다시 했는지\n"
                u"3) 구동점검 폴더의 캡처/OCR결과 파일을 확인하세요."
            )
        return self.rows

    def find(self, code):
        key = normalize_code(code)
        hit = self.by_code.get(key)
        if hit:
            return hit
        # 하이픈/점 제거 + OCR 혼동 문자
        want = set()
        for v in ocr_confusions(key):
            want.add(v)
            want.add(v.replace("-", "").replace(".", ""))
        for c, row in self.by_code.items():
            for v in ocr_confusions(c):
                compact = v.replace("-", "").replace(".", "")
                if v in want or compact in want:
                    return row
        return None

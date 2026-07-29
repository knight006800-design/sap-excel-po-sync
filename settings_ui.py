# -*- coding: utf-8 -*-
"""행구분·수량칸 초기설정 / 재설정 / 클릭 보정."""
from __future__ import print_function

import threading
import time

try:
    import tkinter as tk
    from tkinter import messagebox
except ImportError:
    import Tkinter as tk
    import tkMessageBox as messagebox

import pyautogui

from monitors import (
    enum_monitors,
    is_on_primary,
    place_window_on_secondary,
    secondary_monitors,
)
from paths_util import default_config, load_config, save_config


class RowQtySettings(object):
    """행 Y / 행높이 / 수량 X 설정 창."""

    def __init__(self, parent, on_done=None, logger=None):
        self.parent = parent
        self.on_done = on_done
        self.logger = logger
        self.win = None
        self.monitors = enum_monitors()
        self._stop = False

    def _log(self, msg):
        if self.logger:
            self.logger.info(msg)

    def start(self):
        cfg = load_config()
        sap = cfg.get("sap") or {}
        self.win = tk.Toplevel(self.parent)
        self.win.title(u"행/수량 초기설정")
        self.win.attributes("-topmost", True)
        place_window_on_secondary(self.win, width=440, height=400, margin=24)
        bg = "#F4F5F7"
        self.win.configure(bg=bg)

        tk.Label(
            self.win,
            text=u"주모니터 SAP에서 클릭지정만 하세요.\n값은 자동으로 채워집니다. (숫자 직접 입력 없음)\n설정은 config.json에 저장·재사용됩니다.",
            font=("Segoe UI", 10),
            bg=bg,
            fg="#1A1D23",
            justify="left",
        ).pack(anchor="w", padx=16, pady=14)

        self.var_first = tk.StringVar(value=str(sap.get("first_row_y", 240)))
        self.var_rh = tk.StringVar(value=str(sap.get("row_height", 28)))
        self.var_qty = tk.StringVar(value=str(sap.get("qty_center_x", 1180)))

        self._row(u"첫 행 Y (first_row_y)", self.var_first, self._click_first_y)
        self._row(u"행 높이 (row_height)", self.var_rh, self._click_row_height)
        self._row(u"수량 칸 X (qty_center_x)", self.var_qty, self._click_qty_x)

        bf = tk.Frame(self.win, bg=bg)
        bf.pack(pady=16)
        tk.Button(
            bf,
            text=u"저장",
            width=10,
            command=self._save,
            bg="#2563EB",
            fg="#FFFFFF",
            activebackground="#1D4ED8",
            activeforeground="#FFFFFF",
            relief="flat",
            font=("Segoe UI", 10, "bold"),
            padx=8,
            pady=6,
        ).pack(side="left", padx=4)
        tk.Button(
            bf,
            text=u"재설정",
            width=10,
            command=self._reset,
            bg="#FFFFFF",
            fg="#1A1D23",
            relief="flat",
            font=("Segoe UI", 10),
            padx=8,
            pady=6,
        ).pack(side="left", padx=4)
        tk.Button(
            bf,
            text=u"닫기",
            width=10,
            command=self._close,
            bg="#FFFFFF",
            fg="#1A1D23",
            relief="flat",
            font=("Segoe UI", 10),
            padx=8,
            pady=6,
        ).pack(side="left", padx=4)

        self.status = tk.Label(
            self.win, text=u"", fg="#1E3A8A", bg=bg, font=("Segoe UI", 9)
        )
        self.status.pack(pady=4)
        self.win.protocol("WM_DELETE_WINDOW", self._close)

    def _row(self, label, var, click_cmd):
        bg = "#F4F5F7"
        f = tk.Frame(self.win, bg=bg)
        f.pack(fill="x", padx=16, pady=6)
        tk.Label(
            f, text=label, width=22, anchor="w", bg=bg, fg="#1A1D23", font=("Segoe UI", 10)
        ).pack(side="left")
        tk.Label(
            f,
            textvariable=var,
            width=10,
            anchor="w",
            bg="#FFFFFF",
            fg="#1A1D23",
            relief="flat",
            highlightthickness=1,
            highlightbackground="#E2E4E8",
            padx=8,
            pady=4,
            font=("Segoe UI", 10),
        ).pack(side="left", padx=6)
        tk.Button(
            f,
            text=u"클릭지정",
            command=click_cmd,
            width=10,
            bg="#FFFFFF",
            fg="#2563EB",
            relief="flat",
            font=("Segoe UI", 9, "bold"),
            padx=6,
            pady=4,
        ).pack(side="left")

    def _esc(self):
        import ctypes

        return bool(ctypes.windll.user32.GetAsyncKeyState(0x1B) & 0x8000)

    def _wait_primary_click(self, prompt):
        self.status.config(text=prompt + u" (ESC=취소)")
        self.win.update()
        has_sec = bool(secondary_monitors(self.monitors))
        time.sleep(0.3)
        import ctypes

        VK = 0x01
        while not self._stop:
            if self._esc():
                return None
            if ctypes.windll.user32.GetAsyncKeyState(VK) & 0x8000:
                while ctypes.windll.user32.GetAsyncKeyState(VK) & 0x8000:
                    time.sleep(0.02)
                x, y = pyautogui.position()
                # 설정창 위 클릭 무시
                try:
                    gx, gy = self.win.winfo_rootx(), self.win.winfo_rooty()
                    gw, gh = self.win.winfo_width(), self.win.winfo_height()
                    if gx <= x < gx + gw and gy <= y < gy + gh:
                        continue
                except Exception:
                    pass
                if has_sec and not is_on_primary(x, y, self.monitors):
                    self.status.config(text=u"주모니터 SAP에서 클릭하세요!")
                    continue
                return (x, y)
            time.sleep(0.02)
        return None

    def _click_first_y(self):
        def work():
            pt = self._wait_primary_click(u"첫 데이터 행 세로 중앙을 클릭")
            if pt:
                self.var_first.set(str(pt[1]))
                self.status.config(text=u"first_row_y={0}".format(pt[1]))
                self._log(u"클릭지정 first_row_y={0}".format(pt[1]))

        threading.Thread(target=work, daemon=True).start()

    def _click_row_height(self):
        def work():
            self.status.config(text=u"먼저 첫 행 중앙 클릭…")
            p1 = self._wait_primary_click(u"첫 행 중앙 클릭")
            if not p1:
                return
            p2 = self._wait_primary_click(u"둘째 행 중앙 클릭")
            if not p2:
                return
            rh = abs(p2[1] - p1[1])
            if rh < 5:
                rh = 28
            self.var_first.set(str(p1[1]))
            self.var_rh.set(str(rh))
            self.status.config(text=u"first_y={0}, row_h={1}".format(p1[1], rh))
            self._log(u"클릭지정 first_row_y={0} row_height={1}".format(p1[1], rh))

        threading.Thread(target=work, daemon=True).start()

    def _click_qty_x(self):
        def work():
            pt = self._wait_primary_click(u"오더수량 칸 정중앙 클릭")
            if pt:
                self.var_qty.set(str(pt[0]))
                self.status.config(text=u"qty_center_x={0}".format(pt[0]))
                self._log(u"클릭지정 qty_center_x={0}".format(pt[0]))

        threading.Thread(target=work, daemon=True).start()

    def _save(self):
        try:
            first_y = int(float(self.var_first.get().strip()))
            rh = int(float(self.var_rh.get().strip()))
            qty_x = int(float(self.var_qty.get().strip()))
        except Exception:
            messagebox.showerror(u"입력 오류", u"숫자만 입력하세요.", parent=self.win)
            return
        if rh < 5:
            messagebox.showerror(u"입력 오류", u"행 높이는 5 이상이어야 합니다.", parent=self.win)
            return
        cfg = load_config()
        sap = cfg.setdefault("sap", {})
        sap["first_row_y"] = first_y
        sap["row_height"] = rh
        sap["qty_center_x"] = qty_x
        # visible_rows 갱신
        top = int(sap.get("material_top", 0))
        bottom = int(sap.get("material_bottom", 0))
        if bottom > top:
            sap["visible_rows"] = max(1, int(round(float(bottom - top) / float(rh))))
        save_config(cfg)
        self._log(
            u"행/수량 저장: first_y={0} row_h={1} qty_x={2}".format(first_y, rh, qty_x)
        )
        if self.on_done:
            self.on_done()
        self._close()

    def _reset(self):
        ok = messagebox.askyesno(
            u"재설정",
            u"행/수량 설정을 기본값으로 되돌릴까요?\n(자재코드 영역 박스는 유지됩니다)",
            parent=self.win,
        )
        if not ok:
            return
        d = default_config()["sap"]
        cfg = load_config()
        sap = cfg.setdefault("sap", {})
        sap["first_row_y"] = d["first_row_y"]
        sap["row_height"] = d["row_height"]
        sap["qty_center_x"] = d["qty_center_x"]
        save_config(cfg)
        self.var_first.set(str(d["first_row_y"]))
        self.var_rh.set(str(d["row_height"]))
        self.var_qty.set(str(d["qty_center_x"]))
        self._log(u"행/수량 재설정(기본값)")
        self.status.config(text=u"기본값으로 재설정됨 — [저장]은 이미 반영됨")
        if self.on_done:
            self.on_done()

    def _close(self):
        self._stop = True
        try:
            if self.win:
                self.win.destroy()
        except Exception:
            pass
        self.win = None


def sap_config_ready(cfg=None):
    """동기화에 필요한 행/수량 좌표가 갖춰졌는지 (OCR 영역 불필요)."""
    cfg = cfg or load_config()
    sap = cfg.get("sap") or {}
    try:
        qty_x = int(sap["qty_center_x"])
        first_y = int(sap["first_row_y"])
        rh = int(sap["row_height"])
    except Exception:
        return False
    if rh < 5 or qty_x < 0 or first_y < 0:
        return False
    return True


def format_sap_status(cfg=None):
    cfg = cfg or load_config()
    sap = cfg.get("sap") or {}
    ready = sap_config_ready(cfg)
    return (
        u"설정상태: {0} | 행Y={1} 행H={2} 수량X={3}".format(
            u"준비됨" if ready else u"미완료(행/수량 설정 필요)",
            sap.get("first_row_y"),
            sap.get("row_height"),
            sap.get("qty_center_x"),
        )
    )

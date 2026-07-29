# -*- coding: utf-8 -*-
"""자재코드 영역 — 이동/크기조절 가능 박스 오버레이."""
from __future__ import print_function

import re

try:
    import tkinter as tk
    from tkinter import messagebox
except ImportError:
    import Tkinter as tk
    import tkMessageBox as messagebox

from monitors import place_window_on_secondary, primary_monitor
from paths_util import load_config, save_config

EDGE = 10
MIN_W = 60
MIN_H = 40


class MaterialRegionBox(object):
    def __init__(self, parent, on_done=None, logger=None):
        self.parent = parent
        self.on_done = on_done
        self.logger = logger
        self.box = None
        self.ctrl = None
        self._mode = None
        self._start = None
        self._geom0 = None

    def _log(self, msg):
        if self.logger:
            self.logger.info(msg)

    def start(self):
        cfg = load_config()
        sap = cfg.get("sap") or {}
        prim = primary_monitor()

        left = int(sap.get("material_left", prim["left"] + 120))
        top = int(sap.get("material_top", prim["top"] + 200))
        right = int(sap.get("material_right", left + 160))
        bottom = int(sap.get("material_bottom", top + 200))
        w = max(MIN_W, right - left)
        h = max(MIN_H, bottom - top)
        left = max(prim["left"], min(left, prim["right"] - w))
        top = max(prim["top"], min(top, prim["bottom"] - h))

        messagebox.showinfo(
            u"자재코드 영역 박스",
            u"주모니터 SAP 위에 빨간 박스가 뜹니다.\n"
            u"· 안쪽 드래그 = 이동\n"
            u"· 테두리/모서리 드래그 = 크기 조절\n"
            u"· 이전 크기·위치를 기억합니다\n"
            u"맞추면 안내창의 [확인]을 누르세요.",
            parent=self.parent,
        )

        self.box = tk.Toplevel(self.parent)
        self.box.overrideredirect(True)
        self.box.attributes("-topmost", True)
        try:
            self.box.attributes("-alpha", 0.4)
        except Exception:
            pass
        self.box.configure(bg="#FF1744")
        self.box.geometry("{0}x{1}+{2}+{3}".format(w, h, left, top))

        self.inner = tk.Frame(self.box, bg="#FFCDD2", cursor="fleur")
        self.inner.pack(fill="both", expand=True, padx=3, pady=3)
        tk.Label(
            self.inner,
            text=u"자재코드 영역\n이동/크기조절",
            bg="#FFCDD2",
            fg="#B71C1C",
            font=("Malgun Gothic", 10, "bold"),
            justify="center",
        ).pack(expand=True)

        for wdg in (self.box, self.inner):
            wdg.bind("<ButtonPress-1>", self._on_press)
            wdg.bind("<B1-Motion>", self._on_drag)
            wdg.bind("<ButtonRelease-1>", self._on_release)
            wdg.bind("<Motion>", self._on_hover)

        self.ctrl = tk.Toplevel(self.parent)
        self.ctrl.title(u"자재코드 영역 — 확인")
        self.ctrl.attributes("-topmost", True)
        place_window_on_secondary(self.ctrl, width=380, height=170, margin=20)
        self.ctrl.configure(bg="#fff8e1")
        tk.Label(
            self.ctrl,
            text=u"빨간 박스를 자재코드 열에 맞춘 뒤 [확인]",
            font=("Malgun Gothic", 11),
            bg="#fff8e1",
        ).pack(padx=12, pady=12)
        self.size_lbl = tk.Label(
            self.ctrl, text=u"", font=("Consolas", 9), bg="#fff8e1", fg="#1565c0"
        )
        self.size_lbl.pack(pady=4)
        bf = tk.Frame(self.ctrl, bg="#fff8e1")
        bf.pack(pady=8)
        tk.Button(bf, text=u"확인", width=12, command=self._confirm, bg="#c8e6c9").pack(
            side="left", padx=6
        )
        tk.Button(bf, text=u"취소", width=12, command=self._cancel).pack(
            side="left", padx=6
        )
        self._tick_lbl()
        self.ctrl.protocol("WM_DELETE_WINDOW", self._cancel)

    def _geom(self):
        self.box.update_idletasks()
        m = re.match(r"(\d+)x(\d+)\+(-?\d+)\+(-?\d+)", self.box.geometry())
        if not m:
            return 100, 100, 0, 0
        return int(m.group(1)), int(m.group(2)), int(m.group(3)), int(m.group(4))

    def _tick_lbl(self):
        try:
            w, h, x, y = self._geom()
            self.size_lbl.config(
                text=u"({0},{1})-({2},{3})  {4}x{5}".format(
                    x, y, x + w, y + h, w, h
                )
            )
        except Exception:
            pass
        if self.box:
            try:
                self.box.after(200, self._tick_lbl)
            except Exception:
                pass

    def _hit_mode(self, event):
        w, h, x, y = self._geom()
        rel_x = event.x_root - x
        rel_y = event.y_root - y
        on_w = rel_x <= EDGE
        on_e = rel_x >= w - EDGE
        on_n = rel_y <= EDGE
        on_s = rel_y >= h - EDGE
        if on_n and on_w:
            return "nw"
        if on_n and on_e:
            return "ne"
        if on_s and on_w:
            return "sw"
        if on_s and on_e:
            return "se"
        if on_n:
            return "n"
        if on_s:
            return "s"
        if on_w:
            return "w"
        if on_e:
            return "e"
        return "move"

    def _on_hover(self, event):
        cursors = {
            "move": "fleur",
            "n": "sb_v_double_arrow",
            "s": "sb_v_double_arrow",
            "e": "sb_h_double_arrow",
            "w": "sb_h_double_arrow",
            "ne": "top_right_corner",
            "nw": "top_left_corner",
            "se": "bottom_right_corner",
            "sw": "bottom_left_corner",
        }
        mode = self._hit_mode(event)
        c = cursors.get(mode, "arrow")
        try:
            self.box.configure(cursor=c)
            self.inner.configure(cursor=c)
        except Exception:
            pass

    def _on_press(self, event):
        self._mode = self._hit_mode(event)
        self._start = (event.x_root, event.y_root)
        self._geom0 = self._geom()

    def _on_drag(self, event):
        if not self._mode or not self._start or not self._geom0:
            return
        dx = event.x_root - self._start[0]
        dy = event.y_root - self._start[1]
        w0, h0, x0, y0 = self._geom0
        x, y, w, h = x0, y0, w0, h0
        m = self._mode
        if m == "move":
            x, y = x0 + dx, y0 + dy
        else:
            if "e" in m:
                w = max(MIN_W, w0 + dx)
            if "s" in m:
                h = max(MIN_H, h0 + dy)
            if "w" in m:
                nw = w0 - dx
                if nw < MIN_W:
                    x = x0 + w0 - MIN_W
                    w = MIN_W
                else:
                    x = x0 + dx
                    w = nw
            if "n" in m:
                nh = h0 - dy
                if nh < MIN_H:
                    y = y0 + h0 - MIN_H
                    h = MIN_H
                else:
                    y = y0 + dy
                    h = nh
        try:
            self.box.geometry("{0}x{1}+{2}+{3}".format(int(w), int(h), int(x), int(y)))
        except Exception:
            pass

    def _on_release(self, event):
        self._mode = None
        self._start = None
        self._geom0 = None

    def _confirm(self):
        w, h, x, y = self._geom()
        cfg = load_config()
        sap = cfg.setdefault("sap", {})
        sap["material_left"] = int(x)
        sap["material_top"] = int(y)
        sap["material_right"] = int(x + w)
        sap["material_bottom"] = int(y + h)
        rh = int(sap.get("row_height") or 28)
        if rh >= 5:
            sap["visible_rows"] = max(1, int(round(float(h) / float(rh))))
        save_config(cfg)
        self._log(
            u"자재코드영역 저장: ({0},{1})-({2},{3})".format(
                sap["material_left"],
                sap["material_top"],
                sap["material_right"],
                sap["material_bottom"],
            )
        )
        self._close()
        messagebox.showinfo(u"저장됨", u"자재코드 영역이 저장되었습니다.", parent=self.parent)
        if self.on_done:
            self.on_done()

    def _cancel(self):
        self._log(u"자재코드 영역 박스 취소")
        self._close()

    def _close(self):
        for w in (self.box, self.ctrl):
            try:
                if w:
                    w.destroy()
            except Exception:
                pass
        self.box = None
        self.ctrl = None

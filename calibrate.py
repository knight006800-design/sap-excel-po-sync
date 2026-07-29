# -*- coding: utf-8 -*-
"""SAP 화면 마우스 보정 — 주모니터 안내(빨간점 그림) + 노란＋마커."""
from __future__ import print_function

import os
import sys
import threading
import time

try:
    import tkinter as tk
    from tkinter import messagebox
except ImportError:
    import Tkinter as tk
    import tkMessageBox as messagebox

import pyautogui
from PIL import Image, ImageTk

from monitors import (
    enum_monitors,
    is_on_primary,
    place_window_primary_corner,
    primary_monitor,
    secondary_monitors,
)
from paths_util import app_dir, load_config, save_config


def guide_images_dir():
    local = os.path.join(app_dir(), "guide_images")
    if os.path.isdir(local):
        return local
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        bundled = os.path.join(sys._MEIPASS, "guide_images")
        if os.path.isdir(bundled):
            return bundled
    return local


# (키, 짧은제목, 짧은설명, 그림파일)
STEPS = [
    (
        u"material_tl",
        u"1/5  자재코드 · 왼쪽 위",
        u"SAP에서 빨간 점 위치 = 첫 행 자재코드 칸의 왼쪽 위 모서리",
        u"guide_material_tl.png",
    ),
    (
        u"material_br",
        u"2/5  자재코드 · 오른쪽 아래",
        u"SAP에서 빨간 점 위치 = 마지막 행 자재코드 칸의 오른쪽 아래 모서리",
        u"guide_material_br.png",
    ),
    (
        u"qty_x",
        u"3/5  오더수량 · 정중앙",
        u"SAP에서 빨간 점 위치 = 첫 행 오더수량 숫자 정중앙",
        u"guide_qty_x.png",
    ),
    (
        u"row1_y",
        u"4/5  첫 행 · 세로 중앙",
        u"SAP에서 빨간 점 위치 = 첫 데이터 행의 세로 한가운데",
        u"guide_row1_y.png",
    ),
    (
        u"row2_y",
        u"5/5  둘째 행 · 세로 중앙",
        u"SAP에서 빨간 점 위치 = 둘째 데이터 행의 세로 한가운데",
        u"guide_row2_y.png",
    ),
]


class MouseMarker(object):
    """마우스 옆 노란＋ (투명창 미사용 → 검은박스 방지)."""

    def __init__(self, root):
        self.win = tk.Toplevel(root)
        self.win.overrideredirect(True)
        self.win.attributes("-topmost", True)
        self.win.configure(bg="#FFEB3B")
        tk.Label(
            self.win,
            text=u"＋",
            font=("Arial", 20, "bold"),
            fg="#D50000",
            bg="#FFEB3B",
            bd=1,
            relief="solid",
        ).pack()
        self.win.geometry("+-400+-400")
        self._alive = True
        self.offset_x = 18
        self.offset_y = -34

    def move_to(self, x, y):
        if not self._alive:
            return
        try:
            self.win.geometry("+{0}+{1}".format(x + self.offset_x, y + self.offset_y))
        except Exception:
            pass

    def destroy(self):
        self._alive = False
        try:
            self.win.destroy()
        except Exception:
            pass


class CalibrateWizard(object):
    def __init__(self, parent, on_done=None, logger=None):
        self.parent = parent
        self.on_done = on_done
        self.logger = logger
        self.points = {}
        self._stop = False
        self.win = None
        self.marker = None
        self._photo = None
        self.monitors = enum_monitors()

    def _log(self, msg):
        if self.logger:
            self.logger.info(msg)

    def start(self):
        secs = secondary_monitors(self.monitors)
        prim = primary_monitor(self.monitors)
        msg = (
            u"안내 그림은 주모니터에만 뜹니다. SAP(부모니터)는 가리지 않습니다.\n"
            u"그림의 빨간 점 = 클릭할 위치입니다.\n"
            u"주모니터 클릭은 무시됩니다. 취소: ESC"
        )
        if secs:
            s = secs[0]
            msg = (
                u"주모니터=안내 / 부모니터=SAP\n"
                u"주: ({0},{1})-({2},{3})\n"
                u"부: ({4},{5})-({6},{7})\n\n"
            ).format(
                prim["left"],
                prim["top"],
                prim["right"],
                prim["bottom"],
                s["left"],
                s["top"],
                s["right"],
                s["bottom"],
            ) + msg
        messagebox.showinfo(u"보정 시작", msg, parent=self.parent)

        self.win = tk.Toplevel(self.parent)
        self.win.title(u"보정 안내 (주모니터) — 빨간 점 위치를 클릭")
        self.win.attributes("-topmost", True)
        # 그림이 보이도록 세로로 넉넉히, 주모니터 오른쪽에만
        place_window_primary_corner(self.win, width=480, height=520, side=u"right")
        self.win.configure(bg="#fff8e1")

        self.title_lbl = tk.Label(
            self.win,
            text=STEPS[0][1],
            font=("Malgun Gothic", 12, "bold"),
            bg="#fff8e1",
            fg="#b71c1c",
            wraplength=450,
            justify="left",
        )
        self.title_lbl.pack(anchor="w", padx=10, pady=(8, 2))

        self.detail = tk.Label(
            self.win,
            text=STEPS[0][2],
            font=("Malgun Gothic", 10),
            bg="#fff8e1",
            fg="#222",
            wraplength=450,
            justify="left",
        )
        self.detail.pack(anchor="w", padx=10, pady=2)

        self.img_label = tk.Label(self.win, bg="#dddddd")
        self.img_label.pack(padx=10, pady=6)

        self.coord = tk.Label(
            self.win,
            text=u"마우스: (-, -)",
            font=("Consolas", 10),
            bg="#fff8e1",
            fg="#1565c0",
        )
        self.coord.pack(anchor="w", padx=10)

        self.status = tk.Label(
            self.win,
            text=u"대기 중… 부모니터 SAP에서 빨간 점 위치를 클릭",
            font=("Malgun Gothic", 10, "bold"),
            bg="#fff8e1",
            fg="#0d47a1",
        )
        self.status.pack(anchor="w", padx=10, pady=(2, 8))

        self._show_guide(STEPS[0][3])
        self.marker = MouseMarker(self.parent)
        self._tick_marker()
        threading.Thread(target=self._run, daemon=True).start()

    def _show_guide(self, filename):
        path = os.path.join(guide_images_dir(), filename)
        if not os.path.isfile(path):
            self.img_label.config(
                image="",
                text=u"그림 없음: {0}\n(make_guide_images.py 실행 필요)".format(filename),
            )
            self._photo = None
            return
        try:
            img = Image.open(path)
            max_w = 450
            if img.width > max_w:
                ratio = float(max_w) / float(img.width)
                resample = Image.LANCZOS if hasattr(Image, "LANCZOS") else Image.ANTIALIAS
                img = img.resize(
                    (max_w, max(1, int(img.height * ratio))),
                    resample,
                )
            self._photo = ImageTk.PhotoImage(img)
            self.img_label.config(image=self._photo, text="")
        except Exception as e:
            self.img_label.config(image="", text=u"그림 오류: {0}".format(e))
            self._photo = None

    def _tick_marker(self):
        if self._stop or self.win is None:
            return
        try:
            x, y = pyautogui.position()
            if self.marker:
                self.marker.move_to(x, y)
            on_pri = is_on_primary(x, y, self.monitors)
            where = u"주모니터(무시)" if on_pri else u"부모니터(OK)"
            self.coord.config(text=u"마우스: ({0}, {1})  |  {2}".format(x, y, where))
        except Exception:
            pass
        try:
            self.win.after(30, self._tick_marker)
        except Exception:
            pass

    def _esc(self):
        import ctypes

        return bool(ctypes.windll.user32.GetAsyncKeyState(0x1B) & 0x8000)

    def _wait_click_on_sap(self):
        import ctypes

        VK_LBUTTON = 0x01
        has_secondary = bool(secondary_monitors(self.monitors))
        time.sleep(0.35)
        while not self._stop:
            if self._esc():
                return None
            if ctypes.windll.user32.GetAsyncKeyState(VK_LBUTTON) & 0x8000:
                while ctypes.windll.user32.GetAsyncKeyState(VK_LBUTTON) & 0x8000:
                    if self._esc():
                        return None
                    time.sleep(0.02)
                x, y = pyautogui.position()
                if has_secondary and is_on_primary(x, y, self.monitors):
                    self.win.after(
                        0,
                        lambda: self.status.config(
                            text=u"주모니터 클릭 무시 — SAP(부모니터)에서 클릭!",
                            fg="#c62828",
                        ),
                    )
                    time.sleep(0.2)
                    continue
                time.sleep(0.15)
                return (x, y)
            time.sleep(0.02)
        return None

    def _run(self):
        try:
            for key, title, detail, guide in STEPS:
                if self._stop:
                    break

                def _update(t=title, d=detail, g=guide):
                    self.title_lbl.config(text=t)
                    self.detail.config(text=d)
                    self._show_guide(g)
                    self.status.config(
                        text=u"대기 중… 부모니터 SAP에서 빨간 점 위치를 클릭",
                        fg="#0d47a1",
                    )

                self.win.after(0, _update)
                pt = self._wait_click_on_sap()
                if pt is None:
                    self._log(u"보정 취소(ESC)")
                    self.win.after(0, self._close)
                    return
                self.points[key] = pt
                self._log(u"보정 포인트 {0}: {1}".format(key, pt))
                self.win.after(
                    0,
                    lambda p=pt: self.status.config(
                        text=u"기록됨 ({0}, {1})".format(p[0], p[1]), fg="#2e7d32"
                    ),
                )
                time.sleep(0.4)

            cfg = load_config()
            sap = cfg.setdefault("sap", {})
            tl = self.points["material_tl"]
            br = self.points["material_br"]
            qx = self.points["qty_x"]
            r1 = self.points["row1_y"]
            r2 = self.points["row2_y"]
            sap["material_left"] = min(tl[0], br[0])
            sap["material_top"] = min(tl[1], br[1])
            sap["material_right"] = max(tl[0], br[0])
            sap["material_bottom"] = max(tl[1], br[1])
            sap["qty_center_x"] = qx[0]
            sap["first_row_y"] = r1[1]
            rh = abs(r2[1] - r1[1])
            if rh < 5:
                rh = 28
            sap["row_height"] = rh
            height = sap["material_bottom"] - sap["material_top"]
            sap["visible_rows"] = max(1, int(round(float(height) / float(rh))))
            save_config(cfg)
            self._log(
                u"보정 저장: material=({0},{1})-({2},{3}) qty_x={4} first_y={5} row_h={6} rows={7}".format(
                    sap["material_left"],
                    sap["material_top"],
                    sap["material_right"],
                    sap["material_bottom"],
                    sap["qty_center_x"],
                    sap["first_row_y"],
                    sap["row_height"],
                    sap["visible_rows"],
                )
            )
            self.win.after(
                0,
                lambda: messagebox.showinfo(
                    u"보정 완료",
                    u"저장됨. 행높이={0}, 예상행수={1}".format(
                        sap["row_height"], sap["visible_rows"]
                    ),
                    parent=self.parent,
                ),
            )
            if self.on_done:
                self.win.after(0, self.on_done)
        except Exception as e:
            self._log(u"보정 오류: {0}".format(e))
            if self.logger:
                self.logger.exception(u"보정")
            self.win.after(
                0,
                lambda: messagebox.showerror(u"보정 오류", str(e), parent=self.parent),
            )
        finally:
            self.win.after(0, self._close)

    def _close(self):
        self._stop = True
        try:
            if self.marker:
                self.marker.destroy()
                self.marker = None
        except Exception:
            pass
        try:
            if self.win:
                self.win.destroy()
                self.win = None
        except Exception:
            pass

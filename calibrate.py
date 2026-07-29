# -*- coding: utf-8 -*-
"""SAP 화면 마우스 보정 — 주모니터 안내 + 클릭통과 십자선 (그림 없음)."""
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
    place_window_primary_corner,
    primary_monitor,
    secondary_monitors,
)
from paths_util import load_config, save_config


# (키, 짧은제목, 정밀설명)
STEPS = [
    (
        u"material_tl",
        u"1/5  자재코드 · 왼쪽 위 모서리",
        u"【정확히】부모니터 SAP 표에서\n"
        u"첫 번째 데이터 행의 「자재코드」칸을 보세요.\n"
        u"예: 88512-P9000 글자가 들어있는 칸.\n\n"
        u"그 칸의 ★왼쪽 위 모서리★ (글자 시작점 바로 위-왼쪽 꼭짓점)을\n"
        u"십자선(+) 중심에 맞춘 뒤 클릭하세요.\n"
        u"칸 중앙이 아니라 모서리입니다.",
    ),
    (
        u"material_br",
        u"2/5  자재코드 · 오른쪽 아래 모서리",
        u"【정확히】화면에 보이는 마지막 데이터 행의\n"
        u"「자재코드」칸 ★오른쪽 아래 모서리★ 를 클릭하세요.\n\n"
        u"1번에서 찍은 왼쪽 위와 함께\n"
        u"자재코드 열 전체를 감싸는 사각형이 됩니다.\n"
        u"열 밖(자재명 쪽)으로 나가지 마세요.",
    ),
    (
        u"qty_x",
        u"3/5  오더수량 · 숫자 정중앙",
        u"【정확히】첫 데이터 행의 「오더수량」칸을 찾으세요.\n"
        u"(납품예정일 오른쪽, 숫자만 있는 칸. 예: 360)\n\n"
        u"그 숫자 ★정중앙★ 을 클릭하세요.\n"
        u"여기 클릭의 X좌표만 사용합니다.",
    ),
    (
        u"row1_y",
        u"4/5  첫 행 · 세로 중앙",
        u"【정확히】첫 번째 데이터 행에서\n"
        u"자재코드 글자 높이의 ★세로 한가운데★ 를 클릭하세요.\n\n"
        u"행의 위/아래 선이 아니라, 글자 중간 높이입니다.\n"
        u"여기 클릭의 Y좌표만 사용합니다.",
    ),
    (
        u"row2_y",
        u"5/5  둘째 행 · 세로 중앙",
        u"【정확히】바로 아래 ★두 번째 데이터 행★ 의\n"
        u"세로 한가운데를 클릭하세요.\n\n"
        u"1행과 2행 간격으로 행 높이를 계산합니다.\n"
        u"같은 열(자재코드 근처)에서 찍으면 정확합니다.",
    ),
]


def _make_clickthrough(hwnd):
    """창이 마우스 클릭을 가로채지 않게 (SAP 클릭 가능)."""
    import ctypes

    GWL_EXSTYLE = -20
    WS_EX_LAYERED = 0x00080000
    WS_EX_TRANSPARENT = 0x00000020
    WS_EX_TOOLWINDOW = 0x00000080
    user32 = ctypes.windll.user32
    try:
        style = user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
        user32.SetWindowLongW(
            hwnd,
            GWL_EXSTYLE,
            style | WS_EX_LAYERED | WS_EX_TRANSPARENT | WS_EX_TOOLWINDOW,
        )
    except Exception:
        pass


class Crosshair(object):
    """마우스 위치를 따라가는 빨간 십자선 (클릭 통과)."""

    def __init__(self, root):
        self.win = tk.Toplevel(root)
        self.win.overrideredirect(True)
        self.win.attributes("-topmost", True)
        try:
            self.win.attributes("-transparentcolor", "white")
        except Exception:
            pass
        self.win.configure(bg="white")
        self.size = 48
        self.canvas = tk.Canvas(
            self.win,
            width=self.size,
            height=self.size,
            bg="white",
            highlightthickness=0,
            bd=0,
        )
        self.canvas.pack()
        c = self.size // 2
        self.canvas.create_line(c, 0, c, self.size, fill="red", width=2)
        self.canvas.create_line(0, c, self.size, c, fill="red", width=2)
        self.canvas.create_oval(c - 4, c - 4, c + 4, c + 4, outline="red", width=2)
        self.win.geometry("{0}x{0}+-200+-200".format(self.size))
        self.win.update_idletasks()
        try:
            hwnd = self.win.winfo_id()
            # Tk on Windows: often need GetParent
            import ctypes

            parent = ctypes.windll.user32.GetParent(hwnd)
            _make_clickthrough(parent if parent else hwnd)
            _make_clickthrough(hwnd)
        except Exception:
            pass
        self._alive = True

    def move_to(self, x, y):
        if not self._alive:
            return
        half = self.size // 2
        try:
            self.win.geometry("{0}x{0}+{1}+{2}".format(self.size, x - half, y - half))
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
        self.cross = None
        self.monitors = enum_monitors()

    def _log(self, msg):
        if self.logger:
            self.logger.info(msg)

    def start(self):
        secs = secondary_monitors(self.monitors)
        prim = primary_monitor(self.monitors)
        if not secs:
            messagebox.showwarning(
                u"모니터 확인",
                u"부모니터(보조 모니터)가 감지되지 않았습니다.\n"
                u"듀얼모니터에서 SAP는 부모니터, 이 프로그램은 주모니터에 두세요.\n"
                u"그래도 진행은 가능합니다.",
                parent=self.parent,
            )
        else:
            s = secs[0]
            messagebox.showinfo(
                u"보정 시작",
                u"배치\n"
                u"· 주모니터(프로그램/안내): ({0},{1})-({2},{3})\n"
                u"· 부모니터(SAP): ({4},{5})-({6},{7})\n\n"
                u"안내 창은 주모니터에만 뜹니다. SAP를 가리지 않습니다.\n"
                u"빨간 십자선(+)을 SAP 목표에 맞추고 클릭하세요.\n"
                u"주모니터를 클릭하면 무시됩니다. 취소: ESC".format(
                    prim["left"],
                    prim["top"],
                    prim["right"],
                    prim["bottom"],
                    s["left"],
                    s["top"],
                    s["right"],
                    s["bottom"],
                ),
                parent=self.parent,
            )

        self.win = tk.Toplevel(self.parent)
        self.win.title(u"보정 안내 (주모니터)")
        self.win.attributes("-topmost", True)
        place_window_primary_corner(self.win, width=440, height=320, side=u"right")
        self.win.configure(bg="#fff8e1")

        self.title_lbl = tk.Label(
            self.win,
            text=STEPS[0][1],
            font=("Malgun Gothic", 12, "bold"),
            bg="#fff8e1",
            fg="#b71c1c",
            wraplength=410,
            justify="left",
        )
        self.title_lbl.pack(anchor="w", padx=12, pady=(10, 4))

        self.detail = tk.Label(
            self.win,
            text=STEPS[0][2],
            font=("Malgun Gothic", 10),
            bg="#fff8e1",
            fg="#222",
            wraplength=410,
            justify="left",
        )
        self.detail.pack(anchor="w", padx=12, pady=4)

        self.coord = tk.Label(
            self.win,
            text=u"마우스: (-, -)  |  부모니터에서 클릭",
            font=("Consolas", 10),
            bg="#fff8e1",
            fg="#1565c0",
        )
        self.coord.pack(anchor="w", padx=12, pady=4)

        self.status = tk.Label(
            self.win,
            text=u"대기 중… SAP(부모니터)에서 목표를 클릭하세요",
            font=("Malgun Gothic", 10, "bold"),
            bg="#fff8e1",
            fg="#0d47a1",
        )
        self.status.pack(anchor="w", padx=12, pady=(2, 10))

        tip = tk.Label(
            self.win,
            text=u"※ 이 노란 창은 주모니터 전용입니다. SAP 위로 옮기지 마세요.",
            font=("Malgun Gothic", 9),
            bg="#fff8e1",
            fg="#666",
        )
        tip.pack(anchor="w", padx=12, pady=(0, 8))

        self.cross = Crosshair(self.parent)
        self._tick_crosshair()
        threading.Thread(target=self._run, daemon=True).start()

    def _tick_crosshair(self):
        if self._stop or self.win is None:
            return
        try:
            x, y = pyautogui.position()
            if self.cross:
                self.cross.move_to(x, y)
            on_pri = is_on_primary(x, y, self.monitors)
            where = u"주모니터(클릭무시)" if on_pri else u"부모니터(클릭가능)"
            self.coord.config(text=u"마우스: ({0}, {1})  |  {2}".format(x, y, where))
        except Exception:
            pass
        try:
            self.win.after(30, self._tick_crosshair)
        except Exception:
            pass

    def _esc(self):
        import ctypes

        return bool(ctypes.windll.user32.GetAsyncKeyState(0x1B) & 0x8000)

    def _wait_click_on_sap(self):
        """부모니터(비주모니터) 클릭만 인정. 싱글모니터면 모든 클릭 인정."""
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
                    # 주모니터 클릭 무시
                    self.win.after(
                        0,
                        lambda: self.status.config(
                            text=u"주모니터 클릭은 무시됩니다. SAP(부모니터)에서 클릭!",
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
            for key, title, detail in STEPS:
                if self._stop:
                    break

                def _update(t=title, d=detail):
                    self.title_lbl.config(text=t)
                    self.detail.config(text=d)
                    self.status.config(
                        text=u"대기 중… SAP(부모니터)에서 목표를 클릭하세요",
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
                    u"저장됨.\n행높이={0}, 예상행수={1}\n이제 [동기화 실행]을 누르세요.".format(
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
            if self.cross:
                self.cross.destroy()
                self.cross = None
        except Exception:
            pass
        try:
            if self.win:
                self.win.destroy()
                self.win = None
        except Exception:
            pass

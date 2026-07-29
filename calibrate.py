# -*- coding: utf-8 -*-
"""SAP 화면 마우스 보정 마법사."""
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

from paths_util import load_config, save_config


STEPS = [
    (u"material_tl", u"1/5  자재코드 열의 '왼쪽 위'(첫 데이터 셀 왼쪽 상단)를 클릭하세요"),
    (u"material_br", u"2/5  자재코드 열의 '오른쪽 아래'(마지막 보이는 행 오른쪽 하단)를 클릭하세요"),
    (u"qty_x", u"3/5  오더수량 열의 아무 셀 '중앙'을 클릭하세요 (X좌표 사용)"),
    (u"row1_y", u"4/5  첫 번째 데이터 행의 중앙을 클릭하세요"),
    (u"row2_y", u"5/5  두 번째 데이터 행의 중앙을 클릭하세요 (행 높이 계산)"),
]


class CalibrateWizard(object):
    def __init__(self, parent, on_done=None, logger=None):
        self.parent = parent
        self.on_done = on_done
        self.logger = logger
        self.points = {}
        self._stop = False
        self.win = None

    def _log(self, msg):
        if self.logger:
            self.logger.info(msg)

    def start(self):
        messagebox.showinfo(
            u"SAP 화면 보정",
            u"보정 중에는 마우스로 SAP 화면을 순서대로 클릭합니다.\n"
            u"안내 창의 지시에 따라 진행하세요.\n"
            u"취소: ESC",
            parent=self.parent,
        )
        self.win = tk.Toplevel(self.parent)
        self.win.title(u"보정 진행")
        self.win.attributes("-topmost", True)
        self.win.geometry("520x120+200+50")
        self.label = tk.Label(self.win, text=STEPS[0][1], font=("Malgun Gothic", 11), wraplength=480)
        self.label.pack(expand=True, fill="both", padx=10, pady=10)
        self.status = tk.Label(self.win, text=u"클릭 대기 중...", fg="blue")
        self.status.pack(pady=4)
        threading.Thread(target=self._run, daemon=True).start()

    def _esc(self):
        import ctypes

        return bool(ctypes.windll.user32.GetAsyncKeyState(0x1B) & 0x8000)

    def _wait_click(self):
        """마우스 왼쪽 버튼이 눌렸다가 떼어질 때 좌표 반환."""
        import ctypes

        VK_LBUTTON = 0x01
        # 안내 클릭이 끝난 뒤 안정화
        time.sleep(0.4)
        while not self._stop:
            if self._esc():
                return None
            state = ctypes.windll.user32.GetAsyncKeyState(VK_LBUTTON)
            if state & 0x8000:
                # 누른 동안 대기
                while ctypes.windll.user32.GetAsyncKeyState(VK_LBUTTON) & 0x8000:
                    if self._esc():
                        return None
                    time.sleep(0.02)
                x, y = pyautogui.position()
                time.sleep(0.2)
                return (x, y)
            time.sleep(0.03)
        return None

    def _run(self):
        try:
            for key, text in STEPS:
                if self._stop:
                    break
                self.win.after(0, lambda t=text: self.label.config(text=t))
                self.win.after(0, lambda: self.status.config(text=u"클릭 대기 중...", fg="blue"))
                pt = self._wait_click()
                if pt is None:
                    self._log(u"보정 취소(ESC)")
                    self.win.after(0, self._close)
                    return
                self.points[key] = pt
                self._log(u"보정 포인트 {0}: {1}".format(key, pt))
                self.win.after(
                    0,
                    lambda p=pt: self.status.config(
                        text=u"기록됨: ({0}, {1})".format(p[0], p[1]), fg="green"
                    ),
                )
                time.sleep(0.35)

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
                    u"SAP 화면 좌표가 config.json에 저장되었습니다.\n"
                    u"행 높이={0}, 예상 행 수={1}".format(sap["row_height"], sap["visible_rows"]),
                    parent=self.parent,
                ),
            )
            if self.on_done:
                self.win.after(0, self.on_done)
        except Exception as e:
            self._log(u"보정 오류: {0}".format(e))
            if self.logger:
                self.logger.exception(u"보정")
            self.win.after(0, lambda: messagebox.showerror(u"보정 오류", str(e), parent=self.parent))
        finally:
            self.win.after(0, self._close)

    def _close(self):
        try:
            if self.win:
                self.win.destroy()
        except Exception:
            pass

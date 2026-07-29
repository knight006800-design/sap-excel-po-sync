# -*- coding: utf-8 -*-
"""SAP 오더수량 칸 마우스·키보드 입력."""
from __future__ import print_function

import time

import pyautogui


pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.05


def esc_pressed():
    import ctypes

    return bool(ctypes.windll.user32.GetAsyncKeyState(0x1B) & 0x8000)


def click_and_type_qty(x, y, qty_value, cfg, logger=None):
    """오더수량 셀 클릭 후 값 입력."""
    delay_click = float(cfg.get("delay_after_click", 0.25))
    delay_type = float(cfg.get("delay_after_type", 0.35))

    if logger:
        logger.info(u"클릭 입력: ({0},{1}) ← {2}".format(x, y, qty_value))

    pyautogui.click(x, y)
    time.sleep(delay_click)
    pyautogui.hotkey("ctrl", "a")
    time.sleep(0.08)
    text = str(int(qty_value) if isinstance(qty_value, float) and qty_value == int(qty_value) else qty_value)
    pyautogui.typewrite(text, interval=0.03)
    time.sleep(0.05)
    pyautogui.press("enter")
    time.sleep(delay_type)
    return True

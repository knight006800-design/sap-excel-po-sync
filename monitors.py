# -*- coding: utf-8 -*-
"""모니터 위치 유틸 — 주모니터(프로그램) / 부모니터(SAP)."""
from __future__ import print_function


def enum_monitors():
    """[{left,top,right,bottom,primary}, ...]"""
    monitors = []
    try:
        import win32api

        for h_mon, _hdc, _rect in win32api.EnumDisplayMonitors(None, None):
            info = win32api.GetMonitorInfo(h_mon)
            l, t, r, b = info["Monitor"]
            monitors.append(
                {
                    "left": int(l),
                    "top": int(t),
                    "right": int(r),
                    "bottom": int(b),
                    "primary": bool(info.get("Flags", 0) & 1),
                }
            )
    except Exception:
        import ctypes

        w = int(ctypes.windll.user32.GetSystemMetrics(0))
        h = int(ctypes.windll.user32.GetSystemMetrics(1))
        monitors = [
            {"left": 0, "top": 0, "right": w, "bottom": h, "primary": True}
        ]
    if not monitors:
        monitors = [{"left": 0, "top": 0, "right": 1920, "bottom": 1080, "primary": True}]
    # primary 표시가 없으면 좌상단 (0,0) 포함 모니터를 primary로
    if not any(m.get("primary") for m in monitors):
        for m in monitors:
            if m["left"] <= 0 <= m["right"] and m["top"] <= 0 <= m["bottom"]:
                m["primary"] = True
                break
        else:
            monitors[0]["primary"] = True
    return monitors


def primary_monitor(monitors=None):
    mons = monitors or enum_monitors()
    for m in mons:
        if m.get("primary"):
            return m
    return mons[0]


def secondary_monitors(monitors=None):
    mons = monitors or enum_monitors()
    secs = [m for m in mons if not m.get("primary")]
    return secs


def point_in_monitor(x, y, mon):
    return mon["left"] <= x < mon["right"] and mon["top"] <= y < mon["bottom"]


def is_on_primary(x, y, monitors=None):
    return point_in_monitor(x, y, primary_monitor(monitors))


def place_window_on_primary(win, width=720, height=520, margin=40):
    """Tk 창을 주모니터 안에 배치."""
    mon = primary_monitor()
    x = mon["left"] + margin
    y = mon["top"] + margin
    max_w = max(200, mon["right"] - mon["left"] - margin * 2)
    max_h = max(200, mon["bottom"] - mon["top"] - margin * 2)
    width = min(width, max_w)
    height = min(height, max_h)
    win.geometry("{0}x{1}+{2}+{3}".format(width, height, x, y))


def place_window_primary_corner(win, width=430, height=300, side=u"right"):
    """주모니터 모서리에 작은 안내창 배치."""
    mon = primary_monitor()
    mw = mon["right"] - mon["left"]
    mh = mon["bottom"] - mon["top"]
    width = min(width, mw - 20)
    height = min(height, mh - 20)
    if side == u"right":
        x = mon["right"] - width - 16
    else:
        x = mon["left"] + 16
    y = mon["top"] + 16
    win.geometry("{0}x{1}+{2}+{3}".format(width, height, x, y))


def place_window_on_secondary(win, width=480, height=520, margin=16):
    """보조 모니터가 있으면 그곳(안내창), 없으면 주모니터 모서리.
    SAP는 주모니터에 두고 안내는 서브에 두는 용도.
    """
    secs = secondary_monitors()
    if secs:
        mon = secs[0]
        w = min(width, max(200, mon["right"] - mon["left"] - margin * 2))
        h = min(height, max(200, mon["bottom"] - mon["top"] - margin * 2))
        x = mon["left"] + margin
        y = mon["top"] + margin
        win.geometry("{0}x{1}+{2}+{3}".format(w, h, x, y))
    else:
        place_window_primary_corner(win, width=width, height=height, side=u"right")

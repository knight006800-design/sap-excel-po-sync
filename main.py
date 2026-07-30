# -*- coding: utf-8 -*-
"""SAP–엑셀 자재 발주 프로그램 (웅이전용)."""
from __future__ import print_function

import os
import threading

try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
except ImportError:
    import Tkinter as tk
    import tkFileDialog as filedialog
    import tkMessageBox as messagebox
    import ttk

from datetime import datetime

from drive_check import DriveCheckLog
from monitors import place_window_on_primary
from paths_util import load_config, save_config
from sync_engine import SyncEngine


def today_unmatched_filename():
    """실행일 기준 기본 파일명: 20260730_"""
    return datetime.now().strftime("%Y%m%d_")


# White-first palette — flat surfaces, hairline borders, no gradients
C = {
    "bg": "#F6F6F7",
    "surface": "#FFFFFF",
    "surface_soft": "#FAFAFA",
    "border": "#E6E6E8",
    "border_strong": "#D4D4D8",
    "text": "#111113",
    "muted": "#73737A",
    "label": "#52525B",
    "accent": "#111113",
    "accent_hover": "#2A2A2E",
    "accent_press": "#000000",
    "accent_fg": "#FFFFFF",
    "accent_disabled": "#A1A1AA",
    "ghost_bg": "#FFFFFF",
    "ghost_press": "#F0F0F2",
    "input_bg": "#FFFFFF",
    "result_bg": "#F7FAF8",
    "err_bg": "#FEF3C7",
    "err_fg": "#78350F",
    "sel_bg": "#111113",
    "sel_fg": "#FFFFFF",
    "divider": "#ECECEE",
    "status_ok": "#166534",
    "status_err": "#B45309",
}

FONT = ("Segoe UI", 10)
FONT_SM = ("Segoe UI", 9)
FONT_B = ("Segoe UI", 10, "bold")
FONT_LABEL = ("Segoe UI", 9)
FONT_CTA = ("Segoe UI", 11, "bold")
FONT_MONO = ("Consolas", 10)
FONT_ELLIPSIS = ("Segoe UI", 12, "bold")


class SlimScrollbar(tk.Canvas):
    """화살표 없는 슬림 스크롤바."""

    def __init__(self, parent, command=None, width=7, bg=None, trough=None, thumb=None):
        self._command = command
        self._width = max(4, int(width))
        self._trough = trough or C["surface_soft"]
        self._thumb = thumb or "#C5C5CA"
        self._thumb_active = "#9A9AA1"
        self._lo = 0.0
        self._hi = 1.0
        self._drag_y = None
        self._drag_lo = 0.0
        tk.Canvas.__init__(
            self,
            parent,
            width=self._width,
            highlightthickness=0,
            bd=0,
            bg=self._trough,
            cursor="arrow",
        )
        self.bind("<Configure>", lambda e: self._redraw())
        self.bind("<ButtonPress-1>", self._on_press)
        self.bind("<B1-Motion>", self._on_drag)
        self.bind("<ButtonRelease-1>", self._on_release)
        self.bind("<Enter>", lambda e: self._set_thumb_color(self._thumb_active))
        self.bind("<Leave>", lambda e: self._set_thumb_color(self._thumb))

    def set(self, lo, hi):
        try:
            self._lo = float(lo)
            self._hi = float(hi)
        except Exception:
            self._lo, self._hi = 0.0, 1.0
        self._redraw()

    def _set_thumb_color(self, color):
        self._thumb_draw = color
        self._redraw()

    def _thumb_metrics(self):
        h = max(1, int(self.winfo_height()))
        pad = 2
        usable = max(1, h - pad * 2)
        span = max(0.0, min(1.0, self._hi) - max(0.0, min(1.0, self._lo)))
        if self._hi >= 1.0 and self._lo <= 0.0:
            return None
        th = max(18, int(usable * span))
        top = pad + int(max(0.0, min(1.0, self._lo)) * (usable - th))
        return pad, top, th, usable

    def _redraw(self):
        self.delete("all")
        color = getattr(self, "_thumb_draw", self._thumb)
        m = self._thumb_metrics()
        if m is None:
            return
        pad, top, th, _usable = m
        x0, x1 = 1, self._width - 1
        # 둥근 느낌의 얇은 트랙 썸
        self.create_rectangle(x0, top, x1, top + th, fill=color, outline="", tags="thumb")

    def _on_press(self, event):
        m = self._thumb_metrics()
        if m is None:
            return
        _pad, top, th, usable = m
        if top <= event.y <= top + th:
            self._drag_y = event.y
            self._drag_lo = self._lo
            return
        # 트랙 클릭 → 해당 위치로 점프
        ratio = (event.y - th / 2.0) / float(max(1, usable - th))
        ratio = max(0.0, min(1.0, ratio))
        if self._command:
            self._command("moveto", ratio)

    def _on_drag(self, event):
        if self._drag_y is None:
            return
        m = self._thumb_metrics()
        if m is None:
            return
        _pad, _top, th, usable = m
        delta = (event.y - self._drag_y) / float(max(1, usable - th))
        ratio = max(0.0, min(1.0, self._drag_lo + delta))
        if self._command:
            self._command("moveto", ratio)

    def _on_release(self, _event=None):
        self._drag_y = None


class PressButton(tk.Frame):
    """물리적으로 눌리는 느낌이 나는 flat CTA / ghost 버튼."""

    def __init__(
        self,
        parent,
        text,
        command=None,
        primary=False,
        padx=18,
        pady=8,
        font=None,
        **kwargs
    ):
        tk.Frame.__init__(self, parent, bg=parent.cget("bg"), **kwargs)
        self._command = command
        self._primary = primary
        self._enabled = True
        self._pressed = False
        self._text = text

        if primary:
            self._bg = C["accent"]
            self._fg = C["accent_fg"]
            self._hover = C["accent_hover"]
            self._press = C["accent_press"]
            self._disabled_bg = C["accent_disabled"]
            self._disabled_fg = C["accent_fg"]
            btn_font = font or FONT_CTA
        else:
            self._bg = C["ghost_bg"]
            self._fg = C["text"]
            self._hover = C["surface_soft"]
            self._press = C["ghost_press"]
            self._disabled_bg = C["surface_soft"]
            self._disabled_fg = C["muted"]
            btn_font = font or FONT

        border = C["border_strong"] if not primary else C["accent"]
        self._shell = tk.Frame(self, bg=border, bd=0, highlightthickness=0)
        self._shell.pack(fill="both", expand=True)
        self._inner = tk.Frame(self._shell, bg=self._bg, bd=0, highlightthickness=0)
        self._inner.pack(fill="both", expand=True, padx=1, pady=1)
        self._label = tk.Label(
            self._inner,
            text=text,
            bg=self._bg,
            fg=self._fg,
            font=btn_font,
            cursor="hand2",
            padx=padx,
            pady=pady,
        )
        self._label.pack()

        for w in (self._shell, self._inner, self._label):
            w.bind("<Enter>", self._on_enter)
            w.bind("<Leave>", self._on_leave)
            w.bind("<ButtonPress-1>", self._on_press)
            w.bind("<ButtonRelease-1>", self._on_release)

    def configure_text(self, text):
        self._text = text
        self._label.configure(text=text)

    def set_enabled(self, enabled):
        self._enabled = bool(enabled)
        if self._enabled:
            self._apply_colors(self._bg, self._fg)
            self._label.configure(cursor="hand2")
        else:
            self._apply_colors(self._disabled_bg, self._disabled_fg)
            self._label.configure(cursor="arrow")
            self._pressed = False
            self._inner.pack_configure(padx=1, pady=1)

    def _apply_colors(self, bg, fg):
        self._inner.configure(bg=bg)
        self._label.configure(bg=bg, fg=fg)
        if self._primary and self._enabled:
            self._shell.configure(bg=bg)
        elif self._primary and not self._enabled:
            self._shell.configure(bg=self._disabled_bg)

    def _on_enter(self, _event=None):
        if not self._enabled or self._pressed:
            return
        self._apply_colors(self._hover, self._fg)

    def _on_leave(self, _event=None):
        if not self._enabled:
            return
        self._pressed = False
        self._inner.pack_configure(padx=1, pady=1)
        self._apply_colors(self._bg, self._fg)

    def _on_press(self, _event=None):
        if not self._enabled:
            return
        self._pressed = True
        # 1px 아래로 밀려 물리적으로 눌리는 느낌
        self._inner.pack_configure(padx=1, pady=(2, 0))
        self._apply_colors(self._press, self._fg)

    def _on_release(self, _event=None):
        if not self._enabled:
            return
        was = self._pressed
        self._pressed = False
        self._inner.pack_configure(padx=1, pady=1)
        self._apply_colors(self._hover, self._fg)
        if was and self._command:
            self._command()


class App(object):
    def __init__(self):
        self.root = tk.Tk()
        self.root.title(u"웅이 자재 발주 프로그램")
        self.root.minsize(600, 520)
        self.root.configure(bg=C["bg"])
        place_window_on_primary(self.root, width=680, height=580, margin=28, center=True)

        self.cfg = load_config()
        self._running = False

        self._setup_style()
        self._build()

    def _setup_style(self):
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("TFrame", background=C["bg"])
        style.configure("TLabel", background=C["bg"], foreground=C["text"], font=FONT)
        style.configure(
            "Card.TLabel", background=C["surface"], foreground=C["text"], font=FONT
        )
        style.configure(
            "Muted.TLabel",
            background=C["bg"],
            foreground=C["muted"],
            font=FONT_SM,
        )
        style.configure(
            "Section.TLabel",
            background=C["surface"],
            foreground=C["text"],
            font=FONT_B,
        )
        style.configure(
            "Field.TLabel",
            background=C["surface"],
            foreground=C["label"],
            font=FONT_LABEL,
        )
        style.configure(
            "TEntry",
            fieldbackground=C["input_bg"],
            foreground=C["text"],
            insertcolor=C["text"],
            padding=6,
            bordercolor=C["border"],
            lightcolor=C["border"],
            darkcolor=C["border"],
        )
        style.map(
            "TEntry",
            bordercolor=[("focus", C["text"])],
            lightcolor=[("focus", C["text"])],
            darkcolor=[("focus", C["text"])],
        )

    def _surface(self, parent, padx=14, pady=12):
        outer = tk.Frame(parent, bg=C["border"], bd=0, highlightthickness=0)
        inner = tk.Frame(outer, bg=C["surface"], bd=0, highlightthickness=0)
        inner.pack(fill="both", expand=True, padx=1, pady=1)
        pad = tk.Frame(inner, bg=C["surface"])
        pad.pack(fill="both", expand=True, padx=padx, pady=pady)
        return outer, pad

    def _text_box(self, parent, bg=None, width=28):
        wrap = tk.Frame(parent, bg=C["border"], bd=0, highlightthickness=0)
        inner = tk.Frame(wrap, bg=bg or C["input_bg"], bd=0, highlightthickness=0)
        inner.pack(fill="both", expand=True, padx=1, pady=1)

        box = tk.Text(
            inner,
            height=16,
            width=width,
            font=FONT_MONO,
            bg=bg or C["input_bg"],
            fg=C["text"],
            insertbackground=C["text"],
            selectbackground=C["sel_bg"],
            selectforeground=C["sel_fg"],
            relief="flat",
            highlightthickness=0,
            bd=0,
            padx=10,
            pady=6,
            spacing1=0,
            spacing2=0,
            spacing3=1,
            wrap="none",
        )
        try:
            box.configure(inactiveselectbackground=C["sel_bg"])
        except tk.TclError:
            pass

        scroll = SlimScrollbar(
            inner,
            command=box.yview,
            width=6,
            trough=bg or C["input_bg"],
            thumb="#C9C9CE",
        )
        box.configure(yscrollcommand=scroll.set)

        box.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y", padx=(0, 2), pady=4)

        try:
            box.tag_configure("rowline", underline=True, underlinefg="#C4C4C8")
        except tk.TclError:
            box.tag_configure("rowline", underline=True)
        box.bind("<<Modified>>", lambda e, b=box: self._on_box_modified(b))
        box.bind("<Control-a>", lambda e, b=box: self._select_content(b))
        box.bind("<Control-A>", lambda e, b=box: self._select_content(b))
        # 마우스 휠
        box.bind(
            "<MouseWheel>",
            lambda e, b=box: b.yview_scroll(int(-1 * (e.delta / 120)), "units"),
        )
        wrap.pack(fill="both", expand=True, pady=(8, 0))
        return box

    def _select_content(self, box):
        text = box.get("1.0", "end-1c").rstrip("\r\n")
        box.tag_remove("sel", "1.0", "end")
        if text:
            box.tag_add("sel", "1.0", "1.0+{0}c".format(len(text)))
        return "break"

    def _on_box_modified(self, box):
        try:
            if not box.edit_modified():
                return
            box.edit_modified(False)
            self._apply_row_lines(box)
        except Exception:
            pass

    def _apply_row_lines(self, box):
        try:
            box.tag_remove("rowline", "1.0", "end")
            last = int(float(box.index("end-1c").split(".")[0]))
            for line in range(1, last + 1):
                start = "{0}.0".format(line)
                end = "{0}.end".format(line)
                content = box.get(start, end)
                if not content.strip():
                    continue
                box.tag_add("rowline", start, end)
            try:
                box.tag_raise("err")
            except Exception:
                pass
        except Exception:
            pass

    def _build(self):
        wrap = tk.Frame(self.root, bg=C["bg"])
        wrap.pack(fill="both", expand=True, padx=20, pady=16)

        # Settings surface
        settings, body = self._surface(wrap, padx=14, pady=12)
        settings.pack(fill="x", pady=(0, 10))

        # 경로 행 (폴더와 동일: 앞쪽 라벨)
        file_row = tk.Frame(body, bg=C["surface"])
        file_row.pack(fill="x", pady=(0, 10))
        tk.Label(
            file_row, text=u"경로", bg=C["surface"], fg=C["muted"], font=FONT_SM, width=5
        ).pack(side="left")
        self.excel_var = tk.StringVar(value=self.cfg.get("excel_path") or "")
        ttk.Entry(file_row, textvariable=self.excel_var).pack(
            side="left", fill="x", expand=True, ipady=3
        )
        PressButton(
            file_row,
            text=u"···",
            command=self.browse_excel,
            padx=10,
            pady=4,
            font=FONT_ELLIPSIS,
        ).pack(side="left", padx=(8, 0))

        # Thin divider
        tk.Frame(body, bg=C["divider"], height=1).pack(fill="x", pady=(0, 10))

        # Save location
        tk.Label(
            body,
            text=u"발주 파일 저장위치",
            bg=C["surface"],
            fg=C["label"],
            font=FONT_LABEL,
        ).pack(anchor="w")

        dir_row = tk.Frame(body, bg=C["surface"])
        dir_row.pack(fill="x", pady=(4, 6))
        tk.Label(
            dir_row, text=u"폴더", bg=C["surface"], fg=C["muted"], font=FONT_SM, width=5
        ).pack(side="left")
        self.unmatched_dir_var = tk.StringVar(value=self.cfg.get("unmatched_dir") or "")
        ttk.Entry(dir_row, textvariable=self.unmatched_dir_var).pack(
            side="left", fill="x", expand=True, ipady=3
        )
        PressButton(
            dir_row,
            text=u"···",
            command=self.browse_unmatched_dir,
            padx=10,
            pady=4,
            font=FONT_ELLIPSIS,
        ).pack(side="left", padx=(8, 0))

        name_row = tk.Frame(body, bg=C["surface"])
        name_row.pack(fill="x")
        tk.Label(
            name_row,
            text=u"파일명",
            bg=C["surface"],
            fg=C["muted"],
            font=FONT_SM,
            width=5,
        ).pack(side="left")
        self.unmatched_name_var = tk.StringVar(value=today_unmatched_filename())
        ttk.Entry(name_row, textvariable=self.unmatched_name_var).pack(
            side="left", fill="x", expand=True, ipady=3
        )

        # Action — extract right, status left
        act = tk.Frame(wrap, bg=C["bg"])
        act.pack(fill="x", pady=(0, 10))
        self.status_var = tk.StringVar(value=u"")
        self.status_label = tk.Label(
            act,
            textvariable=self.status_var,
            bg=C["bg"],
            fg=C["muted"],
            font=FONT_SM,
        )
        self.status_label.pack(side="left")
        self.extract_btn = PressButton(
            act,
            text=u"추출",
            command=self.do_run,
            primary=True,
            padx=28,
            pady=9,
        )
        self.extract_btn.pack(side="right")

        # Main panes 3:2
        paste_wrap = tk.Frame(wrap, bg=C["bg"])
        paste_wrap.pack(fill="both", expand=True)
        paste_wrap.columnconfigure(0, weight=3, uniform="paste")
        paste_wrap.columnconfigure(1, weight=2, uniform="paste")
        paste_wrap.rowconfigure(0, weight=1)

        self.code_title_var = tk.StringVar(value=u"SAP 자재코드")
        self.code_box = self._make_col(
            paste_wrap,
            0,
            self.code_title_var,
            u"비우기",
            self.clear_codes,
            padx=(0, 6),
            title_is_var=True,
            box_width=34,
        )
        self.code_box.tag_configure(
            "err", background=C["err_bg"], foreground=C["err_fg"]
        )
        self.code_box.tag_configure(
            "err_sel", background=C["sel_bg"], foreground=C["sel_fg"]
        )

        self.result_box = self._make_col(
            paste_wrap,
            1,
            u"오더수량",
            u"전체선택",
            self.select_result,
            padx=(6, 0),
            bg=C["result_bg"],
            box_width=22,
        )

    def _make_col(
        self,
        parent,
        col,
        title,
        btn_text,
        btn_cmd,
        padx=0,
        bg=None,
        title_is_var=False,
        box_width=28,
    ):
        card, body = self._surface(parent, padx=12, pady=10)
        card.grid(row=0, column=col, sticky="nsew", padx=padx)
        hdr = tk.Frame(body, bg=C["surface"])
        hdr.pack(fill="x")
        if btn_text and btn_cmd:
            PressButton(
                hdr, text=btn_text, command=btn_cmd, padx=10, pady=3
            ).pack(side="right")
        # 왼쪽 정렬 (fill/expand 시 Label 기본 center 방지)
        if title_is_var:
            tk.Label(
                hdr,
                textvariable=title,
                bg=C["surface"],
                fg=C["text"],
                font=FONT_B,
                anchor="w",
                justify="left",
            ).pack(side="left", anchor="w")
        else:
            tk.Label(
                hdr,
                text=title,
                bg=C["surface"],
                fg=C["text"],
                font=FONT_B,
                anchor="w",
                justify="left",
            ).pack(side="left", anchor="w")
        box = self._text_box(body, bg=bg, width=box_width)
        return box

    def _set_status(self, text, kind=u"muted"):
        colors = {
            u"muted": C["muted"],
            u"ok": C["status_ok"],
            u"err": C["status_err"],
            u"busy": C["text"],
        }
        self.status_var.set(text or u"")
        self.status_label.configure(fg=colors.get(kind, C["muted"]))

    def _set_running(self, running):
        self._running = running
        if running:
            self.extract_btn.configure_text(u"추출 중…")
            self.extract_btn.set_enabled(False)
            self._set_status(u"처리 중", u"busy")
            self.root.configure(cursor="watch")
        else:
            self.extract_btn.configure_text(u"추출")
            self.extract_btn.set_enabled(True)
            self.root.configure(cursor="")

    def clear_codes(self):
        self.code_box.tag_remove("err", "1.0", "end")
        self.code_box.delete("1.0", "end")
        self.code_title_var.set(u"SAP 자재코드")
        self._apply_row_lines(self.code_box)
        self._set_status(u"")

    def select_result(self):
        text = self.result_box.get("1.0", "end-1c").rstrip("\r\n")
        self.result_box.tag_remove("sel", "1.0", "end")
        if text:
            self.result_box.tag_add("sel", "1.0", "1.0+{0}c".format(len(text)))
        self.result_box.focus_set()
        try:
            self.root.clipboard_clear()
            self.root.clipboard_append(text)
        except Exception:
            pass

    def _show_results(self, stats):
        rows = stats.get("result_rows") or []
        self.code_box.tag_remove("err", "1.0", "end")
        self.code_box.delete("1.0", "end")
        self.result_box.delete("1.0", "end")

        displays = []
        qtys = []
        err_idx = []
        err_count = int(stats.get("missing_count") or 0)
        for i, r in enumerate(rows):
            code = r.get("code") or u""
            qtys.append(r.get("result_qty_text", u""))
            if code and not r.get("ok"):
                err_idx.append(i)
                sug = r.get("suggestions") or []
                if sug:
                    ratio, scode = sug[0]
                    displays.append(
                        u"{0}  → 유사품번 {1} ({2}%)".format(
                            code, scode, int(round(ratio * 100))
                        )
                    )
                else:
                    displays.append(code)
            else:
                displays.append(code)

        if err_count > 0:
            self.code_title_var.set(u"SAP 자재코드  ·  오류 {0}개".format(err_count))
            self._set_status(u"오류 {0}개".format(err_count), u"err")
        else:
            self.code_title_var.set(u"SAP 자재코드")
            matched = int(stats.get("matched") or 0)
            self._set_status(u"완료  ·  {0}건".format(matched), u"ok")

        if displays:
            self.code_box.insert("1.0", u"\n".join(displays))
        if qtys:
            self.result_box.insert("1.0", u"\n".join(qtys))

        for i in err_idx:
            start = "{0}.0".format(i + 1)
            end = "{0}.end".format(i + 1)
            self.code_box.tag_add("err", start, end)

        self._apply_row_lines(self.code_box)
        self._apply_row_lines(self.result_box)
        try:
            self.code_box.tag_raise("err")
        except Exception:
            pass

    def browse_excel(self):
        path = filedialog.askopenfilename(
            title=u"구매오더 엑셀 선택",
            filetypes=[(u"Excel", "*.xls;*.xlsx"), (u"모든 파일", "*.*")],
        )
        if path:
            self.excel_var.set(path)
            self.cfg = load_config()
            self.cfg["excel_path"] = path
            save_config(self.cfg)

    def browse_unmatched_dir(self):
        path = filedialog.askdirectory(title=u"발주 파일 저장 폴더")
        if path:
            self.unmatched_dir_var.set(path)
            self._save_unmatched_settings()

    def _save_unmatched_settings(self):
        cfg = load_config()
        cfg["excel_path"] = (self.excel_var.get() or "").strip()
        cfg["unmatched_dir"] = (self.unmatched_dir_var.get() or "").strip()
        name = (self.unmatched_name_var.get() or "").strip()
        if not name:
            name = today_unmatched_filename()
            self.unmatched_name_var.set(name)
        cfg["unmatched_filename"] = name
        save_config(cfg)
        self.cfg = cfg

    def _resolve_unmatched_path(self, excel_path):
        folder = (self.unmatched_dir_var.get() or "").strip()
        if not folder:
            folder = os.path.dirname(os.path.abspath(excel_path))
        name = (self.unmatched_name_var.get() or "").strip()
        if not name:
            name = today_unmatched_filename()
            self.unmatched_name_var.set(name)
        base, ext = os.path.splitext(name)
        if not ext:
            ext = os.path.splitext(excel_path)[1] or u".xls"
            name = base + ext
        return os.path.join(folder, name)

    def _box_text(self, box):
        raw = box.get("1.0", "end")
        lines = []
        for line in raw.splitlines():
            s = line.strip()
            if not s or s.startswith("#"):
                continue
            # 유사품번 안내 줄이면 코드만 재사용
            if u"→" in s:
                s = s.split(u"→", 1)[0].strip()
            lines.append(s)
        return u"\n".join(lines)

    def do_run(self):
        if self._running:
            return
        path = self.excel_var.get().strip()
        if not path:
            messagebox.showerror(u"엑셀 필요", u"엑셀 파일을 먼저 연결하세요.")
            return
        if not os.path.isfile(path):
            messagebox.showerror(
                u"파일 없음", u"엑셀 파일을 찾을 수 없습니다:\n{0}".format(path)
            )
            return

        codes = self._box_text(self.code_box)
        if not codes.strip():
            messagebox.showerror(u"코드 필요", u"SAP 자재코드를 Ctrl+V 하세요.")
            return

        self._save_unmatched_settings()
        unmatched_path = self._resolve_unmatched_path(path)

        self._set_running(True)
        threading.Thread(
            target=self._run_thread,
            args=(path, codes, unmatched_path),
            daemon=True,
        ).start()

    def _run_thread(self, path, codes, unmatched_path):
        logger = DriveCheckLog(gui_callback=None)
        try:
            engine = SyncEngine(config=load_config(), logger=logger)
            stats = engine.run(
                excel_path=path,
                codes_text=codes,
                unmatched_path=unmatched_path,
            )
            logger.finish(u"추출 완료")
            if stats:
                self.root.after(0, lambda: self._show_results(stats))
                self.root.after(0, lambda: self._finish_dialog(stats))
        except Exception as e:
            logger.exception(u"추출")
            self.root.after(
                0, lambda: messagebox.showerror(u"오류", u"{0}".format(e))
            )
            self.root.after(0, lambda: self._set_status(u"실패", u"err"))
        finally:
            self.root.after(0, lambda: self._set_running(False))

    def _finish_dialog(self, stats):
        count = int(stats.get("missing_count") or 0)
        if count > 0:
            messagebox.showerror(
                u"오류",
                u"SAP 자재코드 오류 : {0}개".format(count),
            )

    def run(self):
        self.root.mainloop()


def main():
    App().run()


if __name__ == "__main__":
    main()

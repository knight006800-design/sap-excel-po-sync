# -*- coding: utf-8 -*-
"""SAP–엑셀 자재 발주 프로그램 (웅이전용)."""
from __future__ import print_function

import os
import threading

try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, scrolledtext, ttk
except ImportError:
    import Tkinter as tk
    import tkFileDialog as filedialog
    import tkMessageBox as messagebox
    import ScrolledText as scrolledtext
    import ttk

from datetime import datetime

from drive_check import DriveCheckLog
from monitors import place_window_on_primary
from paths_util import load_config, save_config
from sync_engine import SyncEngine


def today_unmatched_filename():
    """실행일 기준 기본 파일명: 20260730_"""
    return datetime.now().strftime("%Y%m%d_")


C = {
    "bg": "#F4F5F7",
    "surface": "#FFFFFF",
    "border": "#E2E4E8",
    "text": "#1A1D23",
    "muted": "#6B7280",
    "accent": "#2563EB",
    "accent_fg": "#FFFFFF",
    "danger": "#DC2626",
    "danger_bg": "#FEF2F2",
    "input_bg": "#FFFFFF",
    "result_bg": "#ECFDF5",
    "err_bg": "#FBBF24",
    "err_fg": "#111827",
    "sel_bg": "#1D4ED8",
    "sel_fg": "#FFFFFF",
}

FONT = ("Segoe UI", 10)
FONT_B = ("Segoe UI", 10, "bold")
FONT_H = ("Segoe UI", 16, "bold")
FONT_MONO = ("Consolas", 10)


class App(object):
    def __init__(self):
        self.root = tk.Tk()
        self.root.title(u"웅이 자재 발주 프로그램")
        self.root.minsize(480, 520)
        self.root.configure(bg=C["bg"])
        place_window_on_primary(self.root, width=520, height=580, margin=28)

        self.cfg = load_config()
        self._stop = False
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
            "CardMuted.TLabel",
            background=C["surface"],
            foreground=C["muted"],
            font=("Segoe UI", 9),
        )
        style.configure(
            "Title.TLabel", background=C["bg"], foreground=C["text"], font=FONT_H
        )
        style.configure(
            "Section.TLabel",
            background=C["surface"],
            foreground=C["text"],
            font=FONT_B,
        )
        style.configure(
            "TEntry",
            fieldbackground=C["input_bg"],
            foreground=C["text"],
            insertcolor=C["text"],
            padding=8,
        )
        style.configure(
            "Primary.TButton",
            background=C["accent"],
            foreground=C["accent_fg"],
            font=FONT_B,
            padding=(16, 10),
            borderwidth=0,
        )
        style.map(
            "Primary.TButton",
            background=[("active", "#1D4ED8"), ("disabled", "#93C5FD")],
        )
        style.configure(
            "Ghost.TButton",
            background=C["surface"],
            foreground=C["text"],
            font=FONT,
            padding=(12, 8),
            borderwidth=1,
        )
        style.map("Ghost.TButton", background=[("active", C["border"])])
        style.configure(
            "Danger.TButton",
            background=C["danger_bg"],
            foreground=C["danger"],
            font=FONT,
            padding=(12, 8),
            borderwidth=0,
        )

    def _card(self, parent, padx=16, pady=14):
        outer = tk.Frame(parent, bg=C["border"], bd=0, highlightthickness=0)
        inner = tk.Frame(outer, bg=C["surface"], bd=0, highlightthickness=0)
        inner.pack(fill="both", expand=True, padx=1, pady=1)
        pad = tk.Frame(inner, bg=C["surface"])
        pad.pack(fill="both", expand=True, padx=padx, pady=pady)
        return outer, pad

    def _text_box(self, parent, bg=None):
        box = scrolledtext.ScrolledText(
            parent,
            height=16,
            width=26,
            font=FONT_MONO,
            bg=bg or C["input_bg"],
            fg=C["text"],
            insertbackground=C["text"],
            selectbackground=C["sel_bg"],
            selectforeground=C["sel_fg"],
            relief="flat",
            highlightthickness=1,
            highlightbackground=C["border"],
            highlightcolor=C["accent"],
            bd=0,
            padx=8,
            pady=4,
            spacing1=1,
            spacing2=0,
            spacing3=2,
        )
        try:
            box.configure(inactiveselectbackground=C["sel_bg"])
        except tk.TclError:
            pass
        # 줄 구분: 선명한 얇은 실선
        try:
            box.tag_configure("rowline", underline=True, underlinefg="#64748B")
        except tk.TclError:
            box.tag_configure("rowline", underline=True)
        box.bind("<<Modified>>", lambda e, b=box: self._on_box_modified(b))
        box.bind("<Control-a>", lambda e, b=box: self._select_content(b))
        box.bind("<Control-A>", lambda e, b=box: self._select_content(b))
        return box

    def _select_content(self, box):
        """맨 아래 빈 줄 없이 내용만 전체 선택."""
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
        """각 줄 아래 얇은 구분선 태그 적용 (빈 줄 제외)."""
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
        wrap = ttk.Frame(self.root, style="TFrame")
        wrap.pack(fill="both", expand=True, padx=20, pady=16)

        top_row = tk.Frame(wrap, bg=C["bg"])
        top_row.pack(fill="x", pady=(0, 8))
        top_row.columnconfigure(0, weight=1)
        top_row.columnconfigure(1, weight=1)

        excel_card, excel = self._card(top_row, padx=12, pady=10)
        excel_card.grid(row=0, column=0, sticky="nsew", padx=(0, 4))
        ttk.Label(excel, text=u"파일", style="Section.TLabel").pack(anchor="w")
        row = tk.Frame(excel, bg=C["surface"])
        row.pack(fill="x", pady=(6, 0))
        self.excel_var = tk.StringVar(value=self.cfg.get("excel_path") or "")
        ttk.Entry(row, textvariable=self.excel_var).pack(
            side="left", fill="x", expand=True, ipady=3
        )
        ttk.Button(
            row, text=u"경로", style="Ghost.TButton", command=self.browse_excel
        ).pack(side="left", padx=(6, 0))

        um_card, um = self._card(top_row, padx=12, pady=10)
        um_card.grid(row=0, column=1, sticky="nsew", padx=(4, 0))
        ttk.Label(um, text=u"발주 파일 저장위치", style="Section.TLabel").pack(
            anchor="w"
        )
        dir_row = tk.Frame(um, bg=C["surface"])
        dir_row.pack(fill="x", pady=(6, 4))
        ttk.Label(dir_row, text=u"폴더", style="Card.TLabel", width=5).pack(side="left")
        self.unmatched_dir_var = tk.StringVar(value=self.cfg.get("unmatched_dir") or "")
        ttk.Entry(dir_row, textvariable=self.unmatched_dir_var).pack(
            side="left", fill="x", expand=True, ipady=3
        )
        ttk.Button(
            dir_row,
            text=u"지정",
            style="Ghost.TButton",
            command=self.browse_unmatched_dir,
        ).pack(side="left", padx=(6, 0))
        name_row = tk.Frame(um, bg=C["surface"])
        name_row.pack(fill="x")
        ttk.Label(name_row, text=u"파일명", style="Card.TLabel", width=6).pack(
            side="left"
        )
        self.unmatched_name_var = tk.StringVar(value=today_unmatched_filename())
        ttk.Entry(name_row, textvariable=self.unmatched_name_var).pack(
            side="left", fill="x", expand=True, ipady=3
        )

        act = tk.Frame(wrap, bg=C["bg"])
        act.pack(fill="x", pady=(0, 8))
        ttk.Button(act, text=u"중지", style="Danger.TButton", command=self.do_stop).pack(
            side="right"
        )
        ttk.Button(
            act, text=u"추출", style="Primary.TButton", command=self.do_run
        ).pack(side="right", padx=(0, 8))

        paste_wrap = tk.Frame(wrap, bg=C["bg"])
        paste_wrap.pack(fill="both", expand=True, anchor="w")
        paste_wrap.columnconfigure(0, weight=1)
        paste_wrap.columnconfigure(1, weight=1)
        paste_wrap.rowconfigure(0, weight=1)

        self.code_title_var = tk.StringVar(value=u"SAP 자재코드")
        self.code_box = self._make_col(
            paste_wrap,
            0,
            self.code_title_var,
            u"비우기",
            self.clear_codes,
            padx=(0, 4),
            title_is_var=True,
        )
        self.code_box.tag_configure(
            "err",
            background=C["err_bg"],
            foreground=C["err_fg"],
        )
        self.code_box.tag_configure(
            "err_sel",
            background=C["sel_bg"],
            foreground=C["sel_fg"],
        )

        self.result_box = self._make_col(
            paste_wrap,
            1,
            u"오더수량(SAP 붙여넣기)",
            u"전체선택",
            self.select_result,
            padx=(4, 0),
            bg=C["result_bg"],
        )

    def _make_col(
        self, parent, col, title, btn_text, btn_cmd, padx=0, bg=None, title_is_var=False
    ):
        card, body = self._card(parent, padx=10, pady=8)
        card.grid(row=0, column=col, sticky="nsew", padx=padx)
        hdr = tk.Frame(body, bg=C["surface"])
        hdr.pack(fill="x")
        if title_is_var:
            ttk.Label(hdr, textvariable=title, style="Section.TLabel").pack(side="left")
        else:
            ttk.Label(hdr, text=title, style="Section.TLabel").pack(side="left")
        if btn_text and btn_cmd:
            ttk.Button(
                hdr, text=btn_text, style="Ghost.TButton", command=btn_cmd
            ).pack(side="right")
        box = self._text_box(body, bg=bg)
        box.pack(fill="both", expand=True, pady=(6, 0))
        return box

    def clear_codes(self):
        self.code_box.tag_remove("err", "1.0", "end")
        self.code_box.delete("1.0", "end")
        self.code_title_var.set(u"SAP 자재코드")
        self._apply_row_lines(self.code_box)

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
            self.code_title_var.set(u"SAP 자재코드  오류 {0}개".format(err_count))
        else:
            self.code_title_var.set(u"SAP 자재코드")

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

    def do_stop(self):
        self._stop = True

    def _box_text(self, box):
        raw = box.get("1.0", "end")
        lines = []
        for line in raw.splitlines():
            s = line.strip()
            if not s or s.startswith("#"):
                continue
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

        self._stop = False
        self._running = True
        threading.Thread(
            target=self._run_thread,
            args=(path, codes, unmatched_path),
            daemon=True,
        ).start()

    def _run_thread(self, path, codes, unmatched_path):
        logger = DriveCheckLog(gui_callback=None)
        try:
            engine = SyncEngine(
                config=load_config(),
                logger=logger,
                stop_flag=lambda: self._stop,
            )
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
        finally:
            self._running = False

    def _finish_dialog(self, stats):
        count = int(stats.get("missing_count") or 0)
        if count > 0:
            messagebox.showerror(
                u"오류",
                u"SAP 자재코드 오류 : {0}개".format(count),
            )
            return

    def run(self):
        self.root.mainloop()


def main():
    App().run()


if __name__ == "__main__":
    main()

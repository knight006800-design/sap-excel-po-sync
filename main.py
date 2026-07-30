# -*- coding: utf-8 -*-
"""SAP–엑셀 구매오더 수량 비교 (웅이전용)."""
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
from paths_util import app_dir, load_config, save_config
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
    "log_bg": "#111827",
    "log_fg": "#E5E7EB",
    "chip_bg": "#EEF2FF",
    "chip_fg": "#1E3A8A",
}

FONT = ("Segoe UI", 10)
FONT_B = ("Segoe UI", 10, "bold")
FONT_H = ("Segoe UI", 16, "bold")
FONT_S = ("Segoe UI", 9)
FONT_MONO = ("Consolas", 10)


class App(object):
    def __init__(self):
        self.root = tk.Tk()
        self.root.title(u"웅이전용 — 구매오더 수량 비교")
        self.root.minsize(900, 620)
        self.root.configure(bg=C["bg"])
        place_window_on_primary(self.root, width=960, height=680, margin=28)

        self.cfg = load_config()
        self._stop = False
        self._running = False

        self._setup_style()
        self._build()

        self.append_log(u"프로그램 폴더: {0}".format(app_dir()))
        if self.excel_var.get():
            self.append_log(u"연결된 엑셀: {0}".format(self.excel_var.get()))

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
            "Muted.TLabel", background=C["bg"], foreground=C["muted"], font=FONT_S
        )
        style.configure(
            "CardMuted.TLabel",
            background=C["surface"],
            foreground=C["muted"],
            font=FONT_S,
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

    def _text_box(self, parent, bg=None, width=None):
        kw = dict(
            height=14,
            font=FONT_MONO,
            bg=bg or C["input_bg"],
            fg=C["text"],
            insertbackground=C["text"],
            relief="flat",
            highlightthickness=1,
            highlightbackground=C["border"],
            highlightcolor=C["accent"],
            bd=0,
            padx=8,
            pady=8,
        )
        if width is not None:
            kw["width"] = width
        return scrolledtext.ScrolledText(parent, **kw)

    def _build(self):
        wrap = ttk.Frame(self.root, style="TFrame")
        wrap.pack(fill="both", expand=True, padx=20, pady=16)

        ttk.Label(wrap, text=u"웅이전용", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            wrap,
            text=u"SAP 자재코드 → 엑셀 수량. 엑셀에 없으면 오류 안내, SAP에 없는 엑셀 코드는 추출.",
            style="Muted.TLabel",
        ).pack(anchor="w", pady=(2, 8))

        top_row = tk.Frame(wrap, bg=C["bg"])
        top_row.pack(fill="x", pady=(0, 8))
        top_row.columnconfigure(0, weight=1)
        top_row.columnconfigure(1, weight=1)

        excel_card, excel = self._card(top_row, padx=12, pady=10)
        excel_card.grid(row=0, column=0, sticky="nsew", padx=(0, 4))
        ttk.Label(excel, text=u"엑셀 파일", style="Section.TLabel").pack(anchor="w")
        row = tk.Frame(excel, bg=C["surface"])
        row.pack(fill="x", pady=(6, 0))
        self.excel_var = tk.StringVar(value=self.cfg.get("excel_path") or "")
        ttk.Entry(row, textvariable=self.excel_var).pack(
            side="left", fill="x", expand=True, ipady=3
        )
        ttk.Button(
            row, text=u"찾아보기", style="Ghost.TButton", command=self.browse_excel
        ).pack(side="left", padx=(6, 0))

        um_card, um = self._card(top_row, padx=12, pady=10)
        um_card.grid(row=0, column=1, sticky="nsew", padx=(4, 0))
        ttk.Label(um, text=u"미매칭 추출", style="Section.TLabel").pack(anchor="w")
        dir_row = tk.Frame(um, bg=C["surface"])
        dir_row.pack(fill="x", pady=(6, 4))
        ttk.Label(dir_row, text=u"폴더", style="Card.TLabel", width=5).pack(side="left")
        self.unmatched_dir_var = tk.StringVar(value=self.cfg.get("unmatched_dir") or "")
        ttk.Entry(dir_row, textvariable=self.unmatched_dir_var).pack(
            side="left", fill="x", expand=True, ipady=3
        )
        ttk.Button(
            dir_row,
            text=u"폴더",
            style="Ghost.TButton",
            command=self.browse_unmatched_dir,
        ).pack(side="left", padx=(6, 0))
        name_row = tk.Frame(um, bg=C["surface"])
        name_row.pack(fill="x")
        ttk.Label(name_row, text=u"파일", style="Card.TLabel", width=5).pack(
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
            act, text=u"비교·결과 생성", style="Primary.TButton", command=self.do_run
        ).pack(side="right", padx=(0, 8))

        self.status_lbl = tk.Label(
            wrap,
            text=u"대기 중 — SAP 자재코드를 붙여넣은 뒤 [비교·결과 생성]",
            font=FONT_S,
            bg=C["chip_bg"],
            fg=C["chip_fg"],
            justify="left",
            anchor="w",
            padx=10,
            pady=6,
        )
        self.status_lbl.pack(fill="x", pady=(0, 8))

        paste_wrap = tk.Frame(wrap, bg=C["bg"])
        paste_wrap.pack(fill="both", expand=True, pady=(0, 10))
        paste_wrap.columnconfigure(0, weight=1)
        paste_wrap.columnconfigure(1, weight=1)
        paste_wrap.columnconfigure(2, weight=1)
        paste_wrap.rowconfigure(0, weight=1)

        self.code_box = self._make_col(
            paste_wrap,
            0,
            u"① SAP 자재코드",
            u"비우기",
            self.clear_codes,
            padx=(0, 4),
        )
        self.result_box = self._make_col(
            paste_wrap,
            1,
            u"② 결과 수량 (엑셀 → SAP 붙여넣기)",
            u"전체선택",
            self.select_result,
            padx=4,
            bg=C["result_bg"],
        )
        self.status_box = self._make_col(
            paste_wrap,
            2,
            u"상태",
            None,
            None,
            padx=(4, 0),
        )

        log_card, log_body = self._card(wrap, padx=12, pady=10)
        log_card.pack(fill="both", expand=True)
        ttk.Label(log_body, text=u"실행 로그", style="Section.TLabel").pack(anchor="w")
        self.log_box = scrolledtext.ScrolledText(
            log_body,
            height=6,
            font=("Consolas", 9),
            bg=C["log_bg"],
            fg=C["log_fg"],
            insertbackground=C["log_fg"],
            relief="flat",
            highlightthickness=0,
            bd=0,
            padx=8,
            pady=8,
        )
        self.log_box.pack(fill="both", expand=True, pady=(6, 0))

    def _make_col(self, parent, col, title, btn_text, btn_cmd, padx=0, bg=None):
        card, body = self._card(parent, padx=12, pady=10)
        card.grid(row=0, column=col, sticky="nsew", padx=padx)
        hdr = tk.Frame(body, bg=C["surface"])
        hdr.pack(fill="x")
        ttk.Label(hdr, text=title, style="Section.TLabel").pack(side="left")
        if btn_text and btn_cmd:
            ttk.Button(
                hdr, text=btn_text, style="Ghost.TButton", command=btn_cmd
            ).pack(side="right")
        box = self._text_box(body, bg=bg)
        box.pack(fill="both", expand=True, pady=(6, 0))
        return box

    def append_log(self, line):
        def _do():
            self.log_box.insert("end", line + "\n")
            self.log_box.see("end")

        self.root.after(0, _do)

    def clear_codes(self):
        self.code_box.delete("1.0", "end")

    def select_result(self):
        self.result_box.tag_add("sel", "1.0", "end")
        self.result_box.focus_set()
        try:
            self.root.clipboard_clear()
            self.root.clipboard_append(self.result_box.get("1.0", "end-1c"))
            self.append_log(u"결과 수량을 클립보드에 복사했습니다.")
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
            self.append_log(u"엑셀 연결 저장: {0}".format(path))

    def browse_unmatched_dir(self):
        path = filedialog.askdirectory(title=u"미매칭 엑셀 저장 폴더")
        if path:
            self.unmatched_dir_var.set(path)
            self._save_unmatched_settings()
            self.append_log(u"미매칭 저장 폴더: {0}".format(path))

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
        self.append_log(u"중지 요청…")

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
        self.status_lbl.config(text=u"비교 중…")
        threading.Thread(
            target=self._run_thread,
            args=(path, codes, unmatched_path),
            daemon=True,
        ).start()

    def _show_results(self, stats):
        self.result_box.delete("1.0", "end")
        self.status_box.delete("1.0", "end")
        for r in stats.get("result_rows") or []:
            self.result_box.insert("end", r.get("result_qty_text", u"") + "\n")
            self.status_box.insert("end", r.get("status", u"") + "\n")

        msg = (
            u"매칭={matched} | 엑셀오류(SAP→없음)={missing_count} | "
            u"엑셀만있음추출={unmatched}"
        ).format(**stats)
        if stats.get("moved_path"):
            msg += u" | 파일: {0}".format(stats["moved_path"])
        self.status_lbl.config(text=msg)

    def _run_thread(self, path, codes, unmatched_path):
        logger = DriveCheckLog(gui_callback=self.append_log)
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
            logger.finish(u"비교 완료")
            if stats:
                self.root.after(0, lambda: self._show_results(stats))
                self.root.after(0, lambda: self._finish_dialog(stats))
        except Exception as e:
            logger.exception(u"비교")
            self.append_log(u"오류: {0}".format(e))
            self.root.after(
                0, lambda: messagebox.showerror(u"오류", u"{0}".format(e))
            )
            self.root.after(
                0,
                lambda: self.status_lbl.config(text=u"오류 — 로그를 확인하세요."),
            )
        finally:
            self._running = False

    def _finish_dialog(self, stats):
        details = stats.get("missing_details") or []
        moved = stats.get("moved_path") or u"(없음)"

        if details:
            blocks = []
            for item in details[:40]:
                c = item["code"]
                sug = item.get("suggestions") or []
                if sug:
                    sug_txt = u", ".join(
                        u"{0} ({1}%)".format(s[1], int(round(s[0] * 100))) for s in sug
                    )
                    blocks.append(u"  · {0}\n      유사 제안: {1}".format(c, sug_txt))
                else:
                    blocks.append(u"  · {0}\n      유사 제안: 없음".format(c))
            more = u""
            if len(details) > 40:
                more = u"\n  … 외 {0}개".format(len(details) - 40)
            messagebox.showerror(
                u"엑셀 코드기입 오류 / 휴먼에러",
                u"SAP에만 있고 엑셀에 없는 코드입니다.\n"
                u"엑셀을 수정한 뒤 다시 비교해 주세요.\n\n"
                u"{0}{1}\n\n"
                u"(엑셀에만 있는 코드 추출: {2})".format(
                    u"\n".join(blocks), more, moved
                ),
            )
            return

        tip = (
            u"비교 완료.\n"
            u"매칭(엑셀수량)={0}\n"
            u"SAP에 없는 엑셀 코드 추출={1}\n"
            u"추출 파일: {2}\n\n"
            u"② 결과 수량을 전체선택 후 SAP 오더수량 열에 Ctrl+V 하세요."
        ).format(stats["matched"], stats["unmatched"], moved)
        messagebox.showinfo(u"완료", tip)

    def run(self):
        self.root.mainloop()


def main():
    App().run()


if __name__ == "__main__":
    main()

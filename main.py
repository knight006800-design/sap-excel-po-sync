# -*- coding: utf-8 -*-
"""SAP–엑셀 구매오더 수량 비교 (웅이전용)."""
from __future__ import print_function

import os
import subprocess
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

from drive_check import DriveCheckLog
from monitors import place_window_on_primary
from paths_util import app_dir, drive_check_dir, load_config, save_config
from sync_engine import SyncEngine


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
        self.root.minsize(1000, 720)
        self.root.configure(bg=C["bg"])
        place_window_on_primary(self.root, width=1080, height=780, margin=28)

        self.cfg = load_config()
        self._stop = False
        self._running = False

        self._setup_style()
        self._build()

        self.append_log(u"프로그램 폴더: {0}".format(app_dir()))
        self.append_log(u"구동점검 폴더: {0}".format(drive_check_dir()))
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

    def _text_box(self, parent, bg=None):
        box = scrolledtext.ScrolledText(
            parent,
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
        return box

    def _build(self):
        wrap = ttk.Frame(self.root, style="TFrame")
        wrap.pack(fill="both", expand=True, padx=20, pady=16)

        ttk.Label(wrap, text=u"웅이전용", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            wrap,
            text=u"SAP 코드·수량을 붙여넣고 엑셀과 비교합니다. 결과 수량을 드래그해 SAP에 직접 Ctrl+V 하세요. (자동 클릭 입력 없음)",
            style="Muted.TLabel",
        ).pack(anchor="w", pady=(2, 12))

        # Excel
        excel_card, excel = self._card(wrap)
        excel_card.pack(fill="x", pady=(0, 10))
        ttk.Label(excel, text=u"엑셀 파일", style="Section.TLabel").pack(anchor="w")
        row = tk.Frame(excel, bg=C["surface"])
        row.pack(fill="x", pady=(8, 0))
        self.excel_var = tk.StringVar(value=self.cfg.get("excel_path") or "")
        ttk.Entry(row, textvariable=self.excel_var).pack(
            side="left", fill="x", expand=True, ipady=4
        )
        ttk.Button(
            row, text=u"찾아보기", style="Ghost.TButton", command=self.browse_excel
        ).pack(side="left", padx=(8, 0))

        # Unmatched
        um_card, um = self._card(wrap)
        um_card.pack(fill="x", pady=(0, 10))
        ttk.Label(um, text=u"엑셀 미매칭 저장 (엑셀에만 있는 코드)", style="Section.TLabel").pack(
            anchor="w"
        )
        ttk.Label(
            um,
            text=u"SAP에 없는 엑셀 행만 새 파일로 복사합니다. 원본 엑셀은 수정하지 않습니다.",
            style="CardMuted.TLabel",
        ).pack(anchor="w", pady=(2, 8))

        dir_row = tk.Frame(um, bg=C["surface"])
        dir_row.pack(fill="x", pady=(0, 6))
        ttk.Label(dir_row, text=u"폴더", style="Card.TLabel", width=8).pack(side="left")
        self.unmatched_dir_var = tk.StringVar(value=self.cfg.get("unmatched_dir") or "")
        ttk.Entry(dir_row, textvariable=self.unmatched_dir_var).pack(
            side="left", fill="x", expand=True, ipady=4
        )
        ttk.Button(
            dir_row, text=u"폴더선택", style="Ghost.TButton", command=self.browse_unmatched_dir
        ).pack(side="left", padx=(8, 0))

        name_row = tk.Frame(um, bg=C["surface"])
        name_row.pack(fill="x")
        ttk.Label(name_row, text=u"파일명", style="Card.TLabel", width=8).pack(side="left")
        self.unmatched_name_var = tk.StringVar(
            value=self.cfg.get("unmatched_filename") or u"미매칭"
        )
        ttk.Entry(name_row, textvariable=self.unmatched_name_var).pack(
            side="left", fill="x", expand=True, ipady=4
        )

        # Actions
        act = tk.Frame(wrap, bg=C["bg"])
        act.pack(fill="x", pady=(0, 10))
        ttk.Button(
            act, text=u"구동점검 폴더", style="Ghost.TButton", command=self.open_log_folder
        ).pack(side="left")
        ttk.Button(
            act, text=u"중지", style="Danger.TButton", command=self.do_stop
        ).pack(side="right")
        ttk.Button(
            act, text=u"비교·결과 생성", style="Primary.TButton", command=self.do_run
        ).pack(side="right", padx=(0, 8))

        self.status_lbl = tk.Label(
            wrap,
            text=u"대기 중 — SAP 코드·수량을 붙여넣은 뒤 [비교·결과 생성]",
            font=FONT_S,
            bg=C["chip_bg"],
            fg=C["chip_fg"],
            justify="left",
            anchor="w",
            padx=10,
            pady=6,
        )
        self.status_lbl.pack(fill="x", pady=(0, 10))

        # 4 columns: code, sap qty, result qty, status
        paste_wrap = tk.Frame(wrap, bg=C["bg"])
        paste_wrap.pack(fill="both", expand=True, pady=(0, 10))
        for i in range(4):
            paste_wrap.columnconfigure(i, weight=1 if i < 3 else 0)
        paste_wrap.rowconfigure(0, weight=1)

        self.code_box = self._make_col(
            paste_wrap, 0, u"① SAP 자재코드", u"붙여넣기", self.clear_codes, padx=(0, 4)
        )
        self.qty_box = self._make_col(
            paste_wrap, 1, u"② SAP 오더수량", u"붙여넣기", self.clear_qtys, padx=4
        )
        self.result_box = self._make_col(
            paste_wrap,
            2,
            u"③ 결과 수량 (SAP에 붙여넣기)",
            u"전체선택",
            self.select_result,
            padx=4,
            bg=C["result_bg"],
        )
        self.status_box = self._make_col(
            paste_wrap, 3, u"상태", u"비우기", self.clear_status, padx=(4, 0), width_hint=10
        )

        # Log
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

    def _make_col(self, parent, col, title, btn_text, btn_cmd, padx=0, bg=None, width_hint=None):
        card, body = self._card(parent, padx=12, pady=10)
        card.grid(row=0, column=col, sticky="nsew", padx=padx)
        hdr = tk.Frame(body, bg=C["surface"])
        hdr.pack(fill="x")
        ttk.Label(hdr, text=title, style="Section.TLabel").pack(side="left")
        ttk.Button(hdr, text=btn_text, style="Ghost.TButton", command=btn_cmd).pack(
            side="right"
        )
        box = self._text_box(body, bg=bg)
        if width_hint:
            box.configure(width=width_hint)
        box.pack(fill="both", expand=True, pady=(6, 0))
        return box

    def append_log(self, line):
        def _do():
            self.log_box.insert("end", line + "\n")
            self.log_box.see("end")

        self.root.after(0, _do)

    def clear_codes(self):
        self.code_box.delete("1.0", "end")

    def clear_qtys(self):
        self.qty_box.delete("1.0", "end")

    def clear_status(self):
        self.status_box.delete("1.0", "end")

    def select_result(self):
        self.result_box.tag_add("sel", "1.0", "end")
        self.result_box.focus_set()
        try:
            self.root.clipboard_clear()
            self.root.clipboard_append(self.result_box.get("1.0", "end-1c"))
            self.append_log(u"결과 수량을 클립보드에 복사했습니다. SAP에 Ctrl+V 하세요.")
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
        name = (self.unmatched_name_var.get() or u"미매칭").strip() or u"미매칭"
        self.unmatched_name_var.set(name)
        cfg["unmatched_filename"] = name
        save_config(cfg)
        self.cfg = cfg

    def _resolve_unmatched_path(self, excel_path):
        folder = (self.unmatched_dir_var.get() or "").strip()
        if not folder:
            folder = os.path.dirname(os.path.abspath(excel_path))
        name = (self.unmatched_name_var.get() or u"미매칭").strip() or u"미매칭"
        base, ext = os.path.splitext(name)
        if not ext:
            ext = os.path.splitext(excel_path)[1] or u".xls"
            name = base + ext
        return os.path.join(folder, name)

    def do_stop(self):
        self._stop = True
        self.append_log(u"중지 요청…")

    def open_log_folder(self):
        folder = drive_check_dir()
        try:
            os.startfile(folder)
        except Exception:
            subprocess.Popen(["explorer", folder])

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
        qtys = self._box_text(self.qty_box)
        if not codes.strip():
            messagebox.showerror(u"코드 필요", u"SAP 자재코드를 Ctrl+V 하세요.")
            return
        if not qtys.strip():
            messagebox.showerror(u"수량 필요", u"SAP 오더수량을 Ctrl+V 하세요.")
            return

        self._save_unmatched_settings()
        unmatched_path = self._resolve_unmatched_path(path)

        self._stop = False
        self._running = True
        self.status_lbl.config(text=u"비교 중…")
        threading.Thread(
            target=self._run_thread,
            args=(path, codes, qtys, unmatched_path),
            daemon=True,
        ).start()

    def _show_results(self, stats):
        self.result_box.delete("1.0", "end")
        self.status_box.delete("1.0", "end")
        rows = stats.get("result_rows") or []
        for r in rows:
            self.result_box.insert("end", r.get("result_qty_text", u"") + "\n")
            self.status_box.insert("end", r.get("status", u"") + "\n")
        msg = (
            u"완료 — 수정={changed} 동일={same} 유지={kept} | 엑셀미매칭추출={unmatched}"
        ).format(**stats)
        if stats.get("moved_path"):
            msg += u" | 파일: {0}".format(stats["moved_path"])
        self.status_lbl.config(text=msg)

    def _run_thread(self, path, codes, qtys, unmatched_path):
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
                qtys_text=qtys,
                unmatched_path=unmatched_path,
            )
            logger.finish(u"비교 완료")
            if stats:
                self.root.after(0, lambda: self._show_results(stats))
                moved = stats.get("moved_path") or u"(없음)"
                tip = (
                    u"결과 생성 완료.\n"
                    u"수정={changed} 동일={same} 유지(엑셀없음)={kept}\n"
                    u"엑셀미매칭추출={unmatched}\n"
                    u"미매칭파일: {moved_path}\n\n"
                    u"③ 결과 수량을 전체선택(또는 버튼) 후\n"
                    u"SAP 오더수량 열에 Ctrl+V 하세요."
                ).format(
                    changed=stats["changed"],
                    same=stats["same"],
                    kept=stats["kept"],
                    unmatched=stats["unmatched"],
                    moved_path=moved,
                )
                self.root.after(0, lambda t=tip: messagebox.showinfo(u"완료", t))
        except Exception as e:
            logger.exception(u"비교")
            self.append_log(u"오류: {0}".format(e))
            self.root.after(
                0,
                lambda: messagebox.showerror(u"오류", u"{0}".format(e)),
            )
            self.root.after(
                0,
                lambda: self.status_lbl.config(text=u"오류 — 로그를 확인하세요."),
            )
        finally:
            self._running = False

    def run(self):
        self.root.mainloop()


def main():
    App().run()


if __name__ == "__main__":
    main()

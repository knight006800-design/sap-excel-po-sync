# -*- coding: utf-8 -*-
"""SAP–엑셀 구매오더 수량 동기화 RPA GUI (웅이전용)."""
from __future__ import print_function

import os
import subprocess
import threading
import time

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
from paths_util import app_dir, default_config, drive_check_dir, load_config, save_config
from settings_ui import RowQtySettings, format_sap_status, sap_config_ready
from sync_engine import SyncEngine


# 플랫 모던 팔레트 (그라데이션 없음)
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
        self.root.title(u"웅이전용 — 구매오더 수량 동기화")
        self.root.minsize(920, 700)
        self.root.configure(bg=C["bg"])
        place_window_on_primary(self.root, width=980, height=760, margin=32)

        self.cfg = load_config()
        self._stop = False
        self._running = False

        self._setup_style()
        self._build()

        self.append_log(u"프로그램 폴더: {0}".format(app_dir()))
        self.append_log(u"구동점검 폴더: {0}".format(drive_check_dir()))
        self.append_log(format_sap_status(self.cfg))
        if self.excel_var.get():
            self.append_log(u"연결된 엑셀: {0}".format(self.excel_var.get()))

    def _setup_style(self):
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("TFrame", background=C["bg"])
        style.configure("Card.TFrame", background=C["surface"])
        style.configure(
            "TLabel",
            background=C["bg"],
            foreground=C["text"],
            font=FONT,
        )
        style.configure(
            "Card.TLabel",
            background=C["surface"],
            foreground=C["text"],
            font=FONT,
        )
        style.configure(
            "Muted.TLabel",
            background=C["bg"],
            foreground=C["muted"],
            font=FONT_S,
        )
        style.configure(
            "CardMuted.TLabel",
            background=C["surface"],
            foreground=C["muted"],
            font=FONT_S,
        )
        style.configure(
            "Title.TLabel",
            background=C["bg"],
            foreground=C["text"],
            font=FONT_H,
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
            focuscolor=C["accent"],
        )
        style.map(
            "Primary.TButton",
            background=[("active", "#1D4ED8"), ("disabled", "#93C5FD")],
            foreground=[("disabled", "#FFFFFF")],
        )
        style.configure(
            "Ghost.TButton",
            background=C["surface"],
            foreground=C["text"],
            font=FONT,
            padding=(12, 8),
            borderwidth=1,
            relief="flat",
        )
        style.map(
            "Ghost.TButton",
            background=[("active", C["border"])],
        )
        style.configure(
            "Danger.TButton",
            background=C["danger_bg"],
            foreground=C["danger"],
            font=FONT,
            padding=(12, 8),
            borderwidth=0,
        )
        style.map(
            "Danger.TButton",
            background=[("active", "#FECACA")],
        )

    def _card(self, parent, padx=16, pady=14):
        outer = tk.Frame(parent, bg=C["border"], bd=0, highlightthickness=0)
        inner = tk.Frame(outer, bg=C["surface"], bd=0, highlightthickness=0)
        inner.pack(fill="both", expand=True, padx=1, pady=1)
        pad = tk.Frame(inner, bg=C["surface"])
        pad.pack(fill="both", expand=True, padx=padx, pady=pady)
        return outer, pad

    def _build(self):
        wrap = ttk.Frame(self.root, style="TFrame")
        wrap.pack(fill="both", expand=True, padx=20, pady=16)

        # Header
        ttk.Label(wrap, text=u"웅이전용", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            wrap,
            text=u"SAP 코드·수량을 붙여넣고 엑셀과 비교한 뒤, 수량 수정과 미매칭 추출을 자동 실행합니다.",
            style="Muted.TLabel",
        ).pack(anchor="w", pady=(2, 12))

        # Excel card
        excel_card, excel = self._card(wrap)
        excel_card.pack(fill="x", pady=(0, 10))
        ttk.Label(excel, text=u"엑셀 파일", style="Section.TLabel").pack(anchor="w")
        row = tk.Frame(excel, bg=C["surface"])
        row.pack(fill="x", pady=(8, 0))
        self.excel_var = tk.StringVar(value=self.cfg.get("excel_path") or "")
        ent = ttk.Entry(row, textvariable=self.excel_var)
        ent.pack(side="left", fill="x", expand=True, ipady=4)
        ttk.Button(row, text=u"찾아보기", style="Ghost.TButton", command=self.browse_excel).pack(
            side="left", padx=(8, 0)
        )

        # Actions card
        act_card, act = self._card(wrap)
        act_card.pack(fill="x", pady=(0, 10))
        ttk.Label(act, text=u"SAP 클릭 위치", style="Section.TLabel").pack(anchor="w")
        ttk.Label(
            act,
            text=u"주모니터 SAP 화면에서 첫 행 Y · 행높이 · 수량 칸 X 를 한 번만 지정하면 됩니다.",
            style="CardMuted.TLabel",
        ).pack(anchor="w", pady=(2, 8))

        btnrow = tk.Frame(act, bg=C["surface"])
        btnrow.pack(fill="x")
        ttk.Button(
            btnrow, text=u"행/수량 초기설정", style="Ghost.TButton", command=self.do_row_qty
        ).pack(side="left")
        ttk.Button(
            btnrow, text=u"재설정", style="Ghost.TButton", command=self.do_reset_row_qty
        ).pack(side="left", padx=(8, 0))
        ttk.Button(
            btnrow, text=u"구동점검 폴더", style="Ghost.TButton", command=self.open_log_folder
        ).pack(side="left", padx=(8, 0))
        ttk.Button(
            btnrow, text=u"중지 ESC", style="Danger.TButton", command=self.do_stop
        ).pack(side="right")
        ttk.Button(
            btnrow, text=u"동기화 실행", style="Primary.TButton", command=self.do_run
        ).pack(side="right", padx=(0, 8))

        self.status_lbl = tk.Label(
            act,
            text=format_sap_status(self.cfg),
            font=FONT_S,
            bg=C["chip_bg"],
            fg=C["chip_fg"],
            justify="left",
            anchor="w",
            padx=10,
            pady=6,
        )
        self.status_lbl.pack(fill="x", pady=(12, 0))

        # Paste cards side by side
        paste_wrap = tk.Frame(wrap, bg=C["bg"])
        paste_wrap.pack(fill="both", expand=True, pady=(0, 10))
        paste_wrap.columnconfigure(0, weight=1)
        paste_wrap.columnconfigure(1, weight=1)
        paste_wrap.rowconfigure(0, weight=1)

        code_card, code_body = self._card(paste_wrap)
        code_card.grid(row=0, column=0, sticky="nsew", padx=(0, 5))
        hdr1 = tk.Frame(code_body, bg=C["surface"])
        hdr1.pack(fill="x")
        ttk.Label(hdr1, text=u"SAP 자재코드", style="Section.TLabel").pack(side="left")
        ttk.Button(
            hdr1, text=u"비우기", style="Ghost.TButton", command=self.clear_codes
        ).pack(side="right")
        ttk.Label(
            code_body,
            text=u"SAP에서 자재번호 열 복사 → 여기 Ctrl+V (한 줄에 하나)",
            style="CardMuted.TLabel",
        ).pack(anchor="w", pady=(4, 6))
        self.code_box = scrolledtext.ScrolledText(
            code_body,
            height=12,
            font=FONT_MONO,
            bg=C["input_bg"],
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
        self.code_box.pack(fill="both", expand=True)

        qty_card, qty_body = self._card(paste_wrap)
        qty_card.grid(row=0, column=1, sticky="nsew", padx=(5, 0))
        hdr2 = tk.Frame(qty_body, bg=C["surface"])
        hdr2.pack(fill="x")
        ttk.Label(hdr2, text=u"SAP 오더수량", style="Section.TLabel").pack(side="left")
        ttk.Button(
            hdr2, text=u"비우기", style="Ghost.TButton", command=self.clear_qtys
        ).pack(side="right")
        ttk.Label(
            qty_body,
            text=u"SAP에서 오더수량 열 복사 → 여기 Ctrl+V (코드와 같은 순서)",
            style="CardMuted.TLabel",
        ).pack(anchor="w", pady=(4, 6))
        self.qty_box = scrolledtext.ScrolledText(
            qty_body,
            height=12,
            font=FONT_MONO,
            bg=C["input_bg"],
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
        self.qty_box.pack(fill="both", expand=True)

        # Log
        log_card, log_body = self._card(wrap, padx=12, pady=10)
        log_card.pack(fill="both", expand=True)
        ttk.Label(log_body, text=u"실행 로그", style="Section.TLabel").pack(anchor="w")
        self.log_box = scrolledtext.ScrolledText(
            log_body,
            height=8,
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

    def refresh_status(self):
        self.cfg = load_config()
        text = format_sap_status(self.cfg)
        self.status_lbl.config(text=text)
        self.append_log(text)

    def append_log(self, line):
        def _do():
            self.log_box.insert("end", line + "\n")
            self.log_box.see("end")

        self.root.after(0, _do)

    def clear_codes(self):
        self.code_box.delete("1.0", "end")

    def clear_qtys(self):
        self.qty_box.delete("1.0", "end")

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

    def do_row_qty(self):
        if self._running:
            messagebox.showwarning(u"실행 중", u"동기화 중에는 설정할 수 없습니다.")
            return

        def on_done():
            self.refresh_status()

        RowQtySettings(self.root, on_done=on_done, logger=_GuiLogger(self)).start()

    def do_reset_row_qty(self):
        if self._running:
            return
        ok = messagebox.askyesno(u"재설정", u"행/수량 설정을 기본값으로 되돌릴까요?")
        if not ok:
            return
        d = default_config()["sap"]
        cfg = load_config()
        sap = cfg.setdefault("sap", {})
        sap["first_row_y"] = d["first_row_y"]
        sap["row_height"] = d["row_height"]
        sap["qty_center_x"] = d["qty_center_x"]
        save_config(cfg)
        self.append_log(u"행/수량 재설정 완료")
        self.refresh_status()

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
            messagebox.showerror(
                u"코드 붙여넣기 필요",
                u"SAP 자재코드를 왼쪽 칸에 Ctrl+V 하세요.",
            )
            return
        if not qtys.strip():
            messagebox.showerror(
                u"수량 붙여넣기 필요",
                u"SAP 오더수량을 오른쪽 칸에 Ctrl+V 하세요.",
            )
            return

        self.cfg = load_config()
        self.cfg["excel_path"] = path
        save_config(self.cfg)

        if not sap_config_ready(self.cfg):
            messagebox.showerror(
                u"설정 미완료",
                u"먼저 [행/수량 초기설정]에서 클릭지정 후 저장하세요.",
            )
            return

        ok = messagebox.askokcancel(
            u"동기화 실행",
            u"비교 후 자동으로 진행합니다.\n"
            u"· 수량 다른 행 → SAP 자동 수정\n"
            u"· 미매칭 엑셀 행 → 새 파일로 추출\n\n"
            u"3초 뒤 시작합니다. 마우스를 만지지 마세요. ESC=중지",
        )
        if not ok:
            return

        self._stop = False
        self._running = True
        threading.Thread(
            target=self._run_thread, args=(path, codes, qtys), daemon=True
        ).start()

    def _run_thread(self, path, codes, qtys):
        logger = DriveCheckLog(gui_callback=self.append_log)
        try:
            logger.info(u"3초 후 시작… (주모니터 SAP 확인)")
            time.sleep(3)
            engine = SyncEngine(
                config=load_config(),
                logger=logger,
                stop_flag=lambda: self._stop,
            )
            stats = engine.run(
                excel_path=path, codes_text=codes, qtys_text=qtys
            )
            logger.finish(u"동기화 완료")
            if stats:
                if not stats.get("moved_path"):
                    stats["moved_path"] = u"(없음)"
                msg = (
                    u"동기화 끝.\n"
                    u"매칭={matched} 수정={updated} 동일={same}\n"
                    u"미매칭추출={unmatched} 스킵={skipped} 실패={failed}\n"
                    u"미매칭파일: {moved_path}"
                ).format(**stats)
                self.root.after(0, lambda m=msg: messagebox.showinfo(u"완료", m))
        except Exception as e:
            logger.exception(u"동기화")
            self.append_log(u"오류: {0}".format(e))
            self.root.after(
                0,
                lambda: messagebox.showerror(
                    u"오류",
                    u"{0}\n\n구동점검 폴더의 로그를 확인하세요.".format(e),
                ),
            )
        finally:
            self._running = False

    def run(self):
        self.root.mainloop()


class _GuiLogger(object):
    def __init__(self, app):
        self.app = app

    def info(self, msg):
        self.app.append_log(msg)

    def exception(self, ctx):
        self.app.append_log(u"예외: {0}".format(ctx))


def main():
    App().run()


if __name__ == "__main__":
    main()

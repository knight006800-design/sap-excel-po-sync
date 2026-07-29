# -*- coding: utf-8 -*-
"""SAP–엑셀 구매오더 수량 동기화 RPA GUI (웅이전용)."""
from __future__ import print_function

import os
import subprocess
import threading

try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, scrolledtext
except ImportError:
    import Tkinter as tk
    import tkFileDialog as filedialog
    import tkMessageBox as messagebox
    import ScrolledText as scrolledtext

from calibrate import CalibrateWizard
from monitors import place_window_on_primary
from paths_util import app_dir, drive_check_dir, load_config, save_config
from sync_engine import SyncEngine


class App(object):
    def __init__(self):
        self.root = tk.Tk()
        self.root.title(u"웅이전용 — 구매오더 수량 동기화 (SAP ↔ Excel)")
        self.root.minsize(640, 440)
        place_window_on_primary(self.root, width=720, height=520, margin=40)

        self.cfg = load_config()
        self._stop = False
        self._running = False

        frm = tk.Frame(self.root, padx=12, pady=10)
        frm.pack(fill="both", expand=True)

        tk.Label(frm, text=u"엑셀 파일 연결", font=("Malgun Gothic", 11, "bold")).pack(
            anchor="w"
        )
        row = tk.Frame(frm)
        row.pack(fill="x", pady=4)
        self.excel_var = tk.StringVar(value=self.cfg.get("excel_path") or "")
        tk.Entry(row, textvariable=self.excel_var).pack(
            side="left", fill="x", expand=True
        )
        tk.Button(row, text=u"찾아보기", command=self.browse_excel, width=10).pack(
            side="left", padx=4
        )

        btn_row = tk.Frame(frm)
        btn_row.pack(fill="x", pady=8)
        tk.Button(
            btn_row, text=u"1. SAP 화면 보정", command=self.do_calibrate, width=18
        ).pack(side="left", padx=2)
        tk.Button(
            btn_row,
            text=u"2. 동기화 실행",
            command=self.do_run,
            width=18,
            bg="#dcebff",
        ).pack(side="left", padx=2)
        tk.Button(btn_row, text=u"중지 (ESC)", command=self.do_stop, width=12).pack(
            side="left", padx=2
        )
        tk.Button(
            btn_row, text=u"구동점검 폴더", command=self.open_log_folder, width=14
        ).pack(side="left", padx=2)

        tk.Label(
            frm,
            text=u"배치: 주모니터=SAP「구매오더 생성」 / 서브모니터=안내창(있으면).\n"
            u"보정·실행 모두 주모니터 SAP 기준으로 동작합니다. 보정은 다시 하세요.\n"
            u"실행 중 마우스 금지 · ESC 중지 · 로그는 「구동점검」폴더.",
            justify="left",
            fg="#333",
        ).pack(anchor="w", pady=4)

        tk.Label(frm, text=u"실행 로그", font=("Malgun Gothic", 10, "bold")).pack(
            anchor="w"
        )
        self.log_box = scrolledtext.ScrolledText(frm, height=18, font=("Consolas", 9))
        self.log_box.pack(fill="both", expand=True, pady=4)

        self.append_log(u"프로그램 폴더: {0}".format(app_dir()))
        self.append_log(u"구동점검 폴더: {0}".format(drive_check_dir()))
        if self.excel_var.get():
            self.append_log(u"연결된 엑셀: {0}".format(self.excel_var.get()))

    def append_log(self, line):
        def _do():
            self.log_box.insert("end", line + "\n")
            self.log_box.see("end")

        self.root.after(0, _do)

    def browse_excel(self):
        path = filedialog.askopenfilename(
            title=u"구매오더 엑셀 선택",
            filetypes=[
                (u"Excel", "*.xls;*.xlsx"),
                (u"모든 파일", "*.*"),
            ],
        )
        if path:
            self.excel_var.set(path)
            self.cfg = load_config()
            self.cfg["excel_path"] = path
            save_config(self.cfg)
            self.append_log(u"엑셀 연결 저장: {0}".format(path))

    def do_calibrate(self):
        if self._running:
            messagebox.showwarning(
                u"실행 중", u"동기화 실행 중에는 보정할 수 없습니다."
            )
            return

        def on_done():
            self.cfg = load_config()
            self.append_log(u"보정 완료 — config.json 저장됨")

        CalibrateWizard(self.root, on_done=on_done, logger=_GuiLogger(self)).start()

    def do_stop(self):
        self._stop = True
        self.append_log(u"중지 요청…")

    def open_log_folder(self):
        folder = drive_check_dir()
        try:
            os.startfile(folder)
        except Exception:
            subprocess.Popen(["explorer", folder])

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

        self.cfg = load_config()
        self.cfg["excel_path"] = path
        save_config(self.cfg)

        ok = messagebox.askokcancel(
            u"동기화 실행",
            u"3초 후 SAP 화면을 읽어 수량을 맞춥니다.\n"
            u"SAP 창이 좌측에 보이고, 보정된 상태인지 확인하세요.\n"
            u"실행 중 마우스를 만지지 마세요. ESC로 중지.\n\n계속할까요?",
        )
        if not ok:
            return

        self._stop = False
        self._running = True
        threading.Thread(target=self._run_thread, args=(path,), daemon=True).start()

    def _run_thread(self, path):
        try:
            engine = SyncEngine(gui_log=self.append_log, stop_flag=lambda: self._stop)
            stats = engine.run(excel_path=path)
            if stats:
                self.root.after(
                    0,
                    lambda: messagebox.showinfo(
                        u"완료",
                        u"동기화가 끝났습니다.\n"
                        u"일치={match} 수정={updated} 미존재={missing}\n"
                        u"상세는 「구동점검」폴더의 구동점검.txt 를 확인하세요.".format(
                            **stats
                        ),
                    ),
                )
        except Exception as e:
            self.append_log(u"오류: {0}".format(e))
            self.root.after(
                0,
                lambda: messagebox.showerror(
                    u"오류",
                    u"{0}\n\n「구동점검」폴더의 로그와 캡처 PNG를 가져와 전달해 주세요.".format(
                        e
                    ),
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

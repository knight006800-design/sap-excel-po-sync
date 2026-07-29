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

from monitors import place_window_on_primary
from paths_util import app_dir, drive_check_dir, load_config, save_config
from region_box import MaterialRegionBox
from settings_ui import RowQtySettings, format_sap_status, sap_config_ready
from sync_engine import SyncEngine


class App(object):
    def __init__(self):
        self.root = tk.Tk()
        self.root.title(u"웅이전용 — 구매오더 수량 동기화 (SAP ↔ Excel)")
        self.root.minsize(680, 500)
        place_window_on_primary(self.root, width=740, height=560, margin=40)

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

        tk.Label(frm, text=u"SAP 설정", font=("Malgun Gothic", 11, "bold")).pack(
            anchor="w", pady=(8, 2)
        )
        btn1 = tk.Frame(frm)
        btn1.pack(fill="x", pady=2)
        tk.Button(
            btn1,
            text=u"자재코드 영역 박스지정",
            command=self.do_region_box,
            width=20,
        ).pack(side="left", padx=2)
        tk.Button(
            btn1, text=u"행/수량 초기설정", command=self.do_row_qty, width=16
        ).pack(side="left", padx=2)
        tk.Button(btn1, text=u"재설정", command=self.do_reset_row_qty, width=10).pack(
            side="left", padx=2
        )

        btn2 = tk.Frame(frm)
        btn2.pack(fill="x", pady=6)
        tk.Button(
            btn2, text=u"동기화 실행", command=self.do_run, width=16, bg="#dcebff"
        ).pack(side="left", padx=2)
        tk.Button(btn2, text=u"중지 (ESC)", command=self.do_stop, width=12).pack(
            side="left", padx=2
        )
        tk.Button(
            btn2, text=u"구동점검 폴더", command=self.open_log_folder, width=14
        ).pack(side="left", padx=2)

        self.status_lbl = tk.Label(
            frm,
            text=format_sap_status(self.cfg),
            font=("Malgun Gothic", 9),
            fg="#0d47a1",
            justify="left",
            wraplength=700,
        )
        self.status_lbl.pack(anchor="w", pady=4)

        tk.Label(
            frm,
            text=u"주모니터=SAP / 서브=안내창. 박스로 자재코드 열을 지정하고, 행·수량은 초기설정에서 저장합니다.\n"
            u"저장된 값은 다음 실행에 그대로 재사용됩니다. OCR은 영문·숫자·하이픈만 인식합니다.",
            justify="left",
            fg="#333",
        ).pack(anchor="w", pady=4)

        tk.Label(frm, text=u"실행 로그", font=("Malgun Gothic", 10, "bold")).pack(
            anchor="w"
        )
        self.log_box = scrolledtext.ScrolledText(frm, height=16, font=("Consolas", 9))
        self.log_box.pack(fill="both", expand=True, pady=4)

        self.append_log(u"프로그램 폴더: {0}".format(app_dir()))
        self.append_log(u"구동점검 폴더: {0}".format(drive_check_dir()))
        self.append_log(format_sap_status(self.cfg))
        if self.excel_var.get():
            self.append_log(u"연결된 엑셀: {0}".format(self.excel_var.get()))

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

    def do_region_box(self):
        if self._running:
            messagebox.showwarning(u"실행 중", u"동기화 중에는 설정할 수 없습니다.")
            return

        def on_done():
            self.refresh_status()
            self.append_log(u"자재코드 영역 박스 저장 완료")

        MaterialRegionBox(self.root, on_done=on_done, logger=_GuiLogger(self)).start()

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
        # 설정 창의 재설정과 동일 — 창 열어 재설정 유도 또는 바로 기본값
        from paths_util import default_config

        ok = messagebox.askyesno(
            u"재설정",
            u"행/수량 설정을 기본값으로 되돌릴까요?\n자재코드 영역 박스는 유지됩니다.",
        )
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

        if not sap_config_ready(self.cfg):
            messagebox.showerror(
                u"설정 미완료",
                u"먼저 [자재코드 영역 박스지정]과 [행/수량 초기설정]을 완료하세요.",
            )
            return

        ok = messagebox.askokcancel(
            u"동기화 실행",
            u"저장된 설정으로 3초 후 SAP 화면을 읽습니다.\n"
            u"주모니터에 SAP가 보이는지 확인하세요.\n"
            u"실행 중 마우스를 만지지 마세요. ESC=중지\n\n계속할까요?",
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
                        u"동기화 끝.\n일치={match} 수정={updated} 미존재={missing}\n"
                        u"상세는 구동점검 폴더를 확인하세요.".format(**stats),
                    ),
                )
        except Exception as e:
            self.append_log(u"오류: {0}".format(e))
            self.root.after(
                0,
                lambda: messagebox.showerror(
                    u"오류",
                    u"{0}\n\n구동점검 폴더의 로그/캡처를 전달해 주세요.".format(e),
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

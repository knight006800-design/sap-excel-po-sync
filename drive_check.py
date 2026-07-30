# -*- coding: utf-8 -*-
"""구동점검 로그 — 다른 PC 오류 분석용 (구동점검 폴더)."""
from __future__ import print_function

import os
import shutil
import traceback
from datetime import datetime

from paths_util import drive_check_dir


class DriveCheckLog(object):
    """파일: 구동점검/구동점검.txt"""

    FILENAME = u"구동점검.txt"

    def __init__(self, gui_callback=None):
        self.gui_callback = gui_callback
        self.base = drive_check_dir()
        self.path = os.path.join(self.base, self.FILENAME)
        self._backup_previous()
        self._lines = []
        self._buf = []
        self._write_header()

    def _flush(self):
        if not self._buf:
            return
        chunk = u"\n".join(self._buf) + u"\n"
        self._buf = []
        try:
            with open(self.path, "a", encoding="utf-8") as f:
                f.write(chunk)
        except Exception:
            try:
                with open(self.path, "a") as f:
                    f.write(chunk.encode("utf-8"))
            except Exception:
                pass

    def _backup_previous(self):
        if os.path.isfile(self.path):
            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            bak = os.path.join(self.base, u"구동점검_이전_{0}.txt".format(stamp))
            try:
                shutil.copy2(self.path, bak)
            except Exception:
                pass
            try:
                open(self.path, "w", encoding="utf-8").close()
            except Exception:
                try:
                    open(self.path, "w").close()
                except Exception:
                    pass

    def _write_header(self):
        self.log(u"=" * 60)
        self.log(u"구동점검 시작: {0}".format(datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
        self.log(u"로그폴더: {0}".format(self.base))
        self.log(u"로그파일: {0}".format(self.path))
        self.log(u"=" * 60)

    def log(self, message, level=u"INFO"):
        ts = datetime.now().strftime("%H:%M:%S")
        line = u"[{0}][{1}] {2}".format(ts, level, message)
        self._lines.append(line)
        self._buf.append(line)
        # 버퍼가 커지거나 ERROR면 즉시 기록
        if level == u"ERROR" or len(self._buf) >= 24:
            self._flush()
        if self.gui_callback:
            try:
                self.gui_callback(line)
            except Exception:
                pass

    def info(self, msg):
        self.log(msg, u"INFO")

    def warn(self, msg):
        self.log(msg, u"WARN")

    def error(self, msg):
        self.log(msg, u"ERROR")

    def step(self, name, detail=u""):
        if detail:
            self.log(u"단계: {0} — {1}".format(name, detail), u"STEP")
        else:
            self.log(u"단계: {0}".format(name), u"STEP")

    def exception(self, context=u""):
        tb = traceback.format_exc()
        self.error(u"예외 ({0}):\n{1}".format(context, tb))

    def save_screenshot(self, image, tag=u"오류"):
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        name = u"구동점검_캡처_{0}_{1}.png".format(tag, stamp)
        safe = name.replace(u"/", u"_").replace(u"\\", u"_")
        out = os.path.join(self.base, safe)
        try:
            image.save(out)
            self.info(u"스크린샷 저장: {0}".format(out))
            return out
        except Exception as e:
            self.error(u"스크린샷 저장 실패: {0}".format(e))
            return None

    def save_text(self, tag, content):
        """OCR 등 긴 텍스트를 구동점검 폴더에 저장."""
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        name = u"구동점검_{0}_{1}.txt".format(tag, stamp)
        safe = name.replace(u"/", u"_").replace(u"\\", u"_")
        out = os.path.join(self.base, safe)
        try:
            with open(out, "w", encoding="utf-8") as f:
                f.write(content if content is not None else u"")
            self.info(u"텍스트 저장: {0}".format(out))
            return out
        except Exception as e:
            self.error(u"텍스트 저장 실패: {0}".format(e))
            return None

    def finish(self, summary=u""):
        self.log(u"=" * 60)
        self.log(u"구동점검 종료: {0}".format(datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
        if summary:
            self.log(u"요약: {0}".format(summary))
        self.log(u"=" * 60)
        self._flush()

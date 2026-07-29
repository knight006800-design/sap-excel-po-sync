# -*- coding: utf-8 -*-
"""로컬 시뮬레이션 — 모니터 배치/보정 클릭 필터/경로."""
from __future__ import print_function

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from monitors import (
    enum_monitors,
    is_on_primary,
    primary_monitor,
    secondary_monitors,
)
from calibrate import STEPS


def main():
    print("=== STEPS count ===", len(STEPS))
    assert len(STEPS) == 5
    for i, (k, t, d) in enumerate(STEPS, 1):
        print(i, k, t.encode("utf-8") if isinstance(t, unicode) else t)
        assert u"정확히" in d or "정확히" in d

    mons = enum_monitors()
    print("=== monitors ===", len(mons))
    for m in mons:
        print(m)
    prim = primary_monitor(mons)
    secs = secondary_monitors(mons)
    print("primary", prim)
    print("secondary count", len(secs))

    # primary center should be on primary
    cx = (prim["left"] + prim["right"]) // 2
    cy = (prim["top"] + prim["bottom"]) // 2
    assert is_on_primary(cx, cy, mons) is True
    print("primary center OK", cx, cy)

    # if secondary exists, its center must NOT be primary
    if secs:
        s = secs[0]
        sx = (s["left"] + s["right"]) // 2
        sy = (s["top"] + s["bottom"]) // 2
        assert is_on_primary(sx, sy, mons) is False
        print("secondary center OK (ignored on primary filter)", sx, sy)

    # guide_images should be gone from runtime deps
    assert not os.path.isfile(
        os.path.join(os.path.dirname(__file__), "copy_guides.py")
    )
    print("copy_guides removed OK")

    from paths_util import drive_check_dir, load_config

    d = drive_check_dir()
    assert os.path.isdir(d)
    cfg = load_config()
    assert "sap" in cfg
    print("drive_check_dir", d)
    print("ALL SIMULATION PASSED")
    return 0


if __name__ == "__main__":
    # py2/3 unicode
    try:
        unicode
    except NameError:
        unicode = str  # noqa
    raise SystemExit(main())

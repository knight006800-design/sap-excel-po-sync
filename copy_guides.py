# -*- coding: utf-8 -*-
"""보정 안내 그림을 guide_images 폴더로 복사."""
from __future__ import print_function

import os
import shutil

HERE = os.path.dirname(os.path.abspath(__file__))
DEST = os.path.join(HERE, "guide_images")
NAMES = [
    "guide_material_tl.png",
    "guide_material_br.png",
    "guide_qty_x.png",
    "guide_row1_y.png",
    "guide_row2_y.png",
]
SEARCH = [
    os.path.join(
        os.path.expanduser("~"),
        ".cursor",
        "projects",
        "c-Users-Desktop",
        "assets",
    ),
    os.path.join(HERE, "assets"),
]


def main():
    if not os.path.isdir(DEST):
        os.makedirs(DEST)
    found = 0
    for name in NAMES:
        src = None
        for folder in SEARCH:
            cand = os.path.join(folder, name)
            if os.path.isfile(cand):
                src = cand
                break
        if src:
            shutil.copy2(src, os.path.join(DEST, name))
            print("copied", name)
            found += 1
        else:
            print("MISSING", name)
    print("done", found, "/", len(NAMES))
    return 0 if found == len(NAMES) else 1


if __name__ == "__main__":
    raise SystemExit(main())

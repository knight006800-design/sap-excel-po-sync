# -*- coding: utf-8 -*-
"""경로·설정 공통 유틸."""
from __future__ import print_function

import json
import os
import sys


APP_NAME = u"구매오더수량동기화"


def app_dir():
    """exe 또는 스크립트 기준 실행 폴더."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def config_path():
    return os.path.join(app_dir(), "config.json")


def default_config():
    return {
        "excel_path": "",
        "delay_after_click": 0.25,
        "delay_after_type": 0.35,
        "delay_between_rows": 0.15,
        "ocr_scale": 2.5,
        "tesseract_cmd": "",
        "sap": {
            "material_left": 280,
            "material_top": 220,
            "material_right": 480,
            "material_bottom": 980,
            "qty_center_x": 1180,
            "first_row_y": 240,
            "row_height": 28,
            "visible_rows": 25,
        },
    }


def load_config():
    path = config_path()
    cfg = default_config()
    if os.path.isfile(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            cfg.update(data)
            if "sap" in data and isinstance(data["sap"], dict):
                sap = default_config()["sap"]
                sap.update(data["sap"])
                cfg["sap"] = sap
        except Exception:
            pass
    return cfg


def save_config(cfg):
    path = config_path()
    with open(path, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)

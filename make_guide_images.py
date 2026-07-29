# -*- coding: utf-8 -*-
"""보정 안내 그림 생성 — SAP 표 모사 + 클릭 위치 빨간 점."""
from __future__ import print_function

import os

from PIL import Image, ImageDraw, ImageFont

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "guide_images")
W, H = 640, 360


def font(size, bold=False):
    candidates = [
        r"C:\Windows\Fonts\malgunbd.ttf" if bold else r"C:\Windows\Fonts\malgun.ttf",
        r"C:\Windows\Fonts\gulim.ttc",
        r"C:\Windows\Fonts\arial.ttf",
    ]
    for p in candidates:
        if os.path.isfile(p):
            try:
                return ImageFont.truetype(p, size)
            except Exception:
                pass
    return ImageFont.load_default()


def draw_sap_base():
    """간단한 SAP 구매오더 그리드."""
    img = Image.new("RGB", (W, H), (245, 245, 240))
    d = ImageDraw.Draw(img)
    # 타이틀바
    d.rectangle([0, 0, W, 28], fill=(0, 100, 140))
    d.text((8, 6), u"구매오더 생성 (SAP 예시)", fill="white", font=font(14, True))
    # 툴바
    d.rectangle([0, 28, W, 52], fill=(220, 220, 215))

    # 컬럼: 상태 | 차종 | 자재코드 | 자재명 | ... | 오더수량
    headers = [
        (10, u"상태", 40),
        (55, u"차종", 45),
        (110, u"자재코드", 130),
        (250, u"자재명", 160),
        (420, u"납품예정일", 90),
        (520, u"오더수량", 80),
    ]
    y0 = 52
    d.rectangle([0, y0, W, y0 + 24], fill=(210, 210, 200))
    for x, name, _w in headers:
        d.text((x + 4, y0 + 4), name, fill=(30, 30, 30), font=font(12, True))

    rows = [
        (u"CN7", u"88510-AA000", u"RECLINER ... LH", u"2026.07.29", u"240"),
        (u"CN7", u"88511-AA000", u"RECLINER ... RH", u"2026.07.29", u"480"),
        (u"CN7", u"88512-AA000", u"RECLINER ...", u"2026.07.29", u"360"),
        (u"CN7", u"88513-AA000", u"RECLINER ...", u"2026.07.29", u"600"),
        (u"CN7", u"88514-AA000", u"RECLINER ...", u"2026.07.29", u"120"),
    ]
    row_h = 28
    for i, (car, code, name, date, qty) in enumerate(rows):
        y = y0 + 24 + i * row_h
        bg = (255, 255, 180) if i == 0 else (255, 255, 255)
        d.rectangle([0, y, W, y + row_h], fill=bg, outline=(200, 200, 190))
        d.ellipse([18, y + 8, 30, y + 20], fill=(255, 200, 0), outline=(180, 120, 0))
        d.text((58, y + 6), car, fill=(0, 0, 0), font=font(12))
        d.text((114, y + 6), code, fill=(0, 0, 0), font=font(12))
        d.text((254, y + 6), name, fill=(60, 60, 60), font=font(11))
        d.text((424, y + 6), date, fill=(0, 0, 0), font=font(11))
        d.text((540, y + 6), qty, fill=(0, 0, 0), font=font(12))

    # 레이아웃 좌표 (점 찍기용)
    layout = {
        "header_bottom": y0 + 24,
        "row_h": row_h,
        "code_left": 110,
        "code_right": 248,
        "qty_left": 520,
        "qty_right": 600,
        "qty_cx": 560,
        "n_rows": len(rows),
    }
    return img, layout


def draw_dot(d, x, y, label, r=14):
    """큰 빨간 점 + 흰 테두리 + 번호."""
    d.ellipse([x - r - 2, y - r - 2, x + r + 2, y + r + 2], fill="white")
    d.ellipse([x - r, y - r, x + r, y + r], fill=(220, 0, 0), outline=(120, 0, 0), width=2)
    # 십자
    d.line([x - r + 3, y, x + r - 3, y], fill="white", width=2)
    d.line([x, y - r + 3, x, y + r - 3], fill="white", width=2)
    d.text((x + r + 6, y - 10), label, fill=(200, 0, 0), font=font(16, True))


def caption(d, text):
    d.rectangle([0, H - 40, W, H], fill=(255, 248, 220))
    d.text((10, H - 30), text, fill=(180, 0, 0), font=font(15, True))


def make_all():
    if not os.path.isdir(OUT):
        os.makedirs(OUT)

    specs = []

    # 1 material_tl — 첫 행 자재코드 왼쪽 위
    img, L = draw_sap_base()
    d = ImageDraw.Draw(img)
    x = L["code_left"] + 2
    y = L["header_bottom"] + 2
    draw_dot(d, x, y, u"여기를 클릭!")
    # 칸 테두리 강조
    d.rectangle(
        [L["code_left"], L["header_bottom"], L["code_right"], L["header_bottom"] + L["row_h"]],
        outline=(220, 0, 0),
        width=2,
    )
    caption(d, u"1/5  첫 행 자재코드 칸의 ★왼쪽 위 모서리★ (빨간 점)")
    path = os.path.join(OUT, "guide_material_tl.png")
    img.save(path)
    specs.append(path)

    # 2 material_br — 마지막 행 자재코드 오른쪽 아래
    img, L = draw_sap_base()
    d = ImageDraw.Draw(img)
    last_bottom = L["header_bottom"] + L["n_rows"] * L["row_h"]
    x = L["code_right"] - 2
    y = last_bottom - 2
    draw_dot(d, x, y, u"여기를 클릭!")
    d.rectangle(
        [L["code_left"], L["header_bottom"], L["code_right"], last_bottom],
        outline=(220, 0, 0),
        width=2,
    )
    caption(d, u"2/5  마지막 행 자재코드 칸의 ★오른쪽 아래 모서리★ (빨간 점)")
    path = os.path.join(OUT, "guide_material_br.png")
    img.save(path)
    specs.append(path)

    # 3 qty_x — 오더수량 중앙
    img, L = draw_sap_base()
    d = ImageDraw.Draw(img)
    x = L["qty_cx"]
    y = L["header_bottom"] + L["row_h"] // 2
    draw_dot(d, x, y, u"여기를 클릭!")
    d.rectangle(
        [
            L["qty_left"],
            L["header_bottom"],
            L["qty_right"],
            L["header_bottom"] + L["row_h"],
        ],
        outline=(220, 0, 0),
        width=2,
    )
    caption(d, u"3/5  첫 행 오더수량 숫자 ★정중앙★ (빨간 점)")
    path = os.path.join(OUT, "guide_qty_x.png")
    img.save(path)
    specs.append(path)

    # 4 row1_y — 첫 행 세로 중앙
    img, L = draw_sap_base()
    d = ImageDraw.Draw(img)
    x = (L["code_left"] + L["code_right"]) // 2
    y = L["header_bottom"] + L["row_h"] // 2
    draw_dot(d, x, y, u"여기를 클릭!")
    d.line(
        [L["code_left"], y, L["code_right"], y],
        fill=(220, 0, 0),
        width=2,
    )
    caption(d, u"4/5  첫 데이터 행 ★세로 중앙★ (빨간 점)")
    path = os.path.join(OUT, "guide_row1_y.png")
    img.save(path)
    specs.append(path)

    # 5 row2_y — 둘째 행 세로 중앙
    img, L = draw_sap_base()
    d = ImageDraw.Draw(img)
    x = (L["code_left"] + L["code_right"]) // 2
    y = L["header_bottom"] + L["row_h"] + L["row_h"] // 2
    draw_dot(d, x, y, u"여기를 클릭!")
    d.line(
        [L["code_left"], y, L["code_right"], y],
        fill=(220, 0, 0),
        width=2,
    )
    caption(d, u"5/5  둘째 데이터 행 ★세로 중앙★ (빨간 점)")
    path = os.path.join(OUT, "guide_row2_y.png")
    img.save(path)
    specs.append(path)

    print("created", len(specs))
    for p in specs:
        print(p)
    return 0


if __name__ == "__main__":
    raise SystemExit(make_all())

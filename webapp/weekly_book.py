"""주간 식단표·검식일지 엑셀(.xlsm) 만들기.

신태양에서 쓰는 주별 파일(예: 11월-11.16~11.22식단표_검식일지.xlsm)과 같은 모양으로 만든다.
weekly_template.xlsm 을 바탕으로
  - 설정 시트: 주 시작일(월요일)과 검식자·시간 등
  - 식단표 시트 5~18행: 중식·석식 메뉴와 오전/오후 간식
  - 메뉴DB: 서버에 추가 저장된 메뉴의 검식 항목
  - 주간식단표(사진), 보존식 기록지(사진): 식판 사진·보존식 사진
만 채운다. 보존식기록표·검식일지의 날짜·메뉴·검사항목은 원래 파일처럼 수식이 엑셀에서 계산한다.
"""
from __future__ import annotations

import io
import re
from datetime import date, datetime, timedelta
from pathlib import Path

import openpyxl
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.utils import get_column_letter

HERE = Path(__file__).resolve().parent
TEMPLATE = HERE / "weekly_template.xlsm"
TEMPLATE_MENU_SHEET = "11월 식단표"

ROLES = ["밥", "국", "주메뉴", "조림볶음", "무침", "김치"]
MEALS = ["중식", "석식"]
DB_COLS = ["분량", "영양", "색", "맛", "풍미", "질감", "가열상태", "온도", "용기의 청결", "재가열", "폐기조치"]

# 설정 시트 칸 (키: 화면에서 보내는 이름)
SETTING_CELLS = {
    "inspector": "B5",      # 검식자
    "inspect_time": "B6",   # 검식시간
    "lunch_time": "B7",     # 채취시간(중식)
    "dinner_time": "B8",    # 채취시간(석식)
    "freezer": "B9",        # 냉동고 온도
    "keep_days": "B10",     # 보존식 보관일수
    "keeper": "B11",        # 담당자/채취자
    "sat_keeper": "B12",    # 토요일 채취자
    "sun_keeper": "B13",    # 일요일 채취자
}

EMU_PER_PX = 9525
MAX_PHOTO_BYTES = 8_000_000
PHOTO_SLOT_RE = re.compile(r"^d([0-6])_(lunch|dinner)_(tray|keep)$")


def parse_monday(s: str) -> date:
    d = datetime.strptime(str(s).strip()[:10], "%Y-%m-%d").date()
    if d.weekday() != 0:
        raise ValueError("주 시작일은 월요일이어야 합니다.")
    return d


def output_name(monday: date) -> str:
    sun = monday + timedelta(days=6)
    return f"{monday.month}월-{monday.month}.{monday.day:02d}~{sun.month}.{sun.day:02d}식단표_검식일지.xlsm"


def template_menudb() -> dict[str, list[str]]:
    wb = openpyxl.load_workbook(TEMPLATE, read_only=True)
    ws = wb["메뉴DB"]
    out: dict[str, list[str]] = {}
    for i, r in enumerate(ws.iter_rows(values_only=True)):
        if i == 0 or not r or not r[0]:
            continue
        name = str(r[0]).strip()
        if name:
            out[name] = [("" if v is None else str(v)) for v in list(r[1:12]) + [None] * (11 - len(r[1:12]))]
    wb.close()
    return out


def _clean(v) -> str:
    return re.sub(r"\s+", " ", str(v or "")).strip()[:60]


def _col_px(ws, col: int) -> int:
    w = ws.column_dimensions[get_column_letter(col)].width
    if not w:
        w = 8.43
    return int(w * 7 + 5)


def _row_px(ws, row: int) -> int:
    h = ws.row_dimensions[row].height or ws.sheet_format.defaultRowHeight or 15
    return int(h * 96 / 72)


def _place_photo(ws, data: bytes, row: int, col_from: int, col_to: int, pad: int = 4):
    """사진을 row 행의 col_from~col_to 칸 안에 비율을 지켜 가운데 맞춰 넣는다."""
    from PIL import Image as PILImage, ImageOps

    im = PILImage.open(io.BytesIO(data))
    im = ImageOps.exif_transpose(im).convert("RGB")
    box_w = sum(_col_px(ws, c) for c in range(col_from, col_to + 1)) - pad * 2
    box_h = _row_px(ws, row) - pad * 2
    s = min(box_w / im.width, box_h / im.height)
    w, h = max(1, int(im.width * s)), max(1, int(im.height * s))
    # 파일 크기를 줄이려고 화면 크기의 2배까지만 남긴다
    keep = min(1.0, (w * 2) / im.width)
    if keep < 1.0:
        im = im.resize((max(1, int(im.width * keep)), max(1, int(im.height * keep))), PILImage.LANCZOS)
    buf = io.BytesIO()
    im.save(buf, "JPEG", quality=85)
    buf.seek(0)
    xi = XLImage(buf)
    xi.width, xi.height = w, h
    # 왼쪽 위 칸 기준 여백(가운데 정렬)
    off_x = pad + (box_w - w) // 2
    off_y = pad + (box_h - h) // 2
    c = col_from
    while c < col_to and off_x >= _col_px(ws, c):
        off_x -= _col_px(ws, c)
        c += 1
    marker = AnchorMarker(col=c - 1, colOff=off_x * EMU_PER_PX, row=row - 1, rowOff=off_y * EMU_PER_PX)
    xi.anchor = OneCellAnchor(_from=marker, ext=XDRPositiveSize2D(w * EMU_PER_PX, h * EMU_PER_PX))
    ws.add_image(xi)


def _rename_menu_sheet(wb, new_title: str):
    old = TEMPLATE_MENU_SHEET
    if new_title == old:
        return
    wb[old].title = new_title
    olds = [f"'{old}'!", f"{old}!"]
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                if isinstance(v, str) and v.startswith("=") and old in v:
                    for o in olds:
                        v = v.replace(o, f"'{new_title}'!")
                    cell.value = v
    a1 = wb[new_title]["A1"]
    if isinstance(a1.value, str) and a1.value.strip() == old:
        a1.value = new_title


def build_weekly(monday: date, days: list, settings: dict | None = None,
                 extra_db: dict | None = None, photos: dict | None = None) -> tuple[str, bytes, dict]:
    """주간 엑셀을 만든다. (파일이름, 내용, 요약) 을 돌려준다."""
    settings = settings or {}
    extra_db = extra_db or {}
    photos = photos or {}
    if not isinstance(days, list) or len(days) != 7:
        raise ValueError("7일치 식단이 필요합니다.")

    wb = openpyxl.load_workbook(TEMPLATE, keep_vba=True)
    sheet_title = f"{monday.month}월 식단표"
    _rename_menu_sheet(wb, sheet_title)

    st = wb["설정"]
    st["B3"].value = datetime(monday.year, monday.month, monday.day)
    for key, cell in SETTING_CELLS.items():
        if key in settings and str(settings[key]).strip() != "":
            v = str(settings[key]).strip()[:40]
            if key == "keep_days":
                try:
                    v = max(1, min(30, int(float(v))))
                except ValueError:
                    continue
            st[cell].value = v

    ms = wb[sheet_title]
    mains = []
    for di, d in enumerate(days):
        d = d or {}
        col = 2 + di
        for mi, meal in enumerate(MEALS):
            m = d.get(meal) or {}
            for ri, role in enumerate(ROLES):
                v = _clean(m.get(role))
                ms.cell(5 + mi * 6 + ri, col).value = v or None
            main = _clean(m.get("주메뉴"))
            if main:
                mains.append(main)
        ms.cell(17, col).value = _clean(d.get("오전")) or None
        ms.cell(18, col).value = _clean(d.get("오후")) or None

    # 메뉴DB: 템플릿 + 서버에 저장된 추가 메뉴 (같은 이름이면 추가 저장값이 우선)
    db = wb["메뉴DB"]
    row_of = {}
    for r in range(2, db.max_row + 1):
        n = db.cell(r, 1).value
        if n:
            row_of[str(n).strip()] = r
    style = [db.cell(2, c)._style for c in range(1, 13)]
    nxt = db.max_row + 1
    for name, vals in extra_db.items():
        name = _clean(name)
        if not name:
            continue
        vals = [(_clean(v) if v is not None else "") for v in list(vals)[:11]] + [""] * (11 - len(list(vals)[:11]))
        r = row_of.get(name)
        if r is None:
            r = nxt
            nxt += 1
            row_of[name] = r
        db.cell(r, 1).value = name
        for c, v in enumerate(vals, 2):
            db.cell(r, c).value = v or None
        for c in range(1, 13):
            db.cell(r, c)._style = style[c - 1]
    missing = sorted({m for m in mains if m not in row_of})

    # 사진
    pw = wb["주간식단표(사진)"]
    pk = wb["보존식 기록지(사진)"]
    placed, bad = [], []
    for slot, data in photos.items():
        m = PHOTO_SLOT_RE.match(slot)
        if not m or not data:
            continue
        di, meal, kind = int(m.group(1)), m.group(2), m.group(3)
        try:
            if kind == "tray":
                _place_photo(pw, data, 3 + di, 3 if meal == "lunch" else 5, 3 if meal == "lunch" else 5)
            else:
                block = di * 2 + (0 if meal == "lunch" else 1)
                _place_photo(pk, data, 4 + block * 10, 2, 3)
            placed.append(slot)
        except Exception:
            bad.append(slot)

    wb.calculation.fullCalcOnLoad = True
    buf = io.BytesIO()
    wb.save(buf)
    return output_name(monday), buf.getvalue(), {"missing_menudb": missing, "photos": len(placed), "bad_photos": bad}

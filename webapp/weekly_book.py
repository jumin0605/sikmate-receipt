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
import zipfile
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
# 사용자가 올린 엑셀을 양식으로 쓰기로 하면 여기에 저장된다 (app.py 가 영구 디스크 경로로 정해 준다)
CUSTOM_TEMPLATE: Path | None = None
REQUIRED_SHEETS = ["설정", "메뉴DB", "주간식단표(사진)", "보존식 기록지(사진)"]


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


def active_template() -> Path:
    if CUSTOM_TEMPLATE is not None and Path(CUSTOM_TEMPLATE).is_file():
        return Path(CUSTOM_TEMPLATE)
    return TEMPLATE


def find_menu_sheet(wb) -> str:
    """'11월 식단표' 처럼 이름이 '식단표'로 끝나는 입력 시트 (주간식단표(사진)은 제외)."""
    for name in wb.sheetnames:
        n = name.strip()
        if n.endswith("식단표") and "주간" not in n and "사진" not in n:
            return name
    raise ValueError("식단표 시트(예: 11월 식단표)를 찾지 못했습니다.")


def template_menudb() -> dict[str, list[str]]:
    wb = openpyxl.load_workbook(active_template(), read_only=True)
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
    old = find_menu_sheet(wb)
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
                 extra_db: dict | None = None, photos: dict | None = None,
                 keep_macros: bool = False) -> tuple[str, bytes, dict]:
    """주간 엑셀을 만든다. (파일이름, 내용, 요약) 을 돌려준다."""
    settings = settings or {}
    extra_db = extra_db or {}
    photos = photos or {}
    if not isinstance(days, list) or len(days) != 7:
        raise ValueError("7일치 식단이 필요합니다.")

    wb = openpyxl.load_workbook(active_template(), keep_vba=True)
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
    out = buf.getvalue()
    had = has_vba(out)
    if had and not keep_macros:
        out = strip_vba(out)
    return output_name(monday), out, {"missing_menudb": missing, "photos": len(placed), "bad_photos": bad,
                                       "macros": bool(had and keep_macros), "macros_removed": bool(had and not keep_macros)}


# ---------------------------------------------------------------------------
# 사용자가 만든 주별 엑셀 가져오기
# ---------------------------------------------------------------------------
def _sheet_col_edges(ws, upto: int = 30) -> list[int]:
    edges = [0]
    for c in range(1, upto + 1):
        edges.append(edges[-1] + _col_px(ws, c))
    return edges


def _sheet_row_edges(ws, upto: int) -> list[int]:
    edges = [0]
    for r in range(1, upto + 1):
        edges.append(edges[-1] + _row_px(ws, r))
    return edges


def _image_center(img, cols: list[int], rows: list[int]):
    """그림의 가운데 위치(px). 열·행 번호는 1부터."""
    a = img.anchor
    fr = getattr(a, "_from", None)
    if fr is None:
        return None
    x1 = cols[min(fr.col, len(cols) - 1)] + (fr.colOff or 0) / EMU_PER_PX
    y1 = rows[min(fr.row, len(rows) - 1)] + (fr.rowOff or 0) / EMU_PER_PX
    to = getattr(a, "to", None)
    ext = getattr(a, "ext", None)
    if to is not None:
        x2 = cols[min(to.col, len(cols) - 1)] + (to.colOff or 0) / EMU_PER_PX
        y2 = rows[min(to.row, len(rows) - 1)] + (to.rowOff or 0) / EMU_PER_PX
    elif ext is not None:
        x2 = x1 + (ext.width or 0) / EMU_PER_PX
        y2 = y1 + (ext.height or 0) / EMU_PER_PX
    else:
        x2, y2 = x1 + (img.width or 0), y1 + (img.height or 0)
    return (x1 + x2) / 2, (y1 + y2) / 2


def _locate(edges: list[int], v: float) -> int:
    for i in range(1, len(edges)):
        if v < edges[i]:
            return i
    return len(edges)


def _small_jpeg(data: bytes, side: int = 900) -> bytes | None:
    from PIL import Image as PILImage, ImageOps
    try:
        im = PILImage.open(io.BytesIO(data))
        im = ImageOps.exif_transpose(im).convert("RGB")
        im.thumbnail((side, side))
        b = io.BytesIO()
        im.save(b, "JPEG", quality=85)
        return b.getvalue()
    except Exception:
        return None


def _photos_from(wb) -> tuple[dict, int]:
    """주간식단표(사진)·보존식 기록지(사진) 시트의 그림을 칸(d0_lunch_tray 등)으로 나눈다."""
    out, skipped = {}, 0
    if "주간식단표(사진)" in wb.sheetnames:
        ws = wb["주간식단표(사진)"]
        cols, rows = _sheet_col_edges(ws, 10), _sheet_row_edges(ws, 20)
        for img in ws._images:
            c = _image_center(img, cols, rows)
            if not c:
                skipped += 1
                continue
            col, row = _locate(cols, c[0]), _locate(rows, c[1])
            di = row - 3
            if not 0 <= di <= 6:
                skipped += 1
                continue
            meal = "lunch" if col <= 3 else "dinner"
            slot = f"d{di}_{meal}_tray"
            if slot in out:
                skipped += 1
                continue
            out[slot] = img
    if "보존식 기록지(사진)" in wb.sheetnames:
        ws = wb["보존식 기록지(사진)"]
        cols, rows = _sheet_col_edges(ws, 10), _sheet_row_edges(ws, 150)
        for img in ws._images:
            c = _image_center(img, cols, rows)
            if not c:
                skipped += 1
                continue
            row = _locate(rows, c[1])
            block = (row - 3) // 10
            if not 0 <= block <= 13:
                skipped += 1
                continue
            slot = f"d{block // 2}_{'lunch' if block % 2 == 0 else 'dinner'}_keep"
            if slot in out:
                skipped += 1
                continue
            out[slot] = img
    photos = {}
    for slot, img in out.items():
        try:
            data = img._data()
        except Exception:
            data = None
        small = _small_jpeg(data) if data else None
        if small:
            photos[slot] = small
        else:
            skipped += 1
    return photos, skipped


def _date_from_filename(name: str) -> date | None:
    m = re.search(r"(\d{1,2})\.(\d{1,2})\s*~", name or "")
    if not m:
        return None
    today = date.today()
    for y in (today.year, today.year - 1, today.year + 1):
        try:
            d = date(y, int(m.group(1)), int(m.group(2)))
        except ValueError:
            continue
        if d.weekday() == 0 and abs((d - today).days) < 200:
            return d
    return None


def _txt(v) -> str:
    if v is None:
        return ""
    if isinstance(v, str) and v.startswith("="):
        return ""
    return _clean(v)


def read_weekly(data: bytes, filename: str = "", with_photos: bool = True) -> dict:
    """사용자가 만든 주별 엑셀에서 주 시작일·식단·설정·메뉴DB·사진을 읽는다."""
    try:
        wb = openpyxl.load_workbook(io.BytesIO(data))
    except Exception:
        raise ValueError("엑셀 파일을 열지 못했습니다. (.xlsm / .xlsx 파일인지 확인해 주세요)")
    warnings = []
    menu_name = find_menu_sheet(wb)
    ms = wb[menu_name]
    monday = None
    if "설정" in wb.sheetnames:
        v = wb["설정"]["B3"].value
        if isinstance(v, datetime):
            monday = v.date()
        elif isinstance(v, date):
            monday = v
    if monday is None:
        monday = _date_from_filename(filename)
        if monday:
            warnings.append("설정 시트에서 주 시작일을 못 읽어 파일 이름의 날짜를 썼습니다.")
    if monday is None:
        raise ValueError("주 시작일(설정 시트 B3 또는 파일 이름의 날짜)을 찾지 못했습니다.")
    if monday.weekday() != 0:
        monday = monday - timedelta(days=monday.weekday())
        warnings.append("주 시작일이 월요일이 아니라서 그 주 월요일로 맞췄습니다.")

    days = []
    for di in range(7):
        col = 2 + di
        d = {}
        for mi, meal in enumerate(MEALS):
            d[meal] = {role: _txt(ms.cell(5 + mi * 6 + ri, col).value) for ri, role in enumerate(ROLES)}
        d["오전"] = _txt(ms.cell(17, col).value)
        d["오후"] = _txt(ms.cell(18, col).value)
        days.append(d)

    settings = {}
    if "설정" in wb.sheetnames:
        st = wb["설정"]
        for key, cell in SETTING_CELLS.items():
            v = st[cell].value
            if v is not None and not (isinstance(v, str) and v.startswith("=")):
                settings[key] = _clean(v)

    menudb = {}
    if "메뉴DB" in wb.sheetnames:
        for i, r in enumerate(wb["메뉴DB"].iter_rows(values_only=True)):
            if i == 0 or not r or not r[0] or not str(r[0]).strip():
                continue
            vals = [("" if v is None else _clean(v)) for v in list(r[1:12])]
            vals += [""] * (11 - len(vals))
            menudb[_clean(r[0])] = vals

    photos, skipped = (_photos_from(wb) if with_photos else ({}, 0))
    missing_sheets = [n for n in REQUIRED_SHEETS if n not in wb.sheetnames]
    return {
        "monday": monday.isoformat(),
        "menu_sheet": menu_name,
        "days": days,
        "settings": settings,
        "menudb": menudb,
        "photos": photos,
        "photos_skipped": skipped,
        "sheets": wb.sheetnames,
        "template_ok": not missing_sheets,
        "missing_sheets": missing_sheets,
        "warnings": warnings,
    }


def make_template(data: bytes) -> bytes:
    """사용자 엑셀을 앞으로 쓸 양식으로 만든다: 메뉴·메모·사진만 비우고 나머지(서식·수식·메뉴DB)는 그대로."""
    try:
        wb = openpyxl.load_workbook(io.BytesIO(data), keep_vba=True)
    except Exception:
        raise ValueError("엑셀 파일을 열지 못했습니다.")
    missing = [n for n in REQUIRED_SHEETS if n not in wb.sheetnames]
    if missing:
        raise ValueError("양식으로 쓰려면 이 시트가 있어야 합니다: " + ", ".join(missing))
    ms = wb[find_menu_sheet(wb)]
    for r in range(5, 19):
        for c in range(2, 9):
            ms.cell(r, c).value = None
    for r in range(1, 19):
        for c in range(10, 15):
            ms.cell(r, c).value = None
    pw = wb["주간식단표(사진)"]
    pw._images = []
    for r in range(3, 10):
        for c in (3, 5):
            v = pw.cell(r, c).value
            if not (isinstance(v, str) and v.startswith("=")):
                pw.cell(r, c).value = None
    wb["보존식 기록지(사진)"]._images = []
    wb.calculation.fullCalcOnLoad = True
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def has_vba(data_or_path) -> bool:
    try:
        z = zipfile.ZipFile(io.BytesIO(data_or_path) if isinstance(data_or_path, (bytes, bytearray)) else data_or_path)
        return any(n.lower().startswith("xl/vbaproject") for n in z.namelist())
    except Exception:
        return False


def strip_vba(data: bytes) -> bytes:
    """매크로(vbaProject)만 빼고 나머지는 그대로 둔 .xlsm 을 만든다.
    인터넷에서 받은 파일에 매크로가 있으면 엑셀이 '보안 위험'으로 막기 때문에, 웹에서 받는 파일은 기본으로 매크로를 뺀다."""
    zin = zipfile.ZipFile(io.BytesIO(data))
    drop = {n for n in zin.namelist() if n.lower().startswith("xl/vbaproject")}
    if not drop:
        return data
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            n = info.filename
            if n in drop:
                continue
            b = zin.read(n)
            if n == "[Content_Types].xml":
                t = b.decode("utf-8")
                t = re.sub(r'<Override[^>]*PartName="/xl/vbaProject[^"]*"[^>]*/>', "", t)
                t = re.sub(r'<Default[^>]*Extension="bin"[^>]*ContentType="application/vnd\.ms-office\.vbaProject"[^>]*/>', "", t)
                b = t.encode("utf-8")
            elif n == "xl/_rels/workbook.xml.rels" or (n.startswith("xl/_rels/") and n.endswith(".rels")):
                t = b.decode("utf-8")
                t = re.sub(r'<Relationship[^>]*Target="[^"]*vbaProject[^"]*"[^>]*/>', "", t)
                b = t.encode("utf-8")
            zout.writestr(info, b)
    return out.getvalue()

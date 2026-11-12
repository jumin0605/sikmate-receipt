# -*- coding: utf-8 -*-
"""식자재 검수서(원산지 표시) - 웹용 도우미

원산지 표시 사진을 받아 A4 한 쪽에 3열로 놓은 엑셀(또는 워드)을 만든다.
표 모양은 기존 도구(영수증_정리_도구.py)의 _insp_excel / _insp_word 를 그대로 쓴다.
이 파일은 업로드 사진을 안전하게 준비하는 일(회전 바로잡기, 크기 줄이기, 이름 정리)과
사진 방향 자동 판단(Gemini)만 맡는다. fastapi 없이도 시험할 수 있게 app.py 와 따로 둔다.
"""

from __future__ import annotations

import io
import re
from datetime import date, datetime
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont, ImageOps

# 영수증 올리기와 같은 사진 형식(PDF 제외). HEIC 는 서버에서 열 수 없다.
INSP_ALLOWED_EXT = {".jpg", ".jpeg", ".png", ".webp", ".bmp", ".tif", ".tiff"}
# 엑셀 칸에 들어가는 사진은 155x116px 정도라 긴 변 1600px 이면 충분하다. 메모리도 아낀다.
INSP_MAX_SIDE = 1600
INSP_TITLE = "식자재 검수서 원산지 표시"
ROTATIONS = (0, 90, 180, 270)      # 시계 방향


def natural_key(name: str):
    """'사진2' 가 '사진10' 보다 앞에 오도록 숫자는 숫자로 비교한다."""
    return [int(t) if t.isdigit() else t.lower() for t in re.split(r"(\d+)", name)]


def clean_stem(name: str) -> str:
    """사진 위에 적힐 이름. 파일 이름에서 확장자와 못 쓰는 글자를 뺀다."""
    stem = Path(name or "").stem
    stem = re.sub(r'[\\/:*?"<>|\x00-\x1f]', "_", stem).strip(" .")
    return stem[:60] or "원산지 사진"


def parse_date(text: str | None) -> date | None:
    """'2026-11-12' 를 날짜로. 비었거나 형식이 다르면 None."""
    text = (text or "").strip()
    if not text:
        return None
    try:
        return datetime.strptime(text, "%Y-%m-%d").date()
    except ValueError:
        return None


def title_and_filename(d: date | None, today: date | None = None) -> tuple[str, str]:
    """제목: '식자재 검수서 원산지 표시  2026년 11월 12일' (날짜 없으면 제목만)
    파일 이름: '원산지 표시_26.11.12' (날짜 없으면 오늘 날짜)"""
    if d:
        title = f"{INSP_TITLE}  {d.year}년 {d.month}월 {d.day}일"
    else:
        title = INSP_TITLE
    stamp = (d or today or date.today()).strftime("%y.%m.%d")
    return title, f"원산지 표시_{stamp}"


def norm_rotation(v) -> int:
    """사용자가 고른 회전(시계 방향, 도). 0/90/180/270 이 아니면 0."""
    try:
        v = int(v) % 360
    except (TypeError, ValueError):
        return 0
    return v if v in ROTATIONS else 0


def open_upright(data: bytes) -> Image.Image:
    """사진을 열어 휴대폰 회전 정보(EXIF)를 반영하고 흰 배경 RGB 로 돌려준다.
    열 수 없는 파일이면 예외를 낸다."""
    im = Image.open(io.BytesIO(data))
    im.seek(0)                           # 여러 장짜리 TIF 는 첫 장만
    im = ImageOps.exif_transpose(im)
    if im.mode in ("RGBA", "LA", "P"):
        im = im.convert("RGBA")
        bg = Image.new("RGB", im.size, "white")
        bg.paste(im, mask=im.split()[-1])
        return bg
    return im.convert("RGB")


def rotate_cw(im: Image.Image, deg: int) -> Image.Image:
    deg = norm_rotation(deg)
    return im.rotate(-deg, expand=True) if deg else im   # PIL 은 양수가 반시계 방향


def prepare_image(data: bytes, out_path: Path, rotate: int = 0) -> Path:
    """올린 사진을 똑바로 세우고(휴대폰 회전 정보 + 사용자가 고른 회전) 줄여서 JPEG 로 저장한다."""
    im = open_upright(data)
    im.thumbnail((INSP_MAX_SIDE, INSP_MAX_SIDE), Image.Resampling.LANCZOS)
    im = rotate_cw(im, rotate)
    # with_suffix 는 '1. 돼지고기' 의 '. 돼지고기' 를 확장자로 보고 잘라 버리므로 뒤에 붙인다
    out_path = out_path.parent / (out_path.name + ".jpg")
    im.save(out_path, "JPEG", quality=90)
    return out_path


def unique_stem(stem: str, used: set) -> str:
    """같은 이름이 있으면 이름_2, 이름_3 을 붙인다."""
    new, n = stem, 2
    while new.lower() in used:
        new = f"{stem}_{n}"
        n += 1
    used.add(new.lower())
    return new


def build_inspection(core, uploads, workdir: Path, date_text: str = "", kind: str = "xlsx",
                     log=lambda m: None):
    """uploads: [(원래 파일 이름, bytes[, 시계 방향 회전 도])].
    만든 파일 경로, 받을 때 이름, 건너뛴 파일 목록, 넣은 사진 수를 돌려준다.

    사진 순서는 파일 이름 순(숫자는 숫자로)이다. 기존 도구도 이름 순으로 놓는다.
    """
    if kind not in ("xlsx", "docx"):
        raise ValueError("알 수 없는 파일 형식입니다.")
    pics = workdir / "사진"
    pics.mkdir(parents=True, exist_ok=True)
    used, images, skipped = set(), [], []
    for up in sorted(uploads, key=lambda u: natural_key(Path(u[0] or "").name)):
        name, data = up[0], up[1]
        rot = up[2] if len(up) > 2 else 0
        base = Path(name or "").name
        ext = Path(base).suffix.lower()
        if ext not in INSP_ALLOWED_EXT:
            skipped.append(f"{base} (지원하지 않는 형식)")
            continue
        stem = unique_stem(clean_stem(base), used)
        try:
            images.append(str(prepare_image(data, pics / stem, rot)))
        except Exception:  # noqa: BLE001  깨진 사진 한 장 때문에 전체를 멈추지 않는다
            skipped.append(f"{base} (사진을 열 수 없음)")
    if not images:
        raise ValueError("검수서에 넣을 수 있는 사진이 없습니다. " + ", ".join(skipped))

    title, stem = title_and_filename(parse_date(date_text))
    out = workdir / f"{stem}.{kind}"
    if kind == "xlsx":
        core._insp_excel(images, title, str(out), log=log)
    else:
        core._insp_word(images, title, str(out), log=log)
    return out, out.name, skipped, len(images)


# ---------------------------------------------------------------------------
# 사진 방향 자동 판단
#
# 카카오톡으로 받은 사진은 휴대폰의 회전 정보(EXIF)가 빠져 있어 옆으로 눕거나 뒤집힌 채로 온다.
# 정보가 없으니 사진 내용을 보고 판단해야 한다. 같은 사진을 네 방향(0/90/180/270도)으로 돌려
# 한 장에 A~D 로 붙여 보여 주고, 글자와 물건이 똑바로 선 것을 고르게 한다.
# ("몇 도 돌려야 하나"를 직접 물으면 AI 가 자주 틀리고, 네 개 중에서 고르게 하면 더 정확하다.)
# ---------------------------------------------------------------------------
ORIENT_LETTERS = "ABCD"
ORIENT_PROMPT = (
    "이 그림에는 같은 사진(식자재 포장지나 원산지 표시)을 네 방향으로 돌린 것이 A, B, C, D 칸에 있습니다. "
    "포장지의 글자(한글, 숫자)와 물건이 사람이 평소에 보는 방향으로 똑바로 서 있는 칸을 하나 고르세요. "
    "글자가 있으면 글자가 바로 읽히는 칸을 가장 우선으로 고르세요. 글자가 위아래로 뒤집혀 있거나 옆으로 "
    "누워 있으면 안 됩니다. 칸 모서리의 빨간 A, B, C, D 표시는 판단에 쓰지 마세요."
)
ORIENT_SCHEMA = {"type": "object", "required": ["upright"],
                 "properties": {"upright": {"type": "string", "enum": list(ORIENT_LETTERS)}}}


def _label_font(size):
    try:
        return ImageFont.load_default(size=size)
    except TypeError:                    # 아주 옛 Pillow
        return ImageFont.load_default()


def orientation_sheet(data: bytes, cell: int = 380) -> bytes:
    """사진 하나를 네 방향으로 돌려 2x2 로 붙인 PNG. A=그대로, B=시계 90도, C=180도, D=시계 270도."""
    im = open_upright(data)
    im.thumbnail((cell - 20, cell - 20), Image.Resampling.LANCZOS)
    sheet = Image.new("RGB", (cell * 2, cell * 2), "white")
    draw = ImageDraw.Draw(sheet)
    font = _label_font(34)
    for i, deg in enumerate(ROTATIONS):
        r = rotate_cw(im, deg)
        x0, y0 = (i % 2) * cell, (i // 2) * cell
        sheet.paste(r, (x0 + (cell - r.width) // 2, y0 + (cell - r.height) // 2))
        draw.rectangle([x0, y0, x0 + cell - 1, y0 + cell - 1], outline=(150, 150, 150), width=2)
        draw.rectangle([x0 + 2, y0 + 2, x0 + 46, y0 + 46], fill=(220, 38, 38))
        draw.text((x0 + 12, y0 + 5), ORIENT_LETTERS[i], fill="white", font=font)
    buf = io.BytesIO()
    sheet.save(buf, "PNG")
    return buf.getvalue()


# 방향 고르기는 쉬운 질문이라 깊이 생각(thinking)할 필요가 없다. 생각을 켜 두면 한 장에 수 초씩 걸려
# 40장이면 몇 분이 된다. 모델마다 받는 설정이 달라 차례로 시도하고, 통한 설정을 기억해 둔다.
_THINK_OPTIONS = ({"thinking_level": "minimal"}, {"thinking_level": "low"}, {"thinking_budget": 0}, None)
_think_ok: dict = {}


def _ask_orientation(client, model, png: bytes) -> dict:
    import json
    from google.genai import types
    start = _think_ok.get(model, 0)
    last = None
    for i in range(start, len(_THINK_OPTIONS)):
        opt = _THINK_OPTIONS[i]
        try:
            extra = {"thinking_config": types.ThinkingConfig(**opt)} if opt else {}
            cfg = types.GenerateContentConfig(response_mime_type="application/json",
                                              response_schema=ORIENT_SCHEMA, temperature=0, **extra)
        except Exception as e:  # noqa: BLE001  설치된 라이브러리가 그 설정을 모름
            last = e
            continue
        try:
            resp = client.models.generate_content(
                model=model, config=cfg,
                contents=[types.Part.from_bytes(data=png, mime_type="image/png"), ORIENT_PROMPT])
        except Exception as e:  # noqa: BLE001
            msg = str(e).lower()
            if opt and ("think" in msg or "invalid" in msg or "400" in msg):
                last = e          # 이 모델이 이 생각 설정을 받지 않음 -> 다음 설정으로
                continue
            raise
        _think_ok[model] = i
        return json.loads(resp.text)
    raise RuntimeError(f"방향 판단 호출 실패: {last}")


def detect_rotation(core, client, model, data: bytes) -> int:
    """AI 에게 똑바로 선 칸을 고르게 해서 시계 방향 회전 도를 돌려준다."""
    png = orientation_sheet(data)
    if hasattr(core, "json_client"):          # 진짜 서버: 생각을 줄여 빠르게 묻는다
        try:
            ans = _ask_orientation(client, model, png)
        except Exception:  # noqa: BLE001  한 번 더, 기존 도구의 호출(재시도 포함)로
            ans = core.json_ask(client, png, model, schema=ORIENT_SCHEMA, prompt=ORIENT_PROMPT)
    else:                                     # 시험용 가짜 core
        ans = core.json_ask(client, png, model, schema=ORIENT_SCHEMA, prompt=ORIENT_PROMPT)
    letter = str((ans or {}).get("upright", "A")).strip().upper()[:1]
    return ROTATIONS[ORIENT_LETTERS.index(letter)] if letter in ORIENT_LETTERS else 0

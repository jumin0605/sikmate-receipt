# -*- coding: utf-8 -*-
"""식자재 검수서(원산지 표시) - 웹용 도우미

원산지 표시 사진을 받아 A4 한 쪽에 3열로 놓은 엑셀(또는 워드)을 만든다.
표 모양은 기존 도구(영수증_정리_도구.py)의 _insp_excel / _insp_word 를 그대로 쓴다.
이 파일은 업로드 사진을 안전하게 준비하는 일(회전 바로잡기, 크기 줄이기, 이름 정리)만 맡는다.
fastapi 없이도 시험할 수 있게 app.py 와 따로 둔다.
"""

from __future__ import annotations

import re
from datetime import date, datetime
from pathlib import Path

from PIL import Image, ImageOps

# 영수증 올리기와 같은 사진 형식(PDF 제외). HEIC 는 서버에서 열 수 없다.
INSP_ALLOWED_EXT = {".jpg", ".jpeg", ".png", ".webp", ".bmp", ".tif", ".tiff"}
# 엑셀 칸에 들어가는 사진은 155x116px 정도라 긴 변 1600px 이면 충분하다. 메모리도 아낀다.
INSP_MAX_SIDE = 1600
INSP_TITLE = "식자재 검수서 원산지 표시"


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


def prepare_image(data: bytes, out_path: Path) -> Path:
    """올린 사진을 똑바로 세우고(휴대폰 회전 정보 반영) 줄여서 JPEG 로 저장한다.
    열 수 없는 파일이면 예외를 낸다."""
    import io
    with Image.open(io.BytesIO(data)) as im:
        im.seek(0)                       # 여러 장짜리 TIF 는 첫 장만
        im = ImageOps.exif_transpose(im)
        if im.mode in ("RGBA", "LA", "P"):
            im = im.convert("RGBA")
            bg = Image.new("RGB", im.size, "white")
            bg.paste(im, mask=im.split()[-1])
            im = bg
        else:
            im = im.convert("RGB")
        im.thumbnail((INSP_MAX_SIDE, INSP_MAX_SIDE), Image.Resampling.LANCZOS)
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
    """uploads: [(원래 파일 이름, bytes)]. 만든 파일 경로, 받을 때 이름, 건너뛴 파일 목록을 돌려준다.

    사진 순서는 파일 이름 순(숫자는 숫자로)이다. 기존 도구도 이름 순으로 놓는다.
    """
    if kind not in ("xlsx", "docx"):
        raise ValueError("알 수 없는 파일 형식입니다.")
    pics = workdir / "사진"
    pics.mkdir(parents=True, exist_ok=True)
    used, images, skipped = set(), [], []
    for name, data in sorted(uploads, key=lambda u: natural_key(Path(u[0] or "").name)):
        base = Path(name or "").name
        ext = Path(base).suffix.lower()
        if ext not in INSP_ALLOWED_EXT:
            skipped.append(f"{base} (지원하지 않는 형식)")
            continue
        stem = unique_stem(clean_stem(base), used)
        try:
            images.append(str(prepare_image(data, pics / stem)))
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

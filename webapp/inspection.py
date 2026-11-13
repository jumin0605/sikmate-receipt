# -*- coding: utf-8 -*-
"""식자재 검수서(원산지 표시) - 웹용 도우미

원산지 표시 사진을 받아 A4 한 쪽에 3열로 놓은 엑셀(또는 워드)을 만든다.
표 모양은 기존 도구(영수증_정리_도구.py)의 _insp_excel / _insp_word 를 그대로 쓴다.
이 파일은 업로드 사진을 안전하게 준비하는 일(회전 바로잡기, 크기 줄이기, 이름 정리)과
사진 방향 자동 판단(Gemini)만 맡는다. fastapi 없이도 시험할 수 있게 app.py 와 따로 둔다.
"""

from __future__ import annotations

import io
import json
from concurrent.futures import ThreadPoolExecutor
import re
import time
from datetime import date, datetime, timedelta
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


def inspection_date_stamp(date_text: str, today: date | None = None) -> str:
    """파일 이름에 붙일 'YY.MM.DD'. 날짜를 못 읽으면 오늘."""
    return (parse_date(date_text) or today or date.today()).strftime("%y.%m.%d")


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
                     log=lambda m: None, sort: bool = True):
    """uploads: [(원래 파일 이름, bytes[, 시계 방향 회전 도])].
    만든 파일 경로, 받을 때 이름, 건너뛴 파일 목록, 넣은 사진 수를 돌려준다.

    사진 순서는 파일 이름 순(숫자는 숫자로)이다. 기존 도구도 이름 순으로 놓는다.
    """
    if kind not in ("xlsx", "docx"):
        raise ValueError("알 수 없는 파일 형식입니다.")
    pics = workdir / "사진"
    pics.mkdir(parents=True, exist_ok=True)
    used, images, skipped = set(), [], []
    ordered = sorted(uploads, key=lambda u: natural_key(Path(u[0] or "").name)) if sort else list(uploads)
    for up in ordered:
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
    "누워 있으면 안 됩니다. 글자가 없으면 탁자·바닥·상자가 아래에 있고 물건이 놓인 방향을 보세요. "
    "칸 모서리의 빨간 A, B, C, D 표시는 판단에 쓰지 마세요.\n"
    "글자도 없고 위아래를 알 수 없는 사진(예: 위에서 내려다본 과자·마카롱)이면 sure 를 false 로 하세요."
)
ORIENT_SCHEMA = {"type": "object", "required": ["upright", "sure"],
                 "properties": {"upright": {"type": "string", "enum": list(ORIENT_LETTERS)},
                                "sure": {"type": "boolean"}}}


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
    return _ask_json(client, model, png, "image/png", ORIENT_PROMPT, ORIENT_SCHEMA, "방향")


def _ask_json(client, model, data: bytes, mime: str, prompt: str, schema: dict, tag: str,
              build=None) -> dict:
    """사진 한 장 + 질문을 보내 JSON 답을 받는다. 생각(thinking)을 최소로 줄여 빠르게 묻는다."""
    import json
    from google.genai import types
    start = _think_ok.get(model, 0)
    last = None
    for i in range(start, len(_THINK_OPTIONS)):
        opt = _THINK_OPTIONS[i]
        try:
            extra = {"thinking_config": types.ThinkingConfig(**opt)} if opt else {}
            cfg = types.GenerateContentConfig(response_mime_type="application/json",
                                              response_schema=schema, temperature=0, **extra)
        except Exception as e:  # noqa: BLE001  설치된 라이브러리가 그 설정을 모름
            last = e
            continue
        try:
            t0 = time.time()
            resp = client.models.generate_content(
                model=model, config=cfg,
                contents=(build(types) if build else [types.Part.from_bytes(data=data, mime_type=mime), prompt]))
            print(f"[{tag}] {model} {opt} {time.time() - t0:.1f}초", flush=True)
        except Exception as e:  # noqa: BLE001
            msg = str(e).lower()
            if opt and ("think" in msg or "invalid" in msg or "400" in msg):
                last = e          # 이 모델이 이 생각 설정을 받지 않음 -> 다음 설정으로
                continue
            raise
        _think_ok[model] = i
        return json.loads(resp.text)
    raise RuntimeError(f"{tag} 판단 호출 실패: {last}")


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
    if (ans or {}).get("sure") is False:      # 확신이 없으면 돌리지 않는다 (멀쩡한 사진을 돌려 망치지 않게)
        return 0
    letter = str((ans or {}).get("upright", "A")).strip().upper()[:1]
    return ROTATIONS[ORIENT_LETTERS.index(letter)] if letter in ORIENT_LETTERS else 0


# ---------------------------------------------------------------------------
# 사진 내용 읽기 -> 파일 이름 자동 만들기
#
# 원산지 표시 사진에서 품명과 원산지를 읽어 '참치 통조림 26.11.09 러시아산' 같은 이름을 만든다.
# 날짜는 사진 속 날짜를 우선한다.
#   - 유통기한(소비기한) 날짜가 적혀 있으면 그 날짜
#   - 제조일·포장일·생산일이고 '기준일로부터 24개월' 같은 기간이 있으면 기준일 + 기간 (26.11.13 + 24개월 = 28.11.13)
#   - 기간 없이 제조일만 있으면 그 날짜
#   - 사진에 날짜가 없으면 파일 이름의 날짜 (카카오톡 사진 KakaoTalk_20261109_135452116 -> 26.11.09)
# ---------------------------------------------------------------------------
READ_PROMPT = (
    "이 사진은 식자재를 납품받을 때 찍은 포장지 또는 원산지 표시 사진입니다. 글자를 읽어 알려 주세요.\n"
    "- product: 식자재의 짧은 일반 이름 (예: '참치 통조림', '돼지고기 앞다리', '배추김치', '계란'). "
    "상표·회사 이름, 용량, 규격, 등급은 빼세요. 읽을 수 없거나 식자재가 아니면 null.\n"
    "- origin: 원산지를 '나라 이름+산' 한 가지로 (예: '국내산', '러시아산', '중국산', '미국산', '노르웨이산'). "
    "원료가 여러 나라면 맨 앞에 적힌(가장 많은) 원료의 원산지 하나만. '국산'은 '국내산'으로. "
    "원산지가 보이지 않으면 null. 짐작해서 적지 마세요.\n"
    "- base_date: 포장지에 인쇄된 제조일·포장일·생산일·도축일 등 '기준일'. YYYY-MM-DD 형식, 연도가 두 자리(26.11.13)면 20대로 보세요(2026). 없으면 null.\n"
    "- expiry_date: '유통기한', '소비기한', '품질유지기한' 뒤에 날짜가 직접 적혀 있을 때만, YYYY-MM-DD. 없으면 null.\n"
    "- shelf_months: '제조일로부터 24개월', '포장일로부터 18개월까지', '기준일로부터 12개월' 처럼 기준일에서 몇 개월인지 적혀 있을 때 그 개월 수(숫자). "
    "'일'이나 년 단위면 개월로 바꿔서(1년=12). 없으면 null.\n"
    "'제조연월일 별도표기', '제조일자는 포장에 표기' 라고 적혀 있으면, 포장 앞뒤 어딘가에 점선·잉크로 찍힌 날짜가 제조일입니다. "
    "라벨 없이 찍힌 그 날짜를 base_date 로 하고, '소비기한: 제조일로부터 24개월까지' 같은 말이 있으면 shelf_months 를 24 로 하세요.\n"
    "- all_dates: 사진에 인쇄된 날짜를 모두 YYYY-MM-DD 로(바코드·로트 번호는 제외). 없으면 빈 배열.\n"
    "날짜가 두 개 이상 나란히 있고 이름표가 안 보이면, 이른 날짜는 제조·포장일이고 늦은 날짜가 유통기한입니다. "
    "사진이 옆으로 누웠거나 뒤집혀 있어도 글자를 돌려서 읽으세요.\n"
    "알려진 포장: 노랑·주황색 비닐 봉지에 닭 그림이 있고 빨간 둥근 도장에 'BRASIL' 또는 'SIF' 가 인쇄된 냉동 닭고기는 "
    "product '닭고기', origin '브라질'. 'BRASIL'/'BRAZIL'/'SIF' 검사 도장이 보이면 origin 은 브라질입니다. "
    "영문 포장이라도 닭(FRANGO/CHICKEN)이면 product 는 닭고기입니다(부위가 보이면 '닭가슴살'처럼).\n"
    "- fresh_vegetable: 신선 채소·버섯·나물(표고, 대파, 무, 양파, 브로콜리, 양배추, 오이, 느타리, 청경채, 열무 등)이면 true. "
    "김치·통조림·절임·가루·말린 것·냉동 등 가공품이거나 고기·수산물·계란·과일이면 false.\n"
    "날짜는 보이는 것만 적고 짐작하지 마세요. 글자가 거의 보이지 않으면 모두 null 로 하세요."
)
READ_SCHEMA = {"type": "object", "required": ["product", "origin", "base_date", "expiry_date", "shelf_months", "fresh_vegetable", "all_dates"],
               "properties": {"all_dates": {"type": "array", "items": {"type": "string"}},"product": {"type": "string", "nullable": True},
                              "origin": {"type": "string", "nullable": True},
                              "base_date": {"type": "string", "nullable": True},
                              "expiry_date": {"type": "string", "nullable": True},
                              "shelf_months": {"type": "integer", "nullable": True},
                              "fresh_vegetable": {"type": "boolean", "nullable": True}}}
# 알려진 포장(webapp/refs/). 사용자가 '이 모양의 사진은 ○○'라고 알려 준 것들이다.
# 읽을 사진이 이 중 하나와 같은 포장이면 품명(과 원산지)을 여기 적힌 대로 쓴다.
# 날짜·글자는 항상 읽을 사진에서 직접 읽는다 (예시 답을 베끼면 안 된다).
# 새 포장은 사진을 refs 에 넣고 한 줄 추가하면 된다. 사진은 EXIF 회전을 적용한 똑바른 방향, 긴 변 900px 이하.
REF_DIR = Path(__file__).resolve().parent / "refs"
REFERENCES = (
    # (사진 파일, 품명, 원산지, 사람이 읽는 설명)
    # 원산지: '나라' = AI 가 못 읽었을 때 그 나라로 채움 / None = AI 가 읽은 대로 / '' = 이름에 원산지를 적지 않음
    #   ('' 은 원료마다 원산지가 달라 한 가지로 정할 수 없는 가공식품용)
    ("chicken_brazil.jpg", "닭고기", "브라질", "주황색 냉동 닭고기 봉지(빨간 둥근 BRASIL 도장)"),
    ("mini_jjinppang.jpg", "미니 찐빵", "", "투명 비닐 봉지에 든 하얀 미니 찐빵(냉동)"),
)
_refs_cache: list | None = None


def load_references() -> list:
    """[(사진 bytes, 품명, 원산지, 설명)]. 사진 파일이 없으면 그 항목은 건너뛴다."""
    global _refs_cache
    if _refs_cache is None:
        out = []
        for name, product, origin, desc in REFERENCES:
            try:
                out.append(((REF_DIR / name).read_bytes(), product, origin, desc))
            except OSError:
                continue
        _refs_cache = out
    return _refs_cache


MATCH_PROMPT = (
    "참고 사진 A, B, ... 는 납품받는 특정 식자재 포장입니다. 맨 마지막 '판단할 사진'이 참고 사진 중 하나와 "
    "같은 제품·같은 포장(같은 봉지, 같은 모양의 내용물)이면 그 참고 사진의 글자를, 어느 것과도 같지 않으면 none 을 고르세요. "
    "사진 각도, 크기, 일부만 찍힌 것, 흐린 것은 상관없이 같은 포장이면 같다고 보세요. "
    "하지만 종류만 비슷하고 포장이 다른 제품이면 none 입니다."
)


def _match_contents(types, data: bytes, mime: str, refs: list) -> list:
    out = []
    for i, (ex_bytes, product, origin, desc) in enumerate(refs):
        out += [f"[참고 사진 {chr(65 + i)}] {desc}", types.Part.from_bytes(data=ex_bytes, mime_type="image/jpeg")]
    return out + ["[판단할 사진]", types.Part.from_bytes(data=data, mime_type=mime), MATCH_PROMPT]


def match_reference(client, model, jpg: bytes, refs: list | None = None) -> int | None:
    """알려진 포장 중 같은 것의 번호(0부터). 없으면 None."""
    refs = load_references() if refs is None else refs
    if not refs:
        return None
    letters = [chr(65 + i) for i in range(len(refs))] + ["none"]
    schema = {"type": "object", "required": ["match"],
              "properties": {"match": {"type": "string", "enum": letters}}}
    ans = _ask_json(client, model, jpg, "image/jpeg", MATCH_PROMPT, schema, "포장",
                    build=lambda types: _match_contents(types, jpg, "image/jpeg", refs))
    m = str((ans or {}).get("match", "none")).strip()
    return ord(m) - 65 if len(m) == 1 and 0 <= ord(m) - 65 < len(refs) else None


_BAD_NAME_CHARS = re.compile(r'[\\/:*?"<>|\x00-\x1f]')


def date_from_name(name: str) -> str | None:
    """파일 이름에서 날짜를 찾아 'YY.MM.DD' 로. 못 찾으면 None.
    KakaoTalk_20261109_135452116 / 20261109 / 2026-11-09 / 26.11.09 / 2026.11.09 을 알아본다."""
    stem = Path(name or "").stem if re.search(r"\.[A-Za-z0-9]{2,5}$", name or "") else (name or "")

    def ok(y, m, d):
        try:
            return datetime(y, m, d).strftime("%y.%m.%d")
        except ValueError:
            return None
    m = re.search(r"(?<!\d)(20\d{2})[-._]?(\d{2})[-._]?(\d{2})(?!\d{3,})", stem)
    if m:
        got = ok(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        if got:
            return got
    m = re.match(r"^(\d{2})\.(\d{1,2})\.(\d{1,2})(?!\d)", stem)
    if m:
        return ok(2000 + int(m.group(1)), int(m.group(2)), int(m.group(3)))
    return None


def parse_iso(v) -> date | None:
    """AI 가 준 'YYYY-MM-DD' (혹은 2026.11.13 / 26.11.13). 없거나 이상하면 None."""
    s = str(v or "").strip()
    m = re.fullmatch(r"(\d{2,4})\s*[-./]\s*(\d{1,2})\s*[-./]\s*(\d{1,2})", s)
    if not m:
        return None
    y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
    if y < 100:
        y += 2000
    if not 2000 <= y <= 2100:
        return None
    try:
        return date(y, mo, d)
    except ValueError:
        return None


def add_months(d: date, n: int) -> date:
    """d 에 n개월을 더한다. 날짜가 없는 달이면 그 달 말일로 (2026-08-31 + 6개월 = 2027-02-28)."""
    t = d.year * 12 + (d.month - 1) + n
    y, m = divmod(t, 12)
    m += 1
    last = (date(y + (m == 12), m % 12 + 1, 1) - timedelta(days=1)).day
    return date(y, m, min(d.day, last))


def pick_date(ans: dict, filename: str = "") -> tuple[str | None, str]:
    """사진에서 읽은 날짜로 이름에 넣을 날짜('YY.MM.DD')와 그 출처 설명을 고른다.
    유통기한 날짜 > 기준일+개월 > 기준일 > 파일 이름의 날짜."""
    base, exp = parse_iso(ans.get("base_date")), parse_iso(ans.get("expiry_date"))
    found = sorted({d for d in (parse_iso(v) for v in (ans.get("all_dates") or []) if isinstance(v, str)) if d})
    try:
        _m = int(ans.get("shelf_months")) if ans.get("shelf_months") is not None else None
    except (TypeError, ValueError):
        _m = None
    if not base and not exp and _m and 1 <= _m <= 120 and len(found) == 1:
        base = found[0]           # '제조일로부터 N개월' 인데 날짜가 하나뿐이면 그 스탬프 날짜가 제조일
    if len(found) >= 2:           # 날짜가 두 개 이상 인쇄돼 있으면 가장 늦은 날짜가 유통기한, 이른 날짜가 제조·포장일
        exp = found[-1]
        base = base if (base and base < exp) else found[0]
    try:
        months = int(ans.get("shelf_months")) if ans.get("shelf_months") is not None else None
    except (TypeError, ValueError):
        months = None
    if months is not None and not 1 <= months <= 120:
        months = None
    f = lambda d: d.strftime("%y.%m.%d")
    if exp and not (base and exp < base):
        return f(exp), f"유통기한 {f(exp)}"
    if base and months:
        out = add_months(base, months)
        return f(out), f"기준일 {f(base)} + {months}개월"
    if base:
        return f(base), f"사진에 적힌 날짜 {f(base)}"
    d = date_from_name(filename)
    return d, (f"파일 이름의 날짜 {d}" if d else "날짜 없음")


def normalize_origin(v) -> str | None:
    """'러시아' -> '러시아산', '국산'/'한국산'/'대한민국' -> '국내산'. 비면 None."""
    v = _BAD_NAME_CHARS.sub("", str(v or "")).strip()
    if not v or v.lower() in ("null", "none", "unknown"):
        return None
    v = re.sub(r"^원산지\s*[:：]?\s*", "", v)
    if v in ("국산", "한국산", "한국", "대한민국", "대한민국산", "국내", "국내산"):
        return "국내산"
    if "산" not in v[-2:] and not v.endswith("산"):
        v += "산"
    return v[:20]


def clean_product(v) -> str | None:
    v = re.sub(r"\s+", " ", _BAD_NAME_CHARS.sub("", str(v or ""))).strip(" .")
    if not v or v.lower() in ("null", "none"):
        return None
    return v[:30]


# 신선 채소·버섯류. 품명에 이 말이 들어 있으면 채소로 본다.
VEG_WORDS = (
    "표고", "대파", "쪽파", "실파", "무우", "총각무", "알타리", "양파", "브로콜리", "양배추", "배추", "가지", "오이",
    "느타리", "새송이", "팽이", "양송이", "만가닥", "버섯", "청경채", "치커리", "열무", "얼갈이", "아욱", "유채",
    "상추", "깻잎", "시금치", "부추", "당근", "감자", "고구마", "호박", "마늘", "생강", "고추", "피망", "파프리카",
    "토마토", "셀러리", "샐러리", "미나리", "쑥갓", "참나물", "취나물", "콩나물", "숙주", "도라지", "우엉",
    "연근", "비트", "케일", "로메인", "양상추", "적채", "콜리플라워", "아스파라거스", "냉이", "달래", "두릅",
    "근대", "래디쉬", "옥수수", "쌈채소", "모듬쌈", "풋마늘", "마늘종", "적근대",
    "꽈리", "깐마늘", "새싹", "무순", "청겨자", "적겨자", "파슬리", "바질",
)
# 한 글자라 다른 말에 들어가기 쉬워서, 띄어쓰기로 나눈 한 덩어리가 정확히 같을 때만 채소로 본다
VEG_EXACT = ("무", "파", "갓", "쑥", "쌈", "중파", "상추", "배추")
# 이런 말이 들어 있으면 가공품이라 채소로 안 본다 (원산지를 마음대로 국내산으로 적지 않는다)
VEG_PROCESSED = ("김치", "통조림", "절임", "장아찌", "피클", "소스", "주스", "즙", "가루", "분말", "말랭이", "말린", "마른",
                 "건조", "냉동", "튀김", "볶음", "무침", "다진", "자반", "잼", "페이스트", "기름", "칩", "스낵", "케첩",
                 "케찹", "고추장", "된장", "간장", "쌈장", "전분", "국수", "라면", "샐러드팩", "샐러드")


def is_vegetable_name(product) -> bool:
    """품명이 신선 채소·버섯류인지. '배추김치', '냉동 브로콜리', '건고추' 같은 가공·냉동·건조품은 False."""
    p = clean_product(product)
    if not p:
        return False
    if p.startswith("건") or any(w in p for w in VEG_PROCESSED):
        return False
    tokens = p.split()
    if any(t in VEG_EXACT for t in tokens):
        return True
    return any(w in p for w in VEG_WORDS)


def decide_origin(product, origin, ai_fresh_vegetable=None) -> tuple[str | None, str]:
    """(이름에 넣을 원산지, 설명). 읽은 원산지가 있으면 그대로, 없는데 신선 채소류면 '국내산'."""
    got = normalize_origin(origin)
    if got:
        return got, ""
    p = clean_product(product)
    if not p:
        return None, ""
    processed = p.startswith("건") or any(w in p for w in VEG_PROCESSED)
    if not processed and (is_vegetable_name(p) or ai_fresh_vegetable is True):
        return "국내산", "원산지가 안 보여 채소류라 국내산으로 표기"
    return None, ""


UNKNOWN_PRODUCT = "품명미확인"      # 품명을 못 읽었을 때 이름 자리에 넣는 말


def compose_name(product, date_str, origin) -> str | None:
    """'참치 통조림 26.11.09 러시아산'. 품명을 읽지 못했으면 None (원래 이름을 그대로 쓰게)."""
    product = clean_product(product)
    if not product:
        return None
    return " ".join(p for p in (product, date_str, normalize_origin(origin)) if p)


def shrink_jpeg(jpg: bytes, side: int) -> bytes:
    """JPEG 를 긴 변 side px 로 줄인다 (이미 작으면 그대로)."""
    with Image.open(io.BytesIO(jpg)) as im:
        if max(im.size) <= side:
            return jpg
        im = im.convert("RGB")
        im.thumbnail((side, side), Image.Resampling.LANCZOS)
        buf = io.BytesIO()
        im.save(buf, "JPEG", quality=85)
        return buf.getvalue()


def read_label(core, client, model, data: bytes, filename: str = "", matcher=None) -> dict:
    """사진 한 장을 읽어 {'product', 'origin', 'date', 'name'} 을 돌려준다. name 은 못 읽으면 None."""
    jpg = io.BytesIO()
    open_upright(data).save(jpg, "JPEG", quality=88)
    jpg = jpg.getvalue()
    # 저장된 포장과 같은지 알아보는 질문은 글자 읽기와 동시에 보낸다 (시간이 두 배로 늘지 않게)
    if matcher is None and hasattr(core, "json_client") and load_references():
        small = shrink_jpeg(jpg, 640)
        matcher = lambda: match_reference(client, model, small)       # noqa: E731
    pool = fut = None
    if matcher is not None:
        pool = ThreadPoolExecutor(max_workers=1)
        fut = pool.submit(matcher)
    if hasattr(core, "json_client"):
        try:
            ans = _ask_json(client, model, jpg, "image/jpeg", READ_PROMPT, READ_SCHEMA, "읽기")
        except Exception:  # noqa: BLE001  한 번 더, 기존 도구의 호출(재시도 포함)로
            ans = core.json_ask(client, jpg, model, schema=READ_SCHEMA, prompt=READ_PROMPT)
    else:                                     # 시험용 가짜 core
        ans = core.json_ask(client, jpg, model, schema=READ_SCHEMA, prompt=READ_PROMPT)
    ans = dict(ans or {})
    ref_note = ""
    if fut is not None:
        try:
            hit = fut.result()
        except Exception:  # noqa: BLE001  알아보기에 실패해도 읽은 내용은 그대로 쓴다
            hit = None
        finally:
            pool.shutdown(wait=False)
        refs = load_references()
        if hit is not None and 0 <= hit < len(refs):
            _, ref_product, ref_origin, ref_desc = refs[hit]
            ans["product"] = ref_product                      # 사용자가 알려 준 품명이 우선
            if ref_origin == "":
                ans["origin"] = None                              # 이 포장은 이름에 원산지를 적지 않는다
            elif ref_origin and not normalize_origin(ans.get("origin")):
                ans["origin"] = ref_origin
            ref_note = f"저장된 포장 '{ref_product}'(으)로 인식했습니다"
    d, note = pick_date(ans, filename)
    origin, origin_note = decide_origin(ans.get("product"), ans.get("origin"), ans.get("fresh_vegetable"))
    name, name_note = compose_name(ans.get("product"), d, origin), ""
    if not name:
        # 품명은 못 읽었어도 사진에서 유통기한·날짜나 원산지를 읽었으면 그것만이라도 이름에 적는다
        # (파일 이름에서 가져온 날짜만 있는 경우는 사진에서 읽은 게 없으니 원래 이름 그대로 둔다)
        from_photo = bool(d) and not note.startswith("파일 이름")
        if from_photo or origin:
            name = " ".join(p for p in (UNKNOWN_PRODUCT, d if from_photo else None, origin) if p)
            name_note = "품명을 못 읽어 읽은 날짜·원산지만 적었습니다. 품명은 이름 칸에서 직접 고쳐 주세요"
    return {"product": clean_product(ans.get("product")), "origin": origin, "origin_note": origin_note,
            "date": d, "date_note": note, "name": name, "name_note": name_note, "ref_note": ref_note,
            "raw": {k: ans.get(k) for k in ("product", "origin", "base_date", "expiry_date", "shelf_months", "all_dates")}}


def photos_zip(uploads, zip_path: Path, max_side: int = 3000) -> tuple[int, list]:
    """[(파일 이름, bytes, 회전 도)] 를 이름 그대로(확장자는 .jpg) 한 zip 으로 묶는다. 돌린 방향도 반영한다.
    같은 이름은 _2, _3 을 붙인다. (넣은 장수, 못 연 파일 목록)"""
    import zipfile
    used, skipped, n = set(), [], 0
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_STORED) as z:     # 사진은 이미 압축돼 있다
        for up in uploads:
            name, data = up[0], up[1]
            rot = up[2] if len(up) > 2 else 0
            base = Path(name or "").name
            if Path(base).suffix.lower() not in INSP_ALLOWED_EXT:
                skipped.append(f"{base} (지원하지 않는 형식)")
                continue
            try:
                im = open_upright(data)
                im.thumbnail((max_side, max_side), Image.Resampling.LANCZOS)
                im = rotate_cw(im, rot)
                buf = io.BytesIO()
                im.save(buf, "JPEG", quality=92)
            except Exception:  # noqa: BLE001
                skipped.append(f"{base} (사진을 열 수 없음)")
                continue
            stem = unique_stem(clean_stem(base), used)
            z.writestr(f"{stem}.jpg", buf.getvalue())
            n += 1
    return n, skipped

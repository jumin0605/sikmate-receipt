"""
영수증 정리 도구 - 사진 정리부터 지출현황 장부까지 한 파일로

1. 정리 - 비틀어진 영수증을 반듯하게 펴고(원근 보정 + 미세 기울기 보정),
          영수증 영역만 남기고 배경(책상, 흰 종이 등)을 제거하고,
          저장 이름의 날짜를 'YY.MM.DD' 로 통일합니다.
          (20260810_3 -> 26.08.10_3, 8.21 충북농산 -> 26.08.21 충북농산)
2. 엑셀 - 정리한 사진을 A4 한 쪽에 2열로 묶어 인쇄용 엑셀로 만듭니다.
3. JSON - 영수증 내용을 읽어 파일마다 JSON 하나를 만듭니다. 합계를 검산하고,
          걸린 장만 더 좋은 모델로 다시 읽습니다.
4. 장부 - 그 JSON 을 모아 식자재 지출현황 엑셀 장부를 만듭니다.

사용법:
    python 영수증_정리_도구.py                     # 인자 없이 실행하면 GUI
                                              # (창에 파일/폴더를 끌어다 놓으면 됨)
    python 영수증_정리_도구.py 영수증.jpg
    python 영수증_정리_도구.py 폴더 -o 결과폴더 --bg transparent
    python 영수증_정리_도구.py 영수증.pdf --to pdf
    python 영수증_정리_도구.py 폴더 --to png          # 무손실로 남기고 싶을 때
    python 영수증_정리_도구.py 폴더 --excel           # 정리한 결과를 엑셀(A4 2열)로도 묶기
    python 영수증_정리_도구.py 정리결과 --excel-only  # 이미 정리해 둔 폴더만 엑셀로 묶기
    python 영수증_정리_도구.py 폴더 --json            # 정리한 뒤 영수증 내용을 JSON 으로 뽑기
    python 영수증_정리_도구.py 정리결과 --json-only   # 이미 정리해 둔 폴더만 JSON 으로 뽑기
    python 영수증_정리_도구.py 폴더 --json --장부     # 정리 -> JSON -> 지출현황 장부까지 한 번에
    python 영수증_정리_도구.py 26.8월분 --장부만      # 이미 뽑아 둔 JSON 으로 장부만 만들기
    python 영수증_정리_도구.py 원산지사진 --검수서 --검수날짜 "2026년 7월 20일"
                                              # 식자재 검수서(원산지 표시) 워드·엑셀

마지막 단계인 '장부'는 JSON 을 모아 식자재 지출현황 엑셀(영수증 내역 + 영수증
이미지 두 시트)을 만듭니다. 월별 폴더 하나만 고르면 하위 폴더까지 다 읽습니다.

JSON 추출은 Gemini 를 씁니다. 싼 모델로 한 번 훑고, 합계 검산에 걸린 장만
좋은 모델로 자동으로 다시 읽습니다. 처음 한 번만:
    pip install google-genai
    setx GEMINI_API_KEY "받은키"      (새 명령창을 열어야 적용됩니다)

필요 패키지: pip install opencv-python numpy pillow pymupdf openpyxl google-genai
"""

import argparse
import csv
import gc
import hashlib
import importlib
import io
import json
import logging
import math
import os
import re
import shutil
import sys
import tempfile
import time
from collections import OrderedDict
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

import cv2
import numpy as np
from PIL import Image, ImageOps

# PDF 렌더링은 pymupdf를 먼저 쓰고, 없으면 pypdfium2로 대체한다
try:
    import pymupdf as fitz
except ImportError:
    fitz = None
try:
    import pypdfium2 as pdfium
except ImportError:
    pdfium = None


IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".bmp", ".webp", ".tif", ".tiff"}
PDF_EXTS = {".pdf"}
# PDF 안의 스캔 원본은 대개 150dpi 안팎이라 300으로 그리면 없는 해상도를 만들어내며
# 파일만 커진다. 200이면 원본보다 커서 손실이 없고 용량과 처리 시간이 크게 준다.
PDF_RENDER_DPI = 200
JPEG_QUALITY = 95        # 글자 가장자리가 뭉개지지 않는 선
# 사진을 읽을 때 긴 변을 이 값으로 줄인다 (0이면 줄이지 않음). 휴대폰 원본(4000px 이상)을 그대로
# 처리하면 사진 한 장에 메모리가 600~800MB까지 올라가므로, 메모리가 작은 서버에서만 켠다.
# 판독에는 긴 변 1600px(재판독 2400px)만 쓰므로 3000px이면 글씨 판독에는 영향이 없다.
LOAD_MAX_SIDE = int(os.environ.get("RECEIPT_MAX_SIDE", "0") or 0)


def cap_side(img):
    """BGR ndarray 의 긴 변이 LOAD_MAX_SIDE 보다 크면 비율을 지켜 줄인다."""
    if LOAD_MAX_SIDE <= 0:
        return img
    h, w = img.shape[:2]
    m = max(h, w)
    if m <= LOAD_MAX_SIDE:
        return img
    r = LOAD_MAX_SIDE / m
    return cv2.resize(img, (max(1, round(w * r)), max(1, round(h * r))), interpolation=cv2.INTER_AREA)


# ---------------------------------------------------------------------------
# 입출력


def load_pages(path):
    """파일을 BGR ndarray 리스트로 읽는다. PDF는 페이지마다 한 장."""
    ext = os.path.splitext(path)[1].lower()
    if ext in PDF_EXTS:
        pages = []
        if fitz is not None:
            with fitz.open(path) as doc:
                for page in doc:
                    pix = page.get_pixmap(dpi=PDF_RENDER_DPI)
                    img = np.frombuffer(pix.samples, dtype=np.uint8)
                    img = img.reshape(pix.height, pix.width, pix.n)
                    if pix.n == 4:
                        img = cv2.cvtColor(img, cv2.COLOR_RGBA2BGR)
                    elif pix.n == 3:
                        img = cv2.cvtColor(img, cv2.COLOR_RGB2BGR)
                    else:
                        img = cv2.cvtColor(img, cv2.COLOR_GRAY2BGR)
                    pages.append(img)
        elif pdfium is not None:
            doc = pdfium.PdfDocument(path)
            try:
                for page in doc:
                    pil = page.render(scale=PDF_RENDER_DPI / 72).to_pil().convert("RGB")
                    pages.append(cv2.cvtColor(np.array(pil), cv2.COLOR_RGB2BGR))
            finally:
                doc.close()
        else:
            raise RuntimeError(
                "PDF를 읽으려면 pymupdf 또는 pypdfium2가 필요합니다: pip install pymupdf")
        return [cap_side(p) for p in pages]

    # EXIF 회전 정보를 먼저 적용한 뒤 OpenCV로 넘긴다
    with Image.open(path) as im:
        if LOAD_MAX_SIDE > 0:
            # 원본 크기로 풀어 놓은 뒤 줄이면 그 순간 메모리가 몇 배로 튄다.
            # JPEG 는 풀 때부터 작게(1/2, 1/4) 읽고, 나머지도 numpy 로 넘기기 전에 줄인다.
            r = LOAD_MAX_SIDE / max(im.size)
            if r < 1:
                im.draft("RGB", (max(1, int(im.size[0] * r)), max(1, int(im.size[1] * r))))
            im = ImageOps.exif_transpose(im).convert("RGB")
            if max(im.size) > LOAD_MAX_SIDE:
                im.thumbnail((LOAD_MAX_SIDE, LOAD_MAX_SIDE), Image.LANCZOS)
        else:
            im = ImageOps.exif_transpose(im).convert("RGB")
        bgr = cv2.cvtColor(np.array(im), cv2.COLOR_RGB2BGR)
        del im
        return [cap_side(bgr)]


def save_pages(pages, out_path, to_pdf=False, jpeg_quality=JPEG_QUALITY):
    """BGR 또는 BGRA ndarray 리스트를 저장한다."""
    pils = []
    for img in pages:
        if img.shape[2] == 4:
            pils.append(Image.fromarray(cv2.cvtColor(img, cv2.COLOR_BGRA2RGBA)))
        else:
            pils.append(Image.fromarray(cv2.cvtColor(img, cv2.COLOR_BGR2RGB)))

    if to_pdf:
        # PDF는 알파를 지원하지 않으므로 흰 배경에 합성
        flat = []
        for im in pils:
            if im.mode == "RGBA":
                bg = Image.new("RGB", im.size, (255, 255, 255))
                bg.paste(im, mask=im.split()[3])
                im = bg
            flat.append(im.convert("RGB"))
        flat[0].save(out_path, save_all=True, append_images=flat[1:],
                     resolution=PDF_RENDER_DPI)
        return [out_path]

    ext = os.path.splitext(out_path)[1].lower()
    written = []
    for i, im in enumerate(pils):
        if len(pils) == 1:
            path = out_path
        else:
            stem, e = os.path.splitext(out_path)
            path = f"{stem}_p{i + 1}{e}"
        if ext in (".jpg", ".jpeg"):
            if im.mode == "RGBA":
                bg = Image.new("RGB", im.size, (255, 255, 255))
                bg.paste(im, mask=im.split()[3])
                im = bg
            # subsampling=0: 색을 절반으로 깎지 않아 글자 가장자리가 번지지 않는다
            im.convert("RGB").save(path, quality=jpeg_quality, subsampling=0)
        else:
            im.save(path)
        written.append(path)
    return written


# ---------------------------------------------------------------------------
# 영수증 영역 찾기


def _order_corners(pts):
    """네 점을 좌상-우상-우하-좌하 순서로 정렬."""
    pts = pts.reshape(4, 2).astype(np.float32)
    s = pts.sum(axis=1)
    d = np.diff(pts, axis=1).ravel()
    return np.array(
        [pts[np.argmin(s)], pts[np.argmin(d)], pts[np.argmax(s)], pts[np.argmax(d)]],
        dtype=np.float32,
    )


def background_color(img, band=0.02):
    """가장자리 띠의 중앙값을 배경색으로 본다(스캔 여백, 책상 바닥 등)."""
    h, w = img.shape[:2]
    b = max(2, int(min(h, w) * band))
    edge = np.concatenate([
        img[:b].reshape(-1, 3), img[-b:].reshape(-1, 3),
        img[:, :b].reshape(-1, 3), img[:, -b:].reshape(-1, 3),
    ])
    return np.median(edge, axis=0)


def _receipt_mask(img, tol=10):
    """배경색과 다른 부분(= 영수증)을 흰색으로 표시한 마스크를 만든다.

    흰 A4에 스캔한 영수증처럼 종이도 배경도 밝은 경우가 있어, 밝기 기준이 아니라
    '가장자리 배경색과 얼마나 다른가'로 판단한다.
    """
    blur = cv2.GaussianBlur(img, (5, 5), 0)
    bg = background_color(blur)
    diff = np.abs(blur.astype(np.int16) - bg.astype(np.int16)).max(axis=2)
    diff = np.clip(diff, 0, 255).astype(np.uint8)

    _, mask = cv2.threshold(diff, tol, 255, cv2.THRESH_BINARY)
    # 글자만 잡히고 종이 바탕은 배경과 비슷할 수 있으므로 크게 뭉쳐 한 덩어리로 만든다
    k = max(3, (min(img.shape[:2]) // 100) | 1)
    mask = cv2.morphologyEx(mask, cv2.MORPH_OPEN, np.ones((3, 3), np.uint8))
    mask = cv2.morphologyEx(mask, cv2.MORPH_CLOSE, np.ones((k, k), np.uint8))
    mask = cv2.dilate(mask, np.ones((k // 2 | 1, k // 2 | 1), np.uint8))
    return mask


def _keep_contours(mask, min_share=0.05, grow=0.03, rounds=3):
    """영수증을 이루는 덩어리들을 모은다.

    큰 덩어리(글자가 몰린 부분)에서 출발해, 그 범위를 조금 넓힌 뒤 거기에 걸치는
    작은 조각들을 계속 끌어들인다. 이렇게 해야 인쇄가 흐린 상단 등록번호 줄이나
    아래쪽 테두리처럼 작게 잡히는 부분이 잘려나가지 않는다.
    """
    contours, _ = cv2.findContours(mask, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    if not contours:
        return []
    areas = [cv2.contourArea(c) for c in contours]
    biggest = max(areas)
    if biggest <= 0:
        return []

    h, w = mask.shape[:2]
    dust = max(9.0, biggest * 0.0002)  # 스캐너 먼지, 점 노이즈는 제외
    keep = [c for c, a in zip(contours, areas) if a >= biggest * min_share]
    rest = [c for c, a in zip(contours, areas) if biggest * min_share > a >= dust]

    margin = int(min(h, w) * grow)
    for _ in range(rounds):
        x, y, bw, bh = cv2.boundingRect(np.vstack(keep))
        box = (x - margin, y - margin, x + bw + margin, y + bh + margin)
        near, far = [], []
        for c in rest:
            cx, cy, cw, ch = cv2.boundingRect(c)
            overlaps = (cx <= box[2] and cx + cw >= box[0]
                        and cy <= box[3] and cy + ch >= box[1])
            (near if overlaps else far).append(c)
        if not near:
            break
        keep += near
        rest = far
    return keep


def _content_hull(mask, min_share=0.05):
    """영수증 덩어리들을 모두 감싸는 볼록 외곽선."""
    keep = _keep_contours(mask, min_share)
    if not keep:
        return None
    return cv2.convexHull(np.vstack(keep))


def ink_mask(img, tol=12):
    """인쇄·필기·종이 테두리처럼 '주변보다 진한' 픽셀만 표시한 마스크.

    배경색과의 단순 비교는 스캔본에 흔한 옅은 회색 그림자나 종이 색 차이까지 잉크로
    보기 때문에 페이지 전체가 잡힌다. 주변 평균과의 차이를 보면 완만한 얼룩은 사라지고
    글자와 선만 남는다.
    """
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    sigma = max(3.0, min(img.shape[:2]) / 100.0)
    local = cv2.GaussianBlur(gray, (0, 0), sigma)
    detail = cv2.absdiff(gray, local)
    _, mask = cv2.threshold(detail, tol, 255, cv2.THRESH_BINARY)
    return cv2.morphologyEx(mask, cv2.MORPH_OPEN, np.ones((3, 3), np.uint8))


def document_map(img, target=240, density=8):
    """이미지를 크게 축소해 '문서가 놓인 자리'를 덩어리로 만든다.

    원본 해상도에서는 글자마다 조각이 나서 문서 단위로 묶이지 않는다. 잉크 마스크를
    엄지손톱 크기로 줄이면 글자들이 뭉쳐 한 장의 종이가 하나의 덩어리가 되고,
    종이와 종이 사이 빈틈은 그대로 남아 서로 다른 문서를 구분할 수 있다.
    """
    h, w = img.shape[:2]
    scale = target / max(h, w)
    small = cv2.resize(ink_mask(img), (max(8, int(w * scale)), max(8, int(h * scale))),
                       interpolation=cv2.INTER_AREA)
    small = cv2.GaussianBlur(small, (5, 5), 0)
    _, blobs = cv2.threshold(small, density, 255, cv2.THRESH_BINARY)
    blobs = cv2.morphologyEx(blobs, cv2.MORPH_CLOSE, np.ones((5, 5), np.uint8))
    blobs = cv2.morphologyEx(blobs, cv2.MORPH_OPEN, np.ones((3, 3), np.uint8))
    return blobs, scale


def _split_once(blobs, cut):
    """가장 큰 문서 한 장만 남긴다.

    깎아내(erode) 가느다란 연결을 끊고 가장 큰 속살을 고른 뒤, 그 속살을 감싸는
    직사각형만큼만 되살린다. 종이는 직사각형이므로, 명세표 모서리에 겹쳐 놓인 약관
    조각처럼 사각형 밖으로 삐져나온 부분은 이 단계에서 떨어져 나간다.
    """
    core = cv2.erode(blobs, np.ones((cut, cut), np.uint8))
    num, labels, stats, _ = cv2.connectedComponentsWithStats(core, 8)
    if num <= 1:
        return None  # 너무 얇아 다 깎였다
    main = 1 + int(np.argmax(stats[1:, cv2.CC_STAT_AREA]))

    seed = (labels == main).astype(np.uint8) * 255
    back = cut // 2 + 2  # 깎여 나간 두께(커널 절반)만큼만 되돌린다

    # 되살리는 범위를 두 가지로 제한한다: 속살에서 back 픽셀 이내이면서,
    # 속살을 감싸는 직사각형(= 종이 모양) 안쪽일 것
    grown = cv2.dilate(seed, np.ones((back * 2 + 1, back * 2 + 1), np.uint8))
    (cx, cy), (rw, rh), angle = cv2.minAreaRect(cv2.findNonZero(seed))
    box = cv2.boxPoints(((cx, cy), (rw + back * 2, rh + back * 2), angle))
    area = np.zeros_like(blobs)
    cv2.fillPoly(area, [np.int32(box)], 255)
    return cv2.bitwise_and(blobs, cv2.bitwise_and(grown, area))


def _rect_fill(mask):
    """영역이 얼마나 직사각형에 가까운지(0~1). 문서 한 장이면 1에 가깝다."""
    pts = cv2.findNonZero(mask)
    if pts is None:
        return 0.0
    (_, _), (w, h), _ = cv2.minAreaRect(pts)
    if w < 2 or h < 2:
        return 0.0
    return float(cv2.countNonZero(mask)) / (w * h)


def select_document(blobs, cuts=(5, 9, 13, 17, 21), gain=0.04):
    """여러 장이 겹쳐 잡혔을 때 '문서 한 장'에 해당하는 영역만 고른다.

    문서는 직사각형이므로, 떼어내는 강도를 올려 가며 가장 직사각형에 가까운 결과를
    고른다. 강도를 올려도 모양이 뚜렷이 나아지지 않으면(gain 미만) 더 많이 남는 쪽,
    즉 약한 강도의 결과를 쓴다. 한 장짜리 사진을 괜히 잘라내지 않기 위해서다.
    """
    best, best_fill, base_area = None, -1.0, 0
    for cut in cuts:
        sel = _split_once(blobs, cut)
        if sel is None:
            break
        area = cv2.countNonZero(sel)
        if base_area == 0:
            base_area = area
        elif area < base_area * 0.6:
            break  # 나란히 놓인 다른 한 장까지 잘라내는 상황이므로 멈춘다
        fill = _rect_fill(sel)
        if fill > best_fill + gain:
            best, best_fill = sel, fill
    return best if best is not None else blobs


def _absorb_aligned(blobs, chosen, span=0.5, reach=0.35):
    """고른 문서와 '변이 나란한' 덩어리를 같은 종이로 보고 합친다.

    거래명세표는 가운데 표가 비어 있어 위아래로 끊겨 잡히곤 한다. 이어지는 부분은
    문서와 폭(또는 높이)을 거의 같이 쓰며 가까이 붙어 있고, 옆에 비스듬히 놓인 다른
    종이는 그렇지 않다. 그 차이로 가른다.
    """
    num, labels, stats, _ = cv2.connectedComponentsWithStats(blobs, 8)
    if num <= 1:
        return chosen

    out = chosen.copy()
    changed = True
    while changed:
        changed = False
        pts = cv2.findNonZero(out)
        dx, dy, dw, dh = cv2.boundingRect(pts)
        for i in range(1, num):
            x, y, w, h = stats[i, :4]
            part = (labels == i)
            if out[part].any():
                continue  # 이미 들어와 있다
            x_overlap = max(0, min(x + w, dx + dw) - max(x, dx)) / max(1, dw)
            y_overlap = max(0, min(y + h, dy + dh) - max(y, dy)) / max(1, dh)
            gap_y = max(dy - (y + h), y - (dy + dh), 0)
            gap_x = max(dx - (x + w), x - (dx + dw), 0)
            if ((x_overlap >= span and gap_y <= dh * reach)
                    or (y_overlap >= span and gap_x <= dw * reach)):
                out[part] = 255
                changed = True
    return out


def main_document_mask(img):
    """한 장에 여러 장이 찍혀 있어도 '문서 한 장'만 남긴 마스크를 만든다.

    먼저 조각들을 떼어내 주 문서를 고르고(select_document), 변이 나란히 이어지는
    부분을 같은 종이로 보고 합친다(_absorb_aligned). 비스듬히 어긋나 놓인 약관 조각은
    변이 맞지 않아 들어오지 않는다.
    """
    blobs, _ = document_map(img)
    chosen = select_document(blobs)
    small = _absorb_aligned(blobs, chosen)
    if cv2.countNonZero(small) == 0:
        return None

    small = cv2.dilate(small, np.ones((3, 3), np.uint8))  # 축소하며 깎인 만큼 되돌린다
    return cv2.resize(small, (img.shape[1], img.shape[0]),
                      interpolation=cv2.INTER_NEAREST)


def document_angle(mask, limit=20.0, min_fill=0.55):
    """문서가 얼마나 기울었는지 각도를 구한다. 못 믿을 때는 None.

    문서는 직사각형이므로 외곽선의 최소 회전사각형 각도가 곧 기울기다. 외곽선이
    사각형과 많이 다르면(여러 장이 겹쳤거나 검출이 어긋난 경우) 회전하지 않는다.
    """
    pts = cv2.findNonZero(mask)
    if pts is None or len(pts) < 10:
        return None
    hull = cv2.convexHull(pts)
    (_, _), (rw, rh), angle = cv2.minAreaRect(hull)
    if rw < 10 or rh < 10:
        return None
    if cv2.contourArea(hull) / (rw * rh) < min_fill:
        return None

    angle = (angle + 45) % 90 - 45  # -45 ~ 45 로 정규화
    return angle if abs(angle) <= limit else None


def extend_box(box, fine, cover=0.25, gap_ratio=0.22):
    """문서 상자를 잉크가 이어지는 만큼 바깥으로 넓힌다.

    거래명세표는 가운데 표가 비어 있어 아래쪽 합계란이 따로 떨어져 보인다. 상자 밖으로
    한 줄씩 훑어 나가면서 '상자 폭의 4분의 1 이상을 채우는 줄'이 나오면 같은 종이로 보고
    상자를 늘린다. 옆에 놓인 다른 영수증 조각은 폭을 그만큼 채우지 못해 들어오지 않는다.
    빈 줄이 일정 길이 이상 이어지면 종이가 끝난 것으로 본다.
    """
    h, w = fine.shape[:2]
    x0, y0, x1, y1 = box
    gap_v, gap_h = int(h * gap_ratio), int(w * gap_ratio)

    def covered(line, need):
        return cv2.countNonZero(line) >= need

    # 위/아래로 넓히기
    for _ in range(4):
        need = max(3, int((x1 - x0) * cover))
        blank = 0
        while y0 > 0 and blank <= gap_v:
            row = fine[y0 - 1:y0, x0:x1]
            y0 -= 1
            blank = 0 if covered(row, need) else blank + 1
        y0 = min(y0 + blank, y1 - 1)
        blank = 0
        while y1 < h and blank <= gap_v:
            row = fine[y1:y1 + 1, x0:x1]
            y1 += 1
            blank = 0 if covered(row, need) else blank + 1
        y1 = max(y1 - blank, y0 + 1)

        # 좌/우로 넓히기
        need = max(3, int((y1 - y0) * cover))
        blank = 0
        while x0 > 0 and blank <= gap_h:
            col = fine[y0:y1, x0 - 1:x0]
            x0 -= 1
            blank = 0 if covered(col, need) else blank + 1
        x0 = min(x0 + blank, x1 - 1)
        blank = 0
        while x1 < w and blank <= gap_h:
            col = fine[y0:y1, x1:x1 + 1]
            x1 += 1
            blank = 0 if covered(col, need) else blank + 1
        x1 = max(x1 - blank, x0 + 1)
    return x0, y0, x1, y1


def grow_to_ink(box, fine, need_ratio=0.005, gap_px=3, limit_ratio=0.03):
    """상자 경계가 글자 한가운데를 지나가지 않도록, 잉크가 이어지는 만큼만 넓힌다.

    앞 단계(extend_box)는 '상자 폭의 4분의 1을 채우는 줄'만 같은 종이로 인정하므로,
    영수증 왼쪽 끝의 좁은 글자 열처럼 옅은 부분이 상자 밖에 남아 잘릴 수 있다. 여기서는
    문턱을 아주 낮춰, 바로 옆에 잉크가 붙어 있는 동안만 조금씩 넓힌다. 몇 픽셀만 비어도
    멈추고 넓히는 폭도 이미지의 3%로 제한하므로, 떨어져 있는 다른 조각은 끌어오지 않는다.
    """
    h, w = fine.shape[:2]
    x0, y0, x1, y1 = box
    limit = int(min(h, w) * limit_ratio)
    gap = gap_px

    def line_has_ink(sl, need):
        return cv2.countNonZero(sl) >= need

    need_col = max(3, int((y1 - y0) * need_ratio))
    need_row = max(3, int((x1 - x0) * need_ratio))

    # 경계선이 실제로 글자를 지나가고 있을 때만 그 쪽으로 넓힌다. 이미 흰 여백에 놓인
    # 경계는 건드리지 않으므로, 옆에 놓인 다른 조각 쪽으로 번지지 않는다.
    if line_has_ink(fine[y0:y1, x0:x0 + 1], need_col):
        moved, blank = 0, 0
        while x0 > 0 and moved < limit and blank <= gap:
            blank = 0 if line_has_ink(fine[y0:y1, x0 - 1:x0], need_col) else blank + 1
            x0 -= 1
            moved += 1
        x0 += blank

    if line_has_ink(fine[y0:y1, x1 - 1:x1], need_col):
        moved, blank = 0, 0
        while x1 < w and moved < limit and blank <= gap:
            blank = 0 if line_has_ink(fine[y0:y1, x1:x1 + 1], need_col) else blank + 1
            x1 += 1
            moved += 1
        x1 -= blank

    if line_has_ink(fine[y0:y0 + 1, x0:x1], need_row):
        moved, blank = 0, 0
        while y0 > 0 and moved < limit and blank <= gap:
            blank = 0 if line_has_ink(fine[y0 - 1:y0, x0:x1], need_row) else blank + 1
            y0 -= 1
            moved += 1
        y0 += blank

    if line_has_ink(fine[y1 - 1:y1, x0:x1], need_row):
        moved, blank = 0, 0
        while y1 < h and moved < limit and blank <= gap:
            blank = 0 if line_has_ink(fine[y1:y1 + 1, x0:x1], need_row) else blank + 1
            y1 += 1
            moved += 1
        y1 -= blank
    return x0, y0, x1, y1


def tighten_box(box, fine, cover=0.015, pad_ratio=0.008, min_pad=10):
    """상자 안쪽에서 실제 내용이 시작·끝나는 자리까지 여백을 줄인다."""
    h, w = fine.shape[:2]
    x0, y0, x1, y1 = box
    rows = [cv2.countNonZero(fine[y:y + 1, x0:x1]) for y in range(y0, y1)]
    cols = [cv2.countNonZero(fine[y0:y1, x:x + 1]) for x in range(x0, x1)]
    row_need = max(2, int((x1 - x0) * cover))
    col_need = max(2, int((y1 - y0) * cover))
    ink_rows = [i for i, v in enumerate(rows) if v >= row_need]
    ink_cols = [i for i, v in enumerate(cols) if v >= col_need]
    if not ink_rows or not ink_cols:
        return box

    pad = max(min_pad, int(min(h, w) * pad_ratio))
    ny0 = max(0, y0 + ink_rows[0] - pad)
    ny1 = min(h, y0 + ink_rows[-1] + pad)
    nx0 = max(0, x0 + ink_cols[0] - pad)
    nx1 = min(w, x0 + ink_cols[-1] + pad)
    if nx1 - nx0 < 20 or ny1 - ny0 < 20:
        return box
    return nx0, ny0, nx1, ny1


def mask_box(mask, shape, pad_ratio=0.02, min_pad=12):
    """마스크가 차지하는 범위를 여유를 두고 잘라낼 사각형으로 돌려준다."""
    pts = cv2.findNonZero(mask)
    if pts is None:
        return None
    x, y, w, h = cv2.boundingRect(pts)
    ih, iw = shape[:2]
    pad = max(min_pad, int(min(ih, iw) * pad_ratio))
    x0, y0 = max(0, x - pad), max(0, y - pad)
    x1, y1 = min(iw, x + w + pad), min(ih, y + h + pad)
    if x1 - x0 < 20 or y1 - y0 < 20:
        return None
    return x0, y0, x1, y1


def crop_to_content(img, pad_ratio=0.03, min_pad=16):
    """배경 여백을 잘라내고 영수증 부분만 남긴다."""
    h, w = img.shape[:2]
    keep = _keep_contours(_receipt_mask(img))
    if not keep:
        return img
    box = np.vstack(keep)
    if cv2.contourArea(cv2.convexHull(box)) < h * w * 0.0005:
        return img  # 먼지 수준이면 건드리지 않는다

    x, y, bw, bh = cv2.boundingRect(box)
    # 글자 끝에 딱 붙여 자르면 등록번호 줄이나 아래 테두리가 날아가므로 여유를 준다
    pad = max(min_pad, int(min(h, w) * pad_ratio))
    x0, y0 = max(0, x - pad), max(0, y - pad)
    x1, y1 = min(w, x + bw + pad), min(h, y + bh + pad)
    if x1 - x0 < 20 or y1 - y0 < 20:
        return img
    return img[y0:y1, x0:x1]


def find_receipt_quad(img, min_area_ratio=0.01):
    """영수증의 네 꼭짓점을 찾는다. 보정이 필요 없거나 실패하면 None."""
    scale = 1000.0 / max(img.shape[:2])
    small = cv2.resize(img, None, fx=scale, fy=scale) if scale < 1 else img.copy()
    inv = 1.0 / scale if scale < 1 else 1.0

    page_area = small.shape[0] * small.shape[1]
    hull = _content_hull(_receipt_mask(small))
    if hull is None:
        return None
    area = cv2.contourArea(hull)
    if area < page_area * min_area_ratio:
        return None

    rect = cv2.minAreaRect(hull)
    (_, _), (rw, rh), _ = rect
    if rw < 10 or rh < 10:
        return None
    # 외곽선이 사각형과 많이 다르면(찢긴 영수증, 글자만 잡힌 경우) 억지로 펴지 않는다
    if area / (rw * rh) < 0.75:
        return None

    peri = cv2.arcLength(hull, True)
    approx = cv2.approxPolyDP(hull, 0.02 * peri, True)
    if len(approx) == 4 and cv2.isContourConvex(approx):
        quad = _order_corners(approx)
    else:
        # 사각형으로 안 잡히면 최소 외접 회전사각형으로 대체
        quad = _order_corners(cv2.boxPoints(rect))

    return quad * inv


def expand_quad(quad, shape, ratio=0.04, margin_ratio=0.02):
    """네 꼭짓점을 중심에서 바깥으로 밀어 잘릴 여유를 준다.

    꼭짓점은 잉크가 있는 곳을 기준으로 잡히기 때문에, 그대로 자르면 위쪽 날짜·등록번호
    줄이나 아래 테두리처럼 인쇄가 흐린 부분이 함께 날아간다. 비율만으로는 큰 서류에서
    여유가 모자라서, 이미지 크기에 비례한 고정 여백도 함께 준다.
    """
    center = quad.mean(axis=0)
    out = center + (quad - center) * (1.0 + ratio)

    margin = max(shape[:2]) * margin_ratio
    direction = out - center
    length = np.linalg.norm(direction, axis=1, keepdims=True)
    length[length == 0] = 1.0
    return (out + direction / length * margin).astype(np.float32)


def warp_quad(img, quad):
    """네 꼭짓점을 직사각형으로 펴서 잘라낸다."""
    tl, tr, br, bl = quad
    width = int(max(np.linalg.norm(tr - tl), np.linalg.norm(br - bl)))
    height = int(max(np.linalg.norm(bl - tl), np.linalg.norm(br - tr)))
    if width < 10 or height < 10:
        return img
    dst = np.array(
        [[0, 0], [width - 1, 0], [width - 1, height - 1], [0, height - 1]],
        dtype=np.float32,
    )
    m = cv2.getPerspectiveTransform(quad, dst)
    return cv2.warpPerspective(img, m, (width, height), flags=cv2.INTER_CUBIC,
                               borderMode=cv2.BORDER_REPLICATE)


# ---------------------------------------------------------------------------
# 미세 기울기 보정 / 세로 방향 보정


def estimate_text_angle(img, limit=15.0):
    """글자 줄의 기울기를 추정한다(±limit도 이내)."""
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    gray = cv2.GaussianBlur(gray, (3, 3), 0)
    bw = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                               cv2.THRESH_BINARY_INV, 31, 15)
    # 글자를 가로로 뭉쳐 한 줄짜리 덩어리로 만든다
    kw = max(9, img.shape[1] // 40)
    bw = cv2.morphologyEx(bw, cv2.MORPH_CLOSE,
                          cv2.getStructuringElement(cv2.MORPH_RECT, (kw, 3)))

    contours, _ = cv2.findContours(bw, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    angles, weights = [], []
    for c in contours:
        (_, _), (w, h), a = cv2.minAreaRect(c)
        if w < h:
            w, h = h, w
            a += 90
        if w < img.shape[1] * 0.15 or h == 0 or w / h < 4:
            continue
        a = (a + 45) % 90 - 45  # -45 ~ 45 로 정규화
        if abs(a) <= limit:
            angles.append(a)
            weights.append(w)
    if len(angles) < 3:
        return 0.0
    return float(np.average(angles, weights=weights))


def rotate_keep(img, angle, border):
    """이미지를 회전하되 잘리지 않게 캔버스를 넓힌다."""
    if abs(angle) < 0.1:
        return img
    h, w = img.shape[:2]
    m = cv2.getRotationMatrix2D((w / 2, h / 2), angle, 1.0)
    cos, sin = abs(m[0, 0]), abs(m[0, 1])
    nw, nh = int(h * sin + w * cos), int(h * cos + w * sin)
    m[0, 2] += nw / 2 - w / 2
    m[1, 2] += nh / 2 - h / 2
    return cv2.warpAffine(img, m, (nw, nh), flags=cv2.INTER_CUBIC,
                          borderMode=cv2.BORDER_CONSTANT, borderValue=border)


def upright(img):
    """영수증은 세로로 긴 것이 보통이므로 가로로 누웠으면 세운다."""
    h, w = img.shape[:2]
    if w > h * 1.3:
        return cv2.rotate(img, cv2.ROTATE_90_CLOCKWISE)
    return img


# ---------------------------------------------------------------------------
# 배경 제거


def cut_background(img, mode="white", grow_ratio=0.01, mask=None):
    """영수증 영역 밖을 흰색으로 칠하거나 투명하게 만든다.

    mode: "white" | "transparent" | "none"
    """
    if mode == "none":
        return img

    if mask is None:
        mask = main_document_mask(img)
    if mask is None:
        return img
    area = img.shape[0] * img.shape[1]
    pts = cv2.findNonZero(mask)
    hull = cv2.convexHull(pts) if pts is not None else None
    solid = np.full(mask.shape, 255, np.uint8)
    if hull is not None and cv2.contourArea(hull) > area * 0.01:
        solid = np.zeros(mask.shape, np.uint8)
        cv2.drawContours(solid, [hull], -1, 255, cv2.FILLED)
        # 외곽선은 잉크를 기준으로 잡히므로, 종이 가장자리가 살도록 조금 넓힌다
        grow = max(6, int(min(mask.shape[:2]) * grow_ratio)) | 1
        solid = cv2.dilate(solid, np.ones((grow, grow), np.uint8))

    # 잘린 경계가 톱니처럼 보이지 않도록 부드럽게
    alpha = cv2.GaussianBlur(solid, (5, 5), 0)

    if mode == "transparent":
        out = cv2.cvtColor(img, cv2.COLOR_BGR2BGRA)
        out[:, :, 3] = alpha
        return out

    a = (alpha.astype(np.float32) / 255.0)[:, :, None]
    return (img.astype(np.float32) * a + 255.0 * (1 - a)).astype(np.uint8)


def whiten(img, strength=1.0):
    """종이의 그림자/누런 기를 걷어내 글자를 또렷하게 만든다."""
    has_alpha = img.shape[2] == 4
    alpha = img[:, :, 3] if has_alpha else None
    bgr = img[:, :, :3]

    # 조명 얼룩(=밝기의 큰 흐름)을 축소본에서 구한다. medianBlur는 큰 커널을 못 받는다.
    h, w = bgr.shape[:2]
    sw, sh = max(16, w // 8), max(16, h // 8)
    small = cv2.resize(bgr, (sw, sh), interpolation=cv2.INTER_AREA)
    small = cv2.medianBlur(small, 15)
    bg = cv2.resize(small, (w, h), interpolation=cv2.INTER_LINEAR)

    # 채널별로 나눠서 보정한다. 흑백으로 바꿔 처리하면 파란 서식지, 빨간 도장 같은
    # 옅은 색 인쇄가 통째로 하얗게 날아간다.
    bg = np.maximum(bg.astype(np.float32), 1.0)
    flat = np.clip(bgr.astype(np.float32) / bg * 245.0, 0, 255).astype(np.uint8)

    out = cv2.addWeighted(flat, strength, bgr, 1 - strength, 0)
    if has_alpha:
        out = np.dstack([out, alpha])
    return out


# ---------------------------------------------------------------------------
# 파이프라인


def process_image(img, bg="white", deskew=True, clean=True, rotate_upright=False,
                  perspective=False):
    """영수증 한 장을 정리한다.

    기본 동작은 '기울기만 바로잡고 문서 밖을 지우는' 것이다. 원근 보정과 세로 세우기는
    스캔본을 오히려 일그러뜨리거나 옆으로 눕혀 놓는 일이 있어 옵션으로 뺐다.
    """
    if perspective:
        quad = find_receipt_quad(img)
        if quad is not None:
            img = warp_quad(img, expand_quad(quad, img.shape))

    mask = main_document_mask(img)
    if mask is None:
        return whiten(img) if clean else img

    if deskew:
        angle = document_angle(mask)
        if angle is None:
            angle = estimate_text_angle(img)  # 표 선이 없으면 글자 줄로 판단
        if angle is not None and 0.3 <= abs(angle) <= 20:
            img = rotate_keep(img, angle, border=(255, 255, 255))
            mask = main_document_mask(img)
            if mask is None:
                return whiten(img) if clean else img

    box = mask_box(mask, img.shape)
    if box is not None:
        # 흐린 인쇄 때문에 덩어리에서 빠진 아래쪽 합계란 같은 부분을 되찾고,
        # 남은 빈 여백은 다시 바짝 줄인다
        fine = ink_mask(img)
        # 넓히고(끊긴 아래쪽 되찾기) → 줄이고(빈 여백 제거) → 잉크에 걸치면 다시 조금 넓힌다
        x0, y0, x1, y1 = grow_to_ink(tighten_box(extend_box(box, fine), fine), fine)
        img = img[y0:y1, x0:x1]
        mask = mask[y0:y1, x0:x1]

    if rotate_upright:
        img = upright(img)
        mask = None

    if mask is not None:
        # 배경을 지울 때는 흐려서 덩어리에 못 든 부분(합계란 등)도 문서로 쳐야 한다
        mask = cv2.bitwise_or(mask, ink_mask(img))
    img = cut_background(img, bg, mask=mask)

    if clean:
        img = whiten(img)
    return img


def normalize_stem(stem, default_year=None):
    """파일 이름의 날짜 부분을 'YY.MM.DD' 로 통일한다.

        20260810_3        -> 26.08.10_3
        2026081704        -> 26.08.17_04
        26.8.13 킹마트     -> 26.08.13 킹마트
        8.21 충북농산      -> 26.08.21 충북농산   (연도는 default_year에서 채움)

    날짜로 읽히지 않으면 원래 이름을 그대로 돌려준다.
    """
    s = stem.strip()

    # 20260810 / 2026081704 처럼 붙여 쓴 형태
    m = re.match(r"^(\d{4})(\d{2})(\d{2})(.*)$", s)
    if m and 1 <= int(m.group(2)) <= 12 and 1 <= int(m.group(3)) <= 31:
        year, month, day, rest = m.group(1)[2:], m.group(2), m.group(3), m.group(4)
    else:
        # 26.8.13 / 8.21 / 8-21 처럼 구분자가 있는 형태
        m = re.match(r"^(?:(\d{2}|\d{4})[.\-_/])?(\d{1,2})[.\-_/](\d{1,2})(.*)$", s)
        if not m:
            return stem
        month, day = int(m.group(2)), int(m.group(3))
        if not (1 <= month <= 12 and 1 <= day <= 31):
            return stem
        year = m.group(1)
        if year is None:
            if default_year is None:
                return stem
            year = default_year
        year = year[2:] if len(year) == 4 else year
        month, day, rest = f"{month:02d}", f"{day:02d}", m.group(4)

    rest = rest.strip()
    if not rest:
        return f"{year}.{month}.{day}"
    # 숫자만 남으면 붙어 있던 일련번호이므로 밑줄로 잇는다
    if rest.isdigit():
        return f"{year}.{month}.{day}_{rest}"
    if rest[0] in "_-":
        return f"{year}.{month}.{day}{rest}"
    return f"{year}.{month}.{day} {rest.lstrip('.').strip()}"


def guess_year(paths, folder):
    """파일 이름이나 폴더 이름에서 가장 많이 보이는 두 자리 연도를 찾는다."""
    counts = {}
    for name in [os.path.basename(folder)] + [os.path.basename(p) for p in paths]:
        for y in re.findall(r"\b(?:20)?(\d{2})[.\-_/]?\d{1,2}[.\-_/]\d{1,2}", name):
            counts[y] = counts.get(y, 0) + 1
    if counts:
        return max(counts, key=counts.get)
    return None


def output_path_for(src, out_dir, fmt, default_year=None, taken=None):
    stem = normalize_stem(os.path.splitext(os.path.basename(src))[0], default_year)
    ext = {"pdf": ".pdf", "png": ".png"}.get(fmt, ".jpg")
    name = f"{stem}_정리{ext}"
    # 이름이 겹치면 뒤에 번호를 붙여 덮어쓰지 않게 한다
    if taken is not None:
        i = 2
        while name.lower() in taken:
            name = f"{stem}_정리_{i}{ext}"
            i += 1
        taken.add(name.lower())
    return os.path.join(out_dir, name)


# 하위 폴더를 훑을 때 건너뛸 폴더들. 결과 폴더는 원본 폴더 '옆'에 생기므로,
# 상위 폴더를 통째로 끌어다 놓으면 지난번 결과가 원본인 척 딸려 들어온다.
SKIP_DIR_NAMES = {"json", "정리결과", "결과", "출력", "output", "result", "results"}


def skip_dir(name: str) -> bool:
    low = name.lower()
    return (name.startswith(".") or low in SKIP_DIR_NAMES
            or name.endswith("_정리결과") or low.startswith("~$"))


def collect_files(target, recursive: bool = False):
    """대상 폴더의 사진과 PDF 목록. recursive 면 하위 폴더까지 훑는다."""
    if not os.path.exists(target):
        raise FileNotFoundError(f"경로를 찾을 수 없습니다: {target}")
    if os.path.isfile(target):
        return [target]

    def wanted(name):
        return os.path.splitext(name)[1].lower() in IMAGE_EXTS | PDF_EXTS

    files = []
    if recursive:
        for folder, dirs, names in os.walk(target):
            dirs[:] = sorted(d for d in dirs if not skip_dir(d))
            for name in sorted(names):
                path = os.path.join(folder, name)
                if wanted(name) and os.path.isfile(path):
                    files.append(path)
        return files

    for name in sorted(os.listdir(target)):
        path = os.path.join(target, name)
        if os.path.isfile(path) and wanted(name):
            files.append(path)
    return files


def rename_sources(target, dry_run=False, log=print, recursive: bool = False):
    """원본 파일 이름 자체를 'YY.MM.DD ...' 규칙으로 바꾼다.

    바꾼 내역은 폴더 안 '이름변경기록.csv' 에 남겨서 되돌릴 수 있게 한다.
    """
    files = collect_files(target, recursive)
    folder = target if os.path.isdir(target) else os.path.dirname(os.path.abspath(target))
    year = guess_year(files, folder)
    taken = {os.path.basename(p).lower() for p in files}

    plan = []
    for src in files:
        stem, ext = os.path.splitext(os.path.basename(src))
        new_stem = normalize_stem(stem, year)
        if new_stem == stem:
            continue
        name = new_stem + ext
        i = 2
        while name.lower() in taken:  # 이미 있는 이름이면 번호를 붙인다
            name = f"{new_stem}_{i}{ext}"
            i += 1
        taken.add(name.lower())
        taken.discard(os.path.basename(src).lower())
        plan.append((src, os.path.join(os.path.dirname(src), name)))

    if not plan:
        log("이름을 바꿀 파일이 없습니다.")
        return []

    done = []
    for src, dst in plan:
        if dry_run:
            log(f"[예정] {os.path.basename(src)} -> {os.path.basename(dst)}")
            continue
        try:
            os.rename(src, dst)
            done.append((src, dst))
            log(f"[변경] {os.path.basename(src)} -> {os.path.basename(dst)}")
        except OSError as e:
            log(f"[실패] {os.path.basename(src)}: {e}")

    if done:
        record = os.path.join(folder, "이름변경기록.csv")
        new_file = not os.path.exists(record)
        with open(record, "a", encoding="utf-8-sig", newline="") as f:
            w = csv.writer(f)
            if new_file:
                w.writerow(["시각", "이전 이름", "새 이름"])
            stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            for src, dst in done:
                w.writerow([stamp, os.path.basename(src), os.path.basename(dst)])
        log(f"기록: {record}")
    return done


def run(target, out_dir=None, bg="white", fmt="jpg", deskew=True, clean=True,
        rotate_upright=False, perspective=False, dpi=None, log=print,
        recursive: bool = False):
    global PDF_RENDER_DPI
    if dpi:
        PDF_RENDER_DPI = dpi
    if bg == "transparent" and fmt == "jpg":
        fmt = "png"            # JPEG 는 투명을 담지 못한다
        log("투명 배경이라 PNG 로 저장합니다.")

    files = collect_files(target, recursive)
    if not files:
        log("처리할 파일이 없습니다.")
        return []
    if recursive and os.path.isdir(target):
        subs = {os.path.dirname(f) for f in files} - {os.path.abspath(target), target}
        if subs:
            log(f"하위 폴더 {len(subs)}개까지 훑어 {len(files)}개를 찾았습니다.")

    base = target if os.path.isdir(target) else os.path.dirname(os.path.abspath(target))
    # 결과는 원본 폴더 '안'이 아니라 '옆'에 둔다. 원본 폴더를 다시 훑어도 결과가
    # 딸려 들어오지 않고, 폴더 이름만 봐도 어느 원본의 결과인지 알 수 있다.
    base = os.path.abspath(base)
    out_dir = out_dir or os.path.join(os.path.dirname(base),
                                      f"{os.path.basename(base)}_정리결과")
    os.makedirs(out_dir, exist_ok=True)

    year = guess_year(files, base)
    taken = set()
    written = []
    for src in files:
        try:
            pages = load_pages(src)
            done = [
                process_image(p, bg=bg, deskew=deskew, clean=clean,
                              rotate_upright=rotate_upright, perspective=perspective)
                for p in pages
            ]
            dst = output_path_for(src, out_dir, fmt, year, taken)
            paths = save_pages(done, dst, to_pdf=(fmt == "pdf"))
            written.extend(paths)
            extra = f" 외 {len(paths) - 1}장" if len(paths) > 1 else ""
            log(f"[완료] {os.path.basename(src)} -> {os.path.basename(paths[0])}{extra}")
        except Exception as e:  # 한 장이 실패해도 나머지는 계속
            log(f"[실패] {os.path.basename(src)}: {e}")
        finally:
            # 큰 배열을 바로 돌려줘서 장수가 많아도 메모리가 쌓이지 않게 한다
            pages = done = None
            gc.collect()
    log(f"저장 위치: {out_dir}")
    return written


# ---------------------------------------------------------------------------
# 엑셀로 묶기 (정리결과 폴더의 이미지를 A4 한 쪽에 2열로 배치)
#
# 원래 '1_영수증_이미지_정리.py' 에 따로 있던 코드를 이 파일 안으로 옮겨 왔다.
# 이 프로그램 하나만 있으면 정리부터 엑셀·JSON 까지 다 되게 하려는 것이다.
# 이름이 겹치는 몇 가지는 앞에 xl_ / XL_ 을 붙여 구분했다.
# 엑셀을 만들 때만 openpyxl 이 필요하므로, 불러오기는 쓰는 함수 안에 두었다.


INCH = 96.0


A4_W_IN, A4_H_IN = 8.2677, 11.6929        # A4 210 x 297 mm


MARGIN_L_IN, MARGIN_R_IN = 15 / 25.4, 10 / 25.4  # 좌우 여백. 왼쪽은 제본에 물리는 만큼 더 준다
MARGIN_T_IN, MARGIN_B_IN = 30 / 25.4, 5 / 25.4   # 상단 30mm는 스프링 펀치 구멍 자리, 하단은 최소만
# 배율은 절대 고정하지 말 것. Zoom을 주면 '페이지 폭 맞춤'이 꺼져 좌우로 페이지가 쪼개진다.


WIDTH_SAFETY = 0.94


HEIGHT_SAFETY = 0.96


USABLE_W = int((A4_W_IN - MARGIN_L_IN - MARGIN_R_IN) * INCH * WIDTH_SAFETY)   # 656 px


USABLE_H = int((A4_H_IN - MARGIN_T_IN - MARGIN_B_IN) * INCH * HEIGHT_SAFETY)  # 950 px


COLS = 2              # 좌우 2열 고정 (3열은 영수증 글씨가 안 보인다)


MAX_BANDS = 2         # 한 페이지에 최대 2줄 -> 최대 4장


COL_PX = USABLE_W // COLS                  # 열 하나의 폭


IMG_W = COL_PX - 20                        # 열 안쪽 좌우 여백 10px씩
FULL_W = USABLE_W - 20                     # 한 쪽을 통째로 쓰는 사진의 폭
# 영수증 두 장이 나란히 붙어 들어온 사진은 열 하나에 넣으면 글씨가 안 보인다.
# 가로/세로 비가 이보다 넓으면 두 장짜리로 보고 한 쪽을 통째로 내준다.
# 두 장짜리(0.516)와 한 장짜리(0.514)가 붙어 있어 비율만으로는 딱 갈리지 않는다.
# 한 장짜리가 크게 나오는 것은 종이만 더 쓸 뿐이라, 기준은 낮은 쪽으로 잡는다.
WIDE_RATIO = 0.50


TITLE_PT = 15.0                            # 제목 행 높이(포인트)


TITLE_PX = TITLE_PT * 4 / 3


DEFAULT_GAP_MM = 8.0                       # 위아래 사진 사이 여백 (--gap 으로 조절)


MM_PX = INCH / 25.4                        # 1mm = 3.78px


MAX_ROW_PT = 380.0                         # 엑셀 행 높이 상한은 409.5pt


SUPERSAMPLE = 2                            # 화면 크기의 2배로 넣어 인쇄 화질 확보


XL_JPEG_QUALITY = 88


EXTS = {".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tif", ".tiff", ".webp"}


XL_PDF_EXTS = {".pdf"}


ALL_EXTS = EXTS | XL_PDF_EXTS


@dataclass
class Source:
    """시트에서 한 칸을 차지하는 원본 하나. 사진 파일 하나, 또는 PDF 한 쪽."""
    path: Path
    page: int | None = None    # PDF 쪽 번호(0부터). 사진이면 None
    pages: int = 1             # 그 PDF의 전체 쪽 수 (제목에 1/3 처럼 적는다)


@dataclass
class Item:
    src: Source
    title: str
    w: int          # 시트에 표시될 폭(px)
    h: int          # 시트에 표시될 높이(px)
    wide: bool = False   # 두 장짜리라 한 쪽을 통째로 쓰는 사진


PDF_INSTALL_HINT = ("PDF를 읽으려면 pypdfium2 가 필요합니다.\n"
                    "  처음설치.bat 을 다시 실행하거나, 명령 프롬프트에 아래를 붙여넣으세요.\n"
                    "    python -m pip install pypdfium2")


_PDF = None          # 찾아낸 PDF 라이브러리 모듈


_PDF_KIND = None     # "pdfium" | "fitz" | "" (없음). None 이면 아직 안 찾아봤다는 뜻


_PDF_OPEN = None     # (경로, 문서) - 방금 연 PDF 하나만 들고 있는다


class PdfUnavailable(Exception):
    """PDF를 읽을 라이브러리가 없다."""


def pdf_lib() -> str:
    """PDF 라이브러리를 한 번만 찾아 두고 종류를 알려 준다."""
    global _PDF, _PDF_KIND
    if _PDF_KIND is None:
        _PDF, _PDF_KIND = None, ""
        for kind, name in (("pdfium", "pypdfium2"), ("fitz", "fitz")):
            try:
                _PDF, _PDF_KIND = importlib.import_module(name), kind
                break
            except ImportError:
                continue
    return _PDF_KIND


def pdf_errors() -> tuple:
    """'이 PDF는 못 읽는다' 예외. 라이브러리마다 이름이 달라 여기서 모아 둔다."""
    if pdf_lib() == "pdfium":
        return (_PDF.PdfiumError, OSError)
    return (RuntimeError, ValueError, OSError)   # fitz.FileDataError 는 RuntimeError


def open_pdf(path: Path):
    """PDF 문서를 연다. 한 파일의 여러 쪽을 잇달아 다루므로 마지막 하나를 재사용한다."""
    global _PDF_OPEN
    if _PDF_OPEN and _PDF_OPEN[0] == path:
        return _PDF_OPEN[1]
    close_pdf()
    kind = pdf_lib()
    if not kind:
        raise PdfUnavailable(PDF_INSTALL_HINT)
    if kind == "pdfium":
        doc = _PDF.PdfDocument(path)
        try:
            doc.init_forms()      # 전자세금계산서처럼 입력칸에 값이 든 PDF 대비
        except (RuntimeError, _PDF.PdfiumError):
            pass
    else:
        doc = _PDF.open(path)
    _PDF_OPEN = (path, doc)
    return doc


def close_pdf():
    global _PDF_OPEN
    if _PDF_OPEN:
        _PDF_OPEN[1].close()
        _PDF_OPEN = None


def pdf_page_count(path: Path) -> int:
    return len(open_pdf(path))


def pdf_page_size(path: Path, page: int) -> tuple[float, float]:
    """PDF 한 쪽의 크기(포인트). 눕혀 저장된 쪽은 회전이 반영된 값이 나온다."""
    doc = open_pdf(path)
    if pdf_lib() == "pdfium":
        return doc[page].get_size()
    rect = doc[page].rect
    return rect.width, rect.height


def pdf_render(path: Path, page: int, target_w: int) -> Image.Image:
    """PDF 한 쪽을 가로 target_w 픽셀 그림으로 그린다.

    PDF 글자는 사진이 아니라 선으로 그려져 있다. 작게 그린 뒤 늘리면 뭉개지므로
    처음부터 필요한 크기로 그려야 인쇄했을 때 품목명과 금액이 읽힌다.
    """
    doc = open_pdf(path)
    pg = doc[page]
    if pdf_lib() == "pdfium":
        scale = target_w / max(1.0, pg.get_size()[0])
        return pg.render(scale=scale).to_pil().convert("RGB")
    scale = target_w / max(1.0, pg.rect.width)
    pix = pg.get_pixmap(matrix=_PDF.Matrix(scale, scale), alpha=False)
    return Image.frombytes("RGB", (pix.width, pix.height), pix.samples)


def natural_key(name: str):
    """20260707_2 가 20260707_10 보다 앞에 오도록 숫자를 숫자로 비교한다."""
    return [int(t) if t.isdigit() else t.lower() for t in re.split(r"(\d+)", name)]


def name_date(name: str, year: int | None = None) -> tuple[int, int, int] | None:
    """파일 이름 앞머리에서 날짜를 읽는다. 못 읽으면 None.

    20260805 / 2026-08-05 / 2026081704(뒤에 순번이 붙은 것) 처럼 연도까지 있는 이름과,
    '8.21 충북농산' 처럼 월.일 만 적은 이름을 모두 알아본다. 뒤엣것은 연도가 없으므로
    같은 폴더의 다른 파일에서 알아낸 연도(year)를 빌려 쓴다.
    """
    s = Path(name).stem
    m = re.match(r"\s*(20\d{2})[.\-_ ]?(\d{2})[.\-_ ]?(\d{2})", s)
    if m:
        y, mo, d = (int(x) for x in m.groups())
    else:
        m = re.match(r"\s*(\d{1,2})[.\-_](\d{1,2})", s)
        if not m or year is None:
            return None
        y, mo, d = year, int(m.group(1)), int(m.group(2))
    return (y, mo, d) if 1 <= mo <= 12 and 1 <= d <= 31 else None


def xl_guess_year(files: list[Path]) -> int | None:
    """연도까지 적힌 이름에서 연도를 알아낸다. ('8.21 …' 같은 이름에 빌려준다)"""
    해 = [d[0] for p in files if (d := name_date(p.name))]
    return max(set(해), key=해.count) if 해 else None


def file_order(files: list[Path]) -> list[Path]:
    """영수증 파일을 날짜 -> 이름 순으로 늘어놓는다.

    '20260805.jpg' 와 '8.21 충북농산.pdf' 처럼 이름 짓는 방식이 섞이면 이름만으로
    줄을 세울 수 없다(숫자로 보면 8 이 20260805 보다 작아 PDF가 앞으로 몰린다).
    그래서 이름에서 날짜를 먼저 읽어 날짜순으로 놓고, 같은 날짜끼리만 이름으로 본다.
    날짜를 못 읽은 파일은 맨 뒤로 보낸다.
    """
    year = xl_guess_year(files)

    def key(p: Path):
        d = name_date(p.name, year) or (9999, 99, 99)
        return (d, natural_key(p.name), natural_key(str(p.parent)))
    return sorted(files, key=key)


def xl_collect_files(folder: Path, recursive: bool = False) -> list[Path]:
    """폴더의 영수증 파일(사진 + PDF) 목록. 윈도우 glob은 대소문자를 구분하지 않아
    *.jpg 와 *.JPG 를 같이 돌리면 전부 두 번 잡히므로 한 번만 훑고 확장자로 거른다.

    파일 이름이 날짜라서 폴더보다 이름을 먼저 본다. pdf 폴더에 따로 받아 둔
    20260806.pdf 가 사진들 뒤로 밀리지 않고 8월 6일 자리에 끼어들어야 한다.
    """
    walk = folder.rglob("*") if recursive else folder.iterdir()
    files = [p for p in walk if p.is_file() and p.suffix.lower() in ALL_EXTS]
    return file_order(files)


def sources_from_files(files: list[Path]) -> list[Source]:
    """파일 목록을 시트에 놓을 칸 목록으로 바꾼다. PDF는 쪽마다 한 칸을 차지한다.

    암호가 걸렸거나 깨진 PDF는 알려 주고 건너뛴다. 한 장 때문에 전체가 멈추면
    안 되기 때문이다. 다만 라이브러리가 아예 없으면 조용히 빠뜨리지 않고 멈춘다.

    3단계 도구는 사진이 여러 폴더에 흩어져 있어(사진 폴더 + PDF 폴더) 파일을
    직접 골라 오므로, 폴더가 아니라 파일 목록을 받는 이 함수를 쓴다.
    """
    pdfs = [p for p in files if p.suffix.lower() in XL_PDF_EXTS]
    if pdfs and not pdf_lib():
        raise PdfUnavailable(f"PDF {len(pdfs)}개를 읽을 수 없습니다.\n{PDF_INSTALL_HINT}")

    bad = pdf_errors()
    sources: list[Source] = []
    for p in files:
        if p.suffix.lower() not in XL_PDF_EXTS:
            sources.append(Source(p))
            continue
        try:
            n = pdf_page_count(p)
        except bad as e:
            print(f"! 건너뜀: {p.name} - 열 수 없는 PDF입니다 (암호가 걸렸을 수 있습니다) {e}")
            continue
        if n == 0:
            print(f"! 건너뜀: {p.name} - 빈 PDF입니다")
        sources.extend(Source(p, i, n) for i in range(n))
    return sources


def collect_sources(folder: Path, recursive: bool = False) -> list[Source]:
    """폴더에서 시트에 놓을 순서대로 원본을 늘어놓는다."""
    return sources_from_files(xl_collect_files(folder, recursive))


def oriented_size(path: Path) -> tuple[int, int]:
    """사진을 다 읽지 않고 크기만 구한다. 휴대폰 사진은 눕혀서 저장하고 회전 정보를
    EXIF 에만 넣어두므로, 세워질 사진이면 가로/세로를 바꿔서 돌려준다."""
    with Image.open(path) as im:
        w, h = im.size
        if im.getexif().get(0x0112, 1) in (5, 6, 7, 8):   # 0x0112 = Orientation
            w, h = h, w
    return w, h


def source_size(src: Source) -> tuple[float, float]:
    """칸을 잡는 데 쓸 가로/세로. 비율만 보므로 사진은 픽셀, PDF는 포인트로 준다."""
    if src.page is None:
        return oriented_size(src.path)
    return pdf_page_size(src.path, src.page)


def title_of(src: Source, keep_ext: bool) -> str:
    """사진 위에 적을 이름. 여러 쪽짜리 PDF는 몇 쪽인지도 같이 적는다."""
    base = src.path.name if keep_ext else src.path.stem
    if src.page is None or src.pages <= 1:
        return base
    return f"{base} ({src.page + 1}/{src.pages})"


def collect_sources_safe(folder: Path) -> list[Source]:
    """폴더가 없거나 읽을 수 없어도 예외 없이 빈 목록을 돌려준다."""
    try:
        return collect_sources(folder)
    except (OSError, PdfUnavailable):
        return []


collect_images_safe = collect_sources_safe      # 3단계 도구가 부르던 이름


def deeper_files(folder: Path) -> int:
    """고른 폴더의 '하위 폴더 안'에 있는 영수증 파일 수. 안내 문구에만 쓴다."""
    try:
        return sum(1 for p in folder.rglob("*")
                   if p.parent != folder and p.suffix.lower() in ALL_EXTS and p.is_file())
    except OSError:
        return 0


def want_subfolders(folder: Path, recursive: bool) -> bool:
    """PDF는 보통 pdf 폴더에 따로 받아 두므로, 아래에 더 있으면 넣을지 물어본다.

    콘솔에서 직접 실행한 때만 물어보고, 아니면 안내만 하고 고른 폴더만 쓴다.
    """
    if recursive:
        return True
    n = deeper_files(folder)
    if not n:
        return False
    msg = f"하위 폴더에 사진/PDF {n}개가 더 있습니다"
    try:
        interactive = sys.stdin is not None and sys.stdin.isatty()
    except ValueError:
        interactive = False
    if not interactive:
        print(f"! {msg}. 함께 넣으려면 -r (--하위폴더) 를 붙이세요.")
        return False
    try:
        answer = input(f"{msg}. 함께 넣을까요? [Y/n] ").strip().lower()
    except EOFError:
        print()
        return False
    return answer in ("", "y", "yes", "ㅛ", "네", "예")


def gather(folder: Path, recursive: bool) -> list[Source]:
    """폴더에서 원본을 모으고, 무엇을 몇 개 읽는지 알려 준다."""
    try:
        sources = collect_sources(folder, want_subfolders(folder, recursive))
    except PdfUnavailable as e:
        raise SystemExit(str(e))
    if not sources:
        raise SystemExit(f"이미지나 PDF가 없습니다: {folder}")

    n_img = sum(1 for s in sources if s.page is None)
    pdfs = {s.path for s in sources if s.page is not None}
    parts = [f"사진 {n_img}장"] if n_img else []
    if pdfs:
        쪽 = len(sources) - n_img
        parts.append(f"PDF {len(pdfs)}개" + (f"({쪽}쪽)" if 쪽 != len(pdfs) else ""))
    print(" + ".join(parts) + " 읽는 중...")

    # 같은 영수증을 사진과 PDF로 둘 다 갖고 있으면 장부에 두 번 오르기 쉽다
    같은이름 = sorted({p.stem for p in pdfs} & {s.path.stem for s in sources if s.page is None})
    if 같은이름:
        더 = f" 외 {len(같은이름) - 5}개" if len(같은이름) > 5 else ""
        print(f"! 같은 이름의 사진과 PDF가 함께 있어 둘 다 넣습니다: "
              f"{', '.join(같은이름[:5])}{더}")
    return sources


def measure(sources: list[Source], keep_ext: bool, max_img_h: int,
            full_img_h: int | None = None) -> list[Item]:
    """사진을 시트에 앉힐 크기로 잰다.

    가로로 넓은(영수증 두 장이 붙은) 사진은 열 하나가 아니라 한 쪽 전체를 쓰게
    재 둔다. full_img_h 는 그때 쓸 수 있는 높이다.
    """
    if full_img_h is None:
        full_img_h = max_img_h
    items = []
    for s in sources:
        iw, ih = source_size(s)
        wide = iw / ih >= WIDE_RATIO
        base_w, cap_h = (FULL_W, full_img_h) if wide else (IMG_W, max_img_h)
        w, h = base_w, max(1, round(base_w * ih / iw))
        if h > cap_h:                            # 페이지보다 긴 사진은 높이에 맞춰 줄인다
            h = cap_h
            w = max(1, round(cap_h * iw / ih))
        items.append(Item(s, title_of(s, keep_ext), w, h, wide))
    return items


def fit_full(item: Item, full_img_h: int) -> Item:
    """사진 하나를 쪽 전체(가로 FULL_W, 세로 full_img_h)에 꽉 차게 다시 잰다."""
    iw, ih = source_size(item.src)
    w, h = FULL_W, max(1, round(FULL_W * ih / iw))
    if h > full_img_h:
        h = full_img_h
        w = max(1, round(full_img_h * iw / ih))
    return Item(item.src, item.title, w, h, True)


def plan_pages(items: list[Item], per_page: str, gap_px: float) -> list[list[list[Item]]]:
    """페이지 -> 줄(band) -> 줄 안의 사진 목록"""
    # 두 장짜리 사진은 옆에 아무것도 세우지 않고 혼자 한 줄을 쓴다.
    bands, cur_band = [], []
    for it in items:
        if it.wide:
            if cur_band:
                bands.append(cur_band)
                cur_band = []
            bands.append([it])
            continue
        cur_band.append(it)
        if len(cur_band) == COLS:
            bands.append(cur_band)
            cur_band = []
    if cur_band:
        bands.append(cur_band)

    # 짝이 없어 혼자 남은 사진은 반 칸만 쓰고 작게 나온다. 어차피 옆이 비므로
    # 쪽 전체를 쓰도록 키워 준다.
    full_img_h = int(USABLE_H - TITLE_PX - gap_px)
    for band in bands:
        if len(band) == 1 and not band[0].wide:
            band[0] = fit_full(band[0], full_img_h)

    def alone(band):
        return len(band) == 1 and band[0].wide

    if per_page == "2":                          # 한 페이지에 한 줄 고정
        return [[b] for b in bands]

    if per_page == "4":                          # 두 줄씩 채우되 두 장짜리는 혼자 한 쪽
        pages, cur = [], []
        for band in bands:
            if alone(band):
                if cur:
                    pages.append(cur)
                    cur = []
                pages.append([band])
                continue
            cur.append(band)
            if len(cur) >= 2:
                pages.append(cur)
                cur = []
        if cur:
            pages.append(cur)
        return pages

    pages, cur, cur_h = [], [], 0.0              # auto: 높이가 허락하는 만큼만 채운다
    for band in bands:
        band_h = TITLE_PX + max(it.h for it in band) + gap_px
        if alone(band):                          # 두 장짜리는 앞뒤로 끊어 혼자 한 쪽을 쓴다
            if cur:
                pages.append(cur)
                cur, cur_h = [], 0.0
            pages.append([band])
            continue
        if cur and (len(cur) >= MAX_BANDS or cur_h + band_h > USABLE_H):
            pages.append(cur)
            cur, cur_h = [], 0.0
        cur.append(band)
        cur_h += band_h
    if cur:
        pages.append(cur)
    return pages


def col_width_from_px(px: int) -> float:
    """열 너비(문자 단위) -> 픽셀 환산식 px = round(width * 7) + 5 의 역함수 (맑은 고딕 11pt 기준)"""
    return (px - 5) / 7.0


def px_from_col_width(width: float) -> int:
    return round(width * 7) + 5


def split_row_heights(total_px: float) -> list[float]:
    """사진 한 장의 높이를 행 높이 상한(409.5pt)에 걸리지 않게 여러 행으로 쪼갠다."""
    total_pt = total_px * 0.75
    n = max(1, math.ceil(total_pt / MAX_ROW_PT))
    return [total_pt / n] * n


def render_source(item: Item) -> Image.Image:
    """사진이든 PDF 한 쪽이든 흰 바탕의 RGB 그림 한 장으로 만든다."""
    src = item.src
    if src.page is not None:
        return pdf_render(src.path, src.page, item.w * SUPERSAMPLE)
    with Image.open(src.path) as raw:
        im = ImageOps.exif_transpose(raw)          # EXIF 회전 정보를 실제 화소에 반영
        if im.mode in ("RGBA", "LA") or (im.mode == "P" and "transparency" in im.info):
            im = im.convert("RGBA")                # 투명한 부분은 흰색으로 깔아준다
            bg = Image.new("RGB", im.size, "white")
            bg.paste(im, mask=im.split()[-1])
            return bg
        return im.convert("RGB")


def place_image(ws, tmpdir: Path, idx: int, item: Item, col0: int, row0: int, col_px: int):
    from openpyxl.drawing.image import Image as XLImage
    from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor
    from openpyxl.drawing.xdr import XDRPositiveSize2D
    from openpyxl.utils.units import pixels_to_EMU

    """사진을 셀 좌표가 아니라 픽셀 단위로 정확히 앉힌다(열 안에서 가운데 정렬)."""
    im = render_source(item)
    target = (item.w * SUPERSAMPLE, item.h * SUPERSAMPLE)
    if im.size != target:                          # PDF는 이미 이 크기로 그려져 온다
        im = im.resize(target, Image.LANCZOS)
    tmp = tmpdir / f"{idx:04d}.jpg"
    im.save(tmp, "JPEG", quality=XL_JPEG_QUALITY)
    im.close()

    pic = XLImage(str(tmp))
    pic.width, pic.height = item.w, item.h
    dx = max(0, (col_px - item.w) // 2)
    pic.anchor = OneCellAnchor(
        _from=AnchorMarker(col=col0, colOff=pixels_to_EMU(dx), row=row0, rowOff=0),
        ext=XDRPositiveSize2D(pixels_to_EMU(item.w), pixels_to_EMU(item.h)),
    )
    ws.add_image(pic)


def build_workbook(pages, tmpdir: Path, gap_px: float, sheet_title: str = "영수증 이미지"):
    from openpyxl import Workbook

    wb = Workbook()
    wb.remove(wb.active)
    add_image_sheet(wb, pages, tmpdir, gap_px, sheet_title)
    return wb


def add_image_sheet(wb, pages, tmpdir: Path, gap_px: float, sheet_title: str = "영수증 이미지"):
    from openpyxl.styles import Alignment, Font
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.pagebreak import Break
    from openpyxl.worksheet.properties import PageSetupProperties

    """이미 있는 통합문서에 사진 시트를 하나 추가한다 (3단계 도구가 이 함수를 쓴다)."""
    ws = wb.create_sheet(sheet_title)
    ws.sheet_view.showGridLines = False

    col_width = col_width_from_px(COL_PX)
    actual_col_px = px_from_col_width(col_width)
    for c in range(1, COLS + 1):
        ws.column_dimensions[get_column_letter(c)].width = col_width

    title_font = Font(name="맑은 고딕", size=10, bold=True)
    title_align = Alignment(horizontal="center", vertical="center")

    r = 1
    n_img = 0
    for page_no, page in enumerate(pages):
        for band in page:
            # 1) 제목 행
            ws.row_dimensions[r].height = TITLE_PT
            wide_band = len(band) == 1 and band[0].wide
            for i, item in enumerate(band):
                cell = ws.cell(row=r, column=i + 1)
                cell.value = item.title
                cell.data_type = "s"        # '=' 로 시작하는 이름이 수식으로 저장되는 것 방지
                cell.font = title_font
                cell.alignment = title_align
            if wide_band and COLS > 1:      # 두 장짜리는 제목도 쪽 전체에 걸친다
                ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=COLS)
            r += 1

            # 2) 사진 행 (줄 안에서 가장 긴 사진 기준)
            band_h = max(it.h for it in band)
            heights = split_row_heights(band_h)
            img_row0 = r
            for h_pt in heights:
                ws.row_dimensions[r].height = h_pt
                r += 1
            band_col_px = actual_col_px * COLS if wide_band else actual_col_px
            for i, item in enumerate(band):
                n_img += 1
                place_image(ws, tmpdir, n_img, item, i, img_row0 - 1, band_col_px)

            # 3) 여백 행. 페이지 나눔은 항상 이 행에 걸어 둔다.
            ws.row_dimensions[r].height = gap_px * 0.75
            gap_row = r
            r += 1

        if page_no < len(pages) - 1:
            ws.row_breaks.append(Break(id=gap_row))

    # 인쇄 설정
    ws.print_area = f"A1:{get_column_letter(COLS)}{r - 1}"
    ws.print_options.horizontalCentered = True
    ws.page_setup.orientation = "portrait"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_margins.left, ws.page_margins.right = MARGIN_L_IN, MARGIN_R_IN
    ws.page_margins.top, ws.page_margins.bottom = MARGIN_T_IN, MARGIN_B_IN
    # 머리글/바닥글 여백이 위아래 여백보다 크면 본문이 밀려 첫 줄 제목이 잘린다.
    ws.page_margins.header = ws.page_margins.footer = 0
    # 폭이 조금 넘쳐도 오른쪽 열이 별도 페이지로 떨어지지 않게 하는 보험
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    close_pdf()                 # 사진을 다 앉혔으니 열어 둔 PDF를 놓아준다
    return ws


def plan_sources(sources: list[Source], per_page: str, keep_ext: bool, gap_mm: float):
    """모아 둔 원본들을 재서 페이지 배치까지 끝낸다. (pages, gap_px)

    3단계 도구가 여러 폴더에서 골라 온 사진·PDF를 그대로 넘겨 쓴다.
    """
    gap_px = gap_mm * MM_PX
    bands_per_page = 1 if per_page == "2" else (2 if per_page == "4" else 1)
    max_img_h = int(USABLE_H / bands_per_page - TITLE_PX - gap_px)
    if max_img_h < 100:
        raise SystemExit(f"여백({gap_mm}mm)이 너무 커서 사진이 들어갈 자리가 없습니다.")
    full_img_h = int(USABLE_H - TITLE_PX - gap_px)   # 두 장짜리가 혼자 쓸 때의 높이
    items = measure(sources, keep_ext, max_img_h, full_img_h)
    return plan_pages(items, per_page, gap_px), gap_px


def plan_from_folder(folder: Path, per_page: str, keep_ext: bool, gap_mm: float,
                     recursive: bool = False):
    """폴더의 사진과 PDF를 재서 페이지 배치까지 끝낸다. (pages, gap_px, 칸 수)"""
    sources = gather(folder, recursive)
    pages, gap_px = plan_sources(sources, per_page, keep_ext, gap_mm)
    return pages, gap_px, len(sources)


# '정리결과' 처럼 어느 달에나 똑같이 붙는 이름들. 이 안에 만들면 파일 이름이 매번
# 같아져 월별로 구분이 안 되므로, 한 단계 위 폴더 이름(예: 26.8월분)을 대신 쓴다.
GENERIC_FOLDERS = {"정리결과", "결과", "출력", "output", "result", "results", "pdf", "img",
                   "image", "images", "scan", "scans"}


def excel_output_path(folder):
    """엑셀을 저장할 경로. 이름은 '월별 폴더 이름_이미지정리.xlsx' 가 되게 한다."""
    folder = os.path.abspath(folder)
    name = os.path.basename(folder)
    parent = os.path.basename(os.path.dirname(folder))
    # '26.8월분_정리결과' 처럼 꼬리표가 붙은 폴더는 꼬리표를 떼고 원본 이름을 쓴다.
    if name.endswith("_정리결과") and len(name) > len("_정리결과"):
        name = name[: -len("_정리결과")]
    if name.strip().lower() in GENERIC_FOLDERS and parent:
        name = parent
    return os.path.join(folder, f"{name}_이미지정리.xlsx")


def make_excel(target, out_path=None, per_page="auto", gap_mm=None, keep_ext=False,
               recursive=True, log=print):
    """폴더(보통 '정리결과')의 이미지와 PDF를 엑셀 한 파일로 묶는다.

    target 은 폴더 경로, 또는 파일 경로 목록. 저장한 파일 경로를 돌려준다.
    """
    gap_mm = DEFAULT_GAP_MM if gap_mm is None else gap_mm

    if isinstance(target, (str, os.PathLike)):
        folder = str(target)
        if not os.path.isdir(folder):
            raise NotADirectoryError(f"폴더가 아닙니다: {folder}")
        files = xl_collect_files(Path(folder), recursive)
    else:
        files = file_order([Path(p) for p in target])
        if not files:
            raise FileNotFoundError("엑셀로 만들 이미지가 없습니다.")
        folder = str(files[0].parent)

    if not files:
        raise FileNotFoundError(f"이미지나 PDF가 없습니다: {folder}")

    sources = sources_from_files(files)
    if not sources:
        raise FileNotFoundError(f"읽을 수 있는 이미지나 PDF가 없습니다: {folder}")

    pages, gap_px = plan_sources(sources, per_page, keep_ext, gap_mm)
    log(f"이미지 {len(sources)}장을 엑셀 {len(pages)}쪽으로 배치합니다...")

    out = out_path or excel_output_path(folder)
    with tempfile.TemporaryDirectory() as td:
        # openpyxl 은 저장하는 순간 이미지 파일을 다시 읽으므로 임시 폴더는 저장 뒤에 지운다
        wb = build_workbook(pages, Path(td), gap_px)
        try:
            wb.save(out)
        except PermissionError:
            stem, ext = os.path.splitext(out)
            out = f"{stem}_새파일{ext}"
            log(f"! 파일이 열려 있어 {os.path.basename(out)} 로 저장합니다.")
            wb.save(out)
    log(f"[엑셀] {out}")
    log(f"엑셀 정리가 끝났습니다. (이미지 {len(sources)}장 / {len(pages)}쪽) "
        f"-> {os.path.basename(out)}")
    return out


# ---------------------------------------------------------------------------
# 손으로 다듬기 (자동으로 못 떼어내는 조각을 직접 지우는 창)


def apply_edit(img, box, mode):
    """손질 창에서 그린 사각형을 적용한다. mode: "erase"(흰색으로) | "keep"(잘라내기)"""
    if mode == "keep":
        return img.crop(box)
    out = img.copy()
    white = Image.new("RGB", (box[2] - box[0], box[3] - box[1]), (255, 255, 255))
    out.paste(white, box[:2])
    return out


def apply_perspective(img, points):
    """찍은 네 점을 직사각형으로 펴서 잘라낸다(비스듬히 찍힌 영수증 바로잡기)."""
    quad = _order_corners(np.array(points, dtype=np.float32))
    bgr = cv2.cvtColor(np.array(img.convert("RGB")), cv2.COLOR_RGB2BGR)
    out = warp_quad(bgr, quad)
    return Image.fromarray(cv2.cvtColor(out, cv2.COLOR_BGR2RGB))


def apply_rotation(img, angle, trim=True):
    """이미지를 angle 도만큼 돌린다(양수 = 시계 방향). 빈 자리는 흰색으로 채운다.

    돌리면 네 귀퉁이에 흰 여백이 생기므로, 기본적으로 내용에 맞춰 다시 바짝 자른다.
    """
    if abs(angle) < 0.01:
        return img
    out = img.rotate(-angle, resample=Image.BICUBIC, expand=True,
                     fillcolor=(255, 255, 255))
    if not trim:
        return out
    bgr = cv2.cvtColor(np.array(out), cv2.COLOR_RGB2BGR)
    box = tighten_box((0, 0, out.width, out.height), ink_mask(bgr))
    return out.crop(box)


def edit_list(picks):
    """손질 창에 올릴 목록을 만든다. 한 장만 골랐으면 같은 폴더의 이미지를 모두 올린다.

    이전/다음 단추와 휠로 폴더 안을 오갈 수 있게 하려는 것이다.
    반환값: (목록, 처음 보여줄 위치)
    """
    picks = [p for p in picks if os.path.splitext(p)[1].lower() in IMAGE_EXTS]
    if not picks:
        return [], 0
    if len(picks) > 1:
        return picks, 0

    # 파일 선택 창은 윈도우에서도 슬래시(/) 경로를 돌려주므로, 폴더 목록과 견주려면
    # 양쪽을 같은 모양으로 맞춰야 한다. 대소문자도 무시한다.
    chosen = os.path.normcase(os.path.abspath(picks[0]))
    folder = os.path.dirname(chosen)
    siblings = [os.path.join(folder, n) for n in sorted(os.listdir(folder))
                if os.path.splitext(n)[1].lower() in IMAGE_EXTS]
    keys = [os.path.normcase(os.path.abspath(x)) for x in siblings]
    if chosen in keys:
        return siblings, keys.index(chosen)
    return picks, 0


# -- 크기/형식 바꾸기 (이미지 수정 창에서 쓴다) ------------------------------

RESIZE_FORMATS = ("원본 그대로", "JPG", "PNG", "WEBP")


def scale_image(img, ratio):
    w = max(1, int(round(img.width * ratio)))
    h = max(1, int(round(img.height * ratio)))
    return img.resize((w, h), Image.LANCZOS)


def _encode(img, ext, quality, palette=False, fast=False):
    """메모리에 한 번 저장해 보고 바이트 수를 잰다(용량 맞추기용).

    palette: PNG를 256색으로 줄여 저장한다(영수증은 거의 흑백이라 눈으로는 같다).
    fast: 더듬어 찾는 동안에는 optimize 를 끄고 빨리 잰다.
    """
    buf = io.BytesIO()
    if ext in (".jpg", ".jpeg"):
        img.convert("RGB").save(buf, "JPEG", quality=quality, optimize=not fast)
    elif ext == ".webp":
        img.save(buf, "WEBP", quality=quality)
    else:
        if palette:
            img = img.convert("RGB").quantize(256, method=Image.Quantize.FASTOCTREE)
        img.save(buf, "PNG", optimize=not fast, compress_level=6 if fast else 9)
    return buf.getvalue()


def fit_to_size(img, ext, limit_bytes):
    """정한 용량 안에서 되도록 목표에 가깝게(=최대한 선명하게) 맞춘다.

    1) 원래 크기 그대로 품질(PNG는 256색)만 낮춰서 들어가면 그것을 쓴다.
    2) 안 되면 가로세로 비율을 이분법으로 찾아 '들어가는 가장 큰 크기'를 고르고,
       남는 용량만큼 품질을 다시 올려 목표 용량에 바짝 붙인다.
    """
    lossy = ext in (".jpg", ".jpeg", ".webp")
    if lossy:
        def enc(im, q, fast=True):
            return _encode(im, ext, q, fast=fast)

        def best_quality(im, lo, hi):
            best = None
            while lo <= hi:                      # 품질을 이분법으로 더듬어 찾는다
                mid = (lo + hi) // 2
                data = enc(im, mid)
                if len(data) <= limit_bytes:
                    best, lo = (mid, data), mid + 1
                else:
                    hi = mid - 1
            return best

        found = best_quality(img, 40, 95)
        if found:
            final = enc(img, found[0], fast=False)   # optimize 를 켜면 더 작아진다
            return final if len(final) <= limit_bytes else found[1]
        base_q = 75                               # 크기를 줄일 때 기준 품질
        fits = lambda im: len(enc(im, base_q)) <= limit_bytes
    else:
        def enc(im, q=None, fast=True, palette=True):
            return _encode(im, ext, 0, palette=palette, fast=fast)

        full = enc(img, fast=False, palette=False)   # 손실 없이 들어가면 그대로
        if len(full) <= limit_bytes:
            return full
        full = enc(img, fast=False)                  # 256색으로 줄여 본다
        if len(full) <= limit_bytes:
            return full
        fits = lambda im: len(enc(im)) <= limit_bytes

    # 들어가는 가장 큰 비율을 찾는다. 용량은 대략 면적(비율²)에 비례하므로
    # 첫 짐작을 그렇게 잡고, 그 둘레에서 이분법으로 좁힌다.
    raw = len(enc(img, base_q) if lossy else enc(img))
    guess = min(0.99, math.sqrt(limit_bytes / max(1, raw)))
    lo, hi = 0.02, 1.0
    if fits(scale_image(img, guess)):
        lo = guess
        if guess * 1.15 < hi and not fits(scale_image(img, guess * 1.15)):
            hi = guess * 1.15
    else:
        hi = guess
        if guess * 0.85 > lo and fits(scale_image(img, guess * 0.85)):
            lo = guess * 0.85
    while (hi - lo) / hi > 0.01:             # 1% 차이 안쪽이면 그만 찾는다
        mid = (lo + hi) / 2
        if fits(scale_image(img, mid)):
            lo = mid
        else:
            hi = mid
    work = scale_image(img, lo)

    if lossy:
        found = best_quality(work, base_q, 95) or best_quality(work, 20, base_q - 1)
        if found:
            final = enc(work, found[0], fast=False)
            return final if len(final) <= limit_bytes else found[1]
        return enc(work, 20)
    return enc(work, fast=False)


def edited_dir(src):
    """원본 폴더와 나란히 놓일 '<폴더이름>_수정' 폴더 경로를 돌려준다(없으면 만든다).

        상위폴더/26.8월분/a.jpg  ->  상위폴더/26.8월분_수정/a.jpg

    원본 폴더 '안'이 아니라 '옆'에 만든다. 그래야 다시 정리를 돌릴 때 편집본이
    원본과 섞여 두 번 처리되지 않는다.
    """
    folder = os.path.dirname(os.path.abspath(src))
    parent, name = os.path.split(folder)
    if not name:                       # 드라이브 뿌리에 놓인 파일
        out = os.path.join(folder, "수정")
    elif name.endswith("_수정"):        # 이미 편집 폴더면 그 안에 그대로 둔다
        out = folder
    else:
        out = os.path.join(parent, name + "_수정")
    os.makedirs(out, exist_ok=True)
    return out


def save_changed(img, src, fmt="원본 그대로", mode="none", percent=100.0,
                 kb=0, width=0, height=0, copy_ok=False):
    """손본 이미지를 크기/형식을 바꿔 '<원본폴더이름>_수정' 폴더에 저장한다.

    파일 이름은 원본 그대로 쓰고, 형식이 바뀌면 확장자만 바뀐다.
    img 에 None 을 주면 원본 파일을 그때 읽는다. 손대지 않고 그대로 복사만 하면
    되는 장은 아예 읽지도 않으므로 여러 장을 한꺼번에 내보낼 때 훨씬 빠르다.
    copy_ok 는 '이 장은 화면에서 손본 적이 없다'는 뜻이다.
    반환값: 저장한 파일 경로
    """
    src_ext = os.path.splitext(src)[1].lower()
    ext = {"JPG": ".jpg", "PNG": ".png", "WEBP": ".webp"}.get(fmt, src_ext)
    if ext not in (".jpg", ".jpeg", ".png", ".webp"):
        ext = ".png"

    # 같은 장을 여러 번 저장하면 '_수정' 폴더의 같은 파일을 덮어쓴다.
    # (원본은 그대로 있으므로 언제든 처음부터 다시 손볼 수 있다)
    stem = os.path.splitext(os.path.basename(src))[0]
    out = os.path.join(edited_dir(src), stem + ext)

    if copy_ok and mode in ("none", "") and ext == src_ext:
        # 크기도 형식도 그대로면 다시 압축할 까닭이 없다. 복사가 빠르고 손실도 없다.
        shutil.copy2(src, out)
        return out

    if img is None:
        img = Image.open(src).convert("RGB")

    if mode == "percent" and percent > 0 and abs(percent - 100.0) > 0.01:
        img = scale_image(img, percent / 100.0)
    elif mode == "wh" and (width > 0 or height > 0):
        if width > 0 and height > 0:
            img = img.resize((width, height), Image.LANCZOS)
        elif width > 0:
            img = scale_image(img, width / img.width)
        else:
            img = scale_image(img, height / img.height)

    if mode == "size" and kb > 0:
        with open(out, "wb") as f:
            f.write(fit_to_size(img, ext, int(kb) * 1024))
    elif ext in (".jpg", ".jpeg"):
        img.convert("RGB").save(out, "JPEG", quality=95)
    elif ext == ".webp":
        img.save(out, "WEBP", quality=95)
    else:
        img.save(out)
    return out


def manual_edit(paths, start=0, log=print):
    """결과 이미지를 하나씩 열어 마우스로 손보는 창을 띄운다.

    - 지우기: 끌어서 그린 사각형을 흰색으로 칠한다(겹쳐 놓인 조각 제거용)
    - 남길 부분만: 끌어서 그린 사각형만 남기고 잘라낸다
    - 원근 맞추기: 네 모서리를 차례로 찍으면 반듯한 직사각형으로 펴서 잘라낸다
    휠은 파일 넘기기, Ctrl+휠은 확대/축소, 오른쪽(또는 가운데) 버튼 끌기는 화면 이동.
    저장하면 같은 파일에 덮어쓴다. 원본 스캔 파일은 건드리지 않는다.
    """
    import threading
    import time
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
    from PIL import ImageTk

    # 빈 목록으로도 연다. 창을 띄워 놓고 파일이나 폴더를 끌어다 놓으면 된다.
    paths = [p for p in paths if os.path.splitext(p)[1].lower() in IMAGE_EXTS]

    # 정리 창이 이미 떠 있으면 그 창에 딸린 새 창으로 연다. 창마다 따로 Tk()를 만들면
    # 이미지와 선택 항목이 다른 창에 묶여 화면이 비어 보인다.
    parent = tk._default_root
    root = tk.Toplevel(parent) if parent is not None else tk.Tk()
    root.title("이미지 수정")
    # 화면(작업 표시줄 제외)보다 크면 아래 버튼 줄이 잘린다. 화면에 맞춰 줄여서 띄운다.
    win_w = min(1000, root.winfo_screenwidth() - 80)
    win_h = min(720, root.winfo_screenheight() - 160)
    root.geometry(f"{win_w}x{win_h}+40+30")

    start = min(max(0, start), len(paths) - 1)
    state = {"index": start, "image": None, "shown": None, "dirty": False,
             "zoom": 1.0, "scale": 1.0, "view": (0.0, 0.0), "angle": 0.0,
             "preview": None, "preview_base": None, "repeat": None,
             "undo": [], "start": None, "pan": None, "points": []}

    bar = ttk.Frame(root, padding=6)
    bar.pack(fill="x")
    mode = tk.StringVar(master=root, value="keep")
    for text, val in (("지우기(흰색으로)", "erase"), ("남길 부분만", "keep"),
                      ("원근 맞추기(네 모서리)", "perspective")):
        ttk.Radiobutton(bar, text=text, value=val, variable=mode,
                        command=lambda: reset_points()).pack(side="left", padx=(0, 8))
    title = ttk.Label(bar, text="")
    title.pack(side="right")

    tilt = ttk.Frame(root, padding=(6, 0, 6, 4))
    tilt.pack(fill="x")
    ttk.Label(tilt, text="기울기").pack(side="left")
    angle_var = tk.DoubleVar(master=root, value=0.0)
    angle_scale = ttk.Scale(tilt, from_=-15.0, to=15.0, variable=angle_var,
                            length=320, command=lambda v: on_angle(float(v)))
    angle_scale.pack(side="left", padx=6)
    angle_label = ttk.Label(tilt, text="0.0°", width=7)
    angle_label.pack(side="left")
    for text, step, pad in (("◀ 0.1", -0.1, (6, 2)), ("0.1 ▶", 0.1, (0, 0))):
        b = ttk.Button(tilt, text=text, width=6)
        b.pack(side="left", padx=pad)
        b.bind("<ButtonPress-1>", lambda e, s=step: hold(s)(e))
        b.bind("<ButtonRelease-1>", lambda e: stop_hold(e))
        b.bind("<Leave>", lambda e: stop_hold(e))   # 누른 채 밖으로 나가도 멈춘다
        for key in ("<space>", "<Return>"):         # 키보드로도 한 번씩 움직인다
            b.bind(key, lambda e, s=step: nudge(s))
    ttk.Button(tilt, text="적용", command=lambda: commit_angle()).pack(side="left", padx=6)
    ttk.Button(tilt, text="되돌림", command=lambda: set_angle(0.0)).pack(side="left")
    ttk.Button(tilt, text="↺ 90°", width=6,
               command=lambda: turn(90)).pack(side="left", padx=(16, 2))
    ttk.Button(tilt, text="↻ 90°", width=6, command=lambda: turn(-90)).pack(side="left")

    size = ttk.LabelFrame(root, text="크기·형식 바꾸기 (원본 폴더 옆 '_수정' 폴더에 저장)",
                          padding=(6, 2))
    size.pack(fill="x", padx=6, pady=(0, 4))
    size_mode = tk.StringVar(master=root, value="size")
    percent_var = tk.StringVar(master=root, value="50")
    kb_var = tk.StringVar(master=root, value="200")
    w_var = tk.StringVar(master=root, value="")
    h_var = tk.StringVar(master=root, value="")
    fmt_var = tk.StringVar(master=root, value="원본 그대로")

    ttk.Radiobutton(size, text="크기 그대로", value="none",
                    variable=size_mode).pack(side="left")
    ttk.Radiobutton(size, text="원본 대비", value="percent",
                    variable=size_mode).pack(side="left", padx=(10, 2))
    ttk.Entry(size, textvariable=percent_var, width=5).pack(side="left")
    ttk.Label(size, text="%").pack(side="left")
    ttk.Radiobutton(size, text="용량 맞추기", value="size",
                    variable=size_mode).pack(side="left", padx=(10, 2))
    ttk.Entry(size, textvariable=kb_var, width=6).pack(side="left")
    ttk.Label(size, text="KB").pack(side="left")
    ttk.Radiobutton(size, text="가로×세로", value="wh",
                    variable=size_mode).pack(side="left", padx=(10, 2))
    ttk.Entry(size, textvariable=w_var, width=6).pack(side="left")
    ttk.Label(size, text="×").pack(side="left")
    ttk.Entry(size, textvariable=h_var, width=6).pack(side="left")
    ttk.Label(size, text="형식").pack(side="left", padx=(12, 2))
    ttk.Combobox(size, textvariable=fmt_var, values=list(RESIZE_FORMATS),
                 state="readonly", width=10).pack(side="left")

    # 아래 버튼 줄과 안내를 그림칸보다 먼저 앉힌다. 그래야 창을 줄여도
    # 그림칸만 작아지고 버튼이 잘리지 않는다.
    foot = ttk.Frame(root, padding=6)
    foot.pack(side="bottom", fill="x")

    HINT = ("휠: 이전/다음 파일   |   Ctrl+휠: 확대·축소   |   "
            "오른쪽 버튼 끌기(확대했을 때 그림 밀기)   |   Ctrl+Z: 되돌리기")
    hint = ttk.Label(root, padding=(8, 2), foreground="#444", text=HINT)
    hint.pack(side="bottom", fill="x")

    canvas = tk.Canvas(root, background="#606060", cursor="cross", highlightthickness=0)
    canvas.pack(side="top", fill="both", expand=True)

    # -- 화면 그리기 ---------------------------------------------------------

    def fit_scale():
        img = view_image()
        cw, ch = max(50, canvas.winfo_width()), max(50, canvas.winfo_height())
        return min(cw / img.width, ch / img.height)

    def view_image():
        """화면에 보일 이미지. 기울기를 만지는 중이면 미리보기를 쓴다."""
        return state["preview"] if state["preview"] is not None else state["image"]

    def render():
        img = view_image()
        if img is None:
            canvas.delete("all")
            title.config(text="파일을 끌어다 놓으세요")
            return
        cw, ch = max(50, canvas.winfo_width()), max(50, canvas.winfo_height())
        scale = fit_scale() * state["zoom"]
        vw, vh = min(img.width, cw / scale), min(img.height, ch / scale)
        vx = min(max(0.0, state["view"][0]), max(0.0, img.width - vw))
        vy = min(max(0.0, state["view"][1]), max(0.0, img.height - vh))
        state["view"], state["scale"] = (vx, vy), scale

        region = img.crop((int(vx), int(vy), int(vx + vw), int(vy + vh)))
        disp = region.resize((max(1, int(vw * scale)), max(1, int(vh * scale))),
                             Image.LANCZOS)
        state["shown"] = ImageTk.PhotoImage(disp, master=canvas)
        canvas.delete("all")
        canvas.create_image(0, 0, anchor="nw", image=state["shown"])

        for i, (px, py) in enumerate(state["points"]):   # 원근용으로 찍은 점
            x, y = to_canvas(px, py)
            canvas.create_oval(x - 6, y - 6, x + 6, y + 6, outline="#ff2d55", width=2)
            canvas.create_text(x, y - 14, text=str(i + 1), fill="#ff2d55")
        if len(state["points"]) > 1:
            pts = [c for p in state["points"] for c in to_canvas(*p)]
            canvas.create_line(*pts, fill="#ff2d55", width=2)

        if not paths:
            title.config(text="파일을 끌어다 놓으세요")
            return
        name = os.path.basename(paths[state["index"]])
        mark = " *" if state["dirty"] else ""
        title.config(text=f"{state['index'] + 1}/{len(paths)}  {name}{mark}  "
                          f"({int(state['zoom'] * 100)}%)")

    def to_image(cx, cy):
        vx, vy = state["view"]
        return vx + cx / state["scale"], vy + cy / state["scale"]

    def to_canvas(ix, iy):
        vx, vy = state["view"]
        return (ix - vx) * state["scale"], (iy - vy) * state["scale"]

    def reset_points():
        state["points"] = []
        render()

    def on_angle(value):
        """기울기 손잡이를 움직이면 축소본으로 미리 보여 준다(빠르게)."""
        state["angle"] = round(value, 1)
        angle_label.config(text=f"{state['angle']:+.1f}°")
        if abs(state["angle"]) < 0.05:
            state["preview"] = None
        else:
            if state["preview_base"] is None:   # 축소본은 한 번만 만들어 두고 다시 쓴다
                base = state["image"].copy()
                base.thumbnail((1400, 1400))
                state["preview_base"] = base
            state["preview"] = apply_rotation(state["preview_base"], state["angle"],
                                              trim=False)
        state["zoom"], state["view"] = 1.0, (0.0, 0.0)
        render()

    def set_angle(value):
        angle_var.set(value)
        on_angle(value)

    def nudge(step):
        set_angle(max(-15.0, min(15.0, state["angle"] + step)))

    def hold(step):
        """단추를 누르고 있는 동안 계속 돌아가게 한다(놓으면 멈춘다).

        오래 누르고 있으면 조금씩 빨라져서 큰 각도도 금방 맞출 수 있다.
        """
        def again(count=0):
            nudge(step)
            delay = 110 if count < 8 else 50
            state["repeat"] = root.after(delay, lambda: again(count + 1))

        def press(event):
            stop_hold()
            nudge(step)                       # 한 번 눌렀다 떼면 0.1도만 움직인다
            state["repeat"] = root.after(400, again)   # 계속 누르고 있으면 그때부터 반복

        return press

    def stop_hold(event=None):
        if state["repeat"] is not None:
            root.after_cancel(state["repeat"])
            state["repeat"] = None

    def commit_angle():
        """미리보던 기울기를 실제 이미지에 적용한다."""
        stop_hold()
        if abs(state["angle"]) < 0.05:
            return False
        angle = state["angle"]
        state["preview"] = None
        set_angle(0.0)
        replace(apply_rotation(state["image"], angle))
        flash(f"기울기 {angle:+.1f}° 적용했습니다.")
        return True

    def turn(degrees):
        """90도씩 돌린다(스캔 방향이 통째로 누웠을 때)."""
        state["preview"] = None
        set_angle(0.0)
        rot = Image.ROTATE_90 if degrees > 0 else Image.ROTATE_270
        replace(state["image"].transpose(rot))
        state["zoom"], state["view"] = 1.0, (0.0, 0.0)
        render()

    def load(i):
        state.update(index=i, undo=[], dirty=False, zoom=1.0, view=(0.0, 0.0),
                     points=[], start=None, preview=None, preview_base=None, angle=0.0)
        angle_var.set(0.0)
        angle_label.config(text="+0.0°")
        state["image"] = Image.open(paths[i]).convert("RGB")
        render()

    def replace(img, mark_dirty=True):
        state["undo"].append(state["image"])
        state["image"] = img
        state["preview_base"] = None
        state["dirty"] = mark_dirty
        state["points"] = []
        render()

    # -- 마우스 --------------------------------------------------------------

    def on_press(event):
        if commit_angle():      # 기울기를 손보던 중이면 먼저 확정하고 시작한다
            return
        if mode.get() == "perspective":
            px, py = to_image(event.x, event.y)
            img = view_image()
            # 그림칸에는 이미지 바깥의 빈 자리(회색)도 있다. 거기 찍은 점은
            # 원근을 펼 때 쓸 수 없으므로 아예 받지 않는다.
            if img is None or not (0 <= px < img.width and 0 <= py < img.height):
                flash("모서리는 이미지 안에서 찍어 주세요.")
                return
            state["points"].append((px, py))
            if len(state["points"]) == 4:
                replace(apply_perspective(state["image"], state["points"]))
                state["zoom"], state["view"] = 1.0, (0.0, 0.0)
                render()
            else:
                render()
            return
        state["start"] = (event.x, event.y)

    def on_drag(event):
        if not state["start"]:
            return
        canvas.delete("band")
        x, y = state["start"]
        canvas.create_rectangle(x, y, event.x, event.y, outline="#ff2d55",
                                width=2, tags="band")

    def on_release(event):
        if not state["start"]:
            return
        (x0, y0), (x1, y1) = state["start"], (event.x, event.y)
        state["start"] = None
        canvas.delete("band")
        ax0, ay0 = to_image(min(x0, x1), min(y0, y1))
        ax1, ay1 = to_image(max(x0, x1), max(y0, y1))
        img = state["image"]
        box = (max(0, int(ax0)), max(0, int(ay0)),
               min(img.width, int(ax1)), min(img.height, int(ay1)))
        if box[2] - box[0] < 3 or box[3] - box[1] < 3:
            return
        replace(apply_edit(img, box, mode.get()))

    def on_pan_start(event):
        state["pan"] = (event.x, event.y, *state["view"])

    def on_pan(event):
        if not state["pan"]:
            return
        px, py, vx, vy = state["pan"]
        state["view"] = (vx - (event.x - px) / state["scale"],
                         vy - (event.y - py) / state["scale"])
        render()

    def on_pan_end(event):
        state["pan"] = None

    def zoom_at(cx, cy, factor):
        ix, iy = to_image(cx, cy)
        state["zoom"] = min(12.0, max(1.0, state["zoom"] * factor))
        scale = fit_scale() * state["zoom"]
        state["view"] = (ix - cx / scale, iy - cy / scale)
        render()

    def on_wheel(event):
        if event.state & 0x0004:                    # Ctrl 를 누른 채면 확대·축소
            zoom_at(event.x, event.y, 1.25 if event.delta > 0 else 0.8)
        else:                                        # 그냥 휠이면 파일 넘기기
            go(-1 if event.delta > 0 else 1)
        return "break"   # 창에도 같은 처리가 걸려 있어 한 번 더 도는 것을 막는다

    canvas.bind("<ButtonPress-1>", on_press)
    canvas.bind("<B1-Motion>", on_drag)
    canvas.bind("<ButtonRelease-1>", on_release)
    for button in ("2", "3"):                        # 가운데/오른쪽 버튼으로 화면 이동
        canvas.bind(f"<ButtonPress-{button}>", on_pan_start)
        canvas.bind(f"<B{button}-Motion>", on_pan)
        canvas.bind(f"<ButtonRelease-{button}>", on_pan_end)
    canvas.bind("<MouseWheel>", on_wheel)
    canvas.bind("<Configure>", lambda e: render())
    # 윈도우에서 휠은 '포커스를 가진 위젯'으로 간다. 파일 선택 창을 거쳐 열면 포커스가
    # 다른 창에 남아 휠이 안 먹으므로, 창 전체에 걸어 두고 마우스가 들어오면 포커스를 준다.
    root.bind("<MouseWheel>", on_wheel)
    canvas.configure(takefocus=1)
    canvas.bind("<Enter>", lambda e: canvas.focus_set())

    # -- 버튼 ----------------------------------------------------------------

    def undo(event=None):
        if state["undo"]:
            state["image"] = state["undo"].pop()
            state["points"] = []
            state["preview"] = None
            set_angle(0.0)
            render()

    def save(event=None):
        """손본 결과를 원본 폴더 옆 '_수정' 폴더에 저장한다(원본은 건드리지 않는다).

        위 '크기·형식 바꾸기' 에 적어 둔 값도 함께 적용한다.
        """
        if not paths or state["image"] is None:
            return
        commit_angle()          # 미리보던 기울기가 있으면 반영하고 저장한다
        src = paths[state["index"]]
        try:
            out = save_changed(state["image"], src, **size_args())
        except Exception as e:
            messagebox.showerror("이미지 수정", f"저장하지 못했습니다: {e}", parent=root)
            return
        state["dirty"] = False
        kb = os.path.getsize(out) / 1024
        folder = os.path.basename(os.path.dirname(out))
        log(f"[저장] {folder}/{os.path.basename(out)} ({kb:,.0f} KB)")
        flash(f"{folder} 폴더에 저장했습니다 ({kb:,.0f} KB)")
        render()

    def flash(text):
        """안내줄에 잠깐 메시지를 띄운다."""
        hint.config(text=text)
        root.after(1600, lambda: hint.config(text=HINT))

    def go(step):
        if not paths:
            return
        i = state["index"] + step
        if not 0 <= i < len(paths):
            flash("첫 파일입니다." if step < 0 else "마지막 파일입니다."
                  + ("  (파일이 한 장뿐입니다)" if len(paths) == 1 else ""))
            return
        if state["dirty"]:
            answer = messagebox.askyesnocancel(
                "이미지 수정", "바뀐 내용을 저장할까요?", parent=root)
            if answer is None:
                return
            if answer:
                save()
        load(i)

    def fit_view():
        state["zoom"], state["view"] = 1.0, (0.0, 0.0)
        render()

    def size_args():
        """입력칸에 적힌 숫자를 읽어 온다. 잘못 적었으면 0 으로 본다."""
        def num(var, default=0.0):
            try:
                return float(str(var.get()).replace(",", "").strip())
            except ValueError:
                return default

        return dict(fmt=fmt_var.get(), mode=size_mode.get(),
                    percent=num(percent_var, 100.0), kb=num(kb_var),
                    width=int(num(w_var)), height=int(num(h_var)))

    def export_all():
        """목록에 있는 모든 파일을 같은 설정으로 한꺼번에 내보낸다.

        창이 멈춘 채(‘응답 없음’) 기다리지 않도록 딴 실에서 돌린다. 예전에는
        이 일을 화면 그리는 실에서 같이 하는 바람에, 실제 작업은 2초면 끝나는데도
        창이 얼어붙어 30초씩 걸리는 것처럼 보였다.
        """
        if not paths:
            flash("먼저 이미지를 열어 주세요.")
            return
        if not messagebox.askyesno(
                "이미지 수정",
                f"목록의 {len(paths)}장을 같은 설정으로 내보낼까요?\n"
                "(원본 폴더 옆에 '_수정' 폴더가 생기고 그 안에 저장됩니다)",
                parent=root):
            return
        args = size_args()
        todo = list(paths)
        current = paths[state["index"]] if state["dirty"] else None
        edited = state["image"] if state["dirty"] else None
        export_btn.config(state="disabled")

        def post(fn, *a):
            root.after(0, lambda: fn(*a))

        def work():
            done, started = 0, time.time()
            for i, p in enumerate(todo, 1):
                try:
                    # 지금 화면에서 손보던 장만 그 이미지를 쓰고,
                    # 나머지는 save_changed 가 필요할 때만 읽게 맡긴다.
                    same = (p == current)
                    out = save_changed(edited if same else None, p,
                                       copy_ok=not same, **args)
                    done += 1
                    post(flash, f"내보내는 중... {i}/{len(todo)}  "
                                f"{os.path.basename(out)}")
                    post(log, f"[내보내기] {os.path.basename(out)}")
                except Exception as e:
                    post(log, f"[실패] {os.path.basename(p)}: {e}")
            post(flash, f"{done}장 내보냈습니다. ({time.time() - started:.1f}초)")
            post(export_btn.config, {"state": "normal"})

        threading.Thread(target=work, daemon=True).start()

    def add_files(new):
        """끌어다 놓은 파일/폴더를 목록 뒤에 붙인다(폴더는 안쪽 이미지를 모두)."""
        found = []
        for p in new:
            if os.path.isdir(p):
                for base, _, names in os.walk(p):
                    for n in sorted(names):
                        if os.path.splitext(n)[1].lower() in IMAGE_EXTS:
                            found.append(os.path.join(base, n))
            elif os.path.splitext(p)[1].lower() in IMAGE_EXTS:
                found.append(p)
        keys = {os.path.normcase(os.path.abspath(x)) for x in paths}
        added = [f for f in found
                 if os.path.normcase(os.path.abspath(f)) not in keys]
        if not added:
            flash("새로 추가할 이미지가 없습니다.")
            return
        first = len(paths)
        paths.extend(added)
        log(f"[이미지 수정] {len(added)}장을 목록에 넣었습니다.")
        if state["dirty"]:
            flash(f"{len(added)}장 추가했습니다. (지금 장을 저장한 뒤 넘어가세요)")
            render()
        else:
            load(first)

    def pick_more():
        add_files(filedialog.askopenfilenames(
            title="수정할 이미지 고르기",
            filetypes=[("이미지", "*.png *.jpg *.jpeg *.bmp *.webp *.tif *.tiff"),
                       ("모든 파일", "*.*")], parent=root))

    def pick_folder():
        folder = filedialog.askdirectory(title="수정할 이미지가 든 폴더", parent=root)
        if folder:
            add_files([folder])

    export_btn = ttk.Button(size, text="목록 전체 내보내기", command=export_all)
    export_btn.pack(side="right", padx=(4, 0))

    ttk.Button(foot, text="되돌리기 (Ctrl+Z)", command=undo).pack(side="left")
    ttk.Button(foot, text="화면 맞춤", command=fit_view).pack(side="left", padx=4)
    ttk.Button(foot, text="저장", command=save).pack(side="left", padx=(16, 4))
    ttk.Button(foot, text="저장하고 다음",
               command=lambda: (save(), go(1))).pack(side="left")
    ttk.Button(foot, text="이전", command=lambda: go(-1)).pack(side="left", padx=(16, 2))
    ttk.Button(foot, text="다음", command=lambda: go(1)).pack(side="left")
    ttk.Button(foot, text="닫기", command=root.destroy).pack(side="right")
    ttk.Button(foot, text="폴더 추가...", command=pick_folder).pack(side="right", padx=4)
    ttk.Button(foot, text="파일 추가...", command=pick_more).pack(side="right")

    # 이 창에도 파일/폴더를 바로 끌어다 놓을 수 있게 한다(tkinterdnd2 가 있을 때).
    try:
        from tkinterdnd2 import DND_FILES
        for widget in (root, canvas, foot):
            widget.drop_target_register(DND_FILES)
            widget.dnd_bind("<<Drop>>", lambda e: add_files(_split_drop(e.data)))
        HINT += "   |   파일·폴더를 이 창에 끌어다 놓으면 목록에 더해집니다"
        hint.config(text=HINT)
    except Exception:
        pass

    root.bind("<Control-z>", undo)
    root.bind("<Control-s>", save)
    root.bind("<Escape>", lambda e: reset_points())
    root.bind("<Prior>", lambda e: go(-1))
    root.bind("<Next>", lambda e: go(1))

    if paths:
        load(start)
    else:
        render()
    root.lift()
    root.focus_force()
    if parent is None:
        root.mainloop()   # 단독 실행일 때만 자체 루프를 돈다


# ---------------------------------------------------------------------------
# 식자재 지출현황 장부 만들기 (JSON 을 모아 엑셀 장부로)
#
# 원래 '3_지출현황_만들기.py' 에 있던 코드를 이 파일 안으로 옮겨 왔다.
# 앞 단계에서 만든 사진별 JSON 을 모아 '샘플_8월_식자재_지출현황.xlsx' 모양의
# 장부를 만들고, 사진이 있으면 '영수증 이미지' 시트도 같은 파일에 넣는다.
#
# 장부 반영 규칙(자세한 내용은 장부기재규칙.md):
#   - "장부반영": "제외(중복)" 인 영수증은 건너뛴다 (보류 전표, 합산 계산서 등)
#   - 품목의 "분류"가 식자재가 아니면(주방용품 등) 그 금액만큼 빼고 적는다
#   - "장부날짜"가 있으면 영수증 날짜 대신 그 날짜로 적는다 (결재본을 주문일로 기재)


FONT = "맑은 고딕"


MONEY_FMT = '"₩"#,##0_);[Red]\\("₩"#,##0\\)'


RED = "FFFF0000"                                                      # 총합계 행 글자색


ROW_H, TITLE_H = 20.15, 51.25


NON_FOOD = {"주방용품", "비식자재", "소모품"}


KEY_금액 = ("장부금액", "합계금액", "총금액", "총액", "금액", "결제금액")


KEY_날짜 = ("장부날짜", "영수증일시", "날짜", "거래일자")


KEY_내용 = ("내용", "메모", "비고요약", "적요")


KEY_거래처 = ("장부거래처", "거래처", "상호", "업체명")


def first(d: dict, keys, default=None):
    for k in keys:
        v = d.get(k)
        if v not in (None, ""):
            return v
    return default


def receipts_in(doc, path: Path):
    """JSON 한 개에서 영수증들을 꺼낸다. 세 가지 모양을 모두 받는다.

    ① {"영수증": [...]}          사진별 파일 안에 영수증 여러 장, 또는 한 달치 묶음
    ② [{...}, {...}]             영수증 목록만 있는 파일
    ③ {"거래처":..., "총금액":...}  파일 하나가 영수증 한 장
    """
    if isinstance(doc, list):
        items, parent = doc, {}
    elif isinstance(doc, dict) and isinstance(doc.get("영수증"), list):
        items, parent = doc["영수증"], doc
    elif isinstance(doc, dict) and (first(doc, KEY_금액) is not None
                                    or "_검산" in doc or "품목" in doc):
        items, parent = [doc], {}
    else:
        return []
    return [(r, parent) for r in items if isinstance(r, dict)]


def 요약(names: list[str], n: int = 5) -> str:
    """이름을 몇 개만 보여 주고 나머지는 개수로 적는다."""
    앞 = ", ".join(str(x) for x in names[:n])
    return 앞 + (f" 외 {len(names) - n}개" if len(names) > n else "")


def load_records(files: list[Path], base: Path | None = None) -> tuple[list[dict], list[str]]:
    """JSON 파일들을 읽어 장부에 넣을 행 목록과, 사람이 확인해야 할 경고 목록을 돌려준다.

    여러 폴더에서 모아 올 수 있으므로, 경고에는 고른 폴더 기준 상대경로로 적는다.
    """
    rows, warns, seen = [], [], {}
    판독폴더 = {}          # 원본 파일 이름 -> 그 영수증 판독본이 있던 폴더들

    def 표시(p: Path) -> str:
        try:
            return str(p.relative_to(base)) if base else p.name
        except ValueError:
            return p.name

    def 폴더표시(p: Path) -> str:
        try:
            return str(p.parent.relative_to(base)) if base else p.parent.name
        except ValueError:
            return p.parent.name

    for p in files:
        try:
            doc = json.loads(p.read_text(encoding="utf-8"))
        except (json.JSONDecodeError, UnicodeDecodeError) as e:
            warns.append(f"[읽기 실패] {표시(p)} - {e}")
            continue

        found = receipts_in(doc, p)
        if not found:
            warns.append(f"[모양 모름] {표시(p)} - 영수증을 찾지 못했습니다 (키: {list(doc)[:6]})")
            continue

        for r, parent in found:
            vendor = first(r, KEY_거래처) or first(parent, KEY_거래처) or "(거래처 미상)"
            total = first(r, KEY_금액)
            # '_원본' 은 이 프로그램이 JSON 마다 적어 두는 원본 사진 경로다. 이걸 이름으로
            # 삼아야 사진별 JSON 과 묶음(_모두.json)에서 같은 영수증이 같은 이름이 되어
            # 아래 중복 걸러내기에 걸린다. 이름이 어긋나면 장부에 두 번 오른다.
            원본파일 = str(r.get("파일명") or r.get("_원본")
                       or parent.get("파일명") or p.name)
            이름 = Path(원본파일).stem
            if r.get("사진내순번"):
                이름 = f"{이름}_{r['사진내순번']}"
            돈 = f"{int(total):,}원" if isinstance(total, (int, float)) else "금액 미상"
            # 같은 영수증을 두 번 판독해 두 폴더에 두면(사진 판독본 + PDF 판독본)
            # 내용이 조금만 달라도 아래 중복 걸러내기를 통과해 장부에 두 번 오른다
            판독폴더.setdefault(Path(원본파일).stem, set()).add(폴더표시(p))

            if r.get("장부반영") == "제외(중복)":
                warns.append(f"[제외] {이름} {vendor} {돈} - {r.get('중복사유') or r.get('검산','')}")
                continue
            if not isinstance(total, (int, float)):
                warns.append(f"[건너뜀] {이름} {vendor} - 금액을 찾지 못했습니다")
                continue

            # 묶음 파일과 사진별 파일이 섞여 있거나, 같은 판독본을 여러 폴더에 풀면
            # 같은 영수증이 두 번 잡힌다
            key = (이름, vendor, int(total))
            if key in seen:
                warns.append(f"[중복] {이름} {vendor} {돈} - {seen[key]} 에 이미 있어 건너뜁니다 ({표시(p)})")
                continue
            seen[key] = 표시(p)

            # 주방용품 등 식자재가 아닌 품목은 장부 금액에서 뺀다
            직접 = isinstance(r.get("장부금액"), (int, float))
            amount = int(total)
            빼기 = sum(i.get("금액") or 0 for i in (r.get("품목") or [])
                     if isinstance(i, dict) and i.get("분류") in NON_FOOD)
            if 빼기 and not 직접:
                amount -= 빼기
                warns.append(f"[차감] {이름} {vendor} {int(total):,} - 비식자재 {빼기:,} = {amount:,}원")

            일시 = first(r, KEY_날짜) or first(parent, KEY_날짜) or 이름[:8]
            날짜 = norm_date(일시)
            파일날짜 = stem_date(이름)
            # 사람이 장부날짜를 직접 적어 둔 영수증은 파일 이름보다 그 값을 따른다
            if 파일날짜 and 파일날짜 != 날짜 and not r.get("장부날짜"):
                warns.append(f"[날짜] {이름} {vendor} - 판독한 {날짜} 대신"
                             f" 파일 이름의 {파일날짜} 로 적습니다")
                날짜 = 파일날짜
            rows.append({
                "파일": 이름, "원본": 원본파일, "거래처": vendor,
                "내용": first(r, KEY_내용, ""), "금액": amount,
                "날짜": 날짜, "비고": r.get("장부비고") or "",
                "json폴더": p.parent,      # 같은 이름 사진이 여러 곳에 있을 때 쓸 단서
            })

            # 검산 결과는 예전 도구가 쓰던 문자열일 수도, 이 프로그램의 '_검산' 묶음일 수도 있다
            검산 = r.get("검산")
            if isinstance(r.get("_검산"), dict):
                v = r["_검산"]
                검산 = json_note(v) if v.get("확인필요") else ""
            검산 = str(검산 or "")
            if 검산 and not 검산.startswith("일치"):
                warns.append(f"[확인] {이름} {vendor} {돈} - {검산}")

    두벌 = sorted(k for k, v in 판독폴더.items() if len(v) > 1)
    if 두벌:
        곳 = sorted({d for k in 두벌 for d in 판독폴더[k]})
        warns.append(
            f"[판독본 두 벌] 같은 영수증의 판독본이 여러 폴더에 있습니다 {len(두벌)}건: {요약(두벌)}"
            f"\n      폴더: {' / '.join(곳)}"
            f"\n      금액이나 거래처 이름이 조금이라도 다르면 장부에 두 번 오릅니다."
            f" 한 벌만 남기고 지우세요.")

    rows.sort(key=lambda d: (d["날짜"], d["파일"]))
    return rows, warns


def find_json_files(picked: Path) -> list[Path]:
    """고른 폴더와 그 아래 모든 하위 폴더에서 JSON을 긁어모은다.

    AI에 사진을 여러 번 나눠 올리면 압축을 풀 때 폴더가 여러 개 생기므로
    한 폴더만 보지 않고 아래를 통째로 훑는다. 사진 폴더를 골라도 되고
    (보통 <사진폴더>\\json), 압축 푼 폴더들을 담아 둔 상위 폴더를 골라도 된다.
    같은 판독본이 두 폴더에 있으면 load_records 가 중복으로 걸러 준다.
    """
    files = [p for p in picked.rglob("*.json") if p.is_file()]
    # '_모두.json' 은 이 프로그램이 만든 묶음이라 사진별 JSON 과 내용이 그대로 겹친다.
    # 사진별 JSON 이 한 장이라도 있으면 묶음은 빼야 중복 안내로 화면이 덮이지 않는다.
    낱장 = [p for p in files if p.name != "_모두.json"]
    if 낱장:
        files = 낱장
    if not files:
        raise SystemExit(f"JSON이 없습니다: {picked}\n"
                         f"이 폴더나 그 아래 폴더(예: {picked.name}\\json)에 .json 파일이 있어야 합니다.")
    # 폴더가 달라도 같은 이름끼리 붙여 놓아야 중복 안내가 읽기 쉽다
    return sorted(files, key=lambda p: (p.name, str(p.parent)))


def receipt_files(folder: Path, recursive: bool = False) -> list[Path]:
    """폴더 안의 영수증 파일(사진 + PDF) 목록. PDF를 열지 않아 가볍다."""
    try:
        return xl_collect_files(folder, recursive)
    except OSError:
        return []


def pick_photo_folders(cands, rows: list[dict], recursive: bool = False) -> list[Path]:
    """장부에 오른 영수증이 들어 있는 폴더를 '전부' 사진 폴더로 삼는다.

    8월처럼 5~20일은 사진(jpg), 21~31일은 PDF로 폴더가 갈려 있는 달이 있다.
    전에는 제일 잘 맞는 폴더 하나만 써서, 이미지 시트에 PDF만(또는 사진만) 실리고
    나머지 반달치가 통째로 빠졌다. 그래서 맞는 폴더를 모두 모아 쓴다.
    이름이 하나도 안 맞으면 예전처럼 파일이 있는 첫 폴더 하나만 쓴다.
    """
    장부 = {Path(r["원본"]).stem for r in rows}
    맞는곳, 첫곳 = [], None
    for cand in cands:
        stems = {p.stem for p in receipt_files(cand, recursive)}
        if not stems:
            continue
        if stems & 장부:
            맞는곳.append(cand)
        elif 첫곳 is None:
            첫곳 = cand
    return 맞는곳 or ([첫곳] if 첫곳 else [])


def best_file(cands: list[Path], row: dict) -> Path:
    """같은 이름의 파일이 여러 폴더에 있으면 장부가 쓴 판독본 쪽 파일을 고른다."""
    원본 = Path(row["원본"]).name
    폴더 = row.get("json폴더")

    def 순위(p: Path):
        return (p.name != 원본,                                  # 이름까지 같은 것이 1순위
                p.suffix.lower() != Path(원본).suffix.lower(),   # 그다음 확장자가 같은 것
                p.parent != 폴더,                                # 그다음 판독본과 같은 폴더
                str(p))
    return min(cands, key=순위)


def 사본인가(a: Path, b: Path) -> bool:
    """이름이 같은 두 파일이 같은 영수증을 한 번 더 담아 둔 것인지 본다.

    확장자까지 같으면 다른 폴더에 복사해 둔 같은 파일, 한쪽만 PDF면 사진을 PDF로도
    받아 둔 것이라 한 장만 넣는다. jpg 와 png 처럼 사진끼리 다르면 다른 영수증일 수
    있어(판독 도구가 png 파일도 파일명을 jpg로 적어 준다) 둘 다 남긴다.
    """
    x, y = a.suffix.lower(), b.suffix.lower()
    return x == y or (x in XL_PDF_EXTS) != (y in XL_PDF_EXTS)


def gather_files(folders, rows: list[dict], keep_all: bool = False,
                 recursive: bool = False) -> tuple[list[Path], list[str]]:
    """사진 폴더들에서 영수증 파일을 모아 장부 순서대로 늘어놓는다. (파일, 뺀것)

    폴더를 훑어 이름별로 묶어 두고, 장부 한 줄마다 그 영수증의 원본을 집어 온다.
    사진 폴더와 PDF 폴더가 갈려 있어도 이미지 시트가 장부와 같은 날짜 순서로
    한 벌 만들어진다. 장부에 없는 영수증은 뒤에 붙이고 아래에서 따로 알려 준다.
    """
    묶음, 담은것 = {}, set()                    # 확장자 뺀 이름 -> 그 이름의 파일들
    for f in folders:
        for p in receipt_files(f, recursive):
            if p in 담은것:                      # 위 폴더까지 훑으면 같은 파일이 두 번 잡힌다
                continue
            담은것.add(p)
            묶음.setdefault(p.stem, []).append(p)
    장부 = {Path(r["원본"]).stem for r in rows}

    files, 본것, 뺀것 = [], set(), []
    for r in rows:                              # 장부는 이미 날짜 순으로 정렬돼 있다
        같은이름 = [p for p in 묶음.get(Path(r["원본"]).stem, []) if p not in 본것]
        if not 같은이름:
            continue
        고른것 = best_file(같은이름, r)
        본것.update(같은이름)
        files.append(고른것)
        for p in 같은이름:                       # 같은 이름으로 더 있는 파일
            if p is 고른것:
                continue
            if keep_all or not 사본인가(p, 고른것):
                files.append(p)
            else:
                뺀것.append(p.name)

    남은 = file_order([p for ps in 묶음.values() for p in ps
                         if p not in 본것 and p.stem not in 장부])
    쓴이름 = set()
    for p in 남은:                              # 장부에 없는 영수증 (판독을 안 한 것)
        if p.name.lower() in 쓴이름 and not keep_all:
            뺀것.append(p.name)                 # 여러 폴더에 같은 파일 -> 한 장만
            continue
        쓴이름.add(p.name.lower())
        files.append(p)
    return files, sorted(set(뺀것))


def check_coverage(items, rows: list[dict], folders) -> list[str]:
    """장부와 이미지 시트가 서로 맞는지 대조하고, 빠진 것을 알려 준다."""
    if not folders:               # 사진 폴더가 아예 없다는 안내는 이미 했다
        return []
    warns = []
    장부 = {Path(r["원본"]).stem for r in rows}
    실린것 = {it.src.path.stem for it in items}

    없음 = sorted(장부 - 실린것)
    if 없음:
        warns.append(f"[이미지 없음] 장부에는 있는데 사진/PDF를 못 찾은 영수증 {len(없음)}건: {요약(없음)}")
        # 못 찾은 것이 하위 폴더(예: pdf 폴더)에 들어 있으면 넣는 방법을 알려 준다
        깊은곳 = {p.stem for f in folders for p in receipt_files(f, True)} & set(없음)
        if 깊은곳:
            warns.append(f"      그중 {len(깊은곳)}건은 하위 폴더에 있습니다"
                         f" - -r (--하위폴더) 를 붙이면 함께 들어갑니다: {요약(sorted(깊은곳))}")
    남음 = sorted(실린것 - 장부)
    if 남음:
        warns.append(f"[장부 없음] 이미지 시트에는 있는데 장부에 없는 영수증 {len(남음)}건: {요약(남음)}"
                     f"\n      판독(2단계)을 안 했거나 JSON을 안 풀어 둔 영수증입니다.")

    # 고른 폴더가 한 묶음(예: pdf 폴더)이면 위층에 남은 영수증을 알려 준다
    for 상위 in dict.fromkeys(f.parent for f in folders):
        if 상위 in folders or 상위 == 상위.parent:
            continue
        위 = {p.stem for p in receipt_files(상위)} - 실린것 - 장부
        if 위:
            warns.append(
                f"[위 폴더에 더 있음] '{상위.name}' 폴더에 이 파일에 안 들어간 "
                f"영수증 {len(위)}개가 더 있습니다: {요약(sorted(위))}"
                f"\n      한 달치를 한 파일로 만들려면 위 폴더('{상위.name}')를 골라 실행하세요.")
    return warns


def stem_date(stem: str) -> str:
    """정리해 둔 파일 이름 앞머리의 날짜를 읽는다. 못 읽으면 빈 문자열.

    1단계가 이름을 'YY.MM.DD' 로 통일해 두므로 '26.08.29 호반떡방앗간_정리' 는
    2026-08-29 가 된다. 손글씨 영수증은 연도를 잘못 읽는 일이 잦은데(2026 을 2006
    으로), 파일 이름 쪽은 사람이 확인한 값이라 더 믿을 만하다.
    """
    m = re.match(r"\s*(20\d{2})[.\-_ ]?(\d{2})[.\-_ ]?(\d{2})", stem)
    if m:
        y, mo, d = (int(x) for x in m.groups())
    else:
        m = re.match(r"\s*(\d{2})\.(\d{2})\.(\d{2})", stem)
        if not m:
            return ""
        y, mo, d = 2000 + int(m.group(1)), int(m.group(2)), int(m.group(3))
    return f"{y:04d}-{mo:02d}-{d:02d}" if 1 <= mo <= 12 and 1 <= d <= 31 else ""


def norm_date(v) -> str:
    """'2026-07-31', '2026.07.31', '20260731' 을 모두 2026-07-31 로 맞춘다."""
    s = str(v)[:19].strip()
    digits = "".join(ch for ch in s if ch.isdigit())
    if len(digits) >= 8:
        return f"{digits[:4]}-{digits[4:6]}-{digits[6:8]}"
    return s[:10]


def vendor_key(name: str) -> str:
    """'(주)남한강마트 충주점' 과 '남한강마트' 를 같은 것으로 보기 위한 비교용 이름."""
    s = re.sub(r"\([^)]*\)", "", str(name))     # (주), (이재혁) 처럼 괄호 안은 버린다
    for junk in ("주식회사", "㈜"):
        s = s.replace(junk, "")
    s = "".join(ch for ch in s if not ch.isspace())
    for tail in ("충주제천점", "충주점", "제천점", "본점", "지점"):
        if s.endswith(tail) and len(s) > len(tail) + 1:
            s = s[: -len(tail)]
    return s


def vendor_order(rows: list[dict], sample: Path | None) -> tuple[list[str], list[str]]:
    """거래처 열 순서를 정하고, 참고 파일의 이름으로 통일한다.

    참고 파일(지난달 장부)이 있으면 그 열 순서를 그대로 잇고,
    '(주)남한강마트 충주점' 처럼 길게 적힌 이름은 장부 이름 '남한강마트' 로 바꾼다.
    """
    order, canon, notes = OrderedDict(), {}, []
    if sample and sample.exists():
        from openpyxl import load_workbook

        ws = load_workbook(sample).active
        for cell in ws[2]:
            v = cell.value
            if v and v not in ("날짜/구분", "내용", "합계", "비고"):
                name = str(v).strip()
                order[name] = None
                canon[vendor_key(name)] = name

    장부이름 = dict(canon)              # 참고 장부에 있던 이름만 (아래에서 늘어나기 전에 보관)
    for r in rows:                      # 장부 이름으로 맞추고, 없던 거래처는 뒤에 붙인다
        표준 = match_vendor(vendor_key(r["거래처"]), 장부이름)
        if 표준:
            if 표준 != r["거래처"]:
                notes.append(f"[이름 통일] {r['거래처']} -> {표준}")
            r["거래처"] = 표준
        else:
            if 장부이름:
                notes.append(f"[새 거래처] {r['거래처']} - 참고 장부에 없어 열을 새로 만듭니다")
            k = vendor_key(r["거래처"])
            canon.setdefault(k, r["거래처"])
            r["거래처"] = canon[k]
        order.setdefault(r["거래처"], None)

    used = {r["거래처"] for r in rows}
    return [v for v in order if v in used], sorted(set(notes))


def match_vendor(key: str, canon: dict) -> str | None:
    """비교용 이름으로 장부 거래처를 찾는다. 딱 하나만 걸릴 때만 인정한다."""
    if key in canon:
        return canon[key]
    hits = {c for c in canon if len(c) >= 2 and len(key) >= 2 and (c in key or key in c)}
    return canon[hits.pop()] if len(hits) == 1 else None


def ledger_styles():
    """장부 시트에 쓰는 서식. (A_샘플.xlsx 에서 그대로 가져왔다)"""
    from openpyxl.styles import Alignment, Border, Color, PatternFill, Side

    thin = Side(style="thin")
    return {
        "머리칠": PatternFill("solid", fgColor=Color(theme=7, tint=0.6)),   # 골드 60% 밝게
        "합계칠": PatternFill("solid", fgColor=Color(theme=6, tint=0.8)),   # 회색 80% 밝게
        "테두리": Border(left=thin, right=thin, top=thin, bottom=thin),
        "가운데": Alignment(horizontal="center", vertical="center"),
    }


def build_ledger_sheet(wb, rows: list[dict], vendors: list[str], title: str,
                       prev_total: int | None, sheet_title: str = "영수증 내역"):
    from openpyxl.styles import Alignment, Font
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.properties import PageSetupProperties

    st = ledger_styles()
    HEAD_FILL, SUM_FILL = st["머리칠"], st["합계칠"]
    BOX, CENTER = st["테두리"], st["가운데"]

    ws = wb.create_sheet(sheet_title)
    n_v = len(vendors)
    first_v, last_v = 4, 3 + n_v            # D열부터 거래처
    sum_col = last_v + 1                    # 그 다음이 합계 열
    last_col = sum_col

    # 1행: 제목
    ws.merge_cells(start_row=1, start_column=2, end_row=1, end_column=last_col)
    t = ws.cell(row=1, column=2, value=title)
    t.font = Font(name=FONT, size=24, bold=True)
    t.alignment = CENTER
    ws.row_dimensions[1].height = TITLE_H

    # 2행: 머리글. 거래처명이 길어도 열을 넓히지 않고 두 줄로 접는다(한 페이지에 들어가야 하므로)
    heads = [None, "날짜/구분", "내용"] + vendors + ["합계"]
    ws.row_dimensions[2].height = ROW_H * 1.7
    wrap = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for i, h in enumerate(heads, start=1):
        c = ws.cell(row=2, column=i)
        if h is None:
            continue
        c.value = h
        c.data_type = "s"
        c.font = Font(name=FONT, size=12)
        c.alignment, c.border, c.fill = wrap, BOX, HEAD_FILL

    # 3행부터: 영수증 한 건이 한 행
    col_of = {v: first_v + i for i, v in enumerate(vendors)}
    r = 3
    for row in rows:
        ws.row_dimensions[r].height = ROW_H
        if row["비고"]:
            a = ws.cell(row=r, column=1, value=row["비고"])
            a.data_type = "s"
            a.font = Font(name=FONT, size=9)
            a.alignment = Alignment(vertical="center")

        d = ws.cell(row=r, column=2, value=fmt_date(row["날짜"]))
        d.data_type = "s"
        d.font = Font(name=FONT, size=12)

        c = ws.cell(row=r, column=3, value=row["내용"])
        c.data_type = "s"
        c.font = Font(name=FONT, size=11)

        for col in range(2, last_col + 1):
            cell = ws.cell(row=r, column=col)
            cell.alignment, cell.border = CENTER, BOX

        m = ws.cell(row=r, column=col_of[row["거래처"]], value=row["금액"])
        m.number_format = MONEY_FMT
        m.font = Font(name=FONT, size=12)
        r += 1

    data_end = r - 1

    # 총합계 행
    ws.row_dimensions[r].height = ROW_H
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
    lab = ws.cell(row=r, column=2, value="마트종류별 총합계")
    lab.data_type = "s"
    lab.font = Font(name=FONT, size=11, color=RED)
    for col in range(2, last_col + 1):
        cell = ws.cell(row=r, column=col)
        cell.alignment, cell.border, cell.fill = CENTER, BOX, SUM_FILL
    for v, col in col_of.items():
        L = get_column_letter(col)
        c = ws.cell(row=r, column=col, value=f"=SUM({L}3:{L}{data_end})")
        c.number_format, c.font = MONEY_FMT, Font(name=FONT, size=11)
    L1, L2 = get_column_letter(first_v), get_column_letter(last_v)
    g = ws.cell(row=r, column=sum_col, value=f"=SUM({L1}{r}:{L2}{r})")
    g.number_format, g.font = MONEY_FMT, Font(name=FONT, size=11)
    total_row = r
    r += 1

    # 전월/당월 총합계. A_샘플.xlsx 처럼 '26년06월 총합계' 로 적는다.
    # 전월 금액을 안 줘도 행은 만든다 (사용자가 나중에 손으로 채워 넣을 자리).
    ym = period_of(rows)
    prev_label = ym_label(*prev_month(*ym)) if ym else "전월 총합계"
    cur_label = ym_label(*ym) if ym else "당월 총합계"
    for label, value in ((prev_label, prev_total),
                         (cur_label, f"={get_column_letter(sum_col)}{total_row}")):
        ws.row_dimensions[r].height = ROW_H
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        lab = ws.cell(row=r, column=2, value=label)
        lab.data_type = "s"
        lab.font = Font(name=FONT, size=11, color=RED)
        lab.fill = SUM_FILL          # 병합된 B:C 는 왼쪽 위 칸 서식으로 칠해진다
        lab.alignment = CENTER
        for col in (2, 3, first_v):
            ws.cell(row=r, column=col).border = BOX
        c = ws.cell(row=r, column=first_v, value=value)
        c.number_format, c.font, c.alignment = MONEY_FMT, Font(name=FONT, size=11), CENTER
        r += 1

    # 열 너비 / 인쇄 설정
    ws.column_dimensions["A"].width = 11.9
    ws.column_dimensions["B"].width = 12.9
    ws.column_dimensions["C"].width = 22.2
    for v, col in col_of.items():
        # 한글은 한 글자가 대략 2칸. 다만 18을 넘기지 않고 머리글 줄바꿈에 맡긴다.
        ws.column_dimensions[get_column_letter(col)].width = min(18.0, max(13.1, len(v) * 2.0))
    ws.column_dimensions[get_column_letter(sum_col)].width = 13.1

    ws.freeze_panes = "D3"
    ws.print_area = f"A1:{get_column_letter(last_col)}{r - 1}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    # 이미지 시트와 같은 제본을 하므로 상단 30mm(펀치 구멍)와 왼쪽 여백을 맞춘다
    ws.page_margins.left, ws.page_margins.right = 1.0, 0.75
    ws.page_margins.top, ws.page_margins.bottom = 30 / 25.4, 0.75
    ws.page_margins.header = ws.page_margins.footer = 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_title_rows = "1:2"
    return ws


def fmt_date(s: str) -> str:
    try:
        y, m, d = (int(x) for x in s[:10].split("-"))
        return f"{m}월 {d}일"
    except (ValueError, AttributeError):
        return s or ""


def period_of(rows: list[dict]) -> tuple[int, int] | None:
    """장부의 연·월을 영수증 날짜에서 뽑는다."""
    for r in rows:
        try:
            y, m, _ = (int(x) for x in r["날짜"][:10].split("-"))
            return y, m
        except ValueError:
            continue
    return None


def prev_month(y: int, m: int) -> tuple[int, int]:
    return (y - 1, 12) if m == 1 else (y, m - 1)


def ym_label(y: int, m: int) -> str:
    """A_샘플.xlsx 의 총합계 라벨 형식: 2026년 7월 -> '26년07월 총합계'."""
    return f"{y % 100:02d}년{m:02d}월 총합계"


def guess_title(rows: list[dict]) -> str:
    ym = period_of(rows)
    return f"{ym[0]}년 {ym[1]}월 식자재 지출현황" if ym else "식자재 지출현황"


def make_ledger(json_dir, out_path=None, photos=None, recursive=False, keep_dup=False,
                title=None, sample=None, prev_total=None, gap_mm=None, log=print):
    """JSON 폴더를 모아 '식자재 지출현황' 엑셀 장부를 만든다. 저장한 경로를 돌려준다.

    고른 폴더 아래 하위 폴더까지 전부 훑으므로, 판독을 여러 번 나눠 받아 폴더가
    여러 개여도 그 폴더들을 담은 상위 폴더 하나만 고르면 된다.
    사진이 있으면 '영수증 이미지' 시트도 같은 파일에 넣어 한 파일로 완성한다.
    """
    from openpyxl import Workbook

    gap_mm = DEFAULT_GAP_MM if gap_mm is None else gap_mm
    picked = Path(json_dir).expanduser().resolve()
    if not picked.is_dir():
        raise NotADirectoryError(f"폴더가 아닙니다: {picked}")

    files = find_json_files(picked)      # 고른 폴더 아래 하위 폴더까지 전부 훑는다
    json_dirs = sorted({p.parent for p in files})
    if json_dirs != [picked]:
        log(f"JSON 폴더 {len(json_dirs)}곳 / 파일 {len(files)}개")
        for d in json_dirs:
            n = sum(1 for p in files if p.parent == d)
            log(f"  {d.relative_to(picked) if d != picked else Path('.')}  ({n}개)")

    rows, warns = load_records(files, picked)
    if not rows:                        # 왜 못 찾았는지 알려줘야 고칠 수 있다
        for w in warns:
            log("  " + w)
        raise FileNotFoundError(
            f"장부에 넣을 영수증이 없습니다: {picked}\n"
            "JSON 안에 거래처와 금액이 있어야 합니다. "
            "금액 키는 " + " / ".join(KEY_금액) + " 중 하나여야 합니다.")

    # 사진 폴더를 안 주면 고른 폴더 -> JSON 폴더들 -> 그 위 순으로 후보를 훑는다.
    # 장부를 먼저 읽어 두고, 장부에 오른 영수증이 있는 폴더를 모두 사진 폴더로 쓴다.
    if photos:
        photo_dirs = [Path(photos).expanduser().resolve()]
    else:
        cands = dict.fromkeys([picked, *json_dirs, *(d.parent for d in json_dirs)])
        photo_dirs = pick_photo_folders(cands, rows, recursive)
        if photo_dirs:
            log(f"사진 폴더 자동 인식 {len(photo_dirs)}곳:")
            for d in photo_dirs:
                log(f"  {d.relative_to(picked) if d != picked and picked in d.parents else d}")

    vendors, 이름메모 = vendor_order(rows, Path(sample).resolve() if sample else None)
    warns += 이름메모
    if prev_total is None:
        warns.append("[전월 총합계] 행은 만들었지만 금액은 비어 있습니다 "
                     "- 엑셀에서 직접 넣거나 전월합계를 주세요")
    title = title or guess_title(rows)
    # 엑셀은 고른 폴더에 놓는다. 다만 json 폴더를 콕 집어 골랐으면 사진이 있는 그 위로.
    base = picked.parent if picked.name.lower() == "json" else picked
    out = Path(out_path) if out_path else base / f"{title.replace(' ', '_')}.xlsx"
    if not out.is_absolute():
        out = base / out

    log(f"영수증 {len(rows)}건 / 거래처 {len(vendors)}곳 -> {title}")

    with tempfile.TemporaryDirectory() as td:
        wb = Workbook()
        wb.remove(wb.active)

        items = []
        pics, 뺀것 = gather_files(photo_dirs, rows, keep_dup, recursive)
        if pics:                            # ① 영수증 이미지 시트
            if 뺀것:
                log(f"  같은 영수증의 사진/PDF 사본 {len(뺀것)}장은 뺐습니다"
                    f" (장부에 쓴 판독본만 넣습니다): {요약(뺀것)}")
            sources = sources_from_files(pics)
            n_img = sum(1 for s in sources if s.page is None)
            n_pdf = len({s.path for s in sources if s.page is not None})
            log(f"  사진 {n_img}장 + PDF {n_pdf}개 읽는 중...")
            pages, gap_px = plan_sources(sources, "auto", False, gap_mm)
            add_image_sheet(wb, pages, Path(td), gap_px)
            items = [it for page in pages for band in page for it in band]
            log(f"  이미지 시트: {len(items)}칸(사진·PDF 쪽), {len(pages)}페이지")
        else:
            warns.append("[이미지 없음] 사진/PDF 폴더를 못 찾아 '영수증 이미지' 시트를"
                         " 못 넣었습니다. 사진 폴더를 직접 지정해 주세요.")
        warns += check_coverage(items, rows, photo_dirs)

        build_ledger_sheet(wb, rows, vendors, title, prev_total)   # ② 영수증 내역 시트

        try:
            wb.save(out)
        except PermissionError:
            out = out.with_name(f"{out.stem}_새파일{out.suffix}")
            log(f"! 파일이 열려 있어 {out.name} 로 저장합니다.")
            wb.save(out)

    if warns:
        log("\n확인이 필요한 항목:")
        for w in warns:
            log("  " + w)
    log(f"[장부] {out}")
    return str(out)


# ---------------------------------------------------------------------------
# 영수증 -> JSON 추출 (Gemini)
#
# 정리해 둔 이미지를 Gemini 로 읽어 파일마다 JSON 하나를 만든다.
#   - 결과는 <폴더>/json/<이름>.json, 묶음 _모두.json 과 요약 _요약.csv
#   - 합계 검산: 품목 금액 합 != 총액 이면 "검산불일치"
#   - 검산에 걸린 장만 좋은 모델(Pro)로 자동 재판독한다
#
# 처음 한 번만:
#     pip install google-genai
#     setx GEMINI_API_KEY "받은키"      (새 명령창을 열어야 적용된다)

# exe 로 묶어 창 모드로 실행하면 콘솔이 없어 sys.stdout 이 None 이 된다.
# 그 상태에서 print() 를 부르면 프로그램이 죽는다. 삼킬 곳을 만들어 둔다.
class _Nowhere:
    def write(self, *_):
        return 0

    def flush(self):
        pass


if sys.stdout is None:
    sys.stdout = _Nowhere()
if sys.stderr is None:
    sys.stderr = _Nowhere()


# SDK 가 매번 찍는 AFC 안내문은 이 용도에선 소음이라 끈다
logging.getLogger("google_genai").setLevel(logging.ERROR)
logging.getLogger("google_genai.models").setLevel(logging.ERROR)

JSON_MODEL = os.environ.get("GEMINI_MODEL", "gemini-2.5-flash")
JSON_PRO_MODEL = os.environ.get("GEMINI_PRO_MODEL", "gemini-3.1-pro-preview")
JSON_MAX_SIDE = 1600      # 긴 변을 이만큼으로 줄여 보낸다 (영수증 글씨는 충분히 읽힌다)
JSON_RECHECK_SIDE = 2400  # 재판독은 손글씨 숫자가 뭉개지지 않게 더 크게 보낸다
JSON_RETRY = 3
JSON_RETRY_WAIT = 6       # 초. 실패할 때마다 배로 늘어난다

JSON_SCHEMA = {
    "type": "object",
    "properties": {
        "상호": {"type": "string", "nullable": True},
        "사업자등록번호": {"type": "string", "nullable": True},
        "거래일자": {"type": "string", "nullable": True, "description": "YYYY-MM-DD"},
        "거래시각": {"type": "string", "nullable": True, "description": "HH:MM"},
        "결제수단": {"type": "string", "nullable": True, "description": "현금/카드/계좌이체 등"},
        "카드번호뒤4자리": {"type": "string", "nullable": True},
        "승인번호": {"type": "string", "nullable": True},
        "공급가액": {"type": "integer", "nullable": True},
        "부가세": {"type": "integer", "nullable": True},
        "총액": {"type": "integer", "nullable": True},
        "품목": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "이름": {"type": "string"},
                    "수량": {"type": "number", "nullable": True},
                    "단가": {"type": "integer", "nullable": True},
                    "금액": {"type": "integer", "nullable": True},
                },
                "required": ["이름"],
            },
        },
        "판독주의": {
            "type": "array",
            "description": "흐릿하거나 확신이 없어 비워 둔 항목 이름들",
            "items": {"type": "string"},
        },
        "비고": {"type": "string", "nullable": True},
        "영수증장수": {"type": "integer", "nullable": True,
                  "description": "이미지에 서로 다른 영수증이 몇 장 보이는지. 보통 1"},
    },
    "required": ["총액", "품목", "판독주의"],
}

JSON_PROMPT = """당신은 한국 영수증을 읽어 회계 장부에 옮기는 사람입니다.
주어진 영수증 사진에서 아래 규칙대로 값을 뽑아 JSON 으로만 답하세요.

규칙
1. 금액은 쉼표와 '원' 을 빼고 정수로만 적습니다. (12,300원 -> 12300)
2. 날짜는 YYYY-MM-DD 로 통일합니다. 연도가 없으면 비고에 적고 거래일자는 null 로 둡니다.
3. 글자가 흐릿하거나 두 가지로 읽힐 수 있으면 절대 추측하지 말고 그 값을 null 로 두고,
   그 항목 이름을 "판독주의" 배열에 넣습니다. 추측한 숫자가 빈칸보다 훨씬 나쁩니다.
4. 품목은 영수증에 적힌 순서 그대로, 적힌 이름 그대로 옮깁니다. 이름을 고치지 않습니다.
5. 할인/포인트사용 같은 마이너스 줄도 품목에 음수 금액으로 넣습니다.
6. 영수증이 아니거나 아무것도 읽을 수 없으면 총액을 null 로 두고 비고에 이유를 적습니다.
7. 소계, 합계, 총액, 공급가액, 부가세, 받은금액, 거스름돈 같은 정산 줄은 품목이 아닙니다.
   품목 배열에 넣지 마세요. 정산 줄까지 품목에 넣으면 합계가 두 배가 되어 맞지 않습니다.
8. 한 이미지에 서로 다른 영수증이 둘 이상 보이면(나란히 놓인 경우 등) '영수증장수'에 그 수를 적으세요.
   보통은 1입니다. 두 영수증의 금액을 합치거나 섞어서 읽지 마세요.
9. 손글씨 숫자는 수량×단가=공급가액, 품목 금액의 합=합계 관계로 서로 맞는지 확인해, 앞뒤가 맞는 값을
   고르세요. 그래도 확인이 안 되고 불확실하면 추측하지 말고 null 로 두세요.
10. 합계는 표 위쪽 '합계금액' 칸, 표 아래쪽 '합계' 칸, 카드 승인금액이 있으면 서로 대조해 읽으세요.
11. '판독주의' 에는 항목 이름만 짧게 적으세요. (예: 거래일자, 총액, 품목[2].이름)
    이유, 계산 과정, 설명, 영어 문장은 쓰지 마세요.
"""


class JsonSetupError(Exception):
    """키가 없다거나 패키지가 없다거나 — 시작조차 할 수 없는 상태.

    GUI 에서도 쓰이므로 sys.exit 대신 예외로 올려 보내고, 부르는 쪽에서 안내를 찍는다.
    """


class JsonFatalError(Exception):
    """키가 틀렸다거나 모델 이름이 없다거나 — 다시 시도해도 소용없는 오류.
    이런 건 남은 사진을 계속 돌리지 말고 바로 멈춰야 한다."""


def json_short_err(e):
    """구글이 돌려주는 긴 JSON 오류에서 사람이 읽을 한 줄만 뽑는다."""
    s = str(e)
    m = re.search(r"'message':\s*'([^']+)'", s) or re.search(r'"message":\s*"([^"]+)"', s)
    msg = m.group(1) if m else s
    return msg[:160].replace("\n", " ")


def json_is_fatal(e):
    s = str(e)
    return any(t in s for t in
               ("API_KEY_INVALID", "PERMISSION_DENIED", "API key not valid",
                "NOT_FOUND", "is not found for API version", "billing"))


def json_client():
    key = os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY") or ""
    # 복사/붙여넣기 때 딸려 오는 따옴표·공백·밑줄을 떼어 낸다 (아주 흔한 사고다)
    key = key.strip().strip("\"'").strip().strip("_")
    if not key:
        raise JsonSetupError(
            "GEMINI_API_KEY 가 없습니다.\n"
            "  1) https://aistudio.google.com/apikey 에서 키를 만들고\n"
            '  2) 명령창에  setx GEMINI_API_KEY "받은키"  를 실행한 뒤\n'
            "  3) 명령창(또는 이 프로그램)을 새로 열어 다시 실행하세요."
        )
    try:
        from google import genai
    except ImportError:
        raise JsonSetupError("google-genai 가 없습니다.  pip install google-genai")
    return genai.Client(api_key=key)


JSON_FALLBACK_MODELS = [m.strip() for m in os.environ.get("GEMINI_FALLBACK_MODELS", "").split(",")
                        if m.strip()]


def json_model_gone(e):
    """모델 이름이 없어졌거나 더 이상 지원하지 않는 경우 (키·결제 문제와 구분한다)."""
    s = str(e)
    return any(t in s for t in ("NOT_FOUND", "is not found for API version",
                                "is not supported for generateContent"))


def json_discover_models(client, kind):
    """지금 이 키로 쓸 수 있는 gemini-숫자-flash(또는 -pro) 모델을 최신 순으로 돌려준다.

    preview/lite/image 등 이름에 덧붙는 것은 일부러 뺀다. 이름이 정확히 맞는 정식 모델만 본다.
    """
    try:
        listed = list(client.models.list())
    except Exception:
        return []
    pat = re.compile(r"^(?:models/)?gemini-(\d+(?:\.\d+)?)-%s$" % re.escape(kind))
    found = []
    for m in listed:
        name = getattr(m, "name", "") or ""
        hit = pat.match(name)
        if not hit:
            continue
        actions = getattr(m, "supported_actions", None)
        if actions and "generateContent" not in actions:
            continue
        found.append((float(hit.group(1)), name.replace("models/", "")))
    found.sort(reverse=True)
    return [n for _, n in found]


def json_resolve_model(client, model, log=print):
    """사진을 보내기 전에 키와 모델이 살아 있는지 확인하고, 쓸 모델 이름을 돌려준다.

    모델이 없어졌으면 GEMINI_FALLBACK_MODELS, 그다음 이 키로 쓸 수 있는 최신 정식 모델,
    마지막으로 gemini-flash-latest 같은 별칭 순으로 대신 쓸 모델을 찾는다.
    키가 틀렸거나 결제 문제면 모델을 바꿔도 소용없으므로 바로 멈춘다.
    """
    def ping(name):
        client.models.generate_content(model=name, contents="ping")

    try:
        ping(model)
        return model
    except Exception as e:
        if json_model_gone(e):
            reason = json_short_err(e)
        elif json_is_fatal(e):
            raise JsonSetupError(
                f"API 키 또는 모델을 쓸 수 없습니다: {json_short_err(e)}\n"
                "  - 키가 맞는지 (aistudio.google.com/apikey 에서 복사한 값 그대로인지)\n"
                f"  - 모델 이름({model})이 맞는지  확인하세요."
            )
        else:
            log(f"  ! 사전 확인을 건너뜁니다 ({json_short_err(e)})")
            return model

    kind = "pro" if "pro" in model.lower() else "flash"
    candidates, seen = [], {model}
    for name in (JSON_FALLBACK_MODELS + json_discover_models(client, kind)
                 + [f"gemini-{kind}-latest"]):
        if name not in seen:
            seen.add(name)
            candidates.append(name)
    for name in candidates[:6]:
        try:
            ping(name)
        except Exception as e2:
            if json_is_fatal(e2) and not json_model_gone(e2):
                raise JsonSetupError(f"API 키 또는 모델을 쓸 수 없습니다: {json_short_err(e2)}")
            continue
        log(f"[모델 대체] {model} 을(를) 쓸 수 없어({reason}) {name} 로 대신 읽습니다.")
        return name
    raise JsonSetupError(
        f"모델 {model} 을(를) 쓸 수 없고, 대신 쓸 모델도 찾지 못했습니다: {reason}\n"
        "  - GEMINI_MODEL / GEMINI_FALLBACK_MODELS 환경변수를 확인하세요."
    )


def json_preflight(client, model, log=print):
    """json_resolve_model 의 예전 이름. 쓸 모델 이름을 돌려준다."""
    return json_resolve_model(client, model, log=log)

def json_page_png(img, max_side=None):
    """정리 단계에서 쓰는 BGR ndarray 한 장을 Gemini 에 보낼 PNG 바이트로 만든다."""
    rgb = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
    pil = Image.fromarray(rgb)
    side = max_side or JSON_MAX_SIDE
    if max(pil.size) > side:
        r = side / max(pil.size)
        pil = pil.resize((max(1, int(pil.width * r)), max(1, int(pil.height * r))),
                         Image.LANCZOS)
    buf = io.BytesIO()
    pil.save(buf, "PNG", optimize=True)
    return buf.getvalue()


def json_ask(client, png, model, schema=None, prompt=None):
    from google.genai import types

    cfg = types.GenerateContentConfig(
        response_mime_type="application/json",
        response_schema=schema or JSON_SCHEMA,
        temperature=0,
    )
    last = None
    for attempt in range(1, JSON_RETRY + 1):
        try:
            resp = client.models.generate_content(
                model=model,
                contents=[types.Part.from_bytes(data=png, mime_type="image/png"),
                          prompt or JSON_PROMPT],
                config=cfg,
            )
            return json.loads(resp.text)
        except Exception as e:                       # 통신 오류/할당량/JSON 깨짐 모두
            if json_is_fatal(e):                     # 재시도해도 똑같다. 바로 올려 보낸다
                raise JsonFatalError(json_short_err(e)) from e
            last = e
            if attempt < JSON_RETRY:
                time.sleep(JSON_RETRY_WAIT * (2 ** (attempt - 1)))
    raise RuntimeError(f"{JSON_RETRY}번 모두 실패: {json_short_err(last)}")


def json_clean_cautions(items):
    """판독주의 항목을 짧은 항목 이름만 남기고 정리한다.

    일부 모델이 이유나 계산 과정, 영어 문장을 그대로 적는 일이 있어서 걸러 낸다.
    """
    out = []
    for s in items or []:
        s = str(s).strip()
        s = re.split(r"\s*[(\uff08:\uff1a]|\s+-\s+|\s*\.\s|Notice|[Ww]ait|[Ll]et's", s)[0].strip(" ,;.")
        if not s:
            continue
        if not re.search(r"[\uac00-\ud7a3]", s) and len(s) > 12:   # 한글 없이 긴 영어 문장은 버린다
            continue
        if len(s) > 40:
            s = s[:40].rstrip(" ,;") + "…"
        if s not in out:
            out.append(s)
    return out[:12]


def json_repair(d):
    """모델이 비워 둔 값 중 영수증 안의 다른 숫자로 계산되는 것만 채운다. 채운 곳은 '_보정'에 남긴다."""
    fixes = []
    items = d.get("품목") or []
    for i, it in enumerate(items):
        if not isinstance(it, dict):
            continue
        q, u = it.get("수량"), it.get("단가")
        if it.get("금액") is None and isinstance(q, (int, float)) and isinstance(u, (int, float)) \
                and q > 0 and u > 0:
            it["금액"] = int(round(q * u))
            fixes.append(f"품목[{i}].금액")
    amounts = [it.get("금액") for it in items if isinstance(it, dict)]
    if d.get("총액") is None:
        sup, vat = d.get("공급가액"), d.get("부가세")
        if isinstance(sup, int) and isinstance(vat, int) and sup + vat > 0:
            d["총액"] = sup + vat
            fixes.append("총액")
        elif items and amounts and all(isinstance(a, (int, float)) for a in amounts):
            d["총액"] = int(sum(amounts))
            fixes.append("총액")
    if fixes:
        d["_보정"] = fixes
    return d


def json_verify(data):
    """품목 합과 총액을 맞춰 보고 결과를 돌려준다.

    '확인필요' 는 정말 금액을 믿기 어려울 때만 켠다: 총액이 없거나, 품목 합과 총액이 다르거나,
    총액을 계산으로 채웠거나, 모델이 총액을 불확실하다고 적은 경우다. 날짜나 카드번호처럼
    장부 금액에 영향이 없는 불확실은 '판독주의' 목록에만 남기고 화면에는 참고로 보여 준다.
    """
    kids = data.get("영수증")
    if isinstance(kids, list) and kids:          # 사진 한 장에 영수증 여러 장
        vs = []
        for k in kids:
            if isinstance(k, dict):
                k["_검산"] = json_verify(k)
                vs.append(k["_검산"])
        bad = [v for v in vs if v.get("상태") != "일치"]
        cautions = []
        for v in vs:
            for c in v.get("판독주의") or []:
                if c not in cautions:
                    cautions.append(c)
        return {
            "품목수": sum(v.get("품목수", 0) for v in vs),
            "품목금액합": None,
            "상태": bad[0]["상태"] if bad else "일치",
            "판독주의": cautions,
            "보정": [x for v in vs for x in (v.get("보정") or [])],
            "확인필요": any(v.get("확인필요") for v in vs),
            "영수증수": len(vs),
        }

    items = data.get("품목") or []
    amounts = [i.get("금액") for i in items if isinstance(i.get("금액"), (int, float))]
    total = data.get("총액")
    r = {"품목수": len(items), "품목금액합": sum(amounts) if amounts else None}

    if total is None:
        r["상태"] = "총액없음"
    elif not amounts:
        r["상태"] = "품목금액없음"
    else:
        diff = total - sum(amounts)
        r["차액"] = diff
        # 봉투값/부가세 별도 표기 등으로 몇 원 어긋나는 일이 잦아 1원까지는 맞은 것으로 본다
        r["상태"] = "일치" if abs(diff) <= 1 else "검산불일치"

    sup, vat = data.get("공급가액"), data.get("부가세")
    if isinstance(sup, int) and isinstance(vat, int) and isinstance(total, int):
        r["공급가액+부가세차액"] = total - (sup + vat)

    cautions = json_clean_cautions(data.get("판독주의"))
    data["판독주의"] = cautions                    # 파일에도 정리된 문구만 남긴다
    if isinstance(data.get("비고"), str) and len(data["비고"]) > 300:
        data["비고"] = data["비고"][:300] + "…"
    fixes = list(data.get("_보정") or [])
    r["판독주의"] = cautions
    r["보정"] = fixes
    total_unsure = any(("총액" in c or "합계" in c) for c in cautions)
    r["확인필요"] = (r["상태"] in ("총액없음", "검산불일치")) or ("총액" in fixes) or total_unsure
    return r


def json_note(v):
    """확인이 필요한 이유를 한 줄로 적는다."""
    bits = [v["상태"]] if v.get("상태") not in (None, "일치") else []
    if v.get("차액"):
        bits.append(f"차액 {v['차액']}원")
    if v.get("보정"):
        bits.append("계산으로 채움: " + ",".join(v["보정"]))
    if v.get("판독주의"):
        bits.append("판독주의 " + ",".join(v["판독주의"]))
    return " / ".join(bits)


def json_hint(v):
    """금액은 믿을 수 있지만 알아 두면 좋은 점을 짧게 적는다 (화면의 참고 문구)."""
    cs = " ".join(v.get("판독주의") or [])
    hints = []
    if "거래일자" in cs:
        hints.append("날짜를 읽지 못해 파일 이름의 날짜를 씁니다")
    if any(k in cs for k in ("상호", "사업자", "카드번호", "승인번호", "결제수단")):
        hints.append("상호·번호 일부를 읽지 못했습니다")
    if v.get("상태") == "품목금액없음":
        hints.append("품목별 금액 없이 총액만 읽었습니다")
    if not hints and cs:
        hints.append("일부 글자가 불분명합니다")
    return " · ".join(hints)


def json_score(d):
    """재판독 결과와 1차 결과 중 어느 쪽이 나은지 비교하는 점수 (클수록 좋다)."""
    v = d.get("_검산") or {}
    total = d.get("총액")
    ok_total = isinstance(total, (int, float)) or (
        isinstance(d.get("영수증"), list) and bool(d["영수증"]))
    return (not v.get("확인필요"), ok_total, v.get("상태") == "일치")


def json_pack_multi(kids):
    """한 사진 안의 영수증 여러 장을 장부가 받는 '영수증 목록' 모양으로 묶는다."""
    for n, k in enumerate(kids, 1):
        k["사진내순번"] = n
        json_repair(k)
    vendors = [k.get("상호") for k in kids if k.get("상호")]
    totals = [k.get("총액") for k in kids]
    return {
        "영수증": kids,
        "상호": " / ".join(vendors) if vendors else None,
        "총액": sum(totals) if all(isinstance(t, int) for t in totals) else None,
        "거래일자": next((k.get("거래일자") for k in kids if k.get("거래일자")), None),
        "품목": [], "판독주의": [], "비고": f"한 사진에 영수증 {len(kids)}장",
    }


def json_read(client, page_img, model, max_side=None):
    """사진 한 장을 읽어 판독 결과를 돌려준다. 영수증이 여러 장이면 따로따로 읽는다."""
    png = json_page_png(page_img, max_side)
    data = json_ask(client, png, model)
    n = data.get("영수증장수")
    if isinstance(n, int) and 2 <= n <= 4:
        one = {k: v for k, v in JSON_SCHEMA["properties"].items() if k != "영수증장수"}
        schema = {"type": "object", "required": ["영수증"],
                  "properties": {"영수증": {"type": "array", "items": {
                      "type": "object", "properties": one,
                      "required": ["총액", "품목", "판독주의"]}}}}
        prompt = (JSON_PROMPT + f"\n이 이미지에는 서로 다른 영수증이 {n}장 있습니다. 왼쪽, 위쪽 영수증부터 "
                  "순서대로 각각 따로 읽어 '영수증' 배열에 한 장씩 넣으세요. 금액이나 품목을 합치거나 "
                  "섞지 마세요. 각 영수증의 총액은 그 영수증 자기 합계 줄의 값입니다.")
        try:
            kids = [k for k in json_ask(client, png, model, schema=schema, prompt=prompt).get("영수증", [])
                    if isinstance(k, dict)]
        except JsonFatalError:
            raise
        except Exception:
            kids = []
        if len(kids) >= 2:
            return json_pack_multi(kids)
    return json_repair(data)


def json_file_hash(path):
    h = hashlib.sha1()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def json_dedup(files, log=print):
    """같은 영수증이 jpg 와 pdf 로 두 벌 있으면 한 벌만 남긴다.

    장부에 같은 금액이 두 번 올라가는 사고를 막는 자리다. 이름은 대소문자와 앞뒤
    공백만 무시한다. 한때 이름 속 공백까지 지웠더니 '남한강 마트' 와 '남한강마트' 를
    같은 영수증으로 보고 멀쩡한 한 장을 통째로 빠뜨렸다.
    """
    # 1) 내용이 완전히 같은 파일 (이름만 다르게 복사해 둔 것). 확실한 중복이다.
    keep, seen = [], {}
    for p in files:
        try:
            h = json_file_hash(p)
        except OSError:
            keep.append(p)
            continue
        prev = seen.get(h)
        if prev is None:
            seen[h] = p
            keep.append(p)
        else:
            log(f"  - 건너뜀 {os.path.basename(p)} (내용이 {os.path.basename(prev)} 와 같음)")

    # 2) 이름이 같은데 하나는 그림, 하나는 PDF -> 같은 영수증의 두 형식.
    #    PDF 쪽을 남긴다 (원본 해상도가 높다).
    by_name, final = {}, []
    for p in keep:
        k = os.path.splitext(os.path.basename(p))[0].strip().lower()
        prev = by_name.get(k)
        if prev is None:
            by_name[k] = p
            final.append(p)
            continue
        is_pdf = os.path.splitext(p)[1].lower() in PDF_EXTS
        was_pdf = os.path.splitext(prev)[1].lower() in PDF_EXTS
        if is_pdf == was_pdf:
            # 같은 형식끼리 이름이 겹치는 건 하위 폴더가 다른 것뿐이다. 둘 다 처리한다.
            final.append(p)
            continue
        drop, stay = (prev, p) if is_pdf else (p, prev)
        if is_pdf:
            final[final.index(prev)] = p
            by_name[k] = p
        log(f"  - 건너뜀 {os.path.basename(drop)} "
            f"(같은 이름의 {'PDF' if stay is not drop else '그림'} {os.path.basename(stay)} 사용)")
    return final


def extract_json(target, out_dir=None, model=None, pro_model=None, auto_pro=True,
                 force=False, recheck=False, limit=None, allow_dup=False,
                 sleep=0.0, log=print):
    """영수증 그림/PDF 를 읽어 JSON 을 만든다. 만들어진 JSON 목록을 돌려준다.

    target 은 파일 하나, 폴더, 또는 (정리 직후처럼) 파일 경로 목록이다.
    이미 만들어 둔 JSON 은 건너뛴다. force 면 모두 다시, recheck 면 검산에
    걸린 것만 다시 읽는다.
    """
    model = model or JSON_MODEL
    pro_model = pro_model or JSON_PRO_MODEL

    if isinstance(target, (list, tuple)):
        files = [p for p in target
                 if os.path.splitext(p)[1].lower() in IMAGE_EXTS | PDF_EXTS]
        base = os.path.dirname(os.path.abspath(files[0])) if files else os.getcwd()
    else:
        files = collect_files(target)
        base = (os.path.abspath(target) if os.path.isdir(target)
                else os.path.dirname(os.path.abspath(target)))

    if not allow_dup:
        files = json_dedup(files, log=log)
    if limit:
        files = files[:limit]
    if not files:
        log("JSON 으로 만들 영수증이 없습니다.")
        return []

    out_dir = out_dir or os.path.join(base, "json")
    os.makedirs(out_dir, exist_ok=True)

    log(f"\nJSON 추출 {len(files)}개  ->  {out_dir}")
    client = json_client()
    # 키가 틀렸으면 사진을 보내기 전에 멈춘다. 모델이 없어졌으면 쓸 수 있는 모델로 바꾼다.
    model = json_resolve_model(client, model, log=log)

    # 1차는 싼 모델로 훑고, 검산에 걸린 장만 뒤에서 Pro 로 다시 읽는다.
    use_pro = bool(auto_pro and pro_model and pro_model != model)
    results, labels, made, failed = [], [], [], []
    pro_todo = []              # (results 안 자리, png, 저장경로, 이름)

    def keep(label, dst, png, data):
        results.append(data)
        labels.append(label)
        if use_pro and (data.get("_검산") or {}).get("확인필요") \
                and data.get("_모델") != pro_model:
            pro_todo.append((len(results) - 1, png, dst, label))

    for i, src in enumerate(files, 1):
        name = os.path.basename(src)
        try:
            pages = load_pages(src)
        except Exception as e:
            log(f"[{i}/{len(files)}] 실패 {name}: {e}")
            failed.append((name, str(e)))
            continue

        stem = os.path.splitext(name)[0]
        for pno, page in enumerate(pages, 1):
            label = name if len(pages) == 1 else f"{name} p{pno}"
            dst = os.path.join(out_dir, (f"{stem}.json" if len(pages) == 1
                                         else f"{stem}_p{pno}.json"))
            old = None
            if os.path.exists(dst) and not force:
                try:
                    with open(dst, encoding="utf-8") as f:
                        old = json.load(f)
                except Exception:
                    old = None
                # recheck 는 사람이 확인해야 할 것만 골라 다시 읽는다. 멀쩡히 나온 것까지
                # 다시 부르면 돈과 시간이 아깝고, 결과가 흔들릴 수도 있다.
                redo = recheck and (old is None
                                    or (old.get("_검산") or {}).get("확인필요"))
                if not redo:
                    log(f"[{i}/{len(files)}] 건너뜀(이미 있음) {label}")
                    if old is not None:
                        made.append(dst)
                        # 이미 있던 것도 검산에 걸렸으면 재판독 대상이다
                        keep(label, dst, page, old)
                    continue
                log(f"[{i}/{len(files)}] 다시 읽기 {label}")
            else:
                log(f"[{i}/{len(files)}] {label}")

            try:
                data = json_read(client, page, model)
            except JsonFatalError as e:
                # 키/권한 문제다. 남은 사진을 돌려 봐야 똑같이 실패한다.
                log(f"[중단] {e}")
                log("  이미 만든 JSON 은 남아 있습니다. 원인을 고친 뒤 다시 실행하면"
                    " 남은 것부터 이어서 합니다.")
                return made
            except Exception as e:
                log(f"    실패: {json_short_err(e)}")
                failed.append((label, json_short_err(e)))
                continue

            data["_원본"] = os.path.abspath(src)
            if len(pages) > 1:
                data["_쪽"] = pno
            data["_처리시각"] = datetime.now().isoformat(timespec="seconds")
            data["_모델"] = model
            data["_검산"] = json_verify(data)
            json_write(dst, data)
            made.append(dst)
            keep(label, dst, page, data)

            v = data["_검산"]
            log(f"    ** {json_note(v)}" if v["확인필요"]
                else f"    총액 {data.get('총액')}원 / 품목 {v['품목수']}개")
            if sleep:
                time.sleep(sleep)

    # 검산에 걸린 장만 좋은 모델로 자동 재판독한다.
    # 1차 결과는 _1차판독 에 남겨 두어 나중에 비교할 수 있게 한다.
    if pro_todo:
        try:
            pro_model = json_resolve_model(client, pro_model, log=log)
        except JsonSetupError as e:
            log(f"[재판독 건너뜀] {json_short_err(e)}  (1차 결과를 그대로 둡니다)")
            pro_todo = []
        if pro_model == model:
            pro_todo = []          # 대신 쓸 모델이 1차 모델과 같으면 다시 읽어도 소용없다
        if pro_todo:
            log(f"\n검산에 걸린 {len(pro_todo)}장을 {pro_model} 로 다시 읽습니다.")
        for n, (idx, page, dst, label) in enumerate(pro_todo, 1):
            log(f"[재판독 {n}/{len(pro_todo)}] {label}")
            try:
                data = json_read(client, page, pro_model, JSON_RECHECK_SIDE)
            except JsonFatalError as e:
                log(f"[중단] {e}")
                log("  1차 결과는 그대로 남아 있습니다. 원인을 고친 뒤 다시 실행하면"
                    " 검산에 걸린 장만 골라 이어서 재판독합니다.")
                break
            except Exception as e:
                log(f"    실패: {json_short_err(e)}  (1차 결과를 그대로 둡니다)")
                failed.append((f"{label} (재판독)", json_short_err(e)))
                continue

            before = results[idx]
            data["_원본"] = before.get("_원본")
            if before.get("_쪽"):
                data["_쪽"] = before["_쪽"]
            data["_처리시각"] = datetime.now().isoformat(timespec="seconds")
            data["_모델"] = pro_model
            data["_검산"] = json_verify(data)
            if json_score(before) > json_score(data):
                log("    -> 재판독 결과가 더 나빠 1차 결과를 그대로 둡니다")
                continue
            data["_1차판독"] = {"모델": before.get("_모델"), "총액": before.get("총액"),
                             "검산": before.get("_검산")}
            json_write(dst, data)
            results[idx] = data

            v = data["_검산"]
            log(f"    -> 여전히 확인 필요: {json_note(v)}" if v["확인필요"]
                else f"    -> 검산 일치. 총액 {data.get('총액')}원 / 품목 {v['품목수']}개")
            if sleep:
                time.sleep(sleep)

    json_write(os.path.join(out_dir, "_모두.json"), results)
    json_summary_csv(os.path.join(out_dir, "_요약.csv"), results)

    # 확인 목록은 재판독까지 끝난 마지막 결과로 만든다 (재판독으로 풀린 장은 빠진다)
    need = [(labels[k], d["_검산"]) for k, d in enumerate(results)
            if (d.get("_검산") or {}).get("확인필요")]
    log(f"\nJSON {len(results)}개  ->  {out_dir}")
    log(f"확인이 필요한 영수증 {len(need)}장" + (":" if need else ""))
    for label, v in need:
        log(f"  - {label}: {json_note(v)}")
    if failed:
        log(f"실패 {len(failed)}장 (다시 실행하면 실패분만 재시도합니다):")
        for label, err in failed:
            log(f"  - {label}: {err[:120]}")
    return made


def json_write(path, data):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def json_summary_csv(path, results):
    """엑셀에서 바로 열어 볼 한 장짜리 요약표."""
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["원본", "거래일자", "상호", "총액", "품목수", "검산", "판독주의",
                    "모델", "비고"])
        for d in results:
            v = d.get("_검산", {})
            w.writerow([
                os.path.basename(d.get("_원본", "")), d.get("거래일자") or "",
                d.get("상호") or "",
                d.get("총액") if d.get("총액") is not None else "",
                v.get("품목수", ""), v.get("상태", ""),
                ",".join(v.get("판독주의") or []), d.get("_모델", ""),
                d.get("비고") or "",
            ])


# ---------------------------------------------------------------------------
# GUI (인자 없이 실행했을 때)


def _split_drop(data):
    """드래그 앤 드롭으로 넘어온 경로 문자열을 목록으로 나눈다.

    공백이 든 경로는 '{...}' 로 묶여 오고, 여러 개면 공백으로 이어 붙어 온다.
    """
    paths, buf, brace = [], "", False
    for ch in data:
        if ch == "{":
            brace = True
        elif ch == "}":
            brace = False
        elif ch == " " and not brace:
            if buf:
                paths.append(buf)
                buf = ""
        else:
            buf += ch
    if buf:
        paths.append(buf)
    return [p for p in paths if os.path.exists(p)]


# ---------------------------------------------------------------------------
# 식자재 검수서(원산지 표시) 만들기
#
# 원래 '식자재검수서생성기/image_to_word.py' 에 따로 있던 코드를 이 파일 안으로 옮겨 왔다.
# 원산지 표시 사진들을 한 쪽에 3열로 놓고, 제목을 단 워드와 엑셀을 함께 만든다.
# 결과는 첫 사진이 있는 폴더에 '원산지 표시_YY.MM.DD' 이름으로 저장한다.
# 워드를 만들 때만 python-docx 가 필요하므로, 불러오기는 쓰는 함수 안에 두었다.


INSP_EXTS = (".jpg", ".jpeg", ".png", ".bmp", ".gif")
INSP_TITLE = "식자재 검수서 원산지 표시"


def insp_collect(paths):
    """고른 파일·폴더에서 검수서에 넣을 사진만 골라 경로순으로 돌려준다.

    폴더는 바로 안쪽만 본다(하위 폴더는 보지 않는다).
    """
    seen, found = set(), []

    def add(p):
        key = os.path.normcase(os.path.abspath(p))
        if key not in seen:
            seen.add(key)
            found.append(p)

    for p in paths:
        if os.path.isdir(p):
            for name in sorted(os.listdir(p)):
                full = os.path.join(p, name)
                if os.path.isfile(full) and name.lower().endswith(INSP_EXTS):
                    add(full)
        elif os.path.isfile(p) and p.lower().endswith(INSP_EXTS):
            add(p)
    return sorted(found)


def _insp_unique_path(folder, base, ext):
    """같은 이름이 있으면 _1, _2 를 붙여 덮어쓰지 않게 한다."""
    path = os.path.join(folder, f"{base}{ext}")
    i = 1
    while os.path.exists(path):
        path = os.path.join(folder, f"{base}_{i}{ext}")
        i += 1
    return path


def _insp_thumb(path, w, h, fmt):
    """사진을 w x h 안에 들어가게 줄여 메모리 속 파일로 돌려준다."""
    with Image.open(path) as im:
        im = im.convert("RGB")
        im.thumbnail((w, h), Image.Resampling.LANCZOS)
        buf = io.BytesIO()
        if fmt == "JPEG":
            im.save(buf, "JPEG", quality=90)
        else:
            im.save(buf, fmt)
    buf.seek(0)
    return buf


def _insp_cell_border(cell, left=False, right=False, top=False, bottom=False):
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    tc_pr = cell._element.get_or_add_tcPr()
    borders = OxmlElement("w:tcBorders")
    for side, on in (("top", top), ("left", left), ("bottom", bottom), ("right", right)):
        el = OxmlElement(f"w:{side}")
        if on:
            el.set(qn("w:val"), "single")
            el.set(qn("w:sz"), "10")
            el.set(qn("w:space"), "0")
            el.set(qn("w:color"), "000000")
        else:
            el.set(qn("w:val"), "none")
        borders.append(el)
    tc_pr.append(borders)


def _insp_cell_margins(cell, left=100, right=100):
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    tc_pr = cell._element.get_or_add_tcPr()
    mar = OxmlElement("w:tcMar")
    for side, value in (("left", left), ("right", right)):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:w"), str(value))
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    tc_pr.append(mar)


def _insp_shrink_table(table):
    """표 너비를 글자에 맞춰 줄인다(워드가 제멋대로 쪽 너비로 늘리는 것을 막는다)."""
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    tbl_pr = table._tbl.tblPr
    for tag in ("w:tblW", "w:tblLayout"):
        for el in tbl_pr.findall(qn(tag)):
            tbl_pr.remove(el)
    tbl_w = OxmlElement("w:tblW")
    tbl_w.set(qn("w:w"), "0")
    tbl_w.set(qn("w:type"), "auto")
    tbl_pr.append(tbl_w)
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "autofit")
    tbl_pr.append(layout)
    for row in table.rows:
        for cell in row.cells:
            tc_pr = cell._element.get_or_add_tcPr()
            for el in tc_pr.findall(qn("w:tcW")):
                tc_pr.remove(el)
            tc_w = OxmlElement("w:tcW")
            tc_w.set(qn("w:w"), "0")
            tc_w.set(qn("w:type"), "auto")
            tc_pr.append(tc_w)


def _insp_word(images, title, path, log=print):
    from docx import Document
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Inches, Mm, Pt

    doc = Document()
    section = doc.sections[0]
    # python-docx 기본 서식은 Letter 라서 A4(210 x 297mm)로 바꾼다
    section.page_width, section.page_height = Mm(210), Mm(297)
    section.top_margin = section.bottom_margin = Mm(10)
    section.left_margin = section.right_margin = Mm(10)

    # 제목은 머리글에 테두리 표로 넣어 매 쪽 위에 반복되게 한다
    header = section.header
    header.is_linked_to_previous = False
    for p in header.paragraphs:
        p.text = ""
    t_tbl = header.add_table(rows=1, cols=1, width=Inches(1))
    t_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    _insp_shrink_table(t_tbl)
    t_cell = t_tbl.rows[0].cells[0]
    _insp_cell_border(t_cell, left=True, right=True, top=True, bottom=True)
    _insp_cell_margins(t_cell, left=200, right=200)
    t_para = t_cell.paragraphs[0]
    t_para.text = title
    t_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    t_para.runs[0].font.size = Pt(14)
    t_para.runs[0].font.bold = True
    t_para.paragraph_format.space_before = Pt(4)
    t_para.paragraph_format.space_after = Pt(4)

    n = len(images)
    n_cols = min(3, n)
    n_rows = (n + n_cols - 1) // n_cols
    img_tbl = doc.add_table(rows=n_rows, cols=n_cols)
    img_tbl.autofit = False

    idx = 0
    for row in img_tbl.rows:
        for c_idx, cell in enumerate(row.cells):
            _insp_cell_border(cell, right=c_idx < n_cols - 1)
            if idx >= n:
                continue
            image_path = images[idx]
            idx += 1

            fn_para = cell.paragraphs[0]   # 사진 위에 파일 이름
            fn_para.text = os.path.splitext(os.path.basename(image_path))[0]
            fn_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            fn_para.runs[0].font.size = Pt(9)
            fn_para.paragraph_format.space_before = Pt(4)
            fn_para.paragraph_format.space_after = Pt(2)
            try:
                buf = _insp_thumb(image_path, 300, 225, "JPEG")
                im_para = cell.add_paragraph()
                im_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                im_para.paragraph_format.space_before = Pt(0)
                im_para.paragraph_format.space_after = Pt(6)
                im_para.add_run().add_picture(buf, width=Inches(2.2))
            except Exception as e:  # 한 장이 깨져도 나머지는 넣는다
                log(f"[건너뜀] 워드 사진 {os.path.basename(image_path)}: {e}")
    doc.save(path)


def _insp_excel(images, title, path, log=print):
    from openpyxl import Workbook
    from openpyxl.drawing.image import Image as XLImage
    from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor
    from openpyxl.drawing.xdr import XDRPositiveSize2D
    from openpyxl.styles import Alignment, Font
    from openpyxl.utils.units import pixels_to_EMU
    from openpyxl.worksheet.page import PageMargins

    wb = Workbook()
    ws = wb.active
    ws.title = "식자재검수서"

    col_n = 3
    col_w = 22        # 열 너비 (문자 단위)
    img_w, img_h = 155, 116   # 사진 크기 (px)
    name_h = 18       # 파일 이름 행 높이 (pt)
    img_row = 96      # 사진 행 높이 (pt)
    col_w_px = round(col_w * 7 + 5)
    row_h_px = round(img_row * 4 / 3)

    for letter in ("A", "B", "C"):
        ws.column_dimensions[letter].width = col_w

    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.print_options.horizontalCentered = True
    # 위아래 가운데 맞춤을 켜 두면 상단 여백이 무시되고 쪽 가운데로 내려간다.
    # 제목이 늘 위쪽 35mm 자리에서 시작하도록 끈다.
    ws.print_options.verticalCentered = False
    ws.page_margins = PageMargins(left=0.3937, right=0.3937, top=35 / 25.4, bottom=0.3937)
    ws.print_title_rows = "1:1"

    ws.merge_cells("A1:C1")
    tc = ws["A1"]
    tc.value = title
    tc.font = Font(size=14, bold=True)
    tc.alignment = Alignment(horizontal="center", vertical="center")
    # 제목 상자와 사진 사이 세로선은 인쇄하면 거슬려서 선은 하나도 긋지 않는다
    ws.row_dimensions[1].height = 32
    ws.row_dimensions[2].height = 6

    xl_row = 3
    for g in range(0, len(images), col_n):
        group = images[g:g + col_n]

        ws.row_dimensions[xl_row].height = name_h
        for ci, image_path in enumerate(group):
            c = ws.cell(row=xl_row, column=ci + 1)
            c.value = os.path.splitext(os.path.basename(image_path))[0]
            c.font = Font(size=9)
            c.alignment = Alignment(horizontal="center", vertical="center")
        xl_row += 1

        ws.row_dimensions[xl_row].height = img_row
        for ci, image_path in enumerate(group):
            c = ws.cell(row=xl_row, column=ci + 1)
            c.alignment = Alignment(horizontal="center", vertical="center")
            try:
                xl_img = XLImage(_insp_thumb(image_path, img_w, img_h, "PNG"))
                xl_img.width, xl_img.height = img_w, img_h
                # 칸 가운데에 오도록 밀어 둔다
                marker = AnchorMarker(col=ci, colOff=pixels_to_EMU(max(0, (col_w_px - img_w) // 2)),
                                      row=xl_row - 1, rowOff=pixels_to_EMU(max(0, (row_h_px - img_h) // 2)))
                size = XDRPositiveSize2D(pixels_to_EMU(img_w), pixels_to_EMU(img_h))
                xl_img.anchor = OneCellAnchor(_from=marker, ext=size)
                ws.add_image(xl_img)
            except Exception as e:
                log(f"[건너뜀] 엑셀 사진 {os.path.basename(image_path)}: {e}")
        xl_row += 1
    wb.save(path)


def make_inspection(targets, date_text="", out_dir=None, log=print):
    """원산지 표시 사진으로 식자재 검수서 워드·엑셀을 만든다. (워드 경로, 엑셀 경로)를 돌려준다."""
    images = insp_collect(targets if isinstance(targets, (list, tuple)) else [targets])
    if not images:
        raise ValueError("검수서에 넣을 사진(jpg/png/bmp/gif)을 찾지 못했습니다.")
    date_text = (date_text or "").strip()
    title = f"{INSP_TITLE}  {date_text}" if date_text else INSP_TITLE
    out_dir = out_dir or os.path.dirname(os.path.abspath(images[0]))
    # 파일 이름은 '원산지 표시_입력한날짜'. 날짜를 비우면 오늘 날짜(YY.MM.DD)를 쓴다.
    # 23/08/15 처럼 파일 이름에 못 쓰는 글자는 점으로 바꾼다.
    stamp = re.sub(r'[\\/:*?"<>|]', ".", date_text).strip(" .")
    base = f"원산지 표시_{stamp or datetime.now().strftime('%y.%m.%d')}"

    log(f"검수서: 사진 {len(images)}장")
    word_path = _insp_unique_path(out_dir, base, ".docx")
    _insp_word(images, title, word_path, log=log)
    log(f"[완료] 워드: {word_path}")
    excel_path = _insp_unique_path(out_dir, base, ".xlsx")
    _insp_excel(images, title, excel_path, log=log)
    log(f"[완료] 엑셀: {excel_path}")
    return word_path, excel_path


def gui():
    import threading
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    # 끌어다 놓기는 tkinterdnd2 가 있을 때만 켠다(없어도 나머지는 그대로 동작)
    try:
        from tkinterdnd2 import DND_FILES, TkinterDnD
        root = TkinterDnD.Tk()
        dnd = True
    except ImportError:
        root = tk.Tk()
        dnd = False

    root.title("영수증 정리 도구")
    # 화면이 작은 노트북에서도 아래 버튼 줄이 잘리지 않게 화면 크기에 맞춰 띄운다.
    root.geometry("{}x{}+60+40".format(
        min(720, root.winfo_screenwidth() - 80),
        min(560, root.winfo_screenheight() - 160)))

    # 기본 글씨가 작아 눈이 아프다는 이야기가 있어 한 단계씩 키운다.
    # 위젯마다 크기를 적는 대신 이름 있는 글꼴을 한 번에 손보면 창 전체에 적용된다.
    import tkinter.font as tkfont
    for fname in ("TkDefaultFont", "TkTextFont", "TkMenuFont", "TkHeadingFont",
                  "TkFixedFont", "TkTooltipFont"):
        try:
            f = tkfont.nametofont(fname)
            f.configure(size=abs(f.cget("size")) + 2)
        except tk.TclError:
            pass
    ttk.Style().configure(".", font="TkDefaultFont")

    targets = []  # 처리할 파일/폴더 목록
    results = []  # 방금 만든 결과 파일 (손질 창에서 사용)
    bg_var = tk.StringVar(value="white")
    fmt_var = tk.StringVar(master=root, value="jpg")
    deskew_var = tk.BooleanVar(value=True)
    clean_var = tk.BooleanVar(value=True)
    rename_var = tk.BooleanVar(value=False)
    excel_var = tk.BooleanVar(value=True)   # 정리한 다음 엑셀까지 만드는 것이 보통이다
    sub_var = tk.BooleanVar(value=True)     # 끌어다 놓은 폴더의 하위 폴더까지 훑는다
    json_var = tk.BooleanVar(value=False)   # JSON 은 돈이 드는 일이라 기본은 꺼 둔다
    ledger_var = tk.BooleanVar(value=False)  # 장부는 JSON 이 있어야 만들 수 있다

    # 창을 작게 줄여도 안쪽 내용을 볼 수 있도록 스크롤 막대를 단다.
    # 아래 버튼 줄은 스크롤 밖(항상 보이는 자리)에 먼저 앉힌다.
    # 버튼이 한 줄에 다 들어가지 않아 두 줄로 나눈다. 둘째 줄은 정리와 따로 가는 일.
    bar2 = ttk.Frame(root)
    bar2.pack(side="bottom", pady=(0, 10))
    bar = ttk.Frame(root)
    bar.pack(side="bottom", pady=(0, 4))

    outer = ttk.Frame(root)
    outer.pack(side="top", fill="both", expand=True)
    vbar = ttk.Scrollbar(outer, orient="vertical")
    hbar = ttk.Scrollbar(outer, orient="horizontal")
    sheet = tk.Canvas(outer, highlightthickness=0,
                      yscrollcommand=vbar.set, xscrollcommand=hbar.set)
    vbar.config(command=sheet.yview)
    hbar.config(command=sheet.xview)
    vbar.pack(side="right", fill="y")
    hbar.pack(side="bottom", fill="x")
    sheet.pack(side="left", fill="both", expand=True)

    body = ttk.Frame(sheet)
    body_id = sheet.create_window((0, 0), window=body, anchor="nw")

    def _sheet_resize(event=None):
        sheet.configure(scrollregion=sheet.bbox("all"))
        # 창이 내용보다 넓으면 내용을 창 너비에 맞춰 늘린다(가로 막대는 그때만 쓴다).
        sheet.itemconfigure(body_id, width=max(sheet.winfo_width(),
                                               body.winfo_reqwidth()))

    body.bind("<Configure>", _sheet_resize)
    sheet.bind("<Configure>", _sheet_resize)
    def _sheet_wheel(event):
        # 휠은 창 전체에서 받되, 다른 창(손질 창)이나 스스로 굴러가는 칸은 건드리지 않는다.
        if event.widget.winfo_toplevel() is not root:
            return
        if event.widget in (logbox, listbox):
            return
        sheet.yview_scroll(-1 if event.delta > 0 else 1, "units")

    root.bind_all("<MouseWheel>", _sheet_wheel)

    top = ttk.Frame(body, padding=10)
    top.pack(fill="x")
    hint = ("파일이나 폴더를 이 창에 끌어다 놓으세요 (여러 개 가능)"
            if dnd else "처리할 파일이나 폴더를 고르세요")
    ttk.Label(top, text=hint).grid(row=0, column=0, columnspan=3, sticky="w")

    listbox = tk.Listbox(top, height=4, width=52)
    listbox.grid(row=1, column=0, rowspan=3, padx=(0, 6), pady=(4, 0))

    def refresh():
        listbox.delete(0, "end")
        for t in targets:
            listbox.insert("end", t)

    def add(paths):
        for p in paths:
            if p not in targets:
                targets.append(p)
        refresh()

    def pick_file():
        add(filedialog.askopenfilenames(filetypes=[
            ("영수증 이미지/PDF", "*.png *.jpg *.jpeg *.bmp *.webp *.tif *.tiff *.pdf"),
            ("모든 파일", "*.*"),
        ]))

    def pick_dir():
        path = filedialog.askdirectory()
        if path:
            add([path])

    def drop(event):
        add(_split_drop(event.data))

    ttk.Button(top, text="파일...", command=pick_file).grid(row=1, column=1, sticky="ew")
    ttk.Button(top, text="폴더...", command=pick_dir).grid(row=2, column=1, sticky="ew", pady=2)
    ttk.Button(top, text="목록 비우기",
               command=lambda: (targets.clear(), refresh())).grid(row=3, column=1, sticky="ew")

    opt = ttk.LabelFrame(body, text="옵션", padding=10)
    opt.pack(fill="x", padx=10, pady=6)
    # 배경은 실제로 쓰는 '흰색으로' 하나만 남긴다(투명/그대로 두기는 명령줄에만 남김).
    ttk.Label(opt, text="배경 처리: 흰색으로").grid(row=0, column=0, columnspan=4,
                                                sticky="w")
    ttk.Checkbutton(opt, text="기울기 자동 보정", variable=deskew_var).grid(row=1, column=0, sticky="w")
    ttk.Checkbutton(opt, text="종이 밝기 보정", variable=clean_var).grid(row=1, column=1, sticky="w")
    fmt_box = ttk.Frame(opt)
    fmt_box.grid(row=1, column=2, columnspan=2, sticky="w")
    ttk.Label(fmt_box, text="저장 형식").pack(side="left")
    # 손질 창에서 덮어써도 열화가 없도록, JPEG 로 저장하더라도 품질은 95 로 고정한다.
    for label, val in (("JPG(권장, 품질95)", "jpg"), ("PNG(무손실·용량 큼)", "png")):
        ttk.Radiobutton(fmt_box, text=label, value=val,
                        variable=fmt_var).pack(side="left", padx=2)
    ttk.Checkbutton(opt, text="하위 폴더까지 훑기",
                    variable=sub_var).grid(row=2, column=0, sticky="w")
    ttk.Checkbutton(opt, text="원본 파일 이름도 바꾸기(되돌리기 기록 남김)",
                    variable=rename_var).grid(row=2, column=1, columnspan=3, sticky="w")
    ttk.Checkbutton(opt, text="정리한 뒤 엑셀(A4 2열)로 묶기",
                    variable=excel_var).grid(row=3, column=0, columnspan=3, sticky="w")
    ttk.Checkbutton(opt, text="정리한 뒤 영수증 내용을 JSON 으로 뽑기 (Gemini, 통신 필요)",
                    variable=json_var).grid(row=4, column=0, columnspan=4, sticky="w")

    def json_toggled(*_):
        # 장부는 JSON 을 읽어 만든다. JSON 을 끄면 장부도 함께 꺼야 앞뒤가 맞는다.
        if not json_var.get():
            ledger_var.set(False)
        ledger_chk.config(state="normal" if json_var.get() else "disabled")

    ledger_chk = ttk.Checkbutton(opt, text="  └ 이어서 식자재 지출현황 장부(엑셀)까지 만들기",
                                 variable=ledger_var, state="disabled")
    ledger_chk.grid(row=5, column=0, columnspan=4, sticky="w")
    json_var.trace_add("write", json_toggled)

    logbox = tk.Text(body, height=6)
    logbox.pack(side="top", fill="both", expand=True, padx=10, pady=(0, 6))

    def log(msg):
        logbox.insert("end", f"{msg}\n")
        logbox.see("end")
        logbox.update_idletasks()

    def edit_results():
        """이미지 수정 창을 연다.

        방금 정리한 결과가 있으면 그것을 올리고, 없으면 빈 창을 띄운다.
        빈 창에는 파일이나 폴더를 그대로 끌어다 놓으면 된다.
        """
        picks = [p for p in results if os.path.splitext(p)[1].lower() in IMAGE_EXTS]
        if not picks:
            # 목록에 넣어 둔 것이 있으면 그것을 쓴다(폴더면 안쪽 이미지를 모두).
            for t in targets:
                if os.path.isdir(t):
                    for base, _, names in os.walk(t):
                        picks.extend(os.path.join(base, n) for n in sorted(names)
                                     if os.path.splitext(n)[1].lower() in IMAGE_EXTS)
                elif os.path.splitext(t)[1].lower() in IMAGE_EXTS:
                    picks.append(t)
        files, start = edit_list(picks)
        manual_edit(files, start=start, log=log)

    def excel_results():
        """정리결과 이미지를 엑셀 한 파일로 묶는다.

        방금 정리한 결과가 있으면 그것을 쓰고, 없으면 폴더를 골라 받는다.
        """
        picks = list(results)
        if not picks:
            folder = filedialog.askdirectory(
                title="엑셀로 묶을 이미지 폴더를 고르세요 (예: 정리결과)")
            if not folder:
                return
            picks = folder

        def work():
            excel_btn.config(state="disabled")
            try:
                make_excel(picks, log=log)
            except Exception as e:
                log(f"[실패] 엑셀 만들기: {e}")
            finally:
                excel_btn.config(state="normal")

        threading.Thread(target=work, daemon=True).start()

    def run_json(picks):
        """영수증 내용을 읽어 JSON 으로 만든다. 키가 없으면 안내만 하고 끝낸다."""
        try:
            return extract_json(picks, log=log)
        except JsonSetupError as e:
            log(f"[JSON 추출 못 함] {e}")
        except Exception as e:
            log(f"[실패] JSON 추출: {e}")
        return []

    def json_results():
        """JSON 으로 뽑을 폴더를 골라서 읽는다.

        어느 폴더를 읽는지 눈으로 확인하고 고르는 편이 안전하다. 방금 정리한
        결과가 있으면 그 폴더가 처음부터 열려 있어 그대로 확인만 하면 된다.
        """
        start_dir = os.path.dirname(results[0]) if results else ""
        folder = filedialog.askdirectory(
            title="JSON 으로 뽑을 영수증 폴더를 고르세요 (예: 정리결과)",
            initialdir=start_dir or None, mustexist=True)
        if not folder:
            return
        picks = folder
        log(f"JSON 대상 폴더: {folder}")

        def work():
            json_btn.config(state="disabled")
            try:
                run_json(picks)
            finally:
                json_btn.config(state="normal")

        threading.Thread(target=work, daemon=True).start()

    def run_ledger(folder):
        try:
            make_ledger(folder, log=log)
        except Exception as e:
            log(f"[실패] 장부 만들기: {e}")

    def ledger_results():
        """JSON 이 든 폴더를 골라 식자재 지출현황 장부를 만든다."""
        start_dir = os.path.dirname(results[0]) if results else ""
        folder = filedialog.askdirectory(
            title="JSON 이 든 폴더를 고르세요 (월별 폴더를 고르면 하위까지 다 읽습니다)",
            initialdir=start_dir or None, mustexist=True)
        if not folder:
            return
        log(f"장부 대상 폴더: {folder}")

        def work():
            ledger_btn.config(state="disabled")
            try:
                run_ledger(folder)
            finally:
                ledger_btn.config(state="normal")

        threading.Thread(target=work, daemon=True).start()

    insp_win = {}   # 검수서 창은 하나만 띄운다

    def inspection_window():
        """식자재 검수서(원산지 표시) 창. 사진을 넣고 날짜를 적어 워드·엑셀을 만든다.

        영수증 목록과는 따로 사진 목록을 가진다. 원산지 표시 사진은 영수증이 아니라
        섞이면 안 되기 때문이다.
        """
        if insp_win.get("win") is not None and insp_win["win"].winfo_exists():
            insp_win["win"].lift()
            insp_win["win"].focus_force()
            return

        win = tk.Toplevel(root)
        insp_win["win"] = win
        win.title("식자재 검수서 만들기")
        win.geometry("+{}+{}".format(root.winfo_x() + 40, root.winfo_y() + 40))
        pics = []

        frm = ttk.Frame(win, padding=10)
        frm.pack(fill="both", expand=True)
        ttk.Label(frm, text=("원산지 표시 사진(또는 폴더)을 이 창에 끌어다 놓으세요"
                             if dnd else "원산지 표시 사진이나 폴더를 고르세요")
                  ).grid(row=0, column=0, columnspan=2, sticky="w")

        plist = tk.Listbox(frm, height=8, width=50)
        plist.grid(row=1, column=0, rowspan=3, sticky="nsew", padx=(0, 6), pady=(4, 0))
        count = ttk.Label(frm, text="사진 0장")
        count.grid(row=4, column=0, sticky="w", pady=(2, 0))

        def refresh_pics():
            plist.delete(0, "end")
            for p in pics:
                plist.insert("end", os.path.basename(p))
            count.config(text=f"사진 {len(pics)}장")
            make_btn.config(state="normal" if pics else "disabled")

        def add_pics(paths):
            found = insp_collect(paths)
            if paths and not found:
                messagebox.showerror("검수서", "사진 파일(jpg/png/bmp/gif)을 찾지 못했습니다.",
                                     parent=win)
                return
            pics[:] = sorted(set(pics) | set(found))
            refresh_pics()

        ttk.Button(frm, text="파일...", command=lambda: add_pics(filedialog.askopenfilenames(
            parent=win, filetypes=[("사진", "*.jpg *.jpeg *.png *.bmp *.gif"),
                                   ("모든 파일", "*.*")]))).grid(row=1, column=1, sticky="ew")

        def pick_pic_dir():
            d = filedialog.askdirectory(parent=win)
            if d:
                add_pics([d])

        ttk.Button(frm, text="폴더...", command=pick_pic_dir).grid(row=2, column=1, sticky="new", pady=2)
        ttk.Button(frm, text="목록 비우기", command=lambda: (pics.clear(), refresh_pics())
                   ).grid(row=3, column=1, sticky="new")

        dfrm = ttk.Frame(frm)
        dfrm.grid(row=5, column=0, columnspan=2, sticky="w", pady=(8, 0))
        ttk.Label(dfrm, text="제목에 붙일 날짜").pack(side="left")
        date_var = tk.StringVar(master=win)
        ttk.Entry(dfrm, textvariable=date_var, width=20).pack(side="left", padx=4)
        ttk.Label(dfrm, text="예: 2026년 7월 20일 (비우면 날짜 없이)",
                  foreground="#888").pack(side="left")

        def make():
            win.config(cursor="watch")
            make_btn.config(state="disabled")
            win.update()
            try:
                word, excel = make_inspection(list(pics), date_var.get(), log=log)
                messagebox.showinfo("검수서", "만들었습니다.\n\n워드: {}\n엑셀: {}\n\n저장 위치: {}".format(
                    os.path.basename(word), os.path.basename(excel), os.path.dirname(word)),
                    parent=win)
            except Exception as e:
                log(f"[실패] 검수서 만들기: {e}")
                messagebox.showerror("검수서", f"만들지 못했습니다.\n{e}", parent=win)
            finally:
                win.config(cursor="")
                make_btn.config(state="normal")

        btns = ttk.Frame(frm)
        btns.grid(row=6, column=0, columnspan=2, pady=(10, 0))
        make_btn = ttk.Button(btns, text="워드/엑셀 만들기", command=make, state="disabled")
        make_btn.pack(side="left", padx=4)
        ttk.Button(btns, text="닫기", command=win.destroy).pack(side="left", padx=4)

        frm.columnconfigure(0, weight=1)
        frm.rowconfigure(3, weight=1)
        if dnd:
            for widget in (win, plist):
                widget.drop_target_register(DND_FILES)
                widget.dnd_bind("<<Drop>>", lambda e: add_pics(_split_drop(e.data)))

    def clear_all():
        """작업창 비우기: 기록과 목록을 모두 지운다."""
        logbox.delete("1.0", "end")
        targets.clear()
        refresh()

    def start():
        if not targets:
            messagebox.showwarning("영수증 정리", "처리할 파일이나 폴더를 넣어 주세요.")
            return
        run_btn.config(state="disabled")
        todo = list(targets)

        def work():
            try:
                results.clear()
                for t in todo:
                    log(f"--- {t} ---")
                    sub = sub_var.get()
                    if rename_var.get():
                        rename_sources(t, log=log, recursive=sub)
                    results.extend(
                        run(t, bg=bg_var.get(), fmt=fmt_var.get(),
                            deskew=deskew_var.get(), clean=clean_var.get(),
                            log=log, recursive=sub))
                if excel_var.get():
                    if results:
                        try:
                            make_excel(list(results), log=log)
                        except Exception as e:
                            log(f"[실패] 엑셀 만들기: {e}")
                    else:
                        log("엑셀로 묶을 결과가 없습니다.")
                if json_var.get():
                    if results:
                        made = run_json(list(results))
                        if ledger_var.get():
                            if made:
                                run_ledger(os.path.dirname(made[0]))
                            else:
                                log("장부로 만들 JSON 이 없습니다.")
                    else:
                        log("JSON 으로 뽑을 결과가 없습니다.")
                log("모두 끝났습니다.")
            finally:
                run_btn.config(state="normal")

        threading.Thread(target=work, daemon=True).start()

    run_btn = ttk.Button(bar, text="정리 시작", command=start)
    run_btn.pack(side="left", padx=4)
    ttk.Button(bar, text="손으로 다듬기...", command=edit_results).pack(side="left", padx=4)
    excel_btn = ttk.Button(bar, text="엑셀로 만들기...", command=excel_results)
    excel_btn.pack(side="left", padx=4)
    json_btn = ttk.Button(bar, text="JSON 뽑기...", command=json_results)
    json_btn.pack(side="left", padx=4)
    ledger_btn = ttk.Button(bar, text="장부 만들기...", command=ledger_results)
    ledger_btn.pack(side="left", padx=4)
    ttk.Button(bar2, text="식자재 검수서 만들기...",
               command=inspection_window).pack(side="left", padx=4)
    ttk.Button(bar2, text="작업창 비우기", command=clear_all).pack(side="left", padx=4)

    if dnd:
        # 창 전체와 목록/기록칸 어디에 놓아도 받도록 등록
        for widget in (root, listbox, logbox):
            widget.drop_target_register(DND_FILES)
            widget.dnd_bind("<<Drop>>", drop)

    # 스크롤 막대가 있으므로 최소 크기는 아래 버튼 줄만 보이면 될 만큼으로 잡는다.
    root.update_idletasks()
    root.minsize(420, bar.winfo_reqheight() + bar2.winfo_reqheight() + 200)

    root.mainloop()


def main():
    if len(sys.argv) == 1:
        gui()
        return

    p = argparse.ArgumentParser(description="영수증 이미지/PDF 기울기 보정 및 배경 제거")
    p.add_argument("target", help="이미지/PDF 파일 또는 폴더")
    p.add_argument("-o", "--out", help="저장 폴더 (기본: 원본 폴더 옆 '원본이름_정리결과')")
    p.add_argument("--bg", choices=["white", "transparent", "none"], default="white",
                   help="배경 처리 방식 (기본: white)")
    p.add_argument("--to", choices=["jpg", "png", "pdf"], default="jpg",
                   help="출력 형식 (기본: jpg, 용량이 PNG의 1/3이고 눈으로는 차이가 없다)")
    p.add_argument("--dpi", type=int, default=PDF_RENDER_DPI,
                   help=f"PDF 를 그릴 해상도 (기본: {PDF_RENDER_DPI})")
    p.add_argument("--no-deskew", action="store_true", help="기울기 보정 끄기")
    p.add_argument("--no-clean", action="store_true", help="종이 밝기 보정 끄기")
    p.add_argument("--upright", action="store_true",
                   help="가로로 누운 결과를 세로로 세운다(기본 꺼짐)")
    p.add_argument("--perspective", action="store_true",
                   help="비스듬히 찍은 사진의 원근을 편다(스캔본에는 쓰지 말 것)")
    p.add_argument("--rename-source", action="store_true",
                   help="원본 파일 이름도 같은 규칙으로 바꾼다")
    p.add_argument("--dry-run", action="store_true",
                   help="--rename-source 와 함께: 바꿀 이름만 미리 보여준다")
    p.add_argument("--edit", action="store_true",
                   help="정리한 결과를 손질 창에서 직접 지우거나 잘라낸다")
    p.add_argument("--excel", action="store_true",
                   help="정리한 결과를 A4 2열 엑셀로도 묶는다")
    p.add_argument("--excel-only", action="store_true",
                   help="정리는 건너뛰고, 고른 폴더의 이미지만 엑셀로 묶는다")
    p.add_argument("--excel-out", help="엑셀 저장 경로 (기본: 결과폴더명_이미지정리.xlsx)")
    p.add_argument("--per-page", choices=["auto", "2", "4"], default="auto",
                   help="엑셀 한 쪽에 넣을 사진 수 (기본 auto)")
    p.add_argument("--json", action="store_true",
                   help="정리한 결과에서 영수증 내용을 읽어 JSON 으로 뽑는다")
    p.add_argument("--json-only", action="store_true",
                   help="정리는 건너뛰고, 고른 폴더의 이미지만 JSON 으로 뽑는다")
    p.add_argument("--json-out", help="JSON 저장 폴더 (기본: 대상폴더/json)")
    p.add_argument("--json-model", default=JSON_MODEL,
                   help=f"1차로 읽을 모델 (기본 {JSON_MODEL})")
    p.add_argument("--json-pro-model", default=JSON_PRO_MODEL,
                   help=f"검산에 걸린 장을 다시 읽을 모델 (기본 {JSON_PRO_MODEL})")
    p.add_argument("--no-auto-pro", action="store_true",
                   help="검산에 걸려도 좋은 모델로 다시 읽지 않는다 (1차 결과만)")
    p.add_argument("--json-force", action="store_true", help="이미 만든 JSON 도 다시 만든다")
    p.add_argument("--json-recheck", action="store_true",
                   help="검산에 걸린 것만 골라 다시 읽는다")
    p.add_argument("--장부", action="store_true", dest="ledger",
                   help="JSON 을 모아 식자재 지출현황 엑셀 장부를 만든다")
    p.add_argument("--장부만", action="store_true", dest="ledger_only",
                   help="앞 단계는 건너뛰고, 고른 폴더의 JSON 으로 장부만 만든다")
    p.add_argument("--장부출력", dest="ledger_out", help="장부 저장 경로")
    p.add_argument("--사진", dest="photos",
                   help="영수증 사진/PDF 폴더 (주면 '영수증 이미지' 시트도 넣는다)")
    p.add_argument("--제목", dest="ledger_title", help="장부 제목 (기본: 날짜에서 자동)")
    p.add_argument("--서식참고", dest="sample", help="거래처 열 순서를 가져올 기존 장부 xlsx")
    p.add_argument("--전월합계", dest="prev_total", type=int, help="전월 총합계 금액")
    p.add_argument("-r", "--하위폴더", action="store_true", dest="recursive",
                   help="하위 폴더까지 훑는다 (정리·이름바꾸기·이미지 시트 모두)")
    p.add_argument("--중복포함", action="store_true", dest="keep_dup",
                   help="같은 영수증의 사진과 PDF 를 이미지 시트에 둘 다 넣는다")
    p.add_argument("--검수서", action="store_true", dest="inspection",
                   help="정리는 건너뛰고, 고른 사진/폴더로 식자재 검수서(원산지 표시) 워드·엑셀을 만든다")
    p.add_argument("--검수날짜", dest="insp_date", default="",
                   help="검수서 제목에 붙일 날짜 (예: '2026년 7월 20일')")
    a = p.parse_args()

    if a.inspection:
        make_inspection(a.target, a.insp_date, out_dir=a.out)
        return

    def do_ledger(folder):
        try:
            make_ledger(folder, out_path=a.ledger_out, photos=a.photos,
                        recursive=a.recursive, keep_dup=a.keep_dup,
                        title=a.ledger_title, sample=a.sample,
                        prev_total=a.prev_total)
        except Exception as e:
            print(f"[장부 만들기 실패] {e}")

    def do_json(what):
        try:
            return extract_json(what, out_dir=a.json_out, model=a.json_model,
                                pro_model=a.json_pro_model, auto_pro=not a.no_auto_pro,
                                force=a.json_force, recheck=a.json_recheck)
        except JsonSetupError as e:
            print(f"\n[JSON 추출 중단] {e}")
            return []

    if a.json_only:
        do_json(a.target)
        return

    if a.ledger_only:
        do_ledger(a.target)
        return

    if a.excel_only:
        make_excel(a.target, out_path=a.excel_out, per_page=a.per_page)
        return

    if a.rename_source:
        rename_sources(a.target, dry_run=a.dry_run, recursive=a.recursive)
        if a.dry_run:
            return

    if a.edit and os.path.exists(a.target) and not a.rename_source:
        # 이미 정리해 둔 결과를 손질만 하려는 경우
        picks = collect_files(a.target)
        if picks and all(os.path.splitext(x)[1].lower() in IMAGE_EXTS for x in picks):
            files, start = edit_list(picks)
            manual_edit(files, start=start)
            return

    written = run(a.target, out_dir=a.out, bg=a.bg, fmt=a.to,
                  deskew=not a.no_deskew, clean=not a.no_clean,
                  rotate_upright=a.upright, perspective=a.perspective, dpi=a.dpi,
                  recursive=a.recursive)
    if a.excel and written:
        make_excel(written, out_path=a.excel_out, per_page=a.per_page)
    if a.edit:
        manual_edit(written)
    # JSON 은 마지막에 뽑는다. 손질까지 끝난 그림을 읽어야 판독이 가장 잘 된다.
    if (a.json or a.ledger) and written:
        made = do_json(written)
        # 장부는 방금 만든 JSON 폴더를 그대로 읽는다
        if a.ledger and made:
            do_ledger(os.path.dirname(made[0]))


if __name__ == "__main__":
    main()

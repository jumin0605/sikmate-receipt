# -*- coding: utf-8 -*-
"""내 엑셀 발주서(월별 통합 파일)의 비고 칸에 검수서 내용을 써 넣는다 - 표준 라이브러리만 쓴다.

엑셀을 openpyxl 로 열어 다시 저장하면 검수서 사진·도형·인쇄 설정이 일부 사라질 수 있다.
그래서 시트 XML 에서 바꿀 칸만 직접 고치고, 나머지 파일은 압축 파일 안에서 그대로 옮긴다.

- list_order_sheets(data): '식재료명' / '비고' 머리글과 '입 고 일' 이 있는 시트를 찾아 품목과 비고 칸 주소를 돌려준다.
- fill_order_notes(data, edits): [{sheet, cell, text}] 대로 비고 칸을 바꾼 새 엑셀(bytes)을 돌려준다.
"""

from __future__ import annotations

import html
import io
import re
import zipfile
from datetime import date

_WS = re.compile(r"\s+")


def _norm(s: str) -> str:
    return _WS.sub("", s or "")


def _col_num(letters: str) -> int:
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch.upper()) - 64)
    return n


def _split_ref(ref: str) -> tuple[str, int]:
    m = re.fullmatch(r"([A-Za-z]{1,3})(\d{1,7})", ref or "")
    if not m:
        raise ValueError(f"칸 주소가 올바르지 않습니다: {ref}")
    return m.group(1).upper(), int(m.group(2))


def _esc(s: str) -> str:
    return (s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))


def _shared_strings(z: zipfile.ZipFile) -> list[str]:
    try:
        raw = z.read("xl/sharedStrings.xml").decode("utf-8")
    except KeyError:
        return []
    out = []
    for si in re.findall(r"<si\b[^>]*?(?:/>|>.*?</si>)", raw, flags=re.S):
        out.append(html.unescape("".join(re.findall(r"<t\b[^>]*>(.*?)</t>", si, flags=re.S))))
    return out


def _sheet_files(z: zipfile.ZipFile) -> list[tuple[str, str]]:
    """[(시트 이름, 압축 파일 안 경로)] 를 통합 문서에 적힌 순서대로"""
    wb = z.read("xl/workbook.xml").decode("utf-8")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    rid_to_target = {}
    for m in re.finditer(r"<Relationship\b([^>]*)/?>", rels):
        a = m.group(1)
        rid = re.search(r'\bId="([^"]*)"', a)
        tg = re.search(r'\bTarget="([^"]*)"', a)
        if rid and tg:
            t = tg.group(1)
            rid_to_target[rid.group(1)] = t.lstrip("/") if t.startswith("/") else "xl/" + t
    out = []
    for m in re.finditer(r"<sheet\b([^>]*)/?>", wb):
        a = m.group(1)
        nm = re.search(r'\bname="([^"]*)"', a)
        rid = re.search(r'\br:id="([^"]*)"', a)
        if nm and rid and rid.group(1) in rid_to_target:
            out.append((html.unescape(nm.group(1)), rid_to_target[rid.group(1)]))
    return out


def _cells(xml: str, sst: list[str]):
    """시트 XML 에서 (행번호, 열번호, 열글자, 글) 을 차례로 꺼낸다. 글이 없는 칸은 건너뛴다."""
    for rm in re.finditer(r"<row\b([^>]*?)(?:/>|>(.*?)</row>)", xml, flags=re.S):
        body = rm.group(2)
        if not body:
            continue
        for cm in re.finditer(r"<c\b([^>]*?)(?:/>|>(.*?)</c>)", body, flags=re.S):
            attrs, inner = cm.group(1), cm.group(2) or ""
            ref = re.search(r'\br="([A-Za-z]+)(\d+)"', attrs)
            if not ref:
                continue
            t = re.search(r'\bt="([^"]*)"', attrs)
            kind = t.group(1) if t else "n"
            text = ""
            if kind == "s":
                v = re.search(r"<v>(.*?)</v>", inner, flags=re.S)
                if v and v.group(1).strip().isdigit() and int(v.group(1)) < len(sst):
                    text = sst[int(v.group(1))]
            elif kind == "inlineStr":
                text = html.unescape("".join(re.findall(r"<t\b[^>]*>(.*?)</t>", inner, flags=re.S)))
            elif kind == "str":
                v = re.search(r"<v>(.*?)</v>", inner, flags=re.S)
                text = html.unescape(v.group(1)) if v else ""
            else:
                v = re.search(r"<v>(.*?)</v>", inner, flags=re.S)
                text = v.group(1) if v else ""
            if text != "":
                letters = ref.group(1).upper()
                yield int(ref.group(2)), _col_num(letters), letters, text


def _parse_date(text: str) -> str | None:
    m = re.search(r"(\d{4})\D+(\d{1,2})\D+(\d{1,2})", text or "")
    if not m:
        return None
    try:
        return date(int(m.group(1)), int(m.group(2)), int(m.group(3))).isoformat()
    except ValueError:
        return None


def list_order_sheets(data: bytes) -> list[dict]:
    """발주서로 보이는 시트 목록. [{sheet, date, items:[{cell,row,name,note}]}]"""
    z = zipfile.ZipFile(io.BytesIO(data))
    sst = _shared_strings(z)
    result = []
    for name, path in _sheet_files(z):
        try:
            xml = z.read(path).decode("utf-8")
        except KeyError:
            continue
        if "<row" not in xml:
            continue
        cells = list(_cells(xml, sst))
        if not cells:
            continue
        by_row: dict[int, dict[int, tuple[str, str]]] = {}
        for r, c, letters, text in cells:
            by_row.setdefault(r, {})[c] = (letters, text)
        # 머리글 행: '식재료명' 칸마다 그 오른쪽에서 처음 나오는 '비고' 칸을 짝으로 삼는다
        header_row, pairs = None, []
        for r in sorted(by_row):
            names = [c for c, (_, t) in by_row[r].items() if _norm(t) == "식재료명"]
            if names:
                header_row = r
                for nc in sorted(names):
                    note_cols = sorted(c for c, (_, t) in by_row[r].items() if c > nc and _norm(t) == "비고")
                    if note_cols:
                        pairs.append((nc, note_cols[0]))
                break
        if header_row is None or not pairs:
            continue
        # 입고일: '입 고 일' 칸 오른쪽 칸의 날짜
        ymd = None
        for r in sorted(by_row):
            if r >= header_row:
                break
            for c, (_, t) in sorted(by_row[r].items()):
                if "입고일" in _norm(t):
                    nxt = [by_row[r][k][1] for k in sorted(by_row[r]) if k > c]
                    for cand in nxt:
                        d = _parse_date(cand)
                        if d:
                            ymd = d
                            break
                if ymd:
                    break
            if ymd:
                break
        items = []
        for nc, vc in pairs:
            for r in sorted(by_row):
                if r <= header_row or nc not in by_row[r]:
                    continue
                nm = by_row[r][nc][1].strip()
                if not nm or nm.endswith(">"):
                    continue
                note_letters = None
                # 비고 칸 글자: 머리글 행에서 알아낸 열
                note_letters = by_row[header_row][vc][0]
                items.append({"cell": f"{note_letters}{r}", "row": r, "name": nm,
                              "note": by_row[r].get(vc, ("", ""))[1].strip()})
        result.append({"sheet": name, "date": ymd, "items": items})
    return result


def _set_cell(xml: str, ref: str, text: str, fallback_style: str | None) -> str:
    letters, rown = _split_ref(ref)
    rm = re.search(rf'(<row\b[^>]*?\br="{rown}"[^>]*?)(/>|>(.*?)</row>)', xml, flags=re.S)
    if not rm:
        raise ValueError(f"{ref} 칸이 들어 있는 줄을 찾지 못했습니다.")
    head, tail, body = rm.group(1), rm.group(2), rm.group(3) or ""
    colnum = _col_num(letters)

    def make(style: str | None) -> str:
        s = f' s="{style}"' if style else ""
        if text == "":
            return f'<c r="{ref}"{s}/>'
        return f'<c r="{ref}"{s} t="inlineStr"><is><t xml:space="preserve">{_esc(text)}</t></is></c>'

    cm = re.search(rf'<c\b[^>]*?\br="{ref}"[^>]*?(?:/>|>.*?</c>)', body, flags=re.S)
    if cm:
        old = cm.group(0)
        st = re.search(r'\bs="([^"]*)"', old.split(">")[0])
        new_body = body[:cm.start()] + make(st.group(1) if st else fallback_style) + body[cm.end():]
    else:
        pos = len(body)
        for m in re.finditer(r'<c\b[^>]*?\br="([A-Za-z]+)\d+"', body):
            if _col_num(m.group(1)) > colnum:
                pos = m.start()
                break
        new_body = body[:pos] + make(fallback_style) + body[pos:]
    if tail == "/>":
        new_row = head + ">" + new_body + "</row>"
    else:
        new_row = head + ">" + new_body + "</row>"
    return xml[:rm.start()] + new_row + xml[rm.end():]


def _style_of_column(xml: str, letters: str) -> str | None:
    m = re.search(rf'<c\b[^>]*?\br="{letters}\d+"[^>]*?\bs="([^"]*)"', xml)
    return m.group(1) if m else None


def fill_order_notes(data: bytes, edits: list[dict]) -> bytes:
    """edits: [{sheet, cell, text}]. 시트 이름이 같은 항목끼리 묶어 한 번에 고친다."""
    z = zipfile.ZipFile(io.BytesIO(data))
    files = dict(_sheet_files(z))
    by_sheet: dict[str, list[dict]] = {}
    for e in edits:
        nm = str(e.get("sheet", ""))
        if nm not in files:
            raise ValueError(f"시트를 찾지 못했습니다: {nm}")
        by_sheet.setdefault(nm, []).append(e)
    new_xml: dict[str, bytes] = {}
    for nm, lst in by_sheet.items():
        path = files[nm]
        xml = z.read(path).decode("utf-8")
        for e in lst:
            ref = str(e.get("cell", ""))
            letters, _ = _split_ref(ref)
            text = str(e.get("text", "")).replace("\r", " ").replace("\n", " ").strip()
            xml = _set_cell(xml, ref, text, _style_of_column(xml, letters))
        new_xml[path] = xml.encode("utf-8")
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w") as zo:
        for info in z.infolist():
            buf = new_xml.get(info.filename)
            if buf is None:
                buf = z.read(info.filename)
            zi = zipfile.ZipInfo(info.filename, date_time=info.date_time)
            zi.compress_type = info.compress_type
            zi.external_attr = info.external_attr
            zo.writestr(zi, buf)
    return out.getvalue()

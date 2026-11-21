# -*- coding: utf-8 -*-
"""엑셀 발주서 비고 쓰기 테스트 (python -m unittest test_order_xlsx)"""
import io
import unittest
import zipfile

import openpyxl
from openpyxl.drawing.image import Image as XLImage
from PIL import Image

import order_xlsx as ox


def _png():
    b = io.BytesIO()
    Image.new("RGB", (40, 30), (200, 50, 50)).save(b, "PNG")
    b.seek(0)
    return b


def _book():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "11월20일"
    ws["E5"], ws["F5"] = "입 고 일 :", "2026. 11. 20일"
    ws["L5"], ws["M5"] = "입 고 일 :", "2026. 11. 20일"
    for c, v in zip("ABCDEF", ["순    서", "식재료명", "단    위", "발 주 량", "규    격", "비    고"]):
        ws[f"{c}8"] = v
    for c, v in zip("HIJKLM", ["순    서", "식재료명", "단    위", "발 주 량", "규    격", "비    고"]):
        ws[f"{c}8"] = v
    ws["B9"], ws["F9"] = "동그랑땡", "국내산"
    ws["B10"] = "닭다리살"
    ws["B11"], ws["F11"] = "두부", "직접 쓴 메모"
    ws["I9"], ws["M9"] = "11/23 돈불고기", None
    ws["I10"] = "공산품>"
    ws["I11"] = "다진마늘"
    ws["B13"] = "다진마늘"
    ws["F13"].style = "Normal"
    ws2 = wb.create_sheet("식자재검수서")
    ws2.add_image(XLImage(_png()), "B2")
    ws3 = wb.create_sheet("11월23일")
    ws3["E5"], ws3["F5"] = "입 고 일 :", "2026. 11. 23일"
    ws3["B8"], ws3["F8"] = "식재료명", "비고"
    ws3["B9"] = "오이"
    b = io.BytesIO()
    wb.save(b)
    return b.getvalue()


class OrderXlsxTest(unittest.TestCase):
    def test_list(self):
        sh = ox.list_order_sheets(_book())
        self.assertEqual([s["sheet"] for s in sh], ["11월20일", "11월23일"])
        s = sh[0]
        self.assertEqual(s["date"], "2026-11-20")
        cells = {i["cell"]: i for i in s["items"]}
        self.assertEqual(cells["F9"]["name"], "동그랑땡")
        self.assertEqual(cells["F9"]["note"], "국내산")
        self.assertEqual(cells["M9"]["name"], "11/23 돈불고기")
        self.assertIn("F13", cells)
        self.assertNotIn("M10", cells)  # '공산품>' 같은 구분 줄은 제외

    def test_fill_keeps_everything_else(self):
        data = _book()
        out = ox.fill_order_notes(data, [
            {"sheet": "11월20일", "cell": "F9", "text": "유통 27.08.10 / 국내산"},
            {"sheet": "11월20일", "cell": "F10", "text": "소비 28.09.05 & <국내산>"},
            {"sheet": "11월20일", "cell": "M9", "text": "국내산"},
            {"sheet": "11월20일", "cell": "M11", "text": "유통 26.12.04 / 국내산"},
            {"sheet": "11월20일", "cell": "F11", "text": ""},
        ])
        wb = openpyxl.load_workbook(io.BytesIO(out))
        ws = wb["11월20일"]
        self.assertEqual(ws["F9"].value, "유통 27.08.10 / 국내산")
        self.assertEqual(ws["F10"].value, "소비 28.09.05 & <국내산>")
        self.assertEqual(ws["M9"].value, "국내산")
        self.assertEqual(ws["M11"].value, "유통 26.12.04 / 국내산")
        self.assertIsNone(ws["F11"].value)
        self.assertEqual(ws["B9"].value, "동그랑땡")
        self.assertEqual(len(wb["식자재검수서"]._images), 1)
        z1, z2 = zipfile.ZipFile(io.BytesIO(data)), zipfile.ZipFile(io.BytesIO(out))
        self.assertEqual(z1.namelist(), z2.namelist())
        changed = [n for n in z1.namelist() if z1.read(n) != z2.read(n)]
        self.assertEqual(changed, ["xl/worksheets/sheet1.xml"])

    def test_bad_input(self):
        with self.assertRaises(ValueError):
            ox.fill_order_notes(_book(), [{"sheet": "없는시트", "cell": "F9", "text": "a"}])
        with self.assertRaises(ValueError):
            ox.fill_order_notes(_book(), [{"sheet": "11월20일", "cell": "ZZ", "text": "a"}])


if __name__ == "__main__":
    unittest.main()

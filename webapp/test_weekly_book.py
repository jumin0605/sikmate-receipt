# -*- coding: utf-8 -*-
"""주간 식단표·검식일지 엑셀 만들기 테스트 (python -m unittest test_weekly_book)"""
import io
import unittest
import zipfile
from datetime import date

import openpyxl
from PIL import Image

import weekly_book as wb


def _days(main="제육볶음"):
    d = {"중식": {"밥": "잡곡밥", "국": "된장국", "주메뉴": main, "조림볶음": "감자채볶음", "무침": "호박나물", "김치": "김치"},
         "석식": {"밥": "잡곡밥", "국": "미역국", "주메뉴": "새메뉴테스트", "조림볶음": "", "무침": "", "김치": "김치"},
         "오전": "찐달걀", "오후": "요플레"}
    return [d] * 7


def _jpg():
    b = io.BytesIO()
    Image.new("RGB", (800, 600), (120, 160, 120)).save(b, "JPEG")
    return b.getvalue()


class WeeklyBookTest(unittest.TestCase):
    def test_name_and_monday(self):
        self.assertEqual(wb.output_name(date(2026, 11, 30)), "11월-11.30~12.06식단표_검식일지.xlsm")
        with self.assertRaises(ValueError):
            wb.parse_monday("2026-11-17")

    def test_build_fills_inputs(self):
        name, data, info = wb.build_weekly(date(2026, 11, 23), _days(), {"inspector": "홍길동", "keep_days": "7"},
                                           {"새메뉴테스트": ["1조각"] + [""] * 10},
                                           {"d0_lunch_tray": _jpg(), "d0_lunch_keep": _jpg(), "d6_dinner_tray": _jpg(), "bad": b"x"})
        self.assertEqual(name, "11월-11.23~11.29식단표_검식일지.xlsm")
        self.assertEqual(info["photos"], 3)
        self.assertEqual(info["missing_menudb"], [])
        z = zipfile.ZipFile(io.BytesIO(data))
        self.assertIn("macroEnabled", z.read("[Content_Types].xml").decode())
        book = openpyxl.load_workbook(io.BytesIO(data), keep_vba=True)
        self.assertEqual([s.title for s in book.worksheets][:3], ["설정", "11월 식단표", "메뉴DB"])
        self.assertEqual(book["설정"]["B3"].value.date(), date(2026, 11, 23))
        self.assertEqual(book["설정"]["B5"].value, "홍길동")
        self.assertEqual(book["설정"]["B10"].value, 7)
        ms = book["11월 식단표"]
        self.assertEqual(ms["B7"].value, "제육볶음")
        self.assertEqual(ms["H13"].value, "새메뉴테스트")
        self.assertEqual(ms["B17"].value, "찐달걀")
        self.assertEqual(len(book["주간식단표(사진)"]._images), 2)
        self.assertEqual(len(book["보존식 기록지(사진)"]._images), 1)
        names = [book["메뉴DB"].cell(r, 1).value for r in range(2, book["메뉴DB"].max_row + 1)]
        self.assertIn("새메뉴테스트", names)

    def test_month_rename_updates_formulas(self):
        _, data, info = wb.build_weekly(date(2026, 12, 7), _days(), {}, {}, {})
        self.assertIn("새메뉴테스트", info["missing_menudb"])
        book = openpyxl.load_workbook(io.BytesIO(data))
        self.assertIn("12월 식단표", book.sheetnames)
        self.assertNotIn("11월 식단표", book.sheetnames)
        for ws in book.worksheets:
            for row in ws.iter_rows():
                for c in row:
                    if isinstance(c.value, str) and c.value.startswith("="):
                        self.assertNotIn("11월 식단표", c.value, f"{ws.title}!{c.coordinate}")
        self.assertEqual(book["검식일지(월)"]["B11"].value, "=IF('12월 식단표'!B$7=\"\",\"\",'12월 식단표'!B$7)")


if __name__ == "__main__":
    unittest.main()

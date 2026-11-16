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


class ImportTest(unittest.TestCase):
    def test_read_back_and_template(self):
        days = _days()
        _, data, _ = wb.build_weekly(date(2026, 11, 23), days, {"inspector": "홍길동"}, {}, {"d2_dinner_tray": _jpg(), "d3_lunch_keep": _jpg()})
        info = wb.read_weekly(data, "x.xlsm")
        self.assertEqual(info["monday"], "2026-11-23")
        self.assertEqual(info["days"][0]["중식"]["주메뉴"], "제육볶음")
        self.assertEqual(info["days"][6]["오후"], "요플레")
        self.assertEqual(info["settings"]["inspector"], "홍길동")
        self.assertEqual(sorted(info["photos"]), ["d2_dinner_tray", "d3_lunch_keep"])
        self.assertTrue(info["template_ok"])
        tpl = wb.make_template(data)
        book = openpyxl.load_workbook(io.BytesIO(tpl))
        self.assertIsNone(book["11월 식단표"]["B7"].value)
        self.assertEqual(len(book["주간식단표(사진)"]._images), 0)

    def test_bad_file(self):
        with self.assertRaises(ValueError):
            wb.read_weekly(b"not an excel", "x.xlsm")

class MacroTest(unittest.TestCase):
    def test_strip_vba_keeps_xlsm(self):
        import zipfile as zf
        buf = io.BytesIO()
        with zf.ZipFile(buf, "w") as z:
            z.writestr("[Content_Types].xml", '<Types><Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml"/></Types>')
            z.writestr("xl/_rels/workbook.xml.rels", '<Relationships><Relationship Id="r9" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/></Relationships>')
            z.writestr("xl/vbaProject.bin", b"x")
        out = wb.strip_vba(buf.getvalue())
        z = zf.ZipFile(io.BytesIO(out))
        self.assertNotIn("xl/vbaProject.bin", z.namelist())
        self.assertIn("macroEnabled", z.read("[Content_Types].xml").decode())
        self.assertNotIn("vbaProject", z.read("[Content_Types].xml").decode())
        self.assertNotIn("vbaProject", z.read("xl/_rels/workbook.xml.rels").decode())
        self.assertFalse(wb.has_vba(out))


class LayoutTest(unittest.TestCase):
    def test_photo_boxes_follow_sheet(self):
        book = openpyxl.load_workbook(wb.TEMPLATE)
        kl = wb.keep_layout(book["보존식 기록지(사진)"])
        self.assertEqual(len(kl), 14)
        self.assertEqual(kl["d1_dinner_keep"]["row"], 33)     # 화요일 석식: 10행 간격이 아님
        self.assertEqual((kl["d1_dinner_keep"]["c1"], kl["d1_dinner_keep"]["c2"]), (3, 3))
        tl = wb.tray_layout(book["주간식단표(사진)"])
        self.assertEqual(tl["d6_dinner_tray"], {"row": 9, "c1": 5, "c2": 5})

    def test_photos_land_in_their_block(self):
        photos = {f"d{d}_{m}_keep": _jpg() for d in range(7) for m in ("lunch", "dinner")}
        _, data, info = wb.build_weekly(date(2026, 11, 23), _days(), {}, {}, photos)
        self.assertEqual(info["photos"], 14)
        back = wb.read_weekly(data)
        self.assertEqual(sorted(back["photos"]), sorted(photos))
        book = openpyxl.load_workbook(io.BytesIO(data))
        kl = wb.keep_layout(book["보존식 기록지(사진)"])
        rows = sorted(i.anchor._from.row + 1 for i in book["보존식 기록지(사진)"]._images)
        self.assertEqual(rows, sorted(b["row"] for b in kl.values()))


if __name__ == "__main__":
    unittest.main()

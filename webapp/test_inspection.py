"""식자재 검수서(원산지 표시) 웹 기능이 기존 도구의 표 모양 그대로 엑셀·워드를 만드는지 확인한다."""
import importlib.util
import io
import sys
import tempfile
import unittest
from datetime import date
from pathlib import Path

from PIL import Image

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))
import inspection as insp  # noqa: E402


def load_core():
    for cand in (ROOT.parent / "영수증_정리_도구.py", ROOT / "영수증_정리_도구.py"):
        if cand.exists():
            spec = importlib.util.spec_from_file_location("receipt_core_test", cand)
            mod = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(mod)
            return mod
    raise RuntimeError("영수증_정리_도구.py 를 찾지 못했습니다.")


def jpeg(w=800, h=600, color=(200, 80, 60), orientation=None):
    im = Image.new("RGB", (w, h), color)
    buf = io.BytesIO()
    if orientation:
        exif = Image.Exif()
        exif[0x0112] = orientation
        im.save(buf, "JPEG", exif=exif)
    else:
        im.save(buf, "JPEG")
    return buf.getvalue()


class Names(unittest.TestCase):
    def test_title_and_filename_with_date(self):
        t, f = insp.title_and_filename(date(2026, 7, 20))
        self.assertEqual(t, "식자재 검수서 원산지 표시  2026년 7월 20일")
        self.assertEqual(f, "원산지 표시_26.07.20")

    def test_title_without_date_uses_today_in_filename(self):
        t, f = insp.title_and_filename(None, today=date(2026, 11, 12))
        self.assertEqual(t, "식자재 검수서 원산지 표시")
        self.assertEqual(f, "원산지 표시_26.11.12")

    def test_parse_date(self):
        self.assertEqual(insp.parse_date("2026-11-12"), date(2026, 11, 12))
        self.assertIsNone(insp.parse_date(""))
        self.assertIsNone(insp.parse_date("11/12"))

    def test_clean_stem_and_order(self):
        self.assertEqual(insp.clean_stem("돼지:앞다리.jpg"), "돼지_앞다리")
        names = sorted(["사진10.jpg", "사진2.jpg", "사진1.jpg"], key=insp.natural_key)
        self.assertEqual(names, ["사진1.jpg", "사진2.jpg", "사진10.jpg"])

    def test_phone_rotation_is_applied(self):
        with tempfile.TemporaryDirectory() as d:
            out = insp.prepare_image(jpeg(800, 600, orientation=6), Path(d) / "a")
            with Image.open(out) as im:
                self.assertEqual(im.size, (600, 800))


class Build(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.core = load_core()

    def test_excel(self):
        from openpyxl import load_workbook
        uploads = [(f"품목{i}.jpg", jpeg(color=(i * 20, 100, 150))) for i in (10, 2, 1, 3)]
        uploads.append(("깨진사진.jpg", b"not an image"))
        uploads.append(("문서.pdf", b"%PDF"))
        with tempfile.TemporaryDirectory() as d:
            out, nice, skipped, n = insp.build_inspection(self.core, uploads, Path(d), "2026-11-12", "xlsx")
            self.assertEqual(nice, "원산지 표시_26.11.12.xlsx")
            self.assertEqual(n, 4)
            self.assertEqual(len(skipped), 2)
            ws = load_workbook(out).active
            self.assertEqual(ws["A1"].value, "식자재 검수서 원산지 표시  2026년 11월 12일")
            # 3열: 첫 줄 이름 행, 다음 줄 사진 행, 그다음 둘째 줄 이름 행
            self.assertEqual([ws.cell(3, c).value for c in (1, 2, 3)], ["품목1", "품목2", "품목3"])
            self.assertEqual(ws.cell(5, 1).value, "품목10")
            self.assertEqual(len(ws._images), 4)

    def test_names_with_dots_are_kept(self):
        from openpyxl import load_workbook
        with tempfile.TemporaryDirectory() as d:
            out, *_ = insp.build_inspection(self.core, [("1. 돼지고기 앞다리.jpg", jpeg()), ("2. 쌀.png", jpeg())],
                                            Path(d), "2026-11-12")
            ws = load_workbook(out).active
            self.assertEqual([ws.cell(3, 1).value, ws.cell(3, 2).value], ["1. 돼지고기 앞다리", "2. 쌀"])

    def test_word(self):
        from docx import Document
        with tempfile.TemporaryDirectory() as d:
            out, nice, skipped, n = insp.build_inspection(self.core, [("a.png", jpeg())], Path(d), "", "docx")
            self.assertTrue(nice.endswith(".docx"))
            Document(str(out))

    def test_no_usable_photo(self):
        with tempfile.TemporaryDirectory() as d:
            with self.assertRaises(ValueError):
                insp.build_inspection(self.core, [("x.heic", b"..")], Path(d))


class Wiring(unittest.TestCase):
    def test_app_and_page_are_connected(self):
        app = (ROOT / "app.py").read_text(encoding="utf-8")
        html = (ROOT / "static" / "index.html").read_text(encoding="utf-8")
        self.assertIn('@app.post("/api/inspection")', app)
        self.assertIn("build_inspection(core", app)
        self.assertIn("fetch('/api/inspection'", html)
        self.assertIn('id="insp"', html)


if __name__ == "__main__":
    unittest.main()

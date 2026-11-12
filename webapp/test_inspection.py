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

    def test_portrait_photo_is_not_stretched(self):
        """세로 사진이 가로 상자(4:3)로 늘어나 뚱뚱해지면 안 된다."""
        from openpyxl import load_workbook
        with tempfile.TemporaryDirectory() as d:
            out, *_ = insp.build_inspection(self.core, [("세로.jpg", jpeg(600, 1200)), ("가로.jpg", jpeg(1200, 900))],
                                            Path(d), "2026-11-12")
            imgs = load_workbook(out).active._images
            ratios = sorted(i.width / i.height for i in imgs)
            self.assertAlmostEqual(ratios[0], 0.5, delta=0.02)
            self.assertAlmostEqual(ratios[1], 4 / 3, delta=0.02)

    def test_user_rotation_is_applied(self):
        from openpyxl import load_workbook
        with tempfile.TemporaryDirectory() as d:
            out, *_ = insp.build_inspection(self.core, [("누운사진.jpg", jpeg(1200, 600), 90)], Path(d))
            img = load_workbook(out).active._images[0]
            self.assertLess(img.width, img.height)      # 가로 사진을 90도 돌려 세로가 됐다

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


class Orientation(unittest.TestCase):
    def test_norm_rotation(self):
        self.assertEqual([insp.norm_rotation(v) for v in (0, 90, "180", 270, 360, -90, 45, None, "x")],
                         [0, 90, 180, 270, 0, 270, 0, 0, 0])

    def test_rotation_direction_is_clockwise(self):
        im = Image.new("RGB", (20, 10), "white")
        im.putpixel((0, 0), (255, 0, 0))            # 왼쪽 위 빨간 점
        r = insp.rotate_cw(im, 90)
        self.assertEqual(r.size, (10, 20))
        self.assertEqual(r.getpixel((9, 0)), (255, 0, 0))   # 시계 방향 90도면 오른쪽 위로 간다

    def test_sheet_has_four_cells(self):
        png = insp.orientation_sheet(jpeg(800, 400), cell=200)
        with Image.open(io.BytesIO(png)) as im:
            self.assertEqual(im.size, (400, 400))

    def test_detect_maps_letter_to_degrees(self):
        class FakeCore:
            def __init__(self, letter):
                self.letter = letter

            def json_ask(self, client, png, model, schema=None, prompt=None):
                return {"upright": self.letter}
        got = [insp.detect_rotation(FakeCore(l), None, "m", jpeg(100, 80)) for l in "ABCDz"]
        self.assertEqual(got, [0, 90, 180, 270, 0])


class AutoName(unittest.TestCase):
    def test_date_from_name(self):
        f = insp.date_from_name
        self.assertEqual(f("KakaoTalk_20261109_135452116.jpg"), "26.11.09")
        self.assertEqual(f("KakaoTalk_20261109_173713784_03.jpg"), "26.11.09")
        self.assertEqual(f("20261109.jpg"), "26.11.09")
        self.assertEqual(f("2026-11-09 참치.png"), "26.11.09")
        self.assertEqual(f("26.11.09 참치.jpg"), "26.11.09")
        self.assertEqual(f("26.11.09_2.jpg"), "26.11.09")
        self.assertIsNone(f("IMG_0042.jpg"))
        self.assertIsNone(f("20261399.jpg"))        # 13월 99일
        self.assertIsNone(f(""))

    def test_normalize_origin(self):
        n = insp.normalize_origin
        self.assertEqual([n(v) for v in ("러시아", "러시아산", "국산", "한국산", "대한민국", "원산지: 중국", " 미국 ")],
                         ["러시아산", "러시아산", "국내산", "국내산", "국내산", "중국산", "미국산"])
        self.assertIsNone(n(None))
        self.assertIsNone(n(""))
        self.assertIsNone(n("null"))

    def test_compose_name(self):
        c = insp.compose_name
        self.assertEqual(c("참치 통조림", "26.11.09", "러시아"), "참치 통조림 26.11.09 러시아산")
        self.assertEqual(c("두부", None, "국산"), "두부 국내산")
        self.assertEqual(c("두부", "26.11.09", None), "두부 26.11.09")
        self.assertEqual(c("돼지/고기:*", "26.11.09", "국내산"), "돼지고기 26.11.09 국내산")
        self.assertIsNone(c(None, "26.11.09", "국내산"))      # 품명을 못 읽으면 이름을 바꾸지 않는다
        self.assertIsNone(c("null", "26.11.09", "국내산"))

    def test_read_label_with_fake_ai(self):
        class FakeCore:
            def json_ask(self, client, png, model, schema=None, prompt=None):
                self.png = png
                return {"product": "참치 통조림", "origin": "러시아"}
        fc = FakeCore()
        got = insp.read_label(fc, None, "m", jpeg(300, 200), "KakaoTalk_20261109_135452116.jpg")
        self.assertEqual(got["name"], "참치 통조림 26.11.09 러시아산")
        self.assertEqual((got["date"], got["origin"]), ("26.11.09", "러시아산"))
        self.assertTrue(fc.png.startswith(b"\xff\xd8"))       # JPEG 로 보낸다

    def test_read_label_unreadable(self):
        class FakeCore:
            def json_ask(self, *a, **k):
                return {"product": None, "origin": None}
        got = insp.read_label(FakeCore(), None, "m", jpeg(), "a.jpg")
        self.assertIsNone(got["name"])

    def test_renamed_photos_keep_order_and_names_in_excel(self):
        from openpyxl import load_workbook
        core = load_core()
        ups = [("참치 통조림 26.11.09 러시아산.jpg", jpeg()), ("가지 26.11.09 국내산.jpg", jpeg())]
        with tempfile.TemporaryDirectory() as d:
            out, *_ = insp.build_inspection(core, ups, Path(d), "2026-11-12", sort=False)
            ws = load_workbook(out).active
            # 점(26.11.09)이 들어 있어도 이름이 잘리지 않고, 화면 순서대로 놓인다
            self.assertEqual([ws.cell(3, 1).value, ws.cell(3, 2).value],
                             ["참치 통조림 26.11.09 러시아산", "가지 26.11.09 국내산"])

    def test_photos_zip(self):
        import zipfile
        ups = [("참치 26.11.09 러시아산.jpg", jpeg(800, 400), 90), ("참치 26.11.09 러시아산.png", jpeg(), 0),
               ("문서.pdf", b"%PDF"), ("깨짐.jpg", b"nope")]
        with tempfile.TemporaryDirectory() as d:
            n, skipped = insp.photos_zip(ups, Path(d) / "p.zip")
            self.assertEqual(n, 2)
            self.assertEqual(len(skipped), 2)
            with zipfile.ZipFile(Path(d) / "p.zip") as z:
                self.assertEqual(sorted(z.namelist()), ["참치 26.11.09 러시아산.jpg", "참치 26.11.09 러시아산_2.jpg"])
                with Image.open(io.BytesIO(z.read("참치 26.11.09 러시아산.jpg"))) as im:
                    self.assertEqual(im.size, (400, 800))      # 90도 돌린 방향이 반영됐다

    def test_date_stamp(self):
        self.assertEqual(insp.inspection_date_stamp("2026-11-12"), "26.11.12")
        self.assertEqual(insp.inspection_date_stamp("", today=date(2026, 1, 2)), "26.01.02")


class Wiring(unittest.TestCase):
    def test_app_and_page_are_connected(self):
        app = (ROOT / "app.py").read_text(encoding="utf-8")
        html = (ROOT / "static" / "index.html").read_text(encoding="utf-8")
        self.assertIn('@app.post("/api/inspection")', app)
        self.assertIn("build_inspection(core", app)
        self.assertIn("fetchAndSave('/api/inspection'", html)
        self.assertIn('id="insp"', html)
        self.assertIn('@app.post("/api/inspection/orient")', app)
        self.assertIn('@app.post("/api/inspection/read")', app)
        self.assertIn('@app.post("/api/inspection/zip")', app)
        self.assertIn("fetch('/api/inspection/read'", html)
        self.assertIn("'/api/inspection/zip'", html)
        self.assertIn('id="inspRead"', html)
        self.assertIn('id="inspZip"', html)
        self.assertIn("fetch('/api/inspection/orient'", html)
        self.assertIn("fd.append('rotations'", html)
        self.assertIn(">목록 비우기<", html)
        self.assertNotIn("사진 모두 빼기", html)


if __name__ == "__main__":
    unittest.main()

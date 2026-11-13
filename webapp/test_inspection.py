"""식자재 검수서(원산지 표시) 웹 기능이 기존 도구의 표 모양 그대로 엑셀·워드를 만드는지 확인한다."""
import importlib.util
import io
import re
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

    def test_add_months(self):
        a = insp.add_months
        self.assertEqual(a(date(2026, 11, 13), 24), date(2028, 11, 13))
        self.assertEqual(a(date(2026, 11, 13), 18), date(2028, 5, 13))
        self.assertEqual(a(date(2026, 11, 13), 12), date(2027, 11, 13))
        self.assertEqual(a(date(2026, 11, 13), 15), date(2028, 2, 13))
        self.assertEqual(a(date(2026, 8, 31), 6), date(2027, 2, 28))      # 날짜가 없는 달은 말일로
        self.assertEqual(a(date(2024, 2, 29), 12), date(2025, 2, 28))
        self.assertEqual(a(date(2026, 1, 31), 11), date(2026, 12, 31))

    def test_parse_iso(self):
        p = insp.parse_iso
        self.assertEqual(p("2026-11-13"), date(2026, 11, 13))
        self.assertEqual(p("26.11.13"), date(2026, 11, 13))
        self.assertEqual(p("2026/1/3"), date(2026, 1, 3))
        self.assertIsNone(p("2026-13-01"))
        self.assertIsNone(p("1999-01-01"))
        self.assertIsNone(p(None))
        self.assertIsNone(p("모름"))

    def test_pick_date_rules(self):
        pk = insp.pick_date
        # 유통기한이 직접 적혀 있으면 그 날짜
        self.assertEqual(pk({"expiry_date": "2027-05-01", "base_date": "2026-11-13", "shelf_months": 24}, "x.jpg")[0], "27.05.01")
        # 제조일 + 24개월까지 -> 기준일에 24개월
        self.assertEqual(pk({"base_date": "2026-11-13", "shelf_months": 24}, "x.jpg")[0], "28.11.13")
        self.assertEqual(pk({"base_date": "2026-11-13", "shelf_months": 18}, "x.jpg")[0], "28.05.13")
        self.assertEqual(pk({"base_date": "2026-11-13", "shelf_months": 12}, "x.jpg")[0], "27.11.13")
        self.assertEqual(pk({"base_date": "2026-11-13", "shelf_months": 15}, "x.jpg")[0], "28.02.13")
        # 기간 없이 제조일만 있으면 그 날짜
        self.assertEqual(pk({"base_date": "2026-11-13"}, "x.jpg")[0], "26.11.13")
        # 사진에 날짜가 없으면 파일 이름
        d, note = pk({}, "KakaoTalk_20261109_135452116.jpg")
        self.assertEqual(d, "26.11.09")
        self.assertIn("파일 이름", note)
        self.assertEqual(pk({}, "IMG_1.jpg")[0], None)
        # 엉뚱한 값은 무시 (개월 0, 너무 큰 값, 유통기한이 제조일보다 앞)
        self.assertEqual(pk({"base_date": "2026-11-13", "shelf_months": 0}, "x.jpg")[0], "26.11.13")
        self.assertEqual(pk({"base_date": "2026-11-13", "shelf_months": 999}, "x.jpg")[0], "26.11.13")
        self.assertEqual(pk({"base_date": "2026-11-13", "expiry_date": "2025-01-01", "shelf_months": 24}, "x.jpg")[0], "28.11.13")
        self.assertEqual(pk({"base_date": "2026-11-13", "shelf_months": "24"}, "x.jpg")[0], "28.11.13")

    def test_read_label_uses_date_in_photo(self):
        class FakeCore:
            def json_ask(self, *a, **k):
                return {"product": "두부", "origin": "국산", "base_date": "2026-11-13", "expiry_date": None, "shelf_months": 24}
        got = insp.read_label(FakeCore(), None, "m", jpeg(), "KakaoTalk_20261109_1.jpg")
        self.assertEqual(got["name"], "두부 28.11.13 국내산")      # 파일 이름의 26.11.09 보다 사진 속 날짜가 우선
        self.assertIn("24개월", got["date_note"])

    def test_vegetable_names(self):
        v = insp.is_vegetable_name
        for name in ("표고", "대파", "무우", "무", "양파", "브로콜리", "양배추", "찰가지", "오이", "느타리", "새송이", "청경채", "치커리",
                     "열무", "얼갈이", "아욱", "유채", "느타리버섯", "새송이버섯", "쌈채소", "깻잎", "쑥갓", "콩나물", "파"):
            self.assertTrue(v(name), name)
        for name in ("배추김치", "참치 통조림", "냉동 브로콜리", "건고추", "고추장", "다진 마늘", "감자칩", "돼지고기 앞다리",
                     "계란", "두부", "말린 표고", "토마토 케첩", "참기름", "무말랭이", "무진장"):
            self.assertFalse(v(name), name)
        self.assertFalse(v(None))

    def test_decide_origin(self):
        d = insp.decide_origin
        # 읽은 원산지는 채소여도 그대로 (중국산 양파는 중국산)
        self.assertEqual(d("양파", "중국", None)[0], "중국산")
        # 원산지를 못 읽은 채소는 국내산
        got, note = d("대파", None, None)
        self.assertEqual(got, "국내산")
        self.assertIn("채소", note)
        self.assertEqual(d("아욱", "", False)[0], "국내산")
        # 사전에 없는 채소도 AI 가 신선 채소라고 하면 국내산
        self.assertEqual(d("다래기", None, True)[0], "국내산")
        # 채소가 아니면 원산지를 비워 둔다
        self.assertIsNone(d("돼지고기 앞다리", None, False)[0])
        self.assertIsNone(d("참치 통조림", None, True)[0])        # 가공품은 AI 가 채소라고 해도 안 함
        self.assertIsNone(d("배추김치", None, True)[0])
        self.assertIsNone(d(None, None, True)[0])

    def test_read_label_vegetable_without_origin(self):
        class FakeCore:
            def __init__(self, ans):
                self.ans = ans

            def json_ask(self, *a, **k):
                return self.ans
        got = insp.read_label(FakeCore({"product": "대파", "origin": None, "base_date": None}), None, "m", jpeg(),
                              "KakaoTalk_20261109_1.jpg")
        self.assertEqual(got["name"], "대파 26.11.09 국내산")
        self.assertTrue(got["origin_note"])
        got = insp.read_label(FakeCore({"product": "돼지고기 앞다리", "origin": None}), None, "m", jpeg(), "KakaoTalk_20261109_1.jpg")
        self.assertEqual(got["name"], "돼지고기 앞다리 26.11.09")
        self.assertEqual(got["origin_note"], "")

    def test_two_dates_later_is_expiry(self):
        pk = insp.pick_date
        # 날짜 두 개(이름표 없음): 늦은 날짜가 유통기한. AI 가 이른 날짜를 유통기한이라 해도 바로잡는다
        self.assertEqual(pk({"all_dates": ["2025-11-28", "2027-11-28"]}, "x.jpg")[0], "27.11.28")
        self.assertEqual(pk({"expiry_date": "2025-11-28", "all_dates": ["2027-11-28", "2025-11-28"]}, "x.jpg")[0], "27.11.28")
        self.assertEqual(pk({"base_date": "2025-11-28", "all_dates": ["2025-11-28", "2027-11-28"], "shelf_months": 24}, "x.jpg")[0], "27.11.28")
        # 날짜가 하나뿐이면 기존 규칙 그대로
        self.assertEqual(pk({"base_date": "2026-11-13", "all_dates": ["2026-11-13"], "shelf_months": 24}, "x.jpg")[0], "28.11.13")
        self.assertEqual(pk({"all_dates": ["쉽지않음", None, 5]}, "KakaoTalk_20261109_1.jpg")[0], "26.11.09")

    def test_known_packages(self):
        refs = insp.load_references()
        self.assertEqual([(r[1], r[2]) for r in refs], [("닭고기", "브라질"), ("미니 찐빵", "")])
        for data, *_ in refs:
            self.assertEqual(data[:2], b"\xff\xd8")
            self.assertLess(len(data), 200_000)
            with Image.open(io.BytesIO(data)) as im:
                self.assertLessEqual(max(im.size), 900)

        # AI 에 보내는 내용: 참고 사진 A, B 를 글자와 함께 먼저, 판단할 사진은 마지막에
        class T:
            class Part:
                @staticmethod
                def from_bytes(data, mime_type):
                    return ("img", len(data))
        c = insp._match_contents(T, b"target", "image/jpeg", refs)
        self.assertIn("[참고 사진 A]", c[0])
        self.assertIn("[참고 사진 B]", c[2])
        self.assertEqual(c[-3:-1], ["[판단할 사진]", ("img", 6)])
        self.assertIs(c[-1], insp.MATCH_PROMPT)

    def test_match_reference_letter_to_index(self):
        refs = insp.load_references()
        orig = insp._ask_json
        try:
            for answer, want in (("A", 0), ("B", 1), ("none", None), ("C", None), ("", None)):
                insp._ask_json = lambda *a, _ans=answer, **k: {"match": _ans}
                self.assertEqual(insp.match_reference(None, "m", b"x", refs), want, answer)
        finally:
            insp._ask_json = orig

    def test_matched_package_sets_product_but_reads_date_from_photo(self):
        class FakeCore:
            def json_ask(self, *a, **k):
                # AI 가 품명을 못 읽었고 유통기한만 읽음
                return {"product": None, "origin": None, "expiry_date": "2027-09-29", "all_dates": ["2027-09-29"]}
        got = insp.read_label(FakeCore(), None, "m", jpeg(), "KakaoTalk_20261113_17.jpg", matcher=lambda: 1)
        self.assertEqual(got["name"], "미니 찐빵 27.09.29")          # 품명은 알려진 포장, 날짜는 사진에서
        self.assertIn("미니 찐빵", got["ref_note"])
        # 알려진 포장에 원산지가 있으면 AI 가 못 읽었을 때만 채운다
        got = insp.read_label(FakeCore(), None, "m", jpeg(), "KakaoTalk_20261113_17.jpg", matcher=lambda: 0)
        self.assertEqual(got["name"], "닭고기 27.09.29 브라질산")
        # AI 가 읽은 원산지가 있으면 그대로 둔다
        class Read:
            def json_ask(self, *a, **k):
                return {"product": "냉동 닭고기", "origin": "미국", "expiry_date": "2027-09-29"}
        got = insp.read_label(Read(), None, "m", jpeg(), "a.jpg", matcher=lambda: 0)
        self.assertEqual(got["name"], "닭고기 27.09.29 미국산")
        # 미니 찐빵은 원료마다 원산지가 달라 AI 가 읽은 원산지(중국)를 이름에 적지 않는다
        class Cn:
            def json_ask(self, *a, **k):
                return {"product": None, "origin": "중국", "expiry_date": "2027-09-29"}
        got = insp.read_label(Cn(), None, "m", jpeg(), "a.jpg", matcher=lambda: 1)
        self.assertEqual(got["name"], "미니 찐빵 27.09.29")
        # 같은 포장이 아니면(None) AI 가 읽은 대로, 알아보기에 실패해도(예외) 그대로
        got = insp.read_label(Read(), None, "m", jpeg(), "a.jpg", matcher=lambda: None)
        self.assertEqual(got["name"], "냉동 닭고기 27.09.29 미국산")
        self.assertEqual(got["ref_note"], "")

        def boom():
            raise RuntimeError("x")
        got = insp.read_label(Read(), None, "m", jpeg(), "a.jpg", matcher=boom)
        self.assertEqual(got["name"], "냉동 닭고기 27.09.29 미국산")

    def test_chicken_brazil_name(self):
        class FakeCore:
            def json_ask(self, *a, **k):
                return {"product": "닭고기", "origin": "브라질", "base_date": "2025-11-28", "expiry_date": None,
                        "all_dates": ["2025-11-28", "2027-11-28"]}
        got = insp.read_label(FakeCore(), None, "m", jpeg(), "KakaoTalk_20261113_175723998_07.jpg")
        self.assertEqual(got["name"], "닭고기 27.11.28 브라질산")
        self.assertIn("유통기한", got["date_note"])

    def test_unreadable_product_still_shows_expiry(self):
        class FakeCore:
            def __init__(self, ans):
                self.ans = ans

            def json_ask(self, *a, **k):
                return self.ans
        # 품명은 못 읽었지만 유통기한(2027.09.29까지)은 읽음
        got = insp.read_label(FakeCore({"product": None, "origin": None, "expiry_date": "2027-09-29", "all_dates": ["2027-09-29"]}),
                              None, "m", jpeg(), "KakaoTalk_20261113_175723998_17.jpg")
        self.assertEqual(got["name"], "품명미확인 27.09.29")
        self.assertTrue(got["name_note"])
        # 원산지만 읽었을 때도 적는다 (날짜는 파일 이름에서만 있으니 뽐다)
        got = insp.read_label(FakeCore({"product": None, "origin": "브라질"}), None, "m", jpeg(), "KakaoTalk_20261113_1.jpg")
        self.assertEqual(got["name"], "품명미확인 브라질산")
        # 아무것도 못 읽었고 날짜는 파일 이름뿐 -> 이름을 바꾸지 않는다
        got = insp.read_label(FakeCore({"product": None, "origin": None}), None, "m", jpeg(), "KakaoTalk_20261113_1.jpg")
        self.assertIsNone(got["name"])
        self.assertEqual(got["name_note"], "")
        # 품명을 읽으면 예전처럼
        got = insp.read_label(FakeCore({"product": "두부", "origin": "국산", "expiry_date": "2027-09-29"}), None, "m", jpeg(), "a.jpg")
        self.assertEqual((got["name"], got["name_note"]), ("두부 27.09.29 국내산", ""))

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
        # "검수서 엑셀 받기" 가 버튼 줄의 맨 왜쪽(첫 번째)
        row = html[html.index('id="inspProblems"') - 1500: html.index('id="inspProblems"')]
        ids = re.findall(r'<button[^>]*id="(insp\w+)"', row)
        self.assertEqual(ids[0], "inspGo")
        # 화면 미리보기만 크게 한다. 엑셀 쪽 사진 크기(155x116)는 그대로
        self.assertIn("--thumb", html)
        core_src = (ROOT.parent / "영수증_정리_도구.py").read_text(encoding="utf-8")
        self.assertIn("img_w, img_h = 155, 116", core_src)
        self.assertIn('id="inspZip"', html)
        self.assertIn("fetch('/api/inspection/orient'", html)
        self.assertIn("fd.append('rotations'", html)
        self.assertIn(">목록 비우기<", html)
        self.assertNotIn("사진 모두 빼기", html)


if __name__ == "__main__":
    unittest.main()

"""업로드 파일 이름을 단순하게 정리하는 규칙을 확인한다. (cv2 없이도 돌도록 필요한 정의만 꺼내 쓴다)"""
import ast
import os
import re
import unittest
from pathlib import Path

HERE = Path(__file__).resolve().parent


def extract(path, names):
    tree = ast.parse(path.read_text(encoding="utf-8"))
    body = [n for n in tree.body
            if getattr(n, "name", None) in names
            or any(getattr(t, "id", None) in names for t in getattr(n, "targets", []))]
    return body


CORE_NAMES = ("_KAKAO_PREFIX", "_KAKAO_RE", "_kakao_stem", "normalize_stem", "guess_year")
ns = {"re": re, "os": os}
exec(compile(ast.Module(body=extract(HERE.parent / "영수증_정리_도구.py", CORE_NAMES), type_ignores=[]),
             "core", "exec"), ns)


class Core:
    normalize_stem = staticmethod(ns["normalize_stem"])
    guess_year = staticmethod(ns["guess_year"])


ans = {"re": re, "core": Core, "Path": Path}
exec(compile(ast.Module(body=extract(HERE / "app.py", ("_decode_filename", "_guess_year_hint", "simple_upload_stem")),
                        type_ignores=[]), "app", "exec"), ans)
simple = ans["simple_upload_stem"]


class SimpleNames(unittest.TestCase):
    def test_kakao_variants(self):
        cases = {
            "KakaoTalk_20260813_142355123": "26.08.13",
            "KakaoTalk_20260813_142355123_01": "26.08.13",
            "KakaoTalk_Photo_2026-08-13-14-23-55": "26.08.13",
            "KakaoTalk_Photo_2026-08-13-14-23-55 001": "26.08.13",
            "KakaoTalk_Image_2026-08-13-14-23-55": "26.08.13",
            "카카오톡_20260813_142355": "26.08.13",
            "kakaotalk_20260813_142355 (1)": "26.08.13",
        }
        for src, want in cases.items():
            self.assertEqual(simple(src), want, src)

    def test_same_day_numbering_in_upload(self):
        # create_job 의 번호 붙이는 규칙: 첫 장은 그대로, 다음부터 _2, _3
        used, out = set(), []
        for s in ("KakaoTalk_20260813_142355123", "KakaoTalk_20260813_142401456", "KakaoTalk_Photo_2026-08-13-15-00-00"):
            stem = simple(s)
            final, n = stem, 2
            while final.lower() in used:
                final = f"{stem}_{n}"
                n += 1
            used.add(final.lower())
            out.append(final)
        self.assertEqual(out, ["26.08.13", "26.08.13_2", "26.08.13_3"])

    def test_not_kakao_or_invalid_left_alone(self):
        for s in ("영수증 사진", "IMG_1234", "KakaoTalk_20261399_142355123", "KakaoTalk_20260813_256100"):
            self.assertEqual(simple(s), s)

    def test_date_names_still_work(self):
        self.assertEqual(simple("20260810_3"), "26.08.10_3")
        self.assertEqual(simple("8.21 충북농산", "26"), "26.08.21 충북농산")

    def test_year_hint_from_names(self):
        self.assertEqual(ans["_guess_year_hint"](["26.8.13 킹마트.jpg", "26.8.14.jpg"]), "26")


if __name__ == "__main__":
    unittest.main()

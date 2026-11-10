"""처리 로그에서 '읽지 못한 파일'을 골라내는 규칙을 확인한다."""
import ast
import re
import unittest
from pathlib import Path

SRC = (Path(__file__).resolve().parent / "app.py").read_text(encoding="utf-8")
TREE = ast.parse(SRC)


def load(*names):
    """app.py 전체를 import 하지 않고(서버 스레드가 뜬다) 필요한 정의만 꺼낸다."""
    body = [n for n in TREE.body
            if getattr(n, "name", None) in names
            or any(getattr(t, "id", None) in names for t in getattr(n, "targets", []))]
    ns = {"re": re}
    exec(compile(ast.Module(body=body, type_ignores=[]), "app.py", "exec"), ns)
    return ns


NS = load("SERVICE_TROUBLE_SHORT", "clean_receipt_name", "ProblemTracker", "problem_lines")
ProblemTracker, problem_lines = NS["ProblemTracker"], NS["problem_lines"]


def feed_all(lines):
    t = ProblemTracker()
    for line in lines:
        t.feed(line)
    return t.items


class ProblemTrackerTests(unittest.TestCase):
    def test_cleaning_failure(self):
        items = feed_all(["[완료] a.jpg -> a_정리.jpg", "[실패] b.jpg: cannot identify image file"])
        self.assertEqual(items, [{"name": "b.jpg", "reason": "cannot identify image file"}])

    def test_load_failure_in_reading(self):
        items = feed_all(["[2/5] 실패 26.08.05_정리.jpg: broken"])
        self.assertEqual(items[0]["name"], "26.08.05.jpg")

    def test_ai_failure_uses_previous_label(self):
        items = feed_all(["[1/3] x_정리.jpg", "    총액 100원 / 품목 2개", "[2/3] y_정리.jpg", "    실패: 응답 시간 초과"])
        self.assertEqual(items, [{"name": "y.jpg", "reason": "응답 시간 초과"}])

    def test_recheck_failure_is_not_a_problem(self):
        items = feed_all(["[1/1] x.jpg", "    ** 확인 필요", "[재판독 1/1] x.jpg", "    실패: 시간 초과  (1차 결과를 그대로 둡니다)"])
        self.assertEqual(items, [])

    def test_fatal_stop_is_reported(self):
        items = feed_all(["[중단] API 키가 올바르지 않습니다"])
        self.assertEqual(items[0]["name"], "남은 영수증")

    def test_fatal_stop_hides_technical_reason(self):
        items = feed_all(["[중단] models/x is not found for API version v1beta"])
        self.assertNotIn("models/x", items[0]["reason"])

    def test_stop_during_recheck_is_not_a_problem(self):
        items = feed_all(["[재판독 1/1] x.jpg", "[중단] quota"])
        self.assertEqual(items, [])

    def test_duplicates_are_listed_once(self):
        items = feed_all(["[실패] b.jpg: x", "[실패] b.jpg: x"])
        self.assertEqual(len(items), 1)

    def test_problem_lines_merges_skipped_and_failed(self):
        lines = problem_lines({"skipped": ["big.jpg (10MB 초과)"], "failed": [{"name": "b.jpg", "reason": "x"}]})
        self.assertEqual(lines, ["big.jpg (10MB 초과)", "b.jpg (x)"])
        self.assertEqual(problem_lines({}), [])


if __name__ == "__main__":
    unittest.main()
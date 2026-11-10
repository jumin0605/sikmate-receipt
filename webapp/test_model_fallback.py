"""모델이 없어졌을 때 대신 쓸 모델을 찾는 규칙을 확인한다 (cv2 없이 필요한 함수만 꺼내 시험)."""
import ast
import os
import re
import unittest
from pathlib import Path

SRC = (Path(__file__).resolve().parent.parent / "영수증_정리_도구.py").read_text(encoding="utf-8")
TREE = ast.parse(SRC)
WANT = {"JsonSetupError", "JsonFatalError", "json_short_err", "json_is_fatal", "json_model_gone",
        "json_discover_models", "json_resolve_model", "json_preflight", "JSON_FALLBACK_MODELS"}


def load(fallback_env=""):
    os.environ["GEMINI_FALLBACK_MODELS"] = fallback_env
    body = [n for n in TREE.body
            if getattr(n, "name", None) in WANT
            or any(getattr(t, "id", None) in WANT for t in getattr(n, "targets", []))]
    ns = {"re": re, "os": os}
    exec(compile(ast.Module(body=body, type_ignores=[]), "core", "exec"), ns)
    return ns


class Model:
    def __init__(self, name, actions=("generateContent",)):
        self.name, self.supported_actions = name, list(actions)


class FakeModels:
    def __init__(self, alive, listed=()):
        self.alive, self.listed, self.pinged = set(alive), list(listed), []

    def generate_content(self, model, contents):
        self.pinged.append(model)
        if model not in self.alive:
            raise RuntimeError("404 NOT_FOUND. {'error': {'message': 'models/%s is not found for API version v1beta'}}" % model)

    def list(self):
        return iter(self.listed)


class FakeClient:
    def __init__(self, alive, listed=()):
        self.models = FakeModels(alive, listed)


class ModelFallbackTests(unittest.TestCase):
    def test_working_model_is_used_as_is(self):
        ns = load()
        c = FakeClient({"gemini-3.8-flash"})
        self.assertEqual(ns["json_resolve_model"](c, "gemini-3.8-flash", log=lambda *_: None), "gemini-3.8-flash")
        self.assertEqual(c.models.pinged, ["gemini-3.8-flash"])

    def test_gone_model_falls_back_to_newest_listed_flash(self):
        ns = load()
        listed = [Model("models/gemini-2.5-flash"), Model("models/gemini-4.0-flash"),
                  Model("models/gemini-4.0-flash-lite"), Model("models/gemini-5.0-flash-preview"),
                  Model("models/gemini-4.0-pro"), Model("models/gemini-9-flash", actions=["embedContent"])]
        c = FakeClient({"gemini-4.0-flash", "gemini-2.5-flash"}, listed)
        logs = []
        got = ns["json_resolve_model"](c, "gemini-3.8-flash", log=logs.append)
        self.assertEqual(got, "gemini-4.0-flash")
        self.assertTrue(any("[모델 대체]" in l for l in logs))

    def test_pro_role_looks_for_pro_models(self):
        ns = load()
        listed = [Model("models/gemini-4.0-flash"), Model("models/gemini-3.5-pro")]
        c = FakeClient({"gemini-3.5-pro"}, listed)
        self.assertEqual(ns["json_resolve_model"](c, "gemini-4-argon-pro", log=lambda *_: None), "gemini-3.5-pro")

    def test_env_fallback_list_goes_first(self):
        ns = load("gemini-2.5-flash, gemini-3.0-flash")
        c = FakeClient({"gemini-2.5-flash", "gemini-4.0-flash"}, [Model("models/gemini-4.0-flash")])
        self.assertEqual(ns["json_resolve_model"](c, "gemini-3.8-flash", log=lambda *_: None), "gemini-2.5-flash")

    def test_alias_is_last_resort(self):
        ns = load()
        c = FakeClient({"gemini-flash-latest"})
        self.assertEqual(ns["json_resolve_model"](c, "gemini-3.8-flash", log=lambda *_: None), "gemini-flash-latest")

    def test_nothing_available_raises_setup_error(self):
        ns = load()
        c = FakeClient(set())
        with self.assertRaises(ns["JsonSetupError"]):
            ns["json_resolve_model"](c, "gemini-3.8-flash", log=lambda *_: None)

    def test_bad_key_stops_without_trying_other_models(self):
        ns = load()
        c = FakeClient(set())
        def boom(model, contents):
            c.models.pinged.append(model)
            raise RuntimeError("400 API_KEY_INVALID API key not valid")
        c.models.generate_content = boom
        with self.assertRaises(ns["JsonSetupError"]):
            ns["json_resolve_model"](c, "gemini-3.8-flash", log=lambda *_: None)
        self.assertEqual(c.models.pinged, ["gemini-3.8-flash"])

    def test_temporary_error_does_not_switch_model(self):
        ns = load()
        c = FakeClient(set())
        def flaky(model, contents):
            raise RuntimeError("503 UNAVAILABLE overloaded")
        c.models.generate_content = flaky
        self.assertEqual(ns["json_resolve_model"](c, "gemini-3.8-flash", log=lambda *_: None), "gemini-3.8-flash")


if __name__ == "__main__":
    unittest.main()
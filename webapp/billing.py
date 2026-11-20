# -*- coding: utf-8 -*-
"""토스페이먼츠 결제(토스페이·카드) + 이용권 관리 - 표준 라이브러리만 쓴다.

흐름: 결제창 열기(브라우저, 클라이언트 키) -> 성공 주소로 돌아옴 -> 서버가 결제 승인(시크릿 키) -> 이용 기간 연장
- 이용권은 한 번 결제하면 PLAN_DAYS(기본 30일) 동안 쓸 수 있다. 이미 이용 중이면 남은 기간 뒤에 이어 붙는다.
- 가입하면 TRIAL_DAYS(기본 14일) 무료 체험. 결제를 꼭 해야 하게 만드는 스위치는 BILLING_ENFORCE=1 (기본은 꺼짐).
- 계정별 기록: <사용자 폴더>/billing.json

환경변수
  TOSS_CLIENT_KEY   토스페이먼츠 클라이언트 키 (test_ck_... / live_ck_...)
  TOSS_SECRET_KEY   토스페이먼츠 시크릿 키 (test_sk_... / live_sk_...)  절대 화면·깃허브에 올리지 말 것
  PLAN_NAME         이용권 이름 (기본 식단온 월 이용권)
  PLAN_PRICE        1회 결제 금액, 원 (기본 9900)
  PLAN_DAYS         1회 결제로 늘어나는 일수 (기본 30)
  TRIAL_DAYS        가입 뒤 무료 체험 일수 (기본 14)
  BILLING_ENFORCE   1 이면 체험·이용 기간이 끝난 계정은 결제 전까지 도구를 못 쓴다
  BILLING_FREE_USERS  결제 없이 계속 쓰는 아이디 목록(쉼표로 구분)
"""

from __future__ import annotations

import base64
import json
import os
import secrets
import threading
import time
import urllib.error
import urllib.request
from pathlib import Path

CONFIRM_URL = "https://api.tosspayments.com/v1/payments/confirm"
_lock = threading.RLock()


def _int_env(name: str, default: int) -> int:
    try:
        return int(os.environ.get(name, "").strip() or default)
    except ValueError:
        return default


def cfg() -> dict:
    return {
        "client_key": os.environ.get("TOSS_CLIENT_KEY", "").strip(),
        "secret_key": os.environ.get("TOSS_SECRET_KEY", "").strip(),
        "name": os.environ.get("PLAN_NAME", "").strip() or "식단온 월 이용권",
        "price": max(100, _int_env("PLAN_PRICE", 9900)),
        "days": max(1, _int_env("PLAN_DAYS", 30)),
        "trial": max(0, _int_env("TRIAL_DAYS", 14)),
        "enforce": os.environ.get("BILLING_ENFORCE", "").strip() == "1",
        "free": {u.strip().lower() for u in os.environ.get("BILLING_FREE_USERS", "").split(",") if u.strip()},
    }


def configured() -> bool:
    c = cfg()
    return bool(c["client_key"] and c["secret_key"])


def _path(user_dir: Path) -> Path:
    return Path(user_dir) / "billing.json"


def _load(user_dir: Path) -> dict:
    try:
        d = json.loads(_path(user_dir).read_text(encoding="utf-8"))
        if isinstance(d, dict):
            d.setdefault("paid_until", 0)
            d.setdefault("orders", {})
            return d
    except Exception:  # noqa: BLE001
        pass
    return {"paid_until": 0, "orders": {}}


def _save(user_dir: Path, d: dict):
    p = _path(user_dir)
    p.parent.mkdir(parents=True, exist_ok=True)
    tmp = p.with_suffix(".tmp")
    tmp.write_text(json.dumps(d, ensure_ascii=False, indent=1), encoding="utf-8")
    os.replace(tmp, p)


def status(rec: dict, user_dir: Path, now: float | None = None) -> dict:
    """이 계정의 이용 상태. active=지금 도구를 써도 되는지"""
    now = now or time.time()
    c = cfg()
    d = _load(user_dir)
    trial_end = int(rec.get("created", 0)) + c["trial"] * 86400
    paid_until = int(d.get("paid_until", 0))
    free = rec["username"].lower() in c["free"]
    in_paid = paid_until > now
    in_trial = trial_end > now
    until = max(paid_until, trial_end)
    active = (not c["enforce"]) or free or in_paid or in_trial
    kind = "free" if free else "paid" if in_paid else "trial" if in_trial else "expired"
    return {"active": active, "kind": kind, "until": int(until), "paid_until": paid_until,
            "trial_end": int(trial_end), "days_left": max(0, int((until - now + 86399) // 86400)),
            "enforce": c["enforce"]}


def new_order(user_dir: Path) -> dict:
    """결제 전에 주문을 만든다. 금액은 서버가 정해 두고, 승인할 때 그 금액과 맞는지 확인한다."""
    c = cfg()
    with _lock:
        d = _load(user_dir)
        pend = [k for k, v in d["orders"].items() if v.get("state") == "ready"]
        for k in sorted(pend, key=lambda k: d["orders"][k]["t"])[:-20]:     # 오래된 미결제 주문 정리
            d["orders"].pop(k, None)
        oid = "sm_" + secrets.token_hex(10)
        d["orders"][oid] = {"amount": c["price"], "days": c["days"], "name": c["name"], "state": "ready", "t": int(time.time())}
        _save(user_dir, d)
    return {"orderId": oid, "amount": c["price"], "orderName": c["name"]}


class PayError(ValueError):
    """화면에 그대로 보여 줘도 되는 안내 문구"""


def _toss_confirm(payment_key: str, order_id: str, amount: int) -> dict:
    secret = cfg()["secret_key"]
    auth = base64.b64encode((secret + ":").encode()).decode()
    body = json.dumps({"paymentKey": payment_key, "orderId": order_id, "amount": amount}).encode()
    req = urllib.request.Request(CONFIRM_URL, data=body, method="POST", headers={
        "Authorization": "Basic " + auth, "Content-Type": "application/json", "Idempotency-Key": order_id})
    try:
        with urllib.request.urlopen(req, timeout=30) as r:
            return json.loads(r.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        try:
            msg = json.loads(e.read().decode("utf-8")).get("message") or ""
        except Exception:  # noqa: BLE001
            msg = ""
        raise PayError(msg or "결제 승인에 실패했습니다.") from None
    except Exception:  # noqa: BLE001
        raise PayError("결제 서버와 연결하지 못했습니다. 잠시 뒤 다시 시도해 주세요.") from None


def confirm(user_dir: Path, payment_key: str, order_id: str, amount: int, _confirm=None) -> dict:
    """결제 성공 주소로 돌아왔을 때 호출. 이용 기간을 늘리고 결과를 돌려준다."""
    _confirm = _confirm or _toss_confirm
    with _lock:
        d = _load(user_dir)
        o = d["orders"].get(order_id)
        if o is None:
            raise PayError("주문을 찾지 못했습니다. 다시 결제해 주세요.")
        if o.get("state") == "paid":          # 새로고침 등으로 두 번 들어와도 기간이 두 번 늘지 않는다
            return {"already": True, "paid_until": d["paid_until"], "name": o["name"], "amount": o["amount"]}
        if int(amount) != int(o["amount"]):
            raise PayError("결제 금액이 주문과 맞지 않아 처리하지 않았습니다.")
        res = _confirm(payment_key, order_id, int(o["amount"]))
        if res.get("status") not in ("DONE", None) or int(res.get("totalAmount", o["amount"])) != int(o["amount"]):
            raise PayError("결제가 완료되지 않았습니다.")
        now = int(time.time())
        base = max(now, int(d.get("paid_until", 0)))
        d["paid_until"] = base + int(o["days"]) * 86400
        o.update({"state": "paid", "paidAt": now, "paymentKey": payment_key, "method": res.get("method", "")})
        _save(user_dir, d)
        return {"already": False, "paid_until": d["paid_until"], "name": o["name"], "amount": o["amount"]}


def history(user_dir: Path) -> list[dict]:
    d = _load(user_dir)
    rows = [{"orderId": k, "name": v["name"], "amount": v["amount"], "paidAt": v.get("paidAt", 0), "method": v.get("method", "")}
            for k, v in d["orders"].items() if v.get("state") == "paid"]
    return sorted(rows, key=lambda r: -r["paidAt"])

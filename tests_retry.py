"""Offline tests for gemini_retry (no API key or network needed)."""
import os
from gemini_retry import generate_with_retry


class FakeModels:
    def __init__(self, outcomes):
        self.outcomes = list(outcomes)
        self.calls = []

    def generate_content(self, **kwargs):
        self.calls.append(kwargs["model"])
        outcome = self.outcomes.pop(0)
        if isinstance(outcome, Exception):
            raise outcome
        return outcome


class FakeClient:
    def __init__(self, outcomes):
        self.models = FakeModels(outcomes)


OVERLOADED = RuntimeError("503 UNAVAILABLE. The model is overloaded.")


def test_retries_then_succeeds():
    client = FakeClient([OVERLOADED, OVERLOADED, "ok"])
    assert generate_with_retry(client, model="m", contents="x", sleep=lambda s: None) == "ok"
    assert client.models.calls == ["m", "m", "m"]


def test_other_errors_not_retried():
    client = FakeClient([ValueError("400 bad request")])
    try:
        generate_with_retry(client, model="m", contents="x", sleep=lambda s: None)
    except ValueError:
        assert client.models.calls == ["m"]
    else:
        raise AssertionError("expected ValueError")


def test_fallback_model_used_when_still_overloaded():
    os.environ["GEMINI_FALLBACK_MODEL"] = "backup"
    try:
        client = FakeClient([OVERLOADED, OVERLOADED, OVERLOADED, "ok"])
        assert generate_with_retry(client, model="m", contents="x", sleep=lambda s: None) == "ok"
        assert client.models.calls == ["m", "m", "m", "backup"]
    finally:
        del os.environ["GEMINI_FALLBACK_MODEL"]


def test_raises_after_retries_without_fallback():
    os.environ.pop("GEMINI_FALLBACK_MODEL", None)
    client = FakeClient([OVERLOADED, OVERLOADED, OVERLOADED])
    try:
        generate_with_retry(client, model="m", contents="x", sleep=lambda s: None)
    except RuntimeError as exc:
        assert "503" in str(exc)
    else:
        raise AssertionError("expected RuntimeError")


if __name__ == "__main__":
    for name, fn in list(globals().items()):
        if name.startswith("test_"):
            fn()
            print("PASS", name)

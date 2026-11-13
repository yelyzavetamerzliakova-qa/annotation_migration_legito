"""Retry helper for transient Gemini errors (503 overloaded / UNAVAILABLE)."""
import os
import time

RETRY_DELAYS_SECONDS = (2, 5)
_TRANSIENT_MARKERS = ("503", "UNAVAILABLE", "overloaded")


def is_transient(exc: Exception) -> bool:
    if getattr(exc, "code", None) == 503:
        return True
    text = str(exc)
    return any(marker in text for marker in _TRANSIENT_MARKERS)


def generate_with_retry(client, *, model: str, sleep=time.sleep, **kwargs):
    """Call client.models.generate_content, retrying transient overload errors.

    Tries the requested model up to three times with a short backoff. If it is
    still overloaded and GEMINI_FALLBACK_MODEL is set to a different model ID,
    makes one last attempt with that model. Other errors are raised unchanged.
    """
    last_exc = None
    for attempt in range(len(RETRY_DELAYS_SECONDS) + 1):
        try:
            return client.models.generate_content(model=model, **kwargs)
        except Exception as exc:
            if not is_transient(exc):
                raise
            last_exc = exc
            if attempt < len(RETRY_DELAYS_SECONDS):
                sleep(RETRY_DELAYS_SECONDS[attempt])
    fallback = os.environ.get("GEMINI_FALLBACK_MODEL", "").strip()
    if fallback and fallback != model:
        return client.models.generate_content(model=fallback, **kwargs)
    raise last_exc

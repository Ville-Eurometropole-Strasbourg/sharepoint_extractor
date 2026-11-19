"""
Patch office365-rest-python-client 2.6.2 to send every HTTP call through one
shared requests.Session, with adaptive throttling.

- Session: replaces the library's direct requests.get/post calls. Only the
  transport changes, not the parameters.
- Throttle: a base delay always applies, doubled on strain, eased back on
  success.
- Thread-safe: one lock guards the delay state and the sleep.

Call enable_persistent_session() once, before creating any ClientContext.
"""

import logging
import threading
import time

import requests
from office365.runtime.client_request import ClientRequest
from office365.runtime.http.http_method import HttpMethod

import monitoring

logger = logging.getLogger("persistent_session")

_patched = False
_shared_session = None

BASE_DELAY_SECONDS = 0.3  # ~3 requests/s max
MAX_DELAY_SECONDS = 5.0
# Slower successful responses still count as strain.
SLOW_RESPONSE_THRESHOLD_SECONDS = 3.0
DELAY_INCREASE_FACTOR = 2.0
DELAY_DECREASE_FACTOR = 0.75
# 500s that are app errors, not strain ($expand on a deleted account).
KNOWN_APP_ERROR_MARKERS = (
    "Impossible de trouver l'utilisateur",
    "User cannot be found",
)
# Fast successes needed before easing down.
SUCCESS_STREAK_BEFORE_DECREASE = 5

_current_delay = BASE_DELAY_SECONDS
_success_streak = 0
# Also held during the sleep.
_throttle_lock = threading.Lock()


def get_shared_session():
    """Shared session, created on first call."""
    global _shared_session
    if _shared_session is None:
        _shared_session = requests.Session()
    return _shared_session


def get_current_delay():
    """Current throttle delay."""
    with _throttle_lock:
        return _current_delay


def _record_strain():
    """Back off. Caller holds _throttle_lock."""
    global _current_delay, _success_streak
    _success_streak = 0
    new_delay = min(_current_delay * DELAY_INCREASE_FACTOR, MAX_DELAY_SECONDS)
    if new_delay != _current_delay:
        logger.warning(
            f"Throttle: increasing delay {_current_delay:.2f}s -> {new_delay:.2f}s "
            f"after a slow/failed request."
        )
    _current_delay = new_delay


def _record_success(elapsed_seconds):
    """Ease the delay down after fast successes. Caller holds _throttle_lock."""
    global _current_delay, _success_streak
    if elapsed_seconds >= SLOW_RESPONSE_THRESHOLD_SECONDS:
        _record_strain()
        return
    _success_streak += 1
    if (
        _success_streak >= SUCCESS_STREAK_BEFORE_DECREASE
        and _current_delay > BASE_DELAY_SECONDS
    ):
        _current_delay = max(BASE_DELAY_SECONDS, _current_delay * DELAY_DECREASE_FACTOR)
        _success_streak = 0
        logger.info(f"Throttle: easing delay back down to {_current_delay:.2f}s.")


def enable_persistent_session():
    """Patch ClientRequest.execute_request_direct (idempotent)."""
    global _patched
    if _patched:
        return
    _patched = True

    session = get_shared_session()

    def _execute_request_direct_via_session(self, request):
        self.beforeExecute.notify(request)

        with _throttle_lock:
            time.sleep(_current_delay)

        start = time.monotonic()
        try:
            if request.method == HttpMethod.Post:
                if request.is_bytes or request.is_file:
                    response = session.post(
                        url=request.url,
                        headers=request.headers,
                        data=request.data,
                        auth=request.auth,
                        verify=request.verify,
                        proxies=request.proxies,
                    )
                else:
                    response = session.post(
                        url=request.url,
                        headers=request.headers,
                        json=request.data,
                        auth=request.auth,
                        verify=request.verify,
                        proxies=request.proxies,
                    )
            elif request.method == HttpMethod.Patch:
                response = session.patch(
                    url=request.url,
                    headers=request.headers,
                    json=request.data,
                    auth=request.auth,
                    verify=request.verify,
                    proxies=request.proxies,
                )
            elif request.method == HttpMethod.Delete:
                response = session.delete(
                    url=request.url,
                    headers=request.headers,
                    auth=request.auth,
                    verify=request.verify,
                    proxies=request.proxies,
                )
            elif request.method == HttpMethod.Put:
                response = session.put(
                    url=request.url,
                    data=request.data,
                    headers=request.headers,
                    auth=request.auth,
                    verify=request.verify,
                    proxies=request.proxies,
                )
            else:
                response = session.get(
                    url=request.url,
                    headers=request.headers,
                    auth=request.auth,
                    verify=request.verify,
                    stream=request.stream,
                    proxies=request.proxies,
                )
        except requests.exceptions.RequestException as e:
            with _throttle_lock:
                _record_strain()
            monitoring.record_call(
                method=request.method,
                url=request.url,
                duration_seconds=time.monotonic() - start,
                error=str(e),
            )
            raise

        elapsed = time.monotonic() - start

        # 429 and 5xx are strain, except known app errors.
        with _throttle_lock:
            if response.status_code >= 500 and any(
                m in response.text for m in KNOWN_APP_ERROR_MARKERS
            ):
                pass
            elif response.status_code == 429 or response.status_code >= 500:
                _record_strain()
            else:
                _record_success(elapsed)

        if getattr(request, "stream", False):
            content_length = response.headers.get("Content-Length")
            response_size = int(content_length) if content_length else None
        else:
            response_size = len(response.content) if response.content else 0

        monitoring.record_call(
            method=request.method,
            url=request.url,
            status_code=response.status_code,
            duration_seconds=elapsed,
            response_bytes=response_size,
        )

        response.raise_for_status()
        return response

    ClientRequest.execute_request_direct = _execute_request_direct_via_session
    logger.info(
        f"Persistent HTTP session with adaptive throttling enabled "
        f"(base delay: {BASE_DELAY_SECONDS}s, thread-safe)."
    )


def close_shared_session():
    """Release pooled connections."""
    global _shared_session
    if _shared_session is not None:
        _shared_session.close()
        _shared_session = None

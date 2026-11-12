"""Detect the model provider refusing requests (usage limit, outage) and wait it out.

A rollout or reflection that failed because the provider answered with an API error says
nothing about the skill, so it must never be scored: callers discard it, wait here, retry.
"""

import time

from agy_cli import is_quota_error, model_available

PROBE_POLL_SEC = 300
MAX_WAIT_SEC = 6 * 3600


class ModelUnavailable(RuntimeError):
    """The model stayed unavailable for the whole waiting window."""


def is_provider_failure(summary):
    """A run cut short by a quota error, or one that errored before the agent did any work."""
    if not summary.get("is_error"):
        return False
    return is_quota_error(summary.get("error")) or summary.get("tool_calls", 0) == 0


def wait_until_model_available(models, log=print, poll_sec=PROBE_POLL_SEC, max_wait_sec=MAX_WAIT_SEC):
    started = time.time()
    while not all(model_available(model) for model in models):
        if time.time() - started > max_wait_sec:
            raise ModelUnavailable(f"model unavailable for {max_wait_sec // 3600} h")
        log(f"model provider is refusing requests; probing again in {poll_sec // 60} min")
        time.sleep(poll_sec)

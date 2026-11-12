#!/usr/bin/env python3
"""Detect an Office app that stopped answering document Apple events, and wait for it.

A wedged app (seen with Word stuck inside the save of a MIP-protected document) still
answers `get version` but blocks on anything touching its documents, so the probe asks for
the document count. A stopped app counts as healthy: the runner starts it hidden. The probe
never launches, activates or quits an app.
"""

import argparse
import subprocess
import sys
import time

APP_NAMES = {"word": "Microsoft Word", "excel": "Microsoft Excel", "powerpoint": "Microsoft PowerPoint"}
COLLECTIONS = {"word": "documents", "excel": "workbooks", "powerpoint": "presentations"}
PROBE_TIMEOUT_SEC = 20
POLL_SEC = 120
MAX_WAIT_SEC = 3 * 3600


class OfficeUnavailable(RuntimeError):
    """The app stayed unresponsive for the whole waiting window."""


def app_responsive(app):
    name = APP_NAMES[app]
    script = (
        f'if application "{name}" is running then\n'
        f'\ttell application "{name}" to return (count of {COLLECTIONS[app]}) as text\n'
        "end if\n"
        'return "not running"'
    )
    try:
        completed = subprocess.run(["osascript", "-e", script], capture_output=True, text=True, timeout=PROBE_TIMEOUT_SEC)
    except subprocess.TimeoutExpired:
        return False
    return completed.returncode == 0


def wait_until_responsive(app, log=print, poll_sec=POLL_SEC, max_wait_sec=MAX_WAIT_SEC):
    """Return the seconds waited; raise OfficeUnavailable after ``max_wait_sec``."""
    started = time.time()
    if app_responsive(app):
        return 0.0
    log(f"{APP_NAMES[app]} does not answer document events; waiting (probe every {poll_sec}s, up to {max_wait_sec // 60} min)")
    while time.time() - started < max_wait_sec:
        time.sleep(poll_sec)
        if app_responsive(app):
            waited = time.time() - started
            log(f"{APP_NAMES[app]} answers again after {waited / 60:.1f} min")
            return waited
    raise OfficeUnavailable(f"{APP_NAMES[app]} stayed unresponsive for {max_wait_sec // 60} min")


def main():
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--app", choices=sorted(APP_NAMES), default="word")
    parser.add_argument("--wait", action="store_true", help="Block until the app answers (or give up)")
    args = parser.parse_args()
    if not args.wait:
        healthy = app_responsive(args.app)
        print("responsive" if healthy else "unresponsive")
        return 0 if healthy else 1
    try:
        wait_until_responsive(args.app, log=lambda message: print(message, flush=True))
    except OfficeUnavailable as error:
        print(error, flush=True)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())

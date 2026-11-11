"""Run one AppleScript against an Office app without disturbing the user.

Every Office call of the runner goes through :func:`run_in_app`, which

* serializes calls per app with an ``fcntl`` lock (scripts address ``active document``, so
  two concurrent runs would operate on each other's files);
* starts a stopped app hidden (``open -g -j``) instead of letting AppleScript launch it in
  the foreground, and keeps an app that is hidden hidden;
* never hides an app the user has visible: that app only gets the focus guard;
* closes the documents the run left open, using the staged paths only;
* puts the user's frontmost app back if an Office app took the focus (``focus_restored``);
* remembers which app the runner itself launched, and quits it on a later call once it sat
  unused for ``IDLE_QUIT_SECONDS`` while still hidden and without documents.

Window and focus state is read through NSWorkspace (JXA), which needs no permission;
hiding goes through System Events (Automation permission for System Events).
"""

import fcntl
import json
import os
import subprocess
import tempfile
import time
from contextlib import contextmanager

from .applescript import (
    close_staged_documents,
    run_jxa_json,
    run_osascript,
    run_osascript_file,
    timeout_hint,
)
from .constants import (
    APP_NAMES,
    BUNDLE_IDS,
    DOCUMENT_COLLECTIONS,
    IDLE_QUIT_SECONDS,
    RUNNER_STATE_DIRNAME,
)
from .staging import staging_root

LAUNCH_TIMEOUT = 30

# Replaced in every script by "true" when the app runs hidden for this call, so a script
# can minimize its own window when the app is the user's visible one.
APP_HIDDEN_TOKEN = "@APP_HIDDEN@"

DESKTOP_STATE_JXA = """// mac-office-desktop-state
ObjC.import('AppKit');
function describe(app) {
  return {name: app.localizedName.js, pid: app.processIdentifier, hidden: app.isHidden,
          launched: app.finishedLaunching};
}
const front = $.NSWorkspace.sharedWorkspace.frontmostApplication;
const matches = $.NSRunningApplication.runningApplicationsWithBundleIdentifier('@BUNDLE@');
JSON.stringify({front: front.isNil() ? null : describe(front),
                app: matches.count > 0 ? describe(matches.objectAtIndex(0)) : null});
"""

RESTORE_FOCUS_JXA = """ObjC.import('AppKit');
const app = $.NSRunningApplication.runningApplicationWithProcessIdentifier(@PID@);
JSON.stringify(!app.isNil() && app.activateWithOptions($.NSApplicationActivateAllWindows));
"""


class AppRun:
    """Outcome of :func:`run_in_app`: ``completed`` is None when the script did not finish."""

    def __init__(self):
        self.completed = None
        self.timed_out = False
        self.error = None
        self.error_hint = None
        self.partial_output = ""
        self.focus_restored = False


def run_in_app(app, source, timeout, env=None, staged_paths=()):
    """Run AppleScript ``source`` against ``app`` (see the module docstring)."""
    run = AppRun()
    with app_lock(app):
        quit_idle_apps(skip=app)
        before = desktop_state(app)
        launched_pid = None
        if before is not None and before["app"] is None:
            try:
                launched_pid = launch_hidden(app)
            except RuntimeError as error:
                run.error = str(error)
                return run
        app_hidden = launched_pid is not None or bool(before and before["app"]["hidden"])
        if app_hidden:
            hide_app(app)
        try:
            script = source.replace(APP_HIDDEN_TOKEN, "true" if app_hidden else "false")
            run.completed = _run_script(script, timeout, env)
        except subprocess.TimeoutExpired as error:
            stdout = error.stdout or ""
            run.partial_output = (stdout.decode("utf-8", "replace") if isinstance(stdout, bytes) else stdout).strip()
            run.timed_out = True
            run.error = f"osascript timed out after {timeout}s"
            run.error_hint = timeout_hint(app)
        except OSError as error:
            run.error = f"Cannot run osascript: {error}"
        finally:
            close_staged_documents(app, staged_paths)
            if app_hidden:
                hide_app(app)
            after = desktop_state(app)
            run.focus_restored = guard_focus(before, after)
            remember_ownership(app, before, after, launched_pid)
    return run


def _run_script(source, timeout, env):
    script_path = None
    try:
        with tempfile.NamedTemporaryFile(
            "w", suffix=".applescript", delete=False, encoding="utf-8"
        ) as handle:
            handle.write(source)
            script_path = handle.name
        return run_osascript_file(script_path, timeout=timeout, env=env)
    finally:
        if script_path and os.path.exists(script_path):
            os.unlink(script_path)


def state_dir():
    path = os.path.join(staging_root(), RUNNER_STATE_DIRNAME)
    os.makedirs(path, exist_ok=True)
    return path


@contextmanager
def app_lock(app, blocking=True):
    """Hold the per-app lock; yields False instead of waiting when ``blocking`` is False."""
    with open(os.path.join(state_dir(), f"{app}.lock"), "a") as handle:
        try:
            fcntl.flock(handle, fcntl.LOCK_EX if blocking else fcntl.LOCK_EX | fcntl.LOCK_NB)
        except BlockingIOError:
            yield False
            return
        try:
            yield True
        finally:
            fcntl.flock(handle, fcntl.LOCK_UN)


def desktop_state(app):
    """``{"front": {...} | None, "app": {...} | None}``, or None when it cannot be read."""
    state = run_jxa_json(DESKTOP_STATE_JXA.replace("@BUNDLE@", BUNDLE_IDS[app]))
    if not isinstance(state, dict) or "front" not in state or "app" not in state:
        return None
    return state


def launch_hidden(app):
    """Start ``app`` hidden and in the background; returns its pid once it is scriptable."""
    name = APP_NAMES[app]
    deadline = time.time() + LAUNCH_TIMEOUT
    try:
        subprocess.run(["open", "-g", "-j", "-a", name], check=True, capture_output=True, timeout=LAUNCH_TIMEOUT)
    except (OSError, subprocess.SubprocessError) as error:
        raise RuntimeError(f"Cannot launch {name}: {error}") from None
    while time.time() < deadline:
        state = desktop_state(app)
        if state and state["app"] and state["app"]["launched"]:
            try:
                if run_osascript(f'tell application "{name}" to get name', timeout=10).returncode == 0:
                    return state["app"]["pid"]
            except subprocess.TimeoutExpired:
                pass
        time.sleep(0.2)
    raise RuntimeError(f"{name} did not become scriptable within {LAUNCH_TIMEOUT}s")


def hide_app(app):
    script = f'tell application "System Events" to set visible of process "{APP_NAMES[app]}" to false'
    try:
        run_osascript(script, timeout=10)
    except subprocess.TimeoutExpired:
        pass


def focus_was_stolen(before, after):
    """True when the frontmost app changed from a non-Office app to an Office app."""
    if not before or not after or not before["front"] or not after["front"]:
        return False
    office = set(APP_NAMES.values())
    return (
        after["front"]["pid"] != before["front"]["pid"]
        and after["front"]["name"] in office
        and before["front"]["name"] not in office
    )


def guard_focus(before, after):
    """Give the focus back to the user's app when the run stole it; True when restored."""
    if not focus_was_stolen(before, after):
        return False
    return run_jxa_json(RESTORE_FOCUS_JXA.replace("@PID@", str(int(before["front"]["pid"])))) is True


def _state_path(app):
    return os.path.join(state_dir(), f"{app}.json")


def read_ownership(app):
    """``{"pid", "last_used"}`` of the instance the runner launched, or None."""
    try:
        with open(_state_path(app), encoding="utf-8") as handle:
            state = json.load(handle)
        return {"pid": int(state["pid"]), "last_used": float(state["last_used"])}
    except (OSError, ValueError, KeyError, TypeError):
        return None


def _forget_ownership(app):
    try:
        os.unlink(_state_path(app))
    except FileNotFoundError:
        pass


def runner_owns(ownership, app_state):
    """The running instance is the one the runner launched, and the user has not taken it
    over (a user who shows the app makes it theirs)."""
    return bool(
        ownership and app_state and app_state["pid"] == ownership["pid"] and app_state["hidden"]
    )


def remember_ownership(app, before, after, launched_pid):
    """Refresh ``last_used`` for an instance the runner owns; forget any other instance."""
    app_state = after["app"] if after else None
    owned = app_state is not None and (
        app_state["pid"] == launched_pid
        or runner_owns(read_ownership(app), before["app"] if before else None)
    )
    if not owned:
        _forget_ownership(app)
        return
    with open(_state_path(app), "w", encoding="utf-8") as handle:
        json.dump({"pid": app_state["pid"], "last_used": time.time()}, handle)


def is_idle(ownership, now):
    return bool(ownership) and now - ownership["last_used"] > IDLE_QUIT_SECONDS


def quit_idle_apps(skip):
    """Quit apps the runner launched that sat unused, hidden and empty past the idle limit."""
    now = time.time()
    for app in APP_NAMES:
        if app == skip or not is_idle(read_ownership(app), now):
            continue
        with app_lock(app, blocking=False) as acquired:
            if not acquired:
                continue
            ownership = read_ownership(app)
            state = desktop_state(app)
            if state is None or not is_idle(ownership, now):
                continue
            if runner_owns(ownership, state["app"]) and not _quit_if_empty(app):
                continue
            _forget_ownership(app)


def _quit_if_empty(app):
    """``quit saving no`` when the app holds no documents; True when it was asked to quit."""
    name = APP_NAMES[app]
    collection = DOCUMENT_COLLECTIONS[app]
    script = (
        f'if application "{name}" is running then\n'
        f'\ttell application "{name}"\n'
        f"\t\tif (count of {collection}) is 0 then\n"
        "\t\t\tquit saving no\n"
        '\t\t\treturn "quit"\n'
        "\t\tend if\n"
        "\tend tell\n"
        '\treturn "kept"\n'
        "end if\n"
        'return "quit"'
    )
    try:
        completed = run_osascript(script, timeout=15)
    except subprocess.TimeoutExpired:
        return False
    return completed.returncode == 0 and completed.stdout.strip() == "quit"

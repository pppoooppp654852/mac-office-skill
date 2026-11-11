"""osascript transport (AppleScript + JXA), staged-document cleanup and dialog probes."""

import json
import subprocess

from .constants import (
    ALERT_HANDLERS,
    APP_NAMES,
    DOCUMENT_COLLECTIONS,
    MODAL_ALERT_HINT,
    NO_DIALOG_HINT,
    TCC_HINT,
    UNREADABLE_DIALOG_HINT,
)


def encode_string(value):
    """Render ``value`` as an AppleScript string literal (no raw newlines)."""
    escaped = str(value).replace("\\", "\\\\").replace('"', '\\"')
    return '"' + escaped + '"'


def build_script(app, user_script):
    """Wrap a user script in the per-app alert-suppression preamble and postamble."""
    preamble, postamble = ALERT_HANDLERS[app]
    return preamble + user_script.rstrip() + "\n" + postamble


def _run(command, timeout, env=None):
    """Run osascript; output is decoded as UTF-8 so non-ASCII survives."""
    return subprocess.run(
        command, capture_output=True, text=True, encoding="utf-8", timeout=timeout, env=env
    )


def run_osascript(source, timeout):
    return _run(["osascript", "-e", source], timeout)


def run_osascript_file(script_path, timeout, env=None):
    return _run(["osascript", script_path], timeout, env)


def run_jxa(script_path, arguments, timeout):
    """Run a JXA script file with ``osascript -l JavaScript``."""
    return _run(
        ["osascript", "-l", "JavaScript", script_path, *[str(argument) for argument in arguments]],
        timeout,
    )


def run_jxa_json(source, timeout=10):
    """Run inline JXA that returns JSON; None when it fails or answers something else."""
    try:
        completed = _run(["osascript", "-l", "JavaScript", "-e", source], timeout)
    except subprocess.TimeoutExpired:
        return None
    if completed.returncode != 0:
        return None
    try:
        return json.loads(completed.stdout)
    except ValueError:
        return None


def tcc_hint_for(error_text):
    """Return the Automation-permission hint when the error is a TCC denial."""
    lowered = (error_text or "").lower()
    return TCC_HINT if ("-1743" in lowered or "not authorized" in lowered) else None


def close_staged_documents(app, staged_paths):
    """Close (without saving) every open document whose file is one of ``staged_paths``.

    Staged files live in a fresh per-run directory, so a match can only be a document this
    run opened; anything else the user has open is never touched. Word and Excel report
    ``full name`` as an HFS path and PowerPoint as a POSIX path, so both forms are compared.
    Each document's name is re-read right before it is closed, so a document the user opens
    meanwhile (shifting the indices) is never closed by mistake.
    """
    if not staged_paths:
        return
    collection = DOCUMENT_COLLECTIONS[app]
    item = collection[:-1]
    wanted = ", ".join(encode_string(path) for path in sorted(staged_paths))
    script = (
        f"set wanted to {{{wanted}}}\n"
        "set wantedHFS to {}\n"
        "repeat with p in wanted\n"
        "\tset end of wantedHFS to ((POSIX file (p as text)) as text)\n"
        "end repeat\n"
        f'if application "{APP_NAMES[app]}" is running then\n'
        f'\ttell application "{APP_NAMES[app]}"\n'
        f"\t\trepeat with i from (count of {collection}) to 1 by -1\n"
        "\t\t\ttry\n"
        f"\t\t\t\tset fn to (full name of {item} i) as text\n"
        f"\t\t\t\tif wanted contains fn or wantedHFS contains fn then close {item} i saving no\n"
        "\t\t\tend try\n"
        "\t\tend repeat\n"
        "\tend tell\n"
        "end if\n"
    )
    try:
        run_osascript(script, timeout=15)
    except subprocess.TimeoutExpired:
        pass


def detect_modal_dialog(app):
    """True/False when the app's window state is readable, None when it is not.

    Word's unreadable-content prompt is a standalone ``AXDialog`` window that stays
    listed (and blocks every ``close``) while the app is hidden.
    """
    process_name = APP_NAMES[app]
    script = (
        f'tell application "System Events" to tell process "{process_name}"\n'
        '\tif not (exists window 1) then return false\n'
        '\tset winNames to name of every window\n'
        '\tif winNames contains "Grant File Access" then return true\n'
        '\tif exists (first window whose subrole is "AXDialog") then return true\n'
        '\tif (count of sheets of window 1) > 0 then return true\n'
        '\treturn false\n'
        'end tell'
    )
    try:
        completed = run_osascript(script, timeout=6)
    except subprocess.TimeoutExpired:
        return None
    if completed.returncode != 0:
        return None
    value = completed.stdout.strip().lower()
    if value in ("true", "false"):
        return value == "true"
    return None


def timeout_hint(app):
    """The ``error_hint`` for a run that hit its timeout."""
    modal = detect_modal_dialog(app)
    if modal is None:
        return UNREADABLE_DIALOG_HINT
    return MODAL_ALERT_HINT if modal else NO_DIALOG_HINT

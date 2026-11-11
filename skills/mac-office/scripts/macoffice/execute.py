"""``--app`` execution: alert suppression, per-run staging, assets, lazy write-back."""

import os
import shutil
import sys
import time

from .app_session import run_in_app
from .applescript import build_script, tcc_hint_for
from .constants import APP_NAMES, KIND_FOR_APP
from .container import staged_extension
from .staging import StageDir

def read_user_script(args):
    if args.file:
        with open(args.file, "r", encoding="utf-8") as handle:
            return handle.read()
    return sys.stdin.buffer.read().decode("utf-8", "replace")


def revised_path(target_file):
    stem, extension = os.path.splitext(target_file)
    return f"{stem}_revised{extension}"


def _stat(path):
    return os.stat(path) if path and os.path.isfile(path) else None


def _modified(before, after):
    return bool(
        before
        and after
        and (before.st_size, before.st_mtime_ns) != (after.st_size, after.st_mtime_ns)
    )


def run_execute(args):
    started = time.time()
    app = args.app.lower()
    result = {
        "success": False,
        "mode": "execute",
        "app": app,
        "target_file": None,
        "working_file": None,
        "output_file": None,
        "target_modified": False,
        "output_exists": False,
        "output_modified": False,
        "output": None,
        "focus_restored": False,
        "error": None,
        "error_hint": None,
        "execution_time_sec": 0.0,
    }
    if app not in APP_NAMES:
        result["error"] = f"Unknown app '{args.app}'. Use word, excel, or powerpoint."
        return _finish(result, started)

    try:
        user_script = read_user_script(args)
    except OSError as error:
        result["error"] = f"Cannot read AppleScript: {error}"
        return _finish(result, started)

    assets = [os.path.abspath(os.path.expanduser(path)) for path in (args.asset or [])]
    for asset in assets:
        if not os.path.isfile(asset):
            result["error"] = f"Asset file not found: {asset}"
            return _finish(result, started)

    target_file = os.path.abspath(os.path.expanduser(args.target_file)) if args.target_file else None
    if target_file and not os.path.isfile(target_file):
        result["error"] = f"Target file not found: {target_file}"
        return _finish(result, started)
    output_file = os.path.abspath(os.path.expanduser(args.output_file)) if args.output_file else None
    result["target_file"] = target_file
    result["output_file"] = output_file

    label = os.path.splitext(os.path.basename(target_file or output_file or app))[0]
    with StageDir(label=label) as stage:
        staged_target = None
        if target_file:
            # The staged copy may carry a legacy extension (.doc for a Word 97 payload named
            # .docx); write-back below always uses the user's own file name.
            stem = os.path.splitext(os.path.basename(target_file))[0]
            staged_target = stage.stage(target_file, stem + staged_extension(target_file, KIND_FOR_APP[app]))
        staged_output = None
        if output_file:
            os.makedirs(os.path.dirname(output_file) or ".", exist_ok=True)
            staged_output = stage.path_for(os.path.basename(output_file))
        staged_assets = [stage.stage(asset) for asset in assets]

        environment = dict(os.environ)
        environment["MAC_OFFICE_TARGET_FILE"] = staged_target or ""
        environment["MAC_OFFICE_OUTPUT_FILE"] = staged_output or ""
        environment["MAC_OFFICE_STAGE_DIR"] = stage.path
        environment["MAC_OFFICE_ORIGINAL_FILE"] = target_file or ""
        environment["MAC_OFFICE_ASSET_COUNT"] = str(len(staged_assets))
        for index, asset in enumerate(staged_assets, start=1):
            environment[f"MAC_OFFICE_ASSET_{index}"] = asset

        before_stat = _stat(staged_target)
        run = run_in_app(
            app,
            build_script(app, user_script),
            args.timeout,
            env=environment,
            staged_paths=[path for path in (staged_target, staged_output, *staged_assets) if path],
        )
        result["focus_restored"] = run.focus_restored
        if run.completed is None:
            result["error"] = run.error
            result["error_hint"] = run.error_hint
            result["output"] = run.partial_output
        else:
            result["output"] = (run.completed.stdout or "").strip()
            if run.completed.returncode != 0:
                error_text = (run.completed.stderr or run.completed.stdout or "").strip()
                result["error"] = error_text
                result["error_hint"] = tcc_hint_for(error_text)
            else:
                result["success"] = True

        if staged_target:
            modified = _modified(before_stat, _stat(staged_target))
            result["target_modified"] = modified
            if result["success"] and modified:
                destination = target_file if args.in_place else revised_path(target_file)
                shutil.copy2(staged_target, destination)
                result["working_file"] = destination

        if output_file:
            if result["success"] and os.path.isfile(staged_output) and os.path.getsize(staged_output) > 0:
                shutil.copy2(staged_output, output_file)
                result["output_exists"] = True
                result["output_modified"] = True
            else:
                result["output_exists"] = os.path.isfile(output_file)
        elif result["working_file"]:
            result["output_exists"] = True
            result["output_modified"] = True
    return _finish(result, started)


def _finish(result, started):
    result["execution_time_sec"] = round(time.time() - started, 2)
    return result

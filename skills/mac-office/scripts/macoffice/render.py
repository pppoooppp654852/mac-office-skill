"""``--render``: export an Office file to PDF with its app, then rasterize with PDFKit.

Both halves use Apple stack only: the Office app speaks AppleScript, the PDF is rasterized
and its text lines measured by ``osascript -l JavaScript`` (PDFKit/AppKit), called with
``pdf_render.js`` next to this runner. Text is never re-flowed by us: coordinates come from
PDFKit text selection bounds, in PDF points with a top-left origin.

The exported PDF is cached by the source's sha256 (it does not depend on ``--dpi``), so
rendering an unchanged file again only re-rasterizes and never opens the Office app.
"""

import hashlib
import json
import os
import shutil
import tempfile
import time
from subprocess import TimeoutExpired

from .app_session import run_in_app
from .applescript import build_script, encode_string, run_jxa, tcc_hint_for
from .constants import APP_BY_EXTENSION, APP_NAMES, KIND_FOR_APP, OPEN_FAILED, PDF_CACHE_DIR
from .container import staged_extension
from .staging import StageDir, safe_name

SCRIPT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
JXA_SCRIPT = os.path.join(SCRIPT_DIR, "pdf_render.js")

DEFAULT_DPI = 110
RASTERIZE_TIMEOUT = 300

# Each script opens the staged copy read-only, finds it by its unique staged name (never
# ``active document``: a swallowed open must not make it export or close a user document),
# and minimizes its own window when the app is the user's visible one.
EXPORT_SCRIPTS = {
    "word": """
tell application "Microsoft Word"
	set targetPath to system attribute "MAC_OFFICE_TARGET_FILE"
	set outPath to system attribute "MAC_OFFICE_OUTPUT_FILE"
	open (targetPath as POSIX file) read only true add to recent files false
	if not (exists document @NAME@) then return "@OPEN_FAILED@"
	set d to document @NAME@
	if not @APP_HIDDEN@ then
		try
			set window state of window 1 of d to window state minimize
		end try
	end if
	save as d file name outPath file format format PDF
	try
		close d saving no
	end try
end tell
""",
    "excel": """
tell application "Microsoft Excel"
	set targetPath to system attribute "MAC_OFFICE_TARGET_FILE"
	set outPath to system attribute "MAC_OFFICE_OUTPUT_FILE"
	open workbook workbook file name targetPath read only true add to mru false
	if not (exists workbook @NAME@) then return "@OPEN_FAILED@"
	set wb to workbook @NAME@
	save workbook as wb filename outPath file format PDF file format
	try
		close wb saving no
	end try
end tell
""",
    "powerpoint": """
tell application "Microsoft PowerPoint"
	set targetPath to system attribute "MAC_OFFICE_TARGET_FILE"
	set outPath to system attribute "MAC_OFFICE_OUTPUT_FILE"
	open (targetPath as POSIX file)
	repeat 25 times
		if exists presentation @NAME@ then exit repeat
		delay 0.1
	end repeat
	if not (exists presentation @NAME@) then return "@OPEN_FAILED@"
	set pres to presentation @NAME@
	save pres in (outPath as POSIX file) as save as PDF
	delay 1
	set hasPDF to false
	try
		set pdfAlias to (outPath as POSIX file) as alias
		set hasPDF to true
	end try
	set protectedFlag to false
	try
		set protectedFlag to is protected of pres
	end try
	try
		close pres saving no
	end try
	if hasPDF then return "ok"
	if protectedFlag then return "protected"
	return "missing"
end tell
""",
}


def run_render(args):
    started = time.time()
    path = os.path.abspath(os.path.expanduser(args.render))
    result = {
        "success": False,
        "mode": "render",
        "file": path,
        "pdf": None,
        "cached": False,
        "pages": 0,
        "png": [],
        "lines": [],
        "focus_restored": False,
        "error": None,
        "error_hint": None,
        "execution_time_sec": 0.0,
    }
    if not os.path.isfile(path):
        result["error"] = f"File not found: {path}"
        return _finish(result, started)
    extension = os.path.splitext(path)[1].lower()
    if extension != ".pdf" and extension not in APP_BY_EXTENSION:
        result["error"] = "Unsupported file type. Expected .docx, .xlsx, .pptx, or .pdf."
        return _finish(result, started)

    stem = os.path.splitext(os.path.basename(path))[0]
    if args.output_dir:
        output_dir = os.path.abspath(os.path.expanduser(args.output_dir))
        os.makedirs(output_dir, exist_ok=True)
    else:
        output_dir = tempfile.mkdtemp(prefix=f"mac-office-render-{safe_name(stem, limit=40)}-")
    dpi = max(10, min(args.dpi, 600))

    if extension == ".pdf":
        result["pdf"] = path
    else:
        # Never overwrite a PDF already in the output directory: pick the next free name.
        destination = unused_path(output_dir, stem, ".pdf")
        cached = os.path.join(PDF_CACHE_DIR, file_digest(path) + ".pdf")
        if os.path.isfile(cached):
            shutil.copy2(cached, destination)
            result["cached"] = True
        elif not _export_pdf(APP_BY_EXTENSION[extension], path, destination, cached, args.timeout, result):
            return _finish(result, started)
        result["pdf"] = destination

    png_prefix = os.path.splitext(os.path.basename(result["pdf"]))[0]
    try:
        completed = run_jxa(
            JXA_SCRIPT, [result["pdf"], output_dir, dpi, png_prefix], timeout=RASTERIZE_TIMEOUT
        )
    except TimeoutExpired:
        result["error"] = f"PDF rasterization timed out after {RASTERIZE_TIMEOUT}s"
        return _finish(result, started)
    except OSError as error:
        result["error"] = f"Cannot run osascript: {error}"
        return _finish(result, started)
    try:
        rendered = json.loads(completed.stdout)
    except ValueError:
        result["error"] = (
            completed.stderr.strip()
            or completed.stdout.strip()
            or "PDF rasterization returned no JSON"
        )
        return _finish(result, started)
    result["pages"] = int(rendered["pages"])
    result["png"] = rendered["png"]
    result["lines"] = rendered["lines"]
    result["success"] = True
    return _finish(result, started)


def file_digest(path):
    digest = hashlib.sha256()
    with open(path, "rb") as handle:
        for chunk in iter(lambda: handle.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def unused_path(directory, stem, extension):
    """``directory/stem.ext``, or ``stem-2.ext``, ``stem-3.ext``... when that is taken."""
    candidate = os.path.join(directory, stem + extension)
    counter = 1
    while os.path.exists(candidate):
        counter += 1
        candidate = os.path.join(directory, f"{stem}-{counter}{extension}")
    return candidate


def _export_pdf(app, path, destination, cached, timeout, result):
    """Export ``path`` to PDF at ``destination`` and into the cache; False on error."""
    if not os.path.isdir(f"/Applications/{APP_NAMES[app]}.app"):
        result["error"] = f"{APP_NAMES[app]} is not installed in /Applications"
        return False
    with StageDir(label=os.path.basename(path)) as stage:
        staged_target = stage.stage_uniquely(path, staged_extension(path, KIND_FOR_APP[app]))
        staged_pdf = stage.path_for("export.pdf")
        environment = dict(os.environ)
        environment["MAC_OFFICE_TARGET_FILE"] = staged_target
        environment["MAC_OFFICE_OUTPUT_FILE"] = staged_pdf
        environment["MAC_OFFICE_STAGE_DIR"] = stage.path
        environment["MAC_OFFICE_ORIGINAL_FILE"] = path
        script = (
            EXPORT_SCRIPTS[app]
            .replace("@NAME@", encode_string(os.path.basename(staged_target)))
            .replace("@OPEN_FAILED@", OPEN_FAILED)
        )
        run = run_in_app(
            app,
            build_script(app, script),
            timeout,
            env=environment,
            staged_paths=[staged_target, staged_pdf],
        )
        result["focus_restored"] = run.focus_restored
        if run.completed is None:
            result["error"] = f"PDF export timed out after {timeout}s" if run.timed_out else run.error
            result["error_hint"] = run.error_hint
            return False
        completed = run.completed
        if completed.returncode != 0:
            error_text = (completed.stderr or completed.stdout or "").strip()
            result["error"] = error_text
            result["error_hint"] = tcc_hint_for(error_text)
            return False
        status = (completed.stdout or "").strip()
        if status == OPEN_FAILED:
            result["error"] = (
                f"{APP_NAMES[app]} could not open {path} (it ignored the request); the format "
                "may be blocked by policy or the extension may not match the container"
            )
            return False
        if not os.path.isfile(staged_pdf) or os.path.getsize(staged_pdf) == 0:
            if status == "protected":
                result["error"] = (
                    f"{APP_NAMES[app]} could not export {path}: the presentation is "
                    "MIP/IRM-protected (is protected: true) and the app writes no PDF for it"
                )
                result["error_hint"] = (
                    "Export the deck through PowerPoint's own File > Export UI, or ask for an "
                    "unprotected copy first."
                )
            else:
                result["error"] = f"{APP_NAMES[app]} did not produce a PDF for {path}"
            return False
        # Staged names are ASCII-safe for Office; the deliverable keeps the user's own name.
        shutil.copy2(staged_pdf, destination)
        os.makedirs(PDF_CACHE_DIR, exist_ok=True)
        partial = f"{cached}.{os.getpid()}.part"
        shutil.copy2(staged_pdf, partial)
        os.replace(partial, cached)
        return True


def _finish(result, started):
    result["execution_time_sec"] = round(time.time() - started, 2)
    return result

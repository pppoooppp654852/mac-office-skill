"""``--inspect`` dispatch: headless OpenXML first, native app for OLE2 containers."""

import os
import time
import xml.etree.ElementTree as ET
import zipfile

from . import openxml
from .constants import APP_FOR_KIND, APP_NAMES
from .container import OPENXML, container_hint, detect_container
from .fallback import inspect_via_app

INSPECTORS = {
    ".docx": ("docx", "word/document.xml", openxml.inspect_docx),
    ".xlsx": ("xlsx", "xl/workbook.xml", openxml.inspect_xlsx),
    ".pptx": ("pptx", "ppt/presentation.xml", openxml.inspect_pptx),
}


def run_inspect(args):
    started = time.time()
    path = os.path.abspath(os.path.expanduser(args.inspect))
    result = {
        "success": False,
        "mode": "inspect",
        "file": path,
        "kind": None,
        "container": None,
        "inspector": None,
        "hint": None,
        "focus_restored": False,
        "error": None,
        "error_hint": None,
        "execution_time_sec": 0.0,
    }
    if not os.path.isfile(path):
        result["error"] = f"File not found: {path}"
        return _finish(result, started)
    inspector = INSPECTORS.get(os.path.splitext(path)[1].lower())
    if inspector is None:
        result["error"] = "Unsupported file type. Expected .docx, .xlsx, or .pptx."
        return _finish(result, started)
    kind, required_part, inspect = inspector
    result["kind"] = kind
    container = detect_container(path)
    result["container"] = container
    result["hint"] = container_hint(container)

    # OLE2 first: a MIP container embeds a readable theme zip, so zipfile would succeed and
    # the required-part check would report a bogus "missing word/document.xml".
    if container != OPENXML:
        try:
            inspected, result["focus_restored"] = inspect_via_app(path, kind, args.max_items, args.timeout)
            result.update(inspected)
            result["success"] = True
        except Exception as error:  # noqa: BLE001 - any failure degrades to a report
            result["error"] = (
                f"File is an {container} {kind} container and the "
                f"{APP_NAMES[APP_FOR_KIND[kind]]} AppleScript fallback failed: {error}"
            )
            result["error_hint"] = getattr(error, "hint", None)
        return _finish(result, started)

    try:
        with zipfile.ZipFile(path) as archive:
            if required_part not in archive.namelist():
                result["error"] = f"Not a valid {kind} package: missing {required_part}"
                return _finish(result, started)
            result.update(inspect(archive, args.max_items))
        result["inspector"] = "headless_openxml"
        result["success"] = True
    except zipfile.BadZipFile as error:
        result["error"] = f"Failed to parse {kind} package: {error}"
    except (ET.ParseError, KeyError) as error:
        result["error"] = f"Failed to parse {kind} package: {error}"
    return _finish(result, started)


def _finish(result, started):
    result["execution_time_sec"] = round(time.time() - started, 2)
    return result

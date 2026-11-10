#!/usr/bin/env python3
"""mac-office runner: headless OpenXML inspection and sandboxed AppleScript execution.

Two modes:

  --inspect <file>   Parse .docx/.xlsx/.pptx with zipfile + ElementTree first (milliseconds,
                     no GUI). If the file is wrapped in a Microsoft Purview / RMS OLE2
                     EncryptedPackage container (common on corporate M365 Macs), transparently
                     falls back to native Office AppleScript inspection.
  --app <app>        Compile and run an AppleScript through osascript with per-app alert
                     suppression, Group-Container sandbox staging (avoiding macOS 'Grant File
                     Access' popups), non-destructive working copies, timeout watchdog, and
                     structured JSON reporting.

Environment variables exposed to AppleScript:
  system attribute "MAC_OFFICE_TARGET_FILE"   -> staged working copy path inside Group Container
  system attribute "MAC_OFFICE_OUTPUT_FILE"   -> staged output path for --output-file (new/PDF/CSV)
  system attribute "MAC_OFFICE_STAGE_DIR"     -> sandbox-safe staging directory path
  system attribute "MAC_OFFICE_ORIGINAL_FILE" -> user's original target path
"""

import argparse
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
import uuid
import xml.etree.ElementTree as ET
import zipfile

APP_NAMES = {
    "word": "Microsoft Word",
    "excel": "Microsoft Excel",
    "powerpoint": "Microsoft PowerPoint",
}

GROUP_CONTAINER_DIR = os.path.expanduser(
    "~/Library/Group Containers/UBF8T346G9.Office/mac-office-staging"
)

# Each Office app declares `display alerts` with a different type: Word uses WdAlertLevel,
# Excel uses a boolean, and PowerPoint exposes no such property at all.
ALERT_HANDLERS = {
    "word": (
        'tell application "Microsoft Word"\n\tlaunch\n\tset display alerts to alerts none\nend tell\n',
        'tell application "Microsoft Word"\n\tset display alerts to alerts all\nend tell\n',
    ),
    "excel": (
        'tell application "Microsoft Excel"\n\tlaunch\n\tset display alerts to false\nend tell\n',
        'tell application "Microsoft Excel"\n\tset display alerts to true\nend tell\n',
    ),
    "powerpoint": (
        'tell application "Microsoft PowerPoint"\n'
        '\tlaunch\n'
        '\trepeat 20 times\n'
        '\t\ttry\n'
        '\t\t\tget Version\n'
        '\t\t\texit repeat\n'
        '\t\ton error\n'
        '\t\t\tdelay 0.1\n'
        '\t\tend try\n'
        '\tend repeat\n'
        'end tell\n',
        "",
    ),
}

# Query open documents only if the target application is already running.
DOC_PATHS_SCRIPT = {
    app: (
        f'if application "{name}" is running then\n'
        f'\ttell application "{name}"\n'
        f'\t\tset AppleScript\'s text item delimiters to linefeed\n'
        f'\t\treturn (full name of every {collection}) as text\n'
        f'\tend tell\n'
        f'else\n'
        f'\treturn ""\n'
        f'end if'
    )
    for app, name, collection in (
        ("word", "Microsoft Word", "document"),
        ("excel", "Microsoft Excel", "workbook"),
        ("powerpoint", "Microsoft PowerPoint", "presentation"),
    )
}

CLOSE_BY_INDEX_SCRIPT = {
    "word": (
        'tell application "Microsoft Word"\n'
        '\trepeat with i from (count of documents) to 1 by -1\n'
        '\t\tif i is in {%s} then\n'
        '\t\t\ttry\n'
        '\t\t\t\tclose document i saving no\n'
        '\t\t\tend try\n'
        '\t\tend if\n'
        '\tend repeat\n'
        'end tell\n'
    ),
    "excel": (
        'tell application "Microsoft Excel"\n'
        '\trepeat with i from (count of workbooks) to 1 by -1\n'
        '\t\tif i is in {%s} then\n'
        '\t\t\ttry\n'
        '\t\t\t\tclose workbook i saving no\n'
        '\t\t\tend try\n'
        '\t\tend if\n'
        '\tend repeat\n'
        'end tell\n'
    ),
    "powerpoint": (
        'tell application "Microsoft PowerPoint"\n'
        '\trepeat with i from (count of presentations) to 1 by -1\n'
        '\t\tif i is in {%s} then\n'
        '\t\t\ttry\n'
        '\t\t\t\tclose presentation i saving no\n'
        '\t\t\tend try\n'
        '\t\tend if\n'
        '\tend repeat\n'
        'end tell\n'
    ),
}

TCC_HINT = (
    "Not authorized to send Apple events. Open System Settings > Privacy & Security > "
    "Automation and allow your terminal/IDE to control Microsoft Word, Excel, and "
    "PowerPoint. To re-trigger the prompt: tccutil reset AppleEvents"
)

MODAL_ALERT_HINT = (
    "Target app raised a modal dialog (such as 'Grant File Access' or a save prompt). "
    "Dismiss it in the app and retry."
)

OLE2_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"

W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
S_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
R_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
P_NS = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
A_NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
PKG_REL_NS = "{http://schemas.openxmlformats.org/package/2006/relationships}"

EMU_PER_POINT = 12700


def emit(payload):
    print(json.dumps(payload, ensure_ascii=False, indent=2))
    return 0 if payload.get("success") else 1


def ensure_stage_dir():
    base = os.path.dirname(GROUP_CONTAINER_DIR)
    if os.path.isdir(base):
        os.makedirs(GROUP_CONTAINER_DIR, exist_ok=True)
        return GROUP_CONTAINER_DIR
    fallback = os.path.join(tempfile.gettempdir(), "mac-office-staging")
    os.makedirs(fallback, exist_ok=True)
    return fallback


def make_staged_path(original_path, tag="work"):
    stage_dir = ensure_stage_dir()
    stem, ext = os.path.splitext(os.path.basename(original_path))
    safe_stem = re.sub(r"[^A-Za-z0-9._-]+", "_", stem)[:40] or "file"
    token = uuid.uuid4().hex[:8]
    return os.path.join(stage_dir, f"{safe_stem}_{tag}_{token}{ext}")


def hfs_to_posix(path):
    """Convert an HFS path ("Disk:Users:me:a.docx") to POSIX ("/Users/me/a.docx")."""
    if not path or path in ("missing value", "/missing value"):
        return None
    if path.startswith("/"):
        return path
    return "/" + path.replace(":", "/").split("/", 1)[-1]


def column_index(cell_reference):
    letters = re.match(r"[A-Z]+", cell_reference or "")
    if not letters:
        return None
    index = 0
    for character in letters.group(0):
        index = index * 26 + (ord(character) - ord("A") + 1)
    return index


def column_name(index):
    name = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        name = chr(ord("A") + remainder) + name
    return name


def trim_text(value, limit=200):
    collapsed = " ".join((value or "").split())
    return collapsed if len(collapsed) <= limit else collapsed[: limit - 1] + "…"


def inspect_docx(archive, max_items):
    body = ET.fromstring(archive.read("word/document.xml")).find(W_NS + "body")
    paragraphs, tables = [], []
    for child in body:
        if child.tag == W_NS + "p":
            style_element = child.find(f"{W_NS}pPr/{W_NS}pStyle")
            paragraphs.append(
                {
                    "index": len(paragraphs) + 1,
                    "style": style_element.get(W_NS + "val") if style_element is not None else None,
                    "text": trim_text("".join(node.text or "" for node in child.iter(W_NS + "t"))),
                }
            )
        elif child.tag == W_NS + "tbl":
            rows = child.findall(W_NS + "tr")
            preview = [
                [
                    trim_text("".join(node.text or "" for node in cell.iter(W_NS + "t")), 80)
                    for cell in row.findall(W_NS + "tc")
                ]
                for row in rows[:5]
            ]
            tables.append(
                {
                    "index": len(tables) + 1,
                    "rows": len(rows),
                    "columns": max((len(row.findall(W_NS + "tc")) for row in rows), default=0),
                    "preview": preview,
                }
            )
    return {
        "paragraph_count": len(paragraphs),
        "table_count": len(tables),
        "paragraphs": paragraphs[:max_items],
        "tables": tables[:max_items],
    }


def shared_strings(archive):
    if "xl/sharedStrings.xml" not in archive.namelist():
        return []
    root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    return [
        "".join(node.text or "" for node in item.iter(S_NS + "t"))
        for item in root.findall(S_NS + "si")
    ]


def workbook_sheet_targets(archive):
    relationships = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    targets = {}
    for relationship in relationships.findall(PKG_REL_NS + "Relationship"):
        target = relationship.get("Target").lstrip("/")
        targets[relationship.get("Id")] = target if target.startswith("xl/") else "xl/" + target
    sheets_root = ET.fromstring(archive.read("xl/workbook.xml"))
    return [
        (sheet.get("name"), targets.get(sheet.get(R_NS + "id")))
        for sheet in sheets_root.iter(S_NS + "sheet")
    ]


def sheet_preview(archive, sheet_path, strings, max_items):
    if not sheet_path or sheet_path not in archive.namelist():
        return {"dimension": None, "rows": [], "truncated": False}
    root = ET.fromstring(archive.read(sheet_path))
    rows, max_row, max_column, truncated = [], 0, 0, False
    for row in root.iter(S_NS + "row"):
        max_row = max(max_row, int(row.get("r") or 0))
        if len(rows) >= min(max_items, 20):
            truncated = True
            break
        values = {}
        for cell in row.findall(S_NS + "c"):
            index = column_index(cell.get("r")) or (len(values) + 1)
            max_column = max(max_column, index)
            value_node = cell.find(S_NS + "v")
            if cell.get("t") == "s" and value_node is not None:
                try:
                    value = strings[int(value_node.text)]
                except (ValueError, IndexError):
                    value = value_node.text
            elif cell.get("t") == "inlineStr":
                value = "".join(node.text or "" for node in cell.iter(S_NS + "t"))
            else:
                value = value_node.text if value_node is not None else None
            values[index] = value
        row_values = [values.get(index) for index in range(1, max(values, default=0) + 1)]
        while row_values and row_values[-1] is None:
            row_values.pop()
        rows.append(row_values)
    dimension = root.find(S_NS + "dimension")
    used_range = dimension.get("ref") if dimension is not None else None
    if used_range is None and max_row and max_column:
        used_range = f"A1:{column_name(max_column)}{max_row}"
    return {"dimension": used_range, "rows": rows, "truncated": truncated}


def inspect_xlsx(archive, max_items):
    strings = shared_strings(archive)
    worksheets = []
    for name, target in workbook_sheet_targets(archive):
        preview = sheet_preview(archive, target, strings, max_items)
        worksheets.append(
            {
                "name": name,
                "used_range": preview["dimension"],
                "preview": preview["rows"],
                "truncated": preview["truncated"],
            }
        )
    return {"sheet_count": len(worksheets), "worksheets": worksheets[:max_items]}


def slide_targets(archive):
    relationships = ET.fromstring(archive.read("ppt/_rels/presentation.xml.rels"))
    target_by_id = {
        relationship.get("Id"): "ppt/" + relationship.get("Target").lstrip("/").removeprefix("ppt/")
        for relationship in relationships.findall(PKG_REL_NS + "Relationship")
    }
    presentation = ET.fromstring(archive.read("ppt/presentation.xml"))
    targets = []
    slide_list = presentation.find(P_NS + "sldIdLst")
    for slide_id in (slide_list if slide_list is not None else []):
        target = target_by_id.get(slide_id.get(R_NS + "id"))
        if target:
            targets.append(target)
    if not targets:
        targets = sorted(
            (name for name in archive.namelist() if re.fullmatch(r"ppt/slides/slide\d+\.xml", name)),
            key=lambda name: int(re.search(r"(\d+)", name).group(1)),
        )
    return targets


def shape_records(slide_root):
    records = []
    for shape in slide_root.iter():
        if shape.tag not in (P_NS + "sp", P_NS + "pic", P_NS + "graphicFrame"):
            continue
        name_node = shape.find(f".//{P_NS}cNvPr")
        offset = shape.find(f"{P_NS}spPr/{A_NS}xfrm/{A_NS}off")
        extent = shape.find(f"{P_NS}spPr/{A_NS}xfrm/{A_NS}ext")
        text = " ".join(node.text or "" for node in shape.iter(A_NS + "t"))
        records.append(
            {
                "index": len(records) + 1,
                "id": int(name_node.get("id")) if name_node is not None and name_node.get("id") else None,
                "name": name_node.get("name") if name_node is not None else None,
                "kind": shape.tag.split("}")[-1],
                "left_pt": round(int(offset.get("x")) / EMU_PER_POINT, 2) if offset is not None else None,
                "top_pt": round(int(offset.get("y")) / EMU_PER_POINT, 2) if offset is not None else None,
                "width_pt": round(int(extent.get("cx")) / EMU_PER_POINT, 2) if extent is not None else None,
                "height_pt": round(int(extent.get("cy")) / EMU_PER_POINT, 2) if extent is not None else None,
                "text": trim_text(text, 120),
            }
        )
    return records


def inspect_pptx(archive, max_items):
    slides = []
    for index, target in enumerate(slide_targets(archive), start=1):
        records = shape_records(ET.fromstring(archive.read(target)))
        slides.append(
            {
                "index": index,
                "shape_count": len(records),
                "shapes": records[:max_items],
            }
        )
    return {"slide_count": len(slides), "slides": slides[:max_items]}


def is_ole2_container(path):
    try:
        with open(path, "rb") as handle:
            return handle.read(8) == OLE2_MAGIC
    except OSError:
        return False


def inspect_via_applescript(path, kind, max_items):
    """Fallback inspector for OLE2 / MIP Sensitivity-Label encrypted Office documents."""
    staged = make_staged_path(path, tag="inspect")
    shutil.copy2(path, staged)
    try:
        if kind == "docx":
            script = f'''
            tell application "Microsoft Word"
                launch
                set display alerts to alerts none
                open (POSIX file "{staged}")
                set d to active document
                set pCount to count of paragraphs of d
                set tCount to count of tables of d
                set pLimit to pCount
                if pLimit > {max_items} then set pLimit to {max_items}
                set pOut to {{}}
                repeat with i from 1 to pLimit
                    set txt to content of text object of paragraph i of d
                    set end of pOut to (i as text) & "\\t" & txt
                end repeat
                set tLimit to tCount
                if tLimit > {max_items} then set tLimit to {max_items}
                set tOut to {{}}
                repeat with j from 1 to tLimit
                    set tbl to table j of d
                    set rCount to count of rows of tbl
                    set cCount to count of columns of tbl
                    set c1 to ""
                    try
                        set c1 to content of text object of cell 1 of row 1 of tbl
                    end try
                    set end of tOut to (j as text) & "\\t" & (rCount as text) & "\\t" & (cCount as text) & "\\t" & c1
                end repeat
                close d saving no
                set display alerts to alerts all
                set AppleScript's text item delimiters to "\\n<<<PARA>>>\\n"
                set pStr to pOut as text
                set AppleScript's text item delimiters to "\\n<<<TBL>>>\\n"
                set tStr to tOut as text
                return (pCount as text) & "\\n===SPLIT===\\n" & (tCount as text) & "\\n===SPLIT===\\n" & pStr & "\\n===SPLIT===\\n" & tStr
            end tell
            '''
            proc = subprocess.run(["osascript", "-e", script], capture_output=True, text=True, timeout=30)
            if proc.returncode != 0:
                raise RuntimeError(proc.stderr.strip() or proc.stdout.strip())
            parts = proc.stdout.strip().split("\n===SPLIT===\n")
            p_count = int(parts[0]) if len(parts) > 0 and parts[0].strip().isdigit() else 0
            t_count = int(parts[1]) if len(parts) > 1 and parts[1].strip().isdigit() else 0
            paragraphs = []
            if len(parts) > 2 and parts[2].strip():
                for item in parts[2].split("\n<<<PARA>>>\n"):
                    idx_str, _, txt = item.partition("\t")
                    txt_clean = trim_text(txt.rstrip("\r\n\x07"))
                    if txt_clean or len(paragraphs) < p_count:
                        paragraphs.append({"index": int(idx_str), "style": None, "text": txt_clean})
            tables = []
            if len(parts) > 3 and parts[3].strip():
                for item in parts[3].split("\n<<<TBL>>>\n"):
                    fields = item.split("\t", 3)
                    if len(fields) == 4:
                        tables.append({
                            "index": int(fields[0]),
                            "rows": int(fields[1]),
                            "columns": int(fields[2]),
                            "preview": [[trim_text(fields[3].rstrip("\r\n\x07"), 80)]],
                        })
            return {
                "inspector": "applescript_fallback",
                "paragraph_count": p_count,
                "table_count": t_count,
                "paragraphs": paragraphs[:max_items],
                "tables": tables[:max_items],
            }

        if kind == "xlsx":
            script = f'''
            tell application "Microsoft Excel"
                launch
                set display alerts to false
                open (POSIX file "{staged}")
                set wb to active workbook
                set sCount to count of worksheets of wb
                set sOut to {{}}
                repeat with i from 1 to sCount
                    set sh to worksheet i of wb
                    set sName to name of sh
                    set ur to used range of sh
                    set addr to get address ur
                    set rCount to count of rows of ur
                    set cCount to count of columns of ur
                    set v1 to ""
                    try
                        set v1 to (value of cell 1 of row 1 of ur) as text
                    end try
                    set end of sOut to sName & "\\t" & addr & "\\t" & (rCount as text) & "\\t" & (cCount as text) & "\\t" & v1
                end repeat
                close wb saving no
                set display alerts to true
                set AppleScript's text item delimiters to "\\n<<<SHEET>>>\\n"
                return sOut as text
            end tell
            '''
            proc = subprocess.run(["osascript", "-e", script], capture_output=True, text=True, timeout=30)
            if proc.returncode != 0:
                raise RuntimeError(proc.stderr.strip() or proc.stdout.strip())
            worksheets = []
            raw = proc.stdout.strip()
            if raw:
                for item in raw.split("\n<<<SHEET>>>\n"):
                    fields = item.split("\t", 4)
                    if len(fields) == 5:
                        worksheets.append({
                            "name": fields[0],
                            "used_range": fields[1].replace("$", ""),
                            "preview": [[fields[4]]] if fields[4] else [],
                            "truncated": int(fields[2]) > 1,
                        })
            return {
                "inspector": "applescript_fallback",
                "sheet_count": len(worksheets),
                "worksheets": worksheets[:max_items],
            }

        if kind == "pptx":
            script = f'''
            tell application "Microsoft PowerPoint"
                launch
                open (POSIX file "{staged}")
                repeat 25 times
                    if (count of presentations) > 0 then exit repeat
                    delay 0.1
                end repeat
                set pres to active presentation
                set sCount to count of slides of pres
                set sOut to {{}}
                repeat with i from 1 to sCount
                    set sl to slide i of pres
                    set shCount to count of shapes of sl
                    set shList to {{}}
                    repeat with j from 1 to shCount
                        set sh to shape j of sl
                        set shName to name of sh
                        set lPos to left position of sh
                        set tPos to top of sh
                        set wVal to width of sh
                        set hVal to height of sh
                        set shTxt to ""
                        try
                            if has text frame of sh then
                                set shTxt to content of text range of text frame of sh
                            end if
                        end try
                        set end of shList to (j as text) & "\\t" & shName & "\\t" & (lPos as text) & "\\t" & (tPos as text) & "\\t" & (wVal as text) & "\\t" & (hVal as text) & "\\t" & shTxt
                    end repeat
                    set AppleScript's text item delimiters to "\\n<<<SHAPE>>>\\n"
                    set end of sOut to (i as text) & "\\t" & (shCount as text) & "\\n<<<SHAPES_BODY>>>\\n" & (shList as text)
                end repeat
                close active presentation saving no
                set AppleScript's text item delimiters to "\\n<<<SLIDE>>>\\n"
                return sOut as text
            end tell
            '''
            proc = subprocess.run(["osascript", "-e", script], capture_output=True, text=True, timeout=30)
            if proc.returncode != 0:
                raise RuntimeError(proc.stderr.strip() or proc.stdout.strip())
            slides = []
            raw = proc.stdout.strip("\r\n")
            if raw:
                for slide_block in raw.split("\n<<<SLIDE>>>\n"):
                    header, _, body = slide_block.partition("\n<<<SHAPES_BODY>>>")
                    body = body.lstrip("\r\n")
                    h_parts = header.strip().split("\t")
                    s_idx = int(h_parts[0])
                    sh_count = int(h_parts[1]) if len(h_parts) > 1 else 0
                    shapes = []
                    if body.strip():
                        for sh_item in body.split("\n<<<SHAPE>>>\n"):
                            sf = sh_item.split("\t", 6)
                            if len(sf) == 7:
                                shapes.append({
                                    "index": int(sf[0]),
                                    "id": None,
                                    "name": sf[1],
                                    "kind": "sp",
                                    "left_pt": round(float(sf[2]), 2),
                                    "top_pt": round(float(sf[3]), 2),
                                    "width_pt": round(float(sf[4]), 2),
                                    "height_pt": round(float(sf[5]), 2),
                                    "text": trim_text(sf[6], 120),
                                })
                    slides.append({
                        "index": s_idx,
                        "shape_count": sh_count,
                        "shapes": shapes[:max_items],
                    })
            return {
                "inspector": "applescript_fallback",
                "slide_count": len(slides),
                "slides": slides[:max_items],
            }
    finally:
        if os.path.exists(staged):
            try:
                os.unlink(staged)
            except OSError:
                pass


INSPECTORS = {
    ".docx": ("docx", "word/document.xml", inspect_docx),
    ".xlsx": ("xlsx", "xl/workbook.xml", inspect_xlsx),
    ".pptx": ("pptx", "ppt/presentation.xml", inspect_pptx),
}


def run_inspect(args):
    started = time.time()
    path = os.path.abspath(os.path.expanduser(args.inspect))
    result = {
        "success": False,
        "mode": "inspect",
        "file": path,
        "kind": None,
        "error": None,
        "error_hint": None,
        "execution_time_sec": 0.0,
    }
    if not os.path.isfile(path):
        result["error"] = f"File not found: {path}"
        return emit(result)
    inspector = INSPECTORS.get(os.path.splitext(path)[1].lower())
    if inspector is None:
        result["error"] = "Unsupported file type. Expected .docx, .xlsx, or .pptx."
        return emit(result)
    kind, required_part, inspect = inspector
    result["kind"] = kind
    try:
        with zipfile.ZipFile(path) as archive:
            if required_part not in archive.namelist():
                result["error"] = f"Not a valid {kind} package: missing {required_part}"
                return emit(result)
            result.update(inspect(archive, args.max_items))
        result["inspector"] = "headless_openxml"
        result["success"] = True
    except zipfile.BadZipFile as error:
        if is_ole2_container(path):
            try:
                result.update(inspect_via_applescript(path, kind, args.max_items))
                result["success"] = True
            except Exception as fallback_error:
                result["error"] = (
                    f"File is an OLE2/MIP-encrypted {kind} container and AppleScript fallback "
                    f"failed: {fallback_error}"
                )
        else:
            result["error"] = f"Failed to parse {kind} package: {error}"
    except (ET.ParseError, KeyError) as error:
        result["error"] = f"Failed to parse {kind} package: {error}"
    result["execution_time_sec"] = round(time.time() - started, 2)
    return emit(result)


def prepare_working_copy(target_file, in_place):
    """Return (working_path, backup_path). Backup exists only for in-place runs."""
    if in_place:
        backup_path = target_file + ".bak"
        shutil.copy2(target_file, backup_path)
        return target_file, backup_path
    stem, extension = os.path.splitext(target_file)
    working_path = f"{stem}_revised{extension}"
    shutil.copy2(target_file, working_path)
    return working_path, None


def open_document_paths(app):
    """Return the POSIX paths of open documents in app order (index + 1 == document index)."""
    try:
        completed = subprocess.run(
            ["osascript", "-e", DOC_PATHS_SCRIPT[app]],
            capture_output=True,
            text=True,
            timeout=15,
        )
    except subprocess.TimeoutExpired:
        return []
    if completed.returncode != 0:
        return []
    return [
        converted
        for converted in (hfs_to_posix(item.strip()) for item in completed.stdout.strip().splitlines())
        if converted
    ]


def close_documents_by_index(app, indices):
    if not indices:
        return
    script = CLOSE_BY_INDEX_SCRIPT[app] % ", ".join(str(index) for index in indices)
    try:
        subprocess.run(["osascript", "-e", script], capture_output=True, text=True, timeout=15)
    except subprocess.TimeoutExpired:
        pass


def detect_modal_dialog(app):
    process_name = APP_NAMES[app]
    script = (
        f'tell application "System Events" to tell process "{process_name}"\n'
        '\tif not (exists window 1) then return false\n'
        '\tset winNames to name of every window\n'
        '\tif winNames contains "Grant File Access" then return true\n'
        '\tif (count of sheets of window 1) > 0 then return true\n'
        '\treturn false\n'
        'end tell'
    )
    try:
        completed = subprocess.run(
            ["osascript", "-e", script], capture_output=True, text=True, timeout=6
        )
    except subprocess.TimeoutExpired:
        return None
    if completed.returncode != 0:
        return None
    val = completed.stdout.strip().lower()
    if val in ("true", "false"):
        return val == "true"
    return None


def build_script(app, user_script):
    preamble, postamble = ALERT_HANDLERS[app]
    return preamble + user_script.rstrip() + "\n" + postamble


def read_user_script(args):
    if args.file:
        with open(args.file, "r", encoding="utf-8") as handle:
            return handle.read()
    return sys.stdin.read()


def run_execute(args):
    started = time.time()
    app = args.app.lower()
    result = {
        "success": False,
        "mode": "execute",
        "app": app,
        "target_file": None,
        "output_file": None,
        "backup_file": None,
        "output": None,
        "error": None,
        "error_hint": None,
        "execution_time_sec": 0.0,
    }
    if app not in APP_NAMES:
        result["error"] = f"Unknown app '{args.app}'. Use word, excel, or powerpoint."
        return emit(result)

    try:
        user_script = read_user_script(args)
    except OSError as error:
        result["error"] = f"Cannot read AppleScript: {error}"
        return emit(result)

    stage_dir = ensure_stage_dir()
    target_file = None
    working_path = None
    backup_path = None
    staged_target_path = None
    if args.target_file:
        target_file = os.path.abspath(os.path.expanduser(args.target_file))
        if not os.path.isfile(target_file):
            result["error"] = f"Target file not found: {target_file}"
            return emit(result)
        try:
            working_path, backup_path = prepare_working_copy(target_file, args.in_place)
            staged_target_path = make_staged_path(working_path, tag="target")
            shutil.copy2(working_path, staged_target_path)
        except OSError as error:
            result["error"] = f"Cannot prepare working copy: {error}"
            return emit(result)

    output_file = None
    staged_output_path = None
    if args.output_file:
        output_file = os.path.abspath(os.path.expanduser(args.output_file))
        os.makedirs(os.path.dirname(output_file), exist_ok=True)
        staged_output_path = make_staged_path(output_file, tag="output")

    result["target_file"] = working_path
    result["output_file"] = output_file
    result["backup_file"] = backup_path

    previously_open = set(open_document_paths(app))
    before_stat = (
        os.stat(staged_target_path)
        if staged_target_path and os.path.isfile(staged_target_path)
        else None
    )
    script_path = None
    try:
        with tempfile.NamedTemporaryFile(
            "w", suffix=".applescript", delete=False, encoding="utf-8"
        ) as handle:
            handle.write(build_script(app, user_script))
            script_path = handle.name
        environment = dict(os.environ)
        environment["MAC_OFFICE_TARGET_FILE"] = staged_target_path or ""
        environment["MAC_OFFICE_OUTPUT_FILE"] = staged_output_path or ""
        environment["MAC_OFFICE_STAGE_DIR"] = stage_dir
        environment["MAC_OFFICE_ORIGINAL_FILE"] = target_file or ""
        completed = subprocess.run(
            ["osascript", script_path],
            capture_output=True,
            text=True,
            timeout=args.timeout,
            env=environment,
        )
        result["output"] = completed.stdout.strip()
        if completed.returncode != 0:
            error_text = (completed.stderr or completed.stdout).strip()
            result["error"] = error_text
            if "-1743" in error_text or "not authorized" in error_text.lower():
                result["error_hint"] = TCC_HINT
            result["success"] = False
        else:
            result["success"] = True
    except subprocess.TimeoutExpired as error:
        result["error"] = f"osascript timed out after {args.timeout}s"
        result["output"] = (error.stdout or b"").decode("utf-8", "replace").strip() if isinstance(
            error.stdout, bytes
        ) else (error.stdout or "").strip()
        modal = detect_modal_dialog(app)
        if modal is None:
            result["error_hint"] = (
                "Timed out and dialog state could not be read (grant Accessibility permission "
                "to detect dialogs). Check the app for a blocking dialog, then retry."
            )
        elif modal:
            result["error_hint"] = MODAL_ALERT_HINT
        else:
            result["error_hint"] = (
                "No dialog found. The script may be waiting on a slow operation; raise --timeout."
            )
    except OSError as error:
        result["error"] = f"Cannot run osascript: {error}"
    finally:
        if script_path and os.path.exists(script_path):
            os.unlink(script_path)

    # Close any staged documents left open by the script before moving files back.
    touched_paths = {
        p for p in (staged_target_path, staged_output_path, working_path, target_file) if p
    }
    if touched_paths:
        opened_by_script = [
            index
            for index, path in enumerate(open_document_paths(app), start=1)
            if path in touched_paths and path not in previously_open
        ]
        close_documents_by_index(app, opened_by_script)

    try:
        if working_path and staged_target_path:
            after_stat = (
                os.stat(staged_target_path) if os.path.isfile(staged_target_path) else None
            )
            modified = bool(
                before_stat
                and after_stat
                and (before_stat.st_size, before_stat.st_mtime_ns)
                != (after_stat.st_size, after_stat.st_mtime_ns)
            )
            if result["success"] and modified:
                shutil.copy2(staged_target_path, working_path)
            if args.in_place:
                if result["success"] and backup_path and os.path.exists(backup_path):
                    os.unlink(backup_path)
                    result["backup_file"] = None
                elif not result["success"] and backup_path and os.path.exists(backup_path):
                    shutil.copy2(backup_path, target_file)
                    os.unlink(backup_path)
                    result["backup_file"] = None
                    restore_note = "Original file restored from backup."
                    result["error_hint"] = (
                        f"{result['error_hint']} {restore_note}"
                        if result["error_hint"]
                        else restore_note
                    )
            result["output_exists"] = os.path.isfile(working_path)
            result["output_modified"] = modified

        if output_file and staged_output_path:
            if result["success"] and os.path.isfile(staged_output_path) and os.path.getsize(staged_output_path) > 0:
                shutil.copy2(staged_output_path, output_file)
                result["output_exists"] = True
                result["output_modified"] = True
            else:
                result["output_exists"] = os.path.isfile(output_file)
                if "output_modified" not in result:
                    result["output_modified"] = False
    finally:
        for temp_staged in (staged_target_path, staged_output_path):
            if temp_staged and os.path.exists(temp_staged):
                try:
                    os.unlink(temp_staged)
                except OSError:
                    pass

    result["execution_time_sec"] = round(time.time() - started, 2)
    return emit(result)


def main():
    parser = argparse.ArgumentParser(
        description="Inspect Office files headlessly or run AppleScript against local Office apps."
    )
    parser.add_argument("--inspect", metavar="FILE", help="Headless structure inspection target")
    parser.add_argument("--max-items", type=int, default=50, help="Max items per collection (default 50)")
    parser.add_argument("--app", choices=sorted(APP_NAMES), help="App to drive: word, excel, powerpoint")
    parser.add_argument("--file", help="AppleScript file to execute (default: read stdin)")
    parser.add_argument("--target-file", help="Absolute path of the existing Office file to edit")
    parser.add_argument("--output-file", help="Absolute path for a newly created or exported file (e.g. .docx, .xlsx, .pptx, .pdf, .csv)")
    parser.add_argument("--in-place", action="store_true", help="Edit the target file instead of a _revised copy")
    parser.add_argument("--timeout", type=int, default=30, help="osascript timeout in seconds (default 30, max 180)")
    args = parser.parse_args()
    args.timeout = max(5, min(args.timeout, 180))

    if args.inspect:
        return run_inspect(args)
    if args.app:
        return run_execute(args)
    parser.error("Specify either --inspect FILE or --app <word|excel|powerpoint>")


if __name__ == "__main__":
    sys.exit(main())

"""Shared helpers for the mac-office test suite (standard library only).

Provides:
  * ``run_runner`` — invoke ``scripts/office_runner.py`` as a subprocess and parse its JSON.
  * synthetic OOXML builders (``write_docx``, ``write_xlsx``, ``write_pptx``) plus
    ``write_ole2_file`` for the MIP/RMS fallback path.
  * ``OsascriptShim`` — an executable fake ``osascript`` placed first on ``PATH`` so unit
    tests never touch Microsoft Office.
  * ``MacOfficeTestCase`` / ``ShimTestCase`` — temp-dir isolation plus assertions that no
    files are left behind (Office staging directory, temp ``*.applescript`` scripts).
"""

import atexit
import fcntl
import glob
import json
import os
import shutil
import subprocess
import sys
import tempfile
import unittest
import zipfile
from xml.sax.saxutils import escape

SKILL_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
RUNNER = os.path.join(SKILL_DIR, "scripts", "office_runner.py")

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
S_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

EMU_PER_POINT = 12700
OLE2_MAGIC = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"

XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'


# ---------------------------------------------------------------------------
# runner invocation
# ---------------------------------------------------------------------------


class RunnerResult:
    """Result of one ``office_runner.py`` invocation."""

    def __init__(self, completed):
        self.returncode = completed.returncode
        self.stdout = completed.stdout or ""
        self.stderr = completed.stderr or ""
        self.payload = None
        try:
            self.payload = json.loads(self.stdout)
        except ValueError:
            self.payload = None

    def __getitem__(self, key):
        """Shortcut for ``payload[key]``."""
        return self.payload[key]

    def get(self, key, default=None):
        return (self.payload or {}).get(key, default)

    def __repr__(self):
        return "RunnerResult(rc={!r}, payload={!r}, stderr={!r})".format(
            self.returncode, self.payload, self.stderr[:300]
        )


def run_runner(args, *, env=None, stdin_text=None, timeout=300, cwd=None):
    """Run the runner subprocess and return a :class:`RunnerResult`."""
    completed = subprocess.run(
        [sys.executable, RUNNER, *[str(a) for a in args]],
        input=stdin_text,
        capture_output=True,
        text=True,
        env=env,
        timeout=timeout,
        cwd=cwd,
    )
    return RunnerResult(completed)


_CONSTANTS_MODULE = None


def constants():
    """Import ``macoffice.constants`` for shared values (stage dir, hint texts, preambles)."""
    global _CONSTANTS_MODULE
    if _CONSTANTS_MODULE is None:
        scripts_dir = os.path.join(SKILL_DIR, "scripts")
        if scripts_dir not in sys.path:
            sys.path.insert(0, scripts_dir)
        from macoffice import constants as module  # noqa: E402  (deliberate late import)

        _CONSTANTS_MODULE = module
    return _CONSTANTS_MODULE


def staging_dir():
    """The Office Group-Container staging directory used by the runner."""
    return constants().GROUP_CONTAINER_DIR


# ---------------------------------------------------------------------------
# synthetic Office packages
# ---------------------------------------------------------------------------


def write_package(path, parts):
    """Write a zip package from a {archive member name: text or bytes} mapping."""
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, content in parts.items():
            data = content.encode("utf-8") if isinstance(content, str) else content
            archive.writestr(name, data)
    return path


def write_ole2_file(path, payload=b"\x00" * 128):
    """Write a file that starts with the OLE2 magic (a legacy Word 97-2003 container).

    Pass ``payload=b"EncryptedPackage"`` to simulate a MIP/RMS encrypted container.
    """
    with open(path, "wb") as handle:
        handle.write(OLE2_MAGIC + payload)
    return path


def write_cfb(path, streams):
    """Write a minimal OLE2 (Compound File) container holding ``streams``.

    ``streams`` maps stream name -> bytes. Layout: sector 0 is the FAT, sector 1 the
    directory, then one sector per stream. Good enough for stream-name detection; the
    directory is a flat list rather than a red-black tree.
    """
    import struct

    sector_size = 512
    names = ["Root Entry"] + list(streams)
    payloads = [b""] + [streams[name] for name in names[1:]]
    fat_entries = [0xFFFFFFFD, 0xFFFFFFFE]  # FAT sector, directory sector
    for index in range(len(payloads) - 1):
        fat_entries.append(0xFFFFFFFE)
    while len(fat_entries) < sector_size // 4:
        fat_entries.append(0xFFFFFFFF)
    fat = struct.pack("<%dI" % len(fat_entries), *fat_entries)

    directory = bytearray()
    for index, (name, payload) in enumerate(zip(names, payloads)):
        entry = bytearray(128)
        encoded = name.encode("utf-16-le") + b"\x00\x00"
        entry[0 : len(encoded)] = encoded
        struct.pack_into("<H", entry, 64, len(encoded))
        entry[66] = 5 if index == 0 else 2  # root storage / stream
        entry[67] = 1
        struct.pack_into("<III", entry, 68, 0xFFFFFFFF, 0xFFFFFFFF, 0xFFFFFFFF)
        if index:
            struct.pack_into("<I", entry, 116, 1 + index)  # start sector
            struct.pack_into("<Q", entry, 120, len(payload))
        else:
            struct.pack_into("<I", entry, 116, 0xFFFFFFFE)
            struct.pack_into("<Q", entry, 120, 0)
        directory += entry
    directory += b"\x00" * (sector_size - len(directory) % sector_size)

    header = bytearray(512)
    header[0:8] = OLE2_MAGIC
    struct.pack_into("<HH", header, 24, 0x003E, 0x0003)
    struct.pack_into("<H", header, 28, 0xFFFE)
    struct.pack_into("<H", header, 30, 9)
    struct.pack_into("<H", header, 32, 6)
    struct.pack_into("<I", header, 40, 0)
    struct.pack_into("<I", header, 44, 1)
    struct.pack_into("<I", header, 48, 1)
    struct.pack_into("<I", header, 52, 0)
    struct.pack_into("<I", header, 56, 4096)
    struct.pack_into("<I", header, 60, 0xFFFFFFFE)
    struct.pack_into("<I", header, 64, 0)
    struct.pack_into("<I", header, 68, 0xFFFFFFFE)
    struct.pack_into("<I", header, 72, 0)
    struct.pack_into("<I", header, 76, 0)
    for offset in range(80, 512, 4):
        struct.pack_into("<I", header, offset, 0xFFFFFFFF)

    with open(path, "wb") as handle:
        handle.write(bytes(header))
        handle.write(fat)
        handle.write(bytes(directory))
        for payload in payloads:
            data = payload[:sector_size]
            handle.write(data + b"\x00" * (sector_size - len(data)))
    return path


def _docx_paragraph_xml(text, style):
    style_xml = '<w:pPr><w:pStyle w:val="{}"/></w:pPr>'.format(escape(style)) if style else ""
    return '<w:p>{}<w:r><w:t xml:space="preserve">{}</w:t></w:r></w:p>'.format(
        style_xml, escape(text)
    )


def _docx_table_xml(rows):
    trs = []
    for row in rows:
        cells = "".join(
            '<w:tc><w:p><w:r><w:t xml:space="preserve">{}</w:t></w:r></w:p></w:tc>'.format(
                escape(str(cell))
            )
            for cell in row
        )
        trs.append("<w:tr>{}</w:tr>".format(cells))
    return "<w:tbl>{}</w:tbl>".format("".join(trs))


def docx_body_xml(body):
    """Build ``word/document.xml`` text for a sequence of body items.

    Each item is ``("p", text, style_or_None)`` or ``("tbl", rows)``.
    """
    parts = []
    for item in body:
        if item[0] == "p":
            parts.append(_docx_paragraph_xml(item[1], item[2] if len(item) > 2 else None))
        elif item[0] == "tbl":
            parts.append(_docx_table_xml(item[1]))
        else:
            raise ValueError("unknown docx body item: {!r}".format(item[0]))
    return (
        XML_DECL
        + '<w:document xmlns:w="{}"><w:body>{}</w:body></w:document>'.format(
            W_NS, "".join(parts)
        )
    )


def write_docx(path, body):
    """Write a .docx package containing ``body`` (see :func:`docx_body_xml`)."""
    return write_package(path, {"word/document.xml": docx_body_xml(body)})


def column_name(index):
    name = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        name = chr(ord("A") + remainder) + name
    return name


def _shared_string_item(runs):
    return "<si>{}</si>".format(
        "".join(
            '<r><t xml:space="preserve">{}</t></r>'.format(escape(run)) for run in runs
        )
    )


def write_xlsx(path, sheets):
    """Write a .xlsx package from sheet definitions.

    ``sheets`` is a list of dicts:
        {"name": "Data", "dimension": "A1:C3" or None, "rows": [[cell, ...], ...],
         "rel_target": optional override for the worksheet relationship target}

    Cell values: ``str`` (shared string), ``("runs", [..])`` (rich-text shared string),
    ``("inline", text)`` (inlineStr), ``int``/``float`` (numeric), ``None`` (empty).
    """
    shared = []
    shared_index = {}

    def shared_id(runs):
        key = tuple(runs)
        if key not in shared_index:
            shared_index[key] = len(shared)
            shared.append(list(runs))
        return shared_index[key]

    worksheet_parts = {}
    sheet_entries = []
    rel_entries = []
    for sheet_number, sheet in enumerate(sheets, start=1):
        rows_xml = []
        for row_number, row in enumerate(sheet.get("rows", []), start=1):
            cells = []
            for column_number, value in enumerate(row, start=1):
                if value is None:
                    continue
                reference = "{}{}".format(column_name(column_number), row_number)
                if isinstance(value, tuple) and value[0] == "inline":
                    cells.append(
                        '<c r="{}" t="inlineStr"><is><t xml:space="preserve">{}</t></is></c>'.format(
                            reference, escape(str(value[1]))
                        )
                    )
                elif isinstance(value, tuple) and value[0] == "runs":
                    cells.append(
                        '<c r="{}" t="s"><v>{}</v></c>'.format(reference, shared_id(value[1]))
                    )
                elif isinstance(value, str):
                    cells.append(
                        '<c r="{}" t="s"><v>{}</v></c>'.format(reference, shared_id([value]))
                    )
                else:
                    cells.append('<c r="{}"><v>{}</v></c>'.format(reference, value))
            rows_xml.append('<row r="{}">{}</row>'.format(row_number, "".join(cells)))
        dimension = sheet.get("dimension")
        dimension_xml = '<dimension ref="{}"/>'.format(dimension) if dimension else ""
        worksheet_parts["xl/worksheets/sheet{}.xml".format(sheet_number)] = (
            XML_DECL
            + '<worksheet xmlns="{}">{}{}</worksheet>'.format(
                S_NS,
                dimension_xml,
                "<sheetData>{}</sheetData>".format("".join(rows_xml)) if rows_xml else "",
            )
        )
        sheet_entries.append(
            '<sheet name="{}" sheetId="{}" r:id="rId{}"/>'.format(
                escape(sheet["name"]), sheet_number, sheet_number
            )
        )
        rel_entries.append(
            '<Relationship Id="rId{}" Type="http://schemas.openxmlformats.org/officeDocument/'
            '2006/relationships/worksheet" Target="{}"/>'.format(
                sheet_number,
                sheet.get("rel_target", "worksheets/sheet{}.xml".format(sheet_number)),
            )
        )

    parts = {
        "xl/workbook.xml": XML_DECL
        + '<workbook xmlns="{}" xmlns:r="{}"><sheets>{}</sheets></workbook>'.format(
            S_NS, R_NS, "".join(sheet_entries)
        ),
        "xl/_rels/workbook.xml.rels": XML_DECL
        + '<Relationships xmlns="{}">{}</Relationships>'.format(PKG_REL_NS, "".join(rel_entries)),
    }
    if shared:
        parts["xl/sharedStrings.xml"] = XML_DECL + '<sst xmlns="{}">{}</sst>'.format(
            S_NS, "".join(_shared_string_item(runs) for runs in shared)
        )
    parts.update(worksheet_parts)
    return write_package(path, parts)


def _xfrm_xml(shape):
    return '<a:xfrm><a:off x="{}" y="{}"/><a:ext cx="{}" cy="{}"/></a:xfrm>'.format(
        shape["x"], shape["y"], shape["cx"], shape["cy"]
    )


def shape_xml(shape):
    """Build one ``<p:sp>`` / ``<p:pic>`` / ``<p:graphicFrame>`` element."""
    kind = shape["kind"]
    name = escape(shape.get("name", "Shape 1"))
    shape_id = shape.get("id", 2)
    text = shape.get("text") or ""
    has_xfrm = all(key in shape for key in ("x", "y", "cx", "cy"))

    if kind == "sp":
        sp_pr = "<p:spPr>{}</p:spPr>".format(_xfrm_xml(shape) if has_xfrm else "")
        tx_body = (
            '<p:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US"/><a:t>{}</a:t></a:r></a:p>'
            "</p:txBody>".format(escape(text))
            if text
            else ""
        )
        return (
            '<p:sp><p:nvSpPr><p:cNvPr id="{}" name="{}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
            "{}{}</p:sp>".format(shape_id, name, sp_pr, tx_body)
        )

    if kind == "pic":
        return (
            '<p:pic><p:nvPicPr><p:cNvPr id="{}" name="{}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>'
            "<p:spPr>{}</p:spPr></p:pic>".format(
                shape_id, name, _xfrm_xml(shape) if has_xfrm else ""
            )
        )

    if kind == "graphicFrame":
        rows = "".join(
            "<a:tr>"
            + "".join(
                '<a:tc><a:txBody><a:p><a:r><a:t>{}</a:t></a:r></a:p></a:txBody></a:tc>'.format(
                    escape(str(cell))
                )
                for cell in row
            )
            + "</a:tr>"
            for row in shape.get("cells", [])
        )
        frame = (
            '<p:xfrm><a:off x="{}" y="{}"/><a:ext cx="{}" cy="{}"/></p:xfrm>'.format(
                shape["x"], shape["y"], shape["cx"], shape["cy"]
            )
            if has_xfrm
            else ""
        )
        graphic = (
            '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/'
            '2006/table"><a:tbl>{}</a:tbl></a:graphicData></a:graphic>'.format(rows)
        )
        return (
            '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="{}" name="{}"/>'
            "<p:nvGraphicFramePr/></p:nvGraphicFramePr>{}{}</p:graphicFrame>".format(
                shape_id, name, frame, graphic
            )
        )

    raise ValueError("unknown pptx shape kind: {!r}".format(kind))


def slide_xml(shapes):
    """Build ``ppt/slides/slideN.xml`` text for a list of shapes."""
    return (
        XML_DECL
        + '<p:sld xmlns:a="{}" xmlns:r="{}" xmlns:p="{}"><p:cSld><p:spTree>'
        "<p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>"
        "<p:grpSpPr/>"
        "{}</p:spTree></p:cSld></p:sld>".format(
            A_NS, R_NS, P_NS, "".join(shape_xml(shape) for shape in shapes)
        )
    )


def write_pptx(path, slides, *, use_slide_id_list=True, slide_file_names=None):
    """Write a .pptx package.

    ``slides`` is a list of shape lists; ``slide_file_names`` overrides the member names
    (used to exercise the numeric-order fallback when the presentation has no slide list).
    """
    names = slide_file_names or ["ppt/slides/slide{}.xml".format(i) for i in range(1, len(slides) + 1)]
    parts = {"ppt/presentation.xml": None}
    rel_entries = []
    slide_ids = []
    for index, (name, shapes) in enumerate(zip(names, slides), start=1):
        parts[name] = slide_xml(shapes)
        if use_slide_id_list:
            rel_id = "rId{}".format(index)
            rel_entries.append(
                '<Relationship Id="{}" Type="http://schemas.openxmlformats.org/officeDocument/'
                '2006/relationships/slide" Target="slides/{}"/>'.format(
                    rel_id, os.path.basename(name)
                )
            )
            slide_ids.append('<p:sldId id="{}" r:id="{}"/>'.format(255 + index, rel_id))
    slide_list = "<p:sldIdLst>{}</p:sldIdLst>".format("".join(slide_ids)) if slide_ids else ""
    parts["ppt/presentation.xml"] = (
        XML_DECL
        + '<p:presentation xmlns:a="{}" xmlns:r="{}" xmlns:p="{}">{}<p:sldSz cx="9144000" '
        'cy="6858000"/></p:presentation>'.format(A_NS, R_NS, P_NS, slide_list)
    )
    parts["ppt/_rels/presentation.xml.rels"] = XML_DECL + (
        '<Relationships xmlns="{}">{}</Relationships>'.format(PKG_REL_NS, "".join(rel_entries))
    )
    return write_package(path, parts)


def write_png(path, width=8, height=8, colour=(0xFF, 0x00, 0x00)):
    """Write a minimal solid-colour PNG (no third-party imaging library)."""
    import struct
    import zlib

    def chunk(kind, data):
        return (
            struct.pack(">I", len(data))
            + kind
            + data
            + struct.pack(">I", zlib.crc32(kind + data) & 0xFFFFFFFF)
        )

    row = b"\x00" + bytes(colour) * width
    header = struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0)
    payload = (
        b"\x89PNG\r\n\x1a\n"
        + chunk(b"IHDR", header)
        + chunk(b"IDAT", zlib.compress(row * height))
        + chunk(b"IEND", b"")
    )
    with open(path, "wb") as handle:
        handle.write(payload)
    return path


# ---------------------------------------------------------------------------
# minimal PDF writer (for render tests; standard library only)
# ---------------------------------------------------------------------------

PAGE_WIDTH = 612
PAGE_HEIGHT = 792
PDF_FONT_SIZE = 14
PDF_LEFT = 72
PDF_FIRST_BASELINE = 700


def write_pdf(path, pages=None):
    """Write a deterministic one-font PDF.

    ``pages`` is a list of line lists; every line is drawn with Helvetica 14 at x=72,
    descending 20 pt per line from y=700 on a 612x792 pt page.
    """
    pages = pages or [["Hello Render"]]
    objects = [
        "<< /Type /Catalog /Pages 2 0 R >>",
        None,  # pages node, filled in below
        "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
    ]
    page_numbers = []
    for lines in pages:
        stream_lines = ["BT", f"/F1 {PDF_FONT_SIZE} Tf"]
        for index, line in enumerate(lines):
            escape = line.replace("\\", r"\\").replace("(", r"\(").replace(")", r"\)")
            if index == 0:
                stream_lines.append(f"{PDF_LEFT} {PDF_FIRST_BASELINE} Td")
            else:
                stream_lines.append("0 -20 Td")
            stream_lines.append(f"({escape}) Tj")
        stream_lines.append("ET")
        content = "\n".join(stream_lines).encode("utf-8")
        objects.append(b"<< /Length " + str(len(content)).encode() + b" >>\nstream\n" + content + b"\nendstream")
        content_number = len(objects)
        objects.append(
            f"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 {PAGE_WIDTH} {PAGE_HEIGHT}] "
            f"/Resources << /Font << /F1 3 0 R >> >> /Contents {content_number} 0 R >>"
        )
        page_numbers.append(len(objects))
    objects[1] = "<< /Type /Pages /Kids [{}] /Count {} >>".format(
        " ".join(f"{number} 0 R" for number in page_numbers), len(page_numbers)
    )

    out = bytearray(b"%PDF-1.4\n")
    offsets = []
    for index, body in enumerate(objects, start=1):
        offsets.append(len(out))
        payload = body.encode("utf-8") if isinstance(body, str) else body
        out += f"{index} 0 obj\n".encode() + payload + b"\nendobj\n"
    xref = len(out)
    out += f"xref\n0 {len(objects) + 1}\n".encode() + b"0000000000 65535 f \n"
    for offset in offsets:
        out += f"{offset:010d} 00000 n \n".encode()
    out += (
        f"trailer\n<< /Size {len(objects) + 1} /Root 1 0 R >>\nstartxref\n{xref}\n%%EOF\n"
    ).encode()
    with open(path, "wb") as handle:
        handle.write(bytes(out))
    return path


# ---------------------------------------------------------------------------
# fake osascript
# ---------------------------------------------------------------------------

FAKE_OSASCRIPT_SOURCE = r'''#!@PYTHON@
"""Fake osascript used by the mac-office unit tests (behaviour driven by env vars)."""

import json
import os
import re
import sys
import time


def env(name, default=None):
    return os.environ.get(name, default)


def log(record):
    path = env("MAC_OFFICE_FAKE_LOG")
    if path:
        with open(path, "a", encoding="utf-8") as handle:
            handle.write(json.dumps(record, ensure_ascii=False) + "\n")


def finish(stdout="", exit_code=0):
    override = env("MAC_OFFICE_FAKE_EXIT")
    code = int(override) if override not in (None, "") else exit_code
    stderr = env("MAC_OFFICE_FAKE_STDERR") or ""
    if stdout:
        sys.stdout.write(stdout)
    if stderr:
        sys.stderr.write(stderr)
    sys.stdout.flush()
    sys.stderr.flush()
    sys.exit(code)


def app_state():
    """The fake Office app as NSWorkspace would describe it (see MAC_OFFICE_FAKE_APP_STATE).

    A running app has pid 4242; a "stopped" one appears hidden with pid 4243 once the fake
    ``open`` launched it.
    """
    mode = env("MAC_OFFICE_FAKE_APP_STATE", "visible")
    if mode == "stopped":
        if not os.path.exists(env("MAC_OFFICE_FAKE_LAUNCHED_MARKER", "")):
            return None
        return {"name": "Microsoft Word", "pid": 4243, "hidden": True, "launched": True}
    return {"name": "Microsoft Word", "pid": 4242, "hidden": mode == "hidden", "launched": True}


def next_front():
    """Successive frontmost apps from MAC_OFFICE_FAKE_FRONT ("A,B,..."; the last one repeats)."""
    names = (env("MAC_OFFICE_FAKE_FRONT") or "Terminal").split(",")
    counter_path = env("MAC_OFFICE_FAKE_FRONT_COUNTER")
    calls = 0
    if counter_path and os.path.exists(counter_path):
        with open(counter_path, encoding="utf-8") as handle:
            calls = int(handle.read() or 0)
    if counter_path:
        with open(counter_path, "w", encoding="utf-8") as handle:
            handle.write(str(calls + 1))
    name = names[min(calls, len(names) - 1)]
    return {"name": name, "pid": 100 + names.index(name), "hidden": False, "launched": True}


def fake_jxa(source):
    log({"kind": "jxa", "script": source})
    if "mac-office-desktop-state" in source:
        if env("MAC_OFFICE_FAKE_DESKTOP") == "unreadable":
            finish("", 1)
        finish(json.dumps({"front": next_front(), "app": app_state()}))
    if "activateWithOptions" in source:
        finish("true")
    finish("")


def staged_hfs(source):
    """HFS form (as Word reports ``full name``) of the file an inspect script opens."""
    match = re.search(r'(?:POSIX file|workbook file name) "([^"]+)"', source)
    return ("Macintosh HD" + match.group(1).replace("/", ":")) if match else ""


def main(argv):
    if argv and argv[0] == "-l":
        if "-e" in argv:
            fake_jxa(argv[argv.index("-e") + 1])
        # JXA (the --render rasterizer) needs the real interpreter; PATH is hijacked here.
        os.execv("/usr/bin/osascript", ["osascript", *argv])

    if argv and argv[0] == "-e":
        script = argv[1] if len(argv) > 1 else ""
        log({"kind": "e", "script": script})
        if "set visible of process" in script:
            finish("")
        if "quit saving no" in script:
            finish(env("MAC_OFFICE_FAKE_QUIT_STDOUT") or "quit")
        if script.endswith(" to get name"):
            finish("Microsoft Word")
        if "System Events" in script:
            modal = env("MAC_OFFICE_FAKE_MODAL", "false")
            if modal == "fail":
                finish("", 1)
            if modal == "unreadable":
                finish("not-a-boolean")
            finish("true" if modal == "true" else "false")
        if any(
            marker in script
            for marker in ("close document i", "close workbook i", "close presentation i")
        ):
            finish("")
        finish(env("MAC_OFFICE_FAKE_E_STDOUT") or "")

    delay = env("MAC_OFFICE_FAKE_SLEEP")
    if delay:
        time.sleep(float(delay))

    script_path = argv[0] if argv else ""
    try:
        with open(script_path, encoding="utf-8") as handle:
            source = handle.read()
    except OSError:
        log({"kind": "script", "path": script_path, "script": None, "readable": False})
        finish("", 127)
    log({"kind": "script", "path": script_path, "script": source, "readable": True})
    if "mac-office-inspect" in source:
        finish(
            (env("MAC_OFFICE_FAKE_INSPECT_STDOUT") or "").replace("@STAGED_HFS@", staged_hfs(source)),
            int(env("MAC_OFFICE_FAKE_INSPECT_EXIT") or 0),
        )

    dump = env("MAC_OFFICE_FAKE_ENV_DUMP")
    if dump:
        names = [
            "MAC_OFFICE_TARGET_FILE",
            "MAC_OFFICE_OUTPUT_FILE",
            "MAC_OFFICE_STAGE_DIR",
            "MAC_OFFICE_ORIGINAL_FILE",
            "MAC_OFFICE_ASSET_COUNT",
        ]
        assets = sorted(
            key
            for key in os.environ
            if key.startswith("MAC_OFFICE_ASSET_") and key.rsplit("_", 1)[-1].isdigit()
        )
        with open(dump, "w", encoding="utf-8") as handle:
            json.dump({key: os.environ.get(key, "") for key in names + assets}, handle)

    target = os.environ.get("MAC_OFFICE_TARGET_FILE") or ""
    replacement = env("MAC_OFFICE_FAKE_REPLACE_TARGET")
    if replacement is not None and target:
        with open(target, "w", encoding="utf-8") as handle:
            handle.write(replacement)
    append = env("MAC_OFFICE_FAKE_APPEND_TARGET")
    if append and target:
        with open(target, "a", encoding="utf-8") as handle:
            handle.write(append)

    output = os.environ.get("MAC_OFFICE_OUTPUT_FILE") or ""
    output_content = env("MAC_OFFICE_FAKE_WRITE_OUTPUT")
    if output_content is not None and output:
        directory = os.path.dirname(output)
        if directory:
            os.makedirs(directory, exist_ok=True)
        with open(output, "w", encoding="utf-8") as handle:
            handle.write(output_content)

    # Copy a real file (e.g. a PDF) to the staged output path, byte for byte.
    output_source = env("MAC_OFFICE_FAKE_OUTPUT_SOURCE")
    if output_source and output:
        import shutil

        shutil.copyfile(output_source, output)

    # Copy the first staged asset to a path the test can inspect after the run.
    asset_copy_dest = env("MAC_OFFICE_FAKE_ASSET_COPY_DEST")
    first_asset = os.environ.get("MAC_OFFICE_ASSET_1") or ""
    if asset_copy_dest and first_asset:
        import shutil

        shutil.copyfile(first_asset, asset_copy_dest)

    finish(env("MAC_OFFICE_FAKE_STDOUT") or "")


main(sys.argv[1:])
'''


# Fake ``open`` (the runner starts stopped apps with ``open -g -j -a``): logs the call and
# marks the fake app as launched, so no real application is ever started by a unit test.
FAKE_OPEN_SOURCE = r'''#!@PYTHON@
import json, os, sys
with open(os.environ["MAC_OFFICE_FAKE_LOG"], "a", encoding="utf-8") as handle:
    handle.write(json.dumps({"kind": "open", "argv": sys.argv[1:]}) + "\n")
open(os.environ["MAC_OFFICE_FAKE_LAUNCHED_MARKER"], "w").close()
'''


class OsascriptShim:
    """Executable fake ``osascript`` (and ``open``) plus a log of everything they were asked."""

    def __init__(self, root):
        self.root = root
        self.bin_dir = os.path.join(root, "fakebin")
        os.makedirs(self.bin_dir, exist_ok=True)
        for name, source in (("osascript", FAKE_OSASCRIPT_SOURCE), ("open", FAKE_OPEN_SOURCE)):
            path = os.path.join(self.bin_dir, name)
            with open(path, "w", encoding="utf-8") as handle:
                handle.write(source.replace("@PYTHON@", sys.executable))
            os.chmod(path, 0o755)
        self.log_path = os.path.join(root, "osascript.log.jsonl")
        self.env_dump_path = os.path.join(root, "osascript-env.json")
        self.launched_marker = os.path.join(root, "fake-app-launched")
        self.front_counter = os.path.join(root, "fake-front-counter")

    def environment(self, **extra):
        env = dict(os.environ)
        env["PATH"] = self.bin_dir + os.pathsep + env.get("PATH", "")
        env["MAC_OFFICE_FAKE_LOG"] = self.log_path
        env["MAC_OFFICE_FAKE_ENV_DUMP"] = self.env_dump_path
        env["MAC_OFFICE_FAKE_LAUNCHED_MARKER"] = self.launched_marker
        env["MAC_OFFICE_FAKE_FRONT_COUNTER"] = self.front_counter
        for key, value in extra.items():
            if value is None:
                env.pop(key, None)
            else:
                env[key] = str(value)
        return env

    def calls(self):
        if not os.path.exists(self.log_path):
            return []
        with open(self.log_path, encoding="utf-8") as handle:
            return [json.loads(line) for line in handle.read().splitlines() if line.strip()]

    def scripts(self):
        return [call["script"] for call in self.calls() if call["kind"] == "script"]

    def e_scripts(self):
        return [call["script"] for call in self.calls() if call["kind"] == "e"]

    def jxa_scripts(self):
        return [call["script"] for call in self.calls() if call["kind"] == "jxa"]

    def open_calls(self):
        return [call["argv"] for call in self.calls() if call["kind"] == "open"]

    def last_script(self):
        scripts = self.scripts()
        if not scripts:
            raise AssertionError("osascript was never invoked with a script file")
        return scripts[-1]

    def last_script_path(self):
        for call in reversed(self.calls()):
            if call["kind"] == "script":
                return call["path"]
        raise AssertionError("osascript was never invoked with a script file")

    def env_dump(self):
        with open(self.env_dump_path, encoding="utf-8") as handle:
            return json.load(handle)


# ---------------------------------------------------------------------------
# test bases with isolation / leftover assertions
# ---------------------------------------------------------------------------


def _listing(path):
    try:
        return sorted(os.listdir(path))
    except (FileNotFoundError, NotADirectoryError):
        return []


def staging_leftovers(before):
    """New staging entries, ignoring the runner's persistent per-app lock/lifecycle files."""
    keep = constants().RUNNER_STATE_DIRNAME
    return sorted(name for name in set(_listing(staging_dir())) - set(before) if name != keep)


def _temp_scripts():
    return sorted(glob.glob(os.path.join(tempfile.gettempdir(), "*.applescript")))


_SUITE_LOCK_PATH = os.path.join(tempfile.gettempdir(), "mac-office-tests.lock")
_PROCESS_LOCK = None


class SuiteLock:
    """Cross-process lock around the staging-directory leftover assertions.

    Every runner invocation stages into its own subdirectory and removes it, so two runs
    never collide on paths.  The *assertion* still compares snapshots of the shared staging
    directory, so a second suite running at the same time would make it nondeterministic.
    """

    def __init__(self):
        self.handle = open(_SUITE_LOCK_PATH, "a+")
        fcntl.flock(self.handle, fcntl.LOCK_EX)

    def release(self):
        try:
            fcntl.flock(self.handle, fcntl.LOCK_UN)
        finally:
            self.handle.close()


def process_lock():
    """Hold the suite lock for the whole process, so two suites never interleave."""
    global _PROCESS_LOCK
    if _PROCESS_LOCK is None:
        _PROCESS_LOCK = SuiteLock()
        atexit.register(_PROCESS_LOCK.release)
    return _PROCESS_LOCK


class MacOfficeTestCase(unittest.TestCase):
    """Temp-dir isolation plus 'no leftovers' assertions after every test."""

    def setUp(self):
        super().setUp()
        process_lock()
        self.staging_before = _listing(staging_dir())
        self.temp_scripts_before = _temp_scripts()
        self.addCleanup(self._assert_no_leftovers)
        self.workdir = tempfile.mkdtemp(prefix="mac-office-test-")
        self.addCleanup(shutil.rmtree, self.workdir, True)

    def _assert_no_leftovers(self):
        new_staged = staging_leftovers(self.staging_before)
        self.assertEqual(new_staged, [], "staging directory leftovers: {}".format(new_staged))
        new_scripts = sorted(set(_temp_scripts()) - set(self.temp_scripts_before))
        self.assertEqual(new_scripts, [], "temp AppleScript leftovers: {}".format(new_scripts))

    def path(self, *parts):
        return os.path.join(self.workdir, *parts)

    def write(self, path, content=b"payload"):
        os.makedirs(os.path.dirname(path), exist_ok=True)
        mode = "wb" if isinstance(content, bytes) else "w"
        with open(path, mode) as handle:
            handle.write(content)
        return path


class ShimTestCase(MacOfficeTestCase):
    """Adds the fake ``osascript`` and a ``run`` helper bound to it."""

    def setUp(self):
        super().setUp()
        self.shim = OsascriptShim(self.workdir)
        # Fixtures with equal bytes would otherwise hit PDFs exported by earlier tests.
        shutil.rmtree(constants().PDF_CACHE_DIR, True)

    def invoke(self, *args, shim_env=None, stdin_text=None, timeout=300):
        return run_runner(
            list(args),
            env=self.shim.environment(**(shim_env or {})),
            stdin_text=stdin_text,
            timeout=timeout,
        )

    def write_script(self, name, source):
        return self.write(self.path(name), source)


# ---------------------------------------------------------------------------
# integration gating
# ---------------------------------------------------------------------------

INTEGRATION_REQUESTED = os.environ.get("MAC_OFFICE_INTEGRATION") == "1"
APP_BUNDLES = {
    "word": "Microsoft Word",
    "excel": "Microsoft Excel",
    "powerpoint": "Microsoft PowerPoint",
}


def app_installed(app):
    return os.path.isdir(os.path.join("/Applications", APP_BUNDLES[app] + ".app"))


def integration_reason(app):
    if not INTEGRATION_REQUESTED:
        return "set MAC_OFFICE_INTEGRATION=1 to run real-Office integration tests"
    if not app_installed(app):
        return "{} is not installed in /Applications".format(APP_BUNDLES[app])
    return ""


def require_integration(testcase, app):
    """Skip the calling test unless integration is enabled and the app is installed."""
    reason = integration_reason(app)
    if reason:
        testcase.skipTest(reason)


def is_ole2_container(path):
    with open(path, "rb") as handle:
        return handle.read(8) == OLE2_MAGIC


DESKTOP_SAMPLER_JXA = r"""
ObjC.import('AppKit');
ObjC.import('CoreGraphics');
const OFFICE = ['Microsoft Word', 'Microsoft Excel', 'Microsoft PowerPoint'];
function sample() {
  const workspace = $.NSWorkspace.sharedWorkspace;
  const front = workspace.frontmostApplication;
  const apps = {};
  const running = workspace.runningApplications;
  for (let i = 0; i < running.count; i++) {
    const app = running.objectAtIndex(i);
    const name = app.localizedName.isNil() ? '' : app.localizedName.js;
    if (OFFICE.indexOf(name) >= 0) apps[name] = {hidden: app.isHidden, windows: 0};
  }
  const windows = ObjC.deepUnwrap(ObjC.castRefToObject($.CGWindowListCopyWindowInfo(
    $.kCGWindowListOptionOnScreenOnly | $.kCGWindowListExcludeDesktopElements, 0)));
  for (const window of windows) {
    const owner = apps[window.kCGWindowOwnerName];
    if (owner && window.kCGWindowLayer === 0 && window.kCGWindowAlpha > 0) owner.windows += 1;
  }
  return {front: front.isNil() ? null : front.localizedName.js, apps: apps};
}
function run() {
  while (true) {
    console.log(JSON.stringify(sample()));
    delay(0.1);
  }
}
"""


class DesktopSampler:
    """Samples the frontmost app and each Office app's on-screen windows every ~100 ms.

    A *focus steal* is a sample whose frontmost app is an Office app right after a sample
    whose frontmost app was not. On-screen windows are counted per app, split by whether
    the app was hidden when sampling started (runner-launched apps run hidden).
    """

    def __init__(self):
        self.script = tempfile.NamedTemporaryFile("w", suffix=".js", delete=False)
        self.script.write(DESKTOP_SAMPLER_JXA)
        self.script.close()
        self.process = subprocess.Popen(
            ["/usr/bin/osascript", "-l", "JavaScript", self.script.name],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.PIPE,
            text=True,
        )

    def stop(self):
        self.process.terminate()
        _, output = self.process.communicate(timeout=10)
        os.unlink(self.script.name)
        samples = []
        for line in output.splitlines():
            try:
                samples.append(json.loads(line))
            except ValueError:
                pass
        return summarize_desktop_samples(samples)


def summarize_desktop_samples(samples):
    office = ("Microsoft Word", "Microsoft Excel", "Microsoft PowerPoint")
    steals = 0
    for previous, current in zip(samples, samples[1:]):
        if current["front"] in office and previous["front"] not in office:
            steals += 1
    hidden_at_start = {
        name for name, state in (samples[0]["apps"] if samples else {}).items() if state["hidden"]
    }
    visible = {}
    for sample in samples:
        for name, state in sample["apps"].items():
            if state["windows"] and not state["hidden"]:
                visible[name] = visible.get(name, 0) + 1
    return {
        "samples": len(samples),
        "focus_steals": steals,
        "visible_window_samples": visible,
        "hidden_at_start": sorted(hidden_at_start),
    }


class IntegrationCase(MacOfficeTestCase):
    """Base for real-Office tests: class-level fixtures, per-test isolated work dirs.

    Subclasses set ``app`` and build their fixtures in ``setUpClass`` via ``cls.cli``.
    The Office staging directory and the temp AppleScript directory are checked for
    leftovers once the class finishes.
    """

    app = None

    @classmethod
    def setUpClass(cls):
        if cls.app is None:
            raise ValueError("IntegrationCase subclass must define app")
        reason = integration_reason(cls.app)
        if reason:
            raise unittest.SkipTest(reason)
        process_lock()
        cls.class_dir = tempfile.mkdtemp(prefix="mac-office-integration-")
        cls.class_staging_before = _listing(staging_dir())
        cls.class_scripts_before = _temp_scripts()
        cls.sampler = DesktopSampler()

    @classmethod
    def tearDownClass(cls):
        desktop = cls.sampler.stop()
        sys.stderr.write("\n[desktop sampler] {}: {}\n".format(cls.__name__, json.dumps(desktop)))
        if desktop["focus_steals"]:
            raise AssertionError("Office took the focus {} time(s): {}".format(desktop["focus_steals"], desktop))
        shutil.rmtree(cls.class_dir, True)
        leftovers = staging_leftovers(cls.class_staging_before)
        if leftovers:
            raise AssertionError("staging directory leftovers: {}".format(leftovers))
        scripts = sorted(set(_temp_scripts()) - set(cls.class_scripts_before))
        if scripts:
            raise AssertionError("temp AppleScript leftovers: {}".format(scripts))

    @classmethod
    def artifact(cls, name):
        return os.path.join(cls.class_dir, name)

    @classmethod
    def cli(cls, *args, stdin_text=None, timeout=240):
        return run_runner(list(args), stdin_text=stdin_text, timeout=timeout)


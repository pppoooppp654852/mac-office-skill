"""Native-app AppleScript inspectors for OLE2 containers.

On MIP-managed Macs every Office-saved file is an OLE2 container (Word 97-2003 binary or an
RMS ``EncryptedPackage`` wrapper), so these inspectors are the main path in practice: only
the owning app can read such a file.

The scripts answer with one stdout blob built from ASCII control-character separators
(field ``\\x1f``, record ``\\x1d``, section ``\\x1e``), which cannot occur inside document
text. Values are coerced with explicit ``as text``: coercing a Word ``content`` implicitly
(e.g. ``quoted form of (content of ...)``) yields ``missing value``. Every per-item read sits
in its own ``try`` block so one failed property degrades to null instead of aborting the
inspection.

The staged copy is opened read-only and looked up by its unique staged name. When the app
swallows the open the script answers ``OPEN_FAILED`` instead of falling back to
``active document``, so a document the user has open is never read or closed.
"""

import os

from .app_session import run_in_app
from .applescript import encode_string
from .constants import (
    APP_FOR_KIND,
    APP_NAMES,
    FIELD_SEPARATOR,
    OPEN_FAILED,
    RECORD_SEPARATOR,
    SECTION_SEPARATOR,
)
from .container import staged_extension
from .staging import StageDir, hfs_to_posix
from .text import trim_text

# Inspecting a large document item by item over Apple events is slow, so the fallback never
# gets less than this, even when --timeout is lower.
MIN_INSPECT_TIMEOUT = 60

PREVIEW_CHARS = 80
PREVIEW_ROWS = 5
SHEET_PREVIEW_ROWS = 10
SHEET_PREVIEW_COLUMNS = 20
SHAPE_TEXT_CHARS = 120

HANDLERS = """
on enc(v)
	if v is missing value then return ""
	try
		set v to v as text
	on error
		return ""
	end try
	return v
end enc

on num(v)
	if v is missing value then return ""
	try
		set v to v as real
	on error
		return ""
	end try
	try
		return v as text
	on error
		return ""
	end try
end num
"""

SEPARATORS = (
    "set fSep to ASCII character 31\n"
    "set rSep to ASCII character 29\n"
    "set sSep to ASCII character 30\n"
)

WORD_SCRIPT = (
    "-- mac-office-inspect: docx\n"
    + HANDLERS
    + """
on run
	"""
    + SEPARATORS
    + """	set maxItems to @MAX@
	tell application "Microsoft Word"
		set display alerts to alerts none
		open (POSIX file @STAGED@) read only true add to recent files false
		if not (exists document @NAME@) then return "@OPEN_FAILED@"
		set d to document @NAME@
		if not @APP_HIDDEN@ then
			try
				set window state of window 1 of d to window state minimize
			end try
		end if
		set pCount to (count of paragraphs of d) as integer
		set tCount to (count of tables of d) as integer
		set shCount to (count of shapes of d) as integer
		set isCount to (count of inline shapes of d) as integer
		set meta to (my enc(full name of d)) & fSep & (pCount as text) & fSep & (tCount as text) & fSep & (shCount as text) & fSep & (isCount as text)
		try
			set ps to page setup of d
			set meta to meta & fSep & (my num(page width of ps)) & fSep & (my num(page height of ps))
			set meta to meta & fSep & (my num(top margin of ps)) & fSep & (my num(bottom margin of ps))
			set meta to meta & fSep & (my num(left margin of ps)) & fSep & (my num(right margin of ps))
		on error
			set meta to meta & fSep & fSep & fSep & fSep & fSep & fSep
		end try
		set pLimit to pCount
		if pLimit > maxItems then set pLimit to maxItems
		set pOut to {}
		repeat with i from 1 to pLimit
			set rec to (i as text)
			try
				set tro to text object of paragraph i of d
				set rec to rec & fSep & (my num(start of content of tro)) & fSep & (my num(end of content of tro))
				set rec to rec & fSep & (my enc(alignment of paragraph i of d))
				set rec to rec & fSep & (my num(font size of font object of tro))
				set rec to rec & fSep & (my enc(content of tro))
			on error
				set rec to rec & fSep & fSep & fSep & fSep & fSep
			end try
			set end of pOut to rec
		end repeat
		set AppleScript's text item delimiters to rSep
		set pSection to pOut as text
		set tLimit to tCount
		if tLimit > maxItems then set tLimit to maxItems
		set tOut to {}
		set cOut to {}
		repeat with j from 1 to tLimit
			set tb to table j of d
			set rCount to 0
			set cCount to 0
			try
				set rCount to (count of rows of tb) as integer
				set cCount to (count of columns of tb) as integer
			end try
			set end of tOut to (j as text) & fSep & (rCount as text) & fSep & (cCount as text)
			set cRead to 0
			repeat with r from 1 to rCount
				if cRead >= maxItems then exit repeat
				repeat with c from 1 to cCount
					if cRead >= maxItems then exit repeat
					set cRead to cRead + 1
					set crec to (j as text) & fSep & (r as text) & fSep & (c as text)
					try
						set ctr to text object of cell c of row r of tb
						set crec to crec & fSep & (my num(start of content of ctr)) & fSep & (my num(end of content of ctr))
						set crec to crec & fSep & (my enc(content of ctr))
					on error
						set crec to crec & fSep & fSep & fSep & fSep
					end try
					set end of cOut to crec
				end repeat
			end repeat
		end repeat
		set AppleScript's text item delimiters to rSep
		set tSection to tOut as text
		set cSection to cOut as text
		set shLimit to shCount
		if shLimit > maxItems then set shLimit to maxItems
		set shOut to {}
		repeat with k from 1 to shLimit
			set rec to (k as text)
			try
				set s to shape k of d
				set rec to rec & fSep & (my enc(name of s)) & fSep & (my enc(shape type of s))
				set rec to rec & fSep & (my num(left position of s)) & fSep & (my num(top of s))
				set rec to rec & fSep & (my num(width of s)) & fSep & (my num(height of s))
				set boxText to ""
				try
					set boxText to my enc(content of text range of text frame of s)
				end try
				set rec to rec & fSep & boxText
			on error
				set rec to rec & fSep & fSep & fSep & fSep & fSep & fSep & fSep
			end try
			set end of shOut to rec
		end repeat
		set AppleScript's text item delimiters to rSep
		set shSection to shOut as text
		try
			close d saving no
		end try
		set display alerts to alerts all
	end tell
	set AppleScript's text item delimiters to sSep
	return meta & sSep & pSection & sSep & cSection & sSep & tSection & sSep & shSection
end run
"""
)

EXCEL_SCRIPT = (
    "-- mac-office-inspect: xlsx\n"
    + HANDLERS
    + """
on run
	"""
    + SEPARATORS
    + """	set maxItems to @MAX@
	tell application "Microsoft Excel"
		set display alerts to false
		open workbook workbook file name @STAGED@ read only true add to mru false
		if not (exists workbook @NAME@) then return "@OPEN_FAILED@"
		set wb to workbook @NAME@
		set sCount to (count of worksheets of wb) as integer
		set wbName to my enc(full name of wb)
		set sOut to {}
		set cOut to {}
		set sLimit to sCount
		if sLimit > maxItems then set sLimit to maxItems
		repeat with i from 1 to sLimit
			set rec to my enc(name of worksheet i of wb)
			set addrText to ""
			set urRows to 0
			set urCols to 0
			try
				set ur to used range of worksheet i of wb
				try
					set addrText to my enc(get address ur)
				end try
				try
					set urRows to (count of rows of ur) as integer
					set urCols to (count of columns of ur) as integer
				end try
			end try
			set rec to rec & fSep & addrText & fSep & (urRows as text) & fSep & (urCols as text)
			set end of sOut to rec
			set pRows to urRows
			if pRows > 10 then set pRows to 10
			set pCols to urCols
			if pCols > 20 then set pCols to 20
			repeat with r from 1 to pRows
				repeat with c from 1 to pCols
					set cellText to ""
					try
						set cellText to my enc(value of cell c of row r of ur)
					end try
					set end of cOut to (i as text) & fSep & (r as text) & fSep & (c as text) & fSep & cellText
				end repeat
			end repeat
		end repeat
		set AppleScript's text item delimiters to rSep
		set sSection to sOut as text
		set cSection to cOut as text
		try
			close wb saving no
		end try
		set display alerts to true
	end tell
	return wbName & fSep & (sCount as text) & sSep & sSection & sSep & cSection
end run
"""
)

POWERPOINT_SCRIPT = (
    "-- mac-office-inspect: pptx\n"
    + HANDLERS
    + """
on run
	"""
    + SEPARATORS
    + """	set maxItems to @MAX@
	tell application "Microsoft PowerPoint"
		open (POSIX file @STAGED@)
		repeat 25 times
			if exists presentation @NAME@ then exit repeat
			delay 0.1
		end repeat
		if not (exists presentation @NAME@) then return "@OPEN_FAILED@"
		set pres to presentation @NAME@
		set sCount to (count of slides of pres) as integer
		set presName to my enc(full name of pres)
		set sOut to {}
		set shOut to {}
		set sLimit to sCount
		if sLimit > maxItems then set sLimit to maxItems
		repeat with i from 1 to sLimit
			set sl to slide i of pres
			set shCount to 0
			try
				set shCount to (count of shapes of sl) as integer
			end try
			set end of sOut to (i as text) & fSep & (shCount as text)
			set shLimit to shCount
			if shLimit > maxItems then set shLimit to maxItems
			repeat with j from 1 to shLimit
				set rec to (i as text) & fSep & (j as text)
				try
					set sh to shape j of sl
					set rec to rec & fSep & (my enc(name of sh))
					set rec to rec & fSep & (my num(left position of sh)) & fSep & (my num(top of sh))
					set rec to rec & fSep & (my num(width of sh)) & fSep & (my num(height of sh))
					set shText to ""
					try
						if has text frame of sh then
							set shText to my enc(content of text range of text frame of sh)
						end if
					end try
					set rec to rec & fSep & shText
				on error
					set rec to rec & fSep & fSep & fSep & fSep & fSep & fSep
				end try
				set end of shOut to rec
			end repeat
		end repeat
		set AppleScript's text item delimiters to rSep
		set slSection to sOut as text
		set shSection to shOut as text
		try
			close pres saving no
		end try
	end tell
	return presName & fSep & (sCount as text) & sSep & slSection & sSep & shSection
end run
"""
)

SCRIPTS = {"docx": WORD_SCRIPT, "xlsx": EXCEL_SCRIPT, "pptx": POWERPOINT_SCRIPT}


def _number(value):
    try:
        return round(float(value), 2)
    except (TypeError, ValueError):
        return None


def _integer(value):
    try:
        return int(float(value))
    except (TypeError, ValueError):
        return None


def _integer_or(value, fallback):
    parsed = _integer(value)
    return fallback if parsed is None else parsed


def native_text(value, limit=None):
    """Drop Word/Excel control marks (cell and row marks) and collapse whitespace."""
    cleaned = "".join(character for character in value or "" if character >= " " or character in "\t\n\r")
    return trim_text(cleaned, limit or 200)


def _records(section):
    return [record.split(FIELD_SEPARATOR) for record in section.split(RECORD_SEPARATOR) if record.strip()]


def _sections(stdout):
    parts = [part.lstrip("\r\n") for part in (stdout or "").split(SECTION_SEPARATOR)]
    return parts + [""] * (5 - len(parts))


def _field(fields, index):
    return fields[index] if len(fields) > index else ""


def _capped(total, max_items):
    """True when the returned list was cut by ``--max-items``."""
    return total > max_items


def _unreadable(total, max_items, returned, degraded):
    """Items within the cap that could not be read: missing records plus null-only ones."""
    return max(0, min(total, max_items) - len(returned)) + degraded


def _degraded(shape):
    """A shape whose whole read failed in the app: every property came back empty."""
    return all(shape[key] is None for key in ("name", "left_pt", "top_pt", "width_pt", "height_pt"))


def _verify_opened(meta_fields, staged_path):
    """Fail loudly when the app reports another document than the staged copy."""
    opened = hfs_to_posix(_field(meta_fields, 0))
    if opened != staged_path:
        raise RuntimeError(
            f"the app did not open {os.path.basename(staged_path)} (opened: "
            f"{opened or 'nothing'}); the file format may be blocked or its extension may not "
            "match the container"
        )


def parse_word(stdout, max_items, staged_path):
    meta, paragraph_section, cell_section, table_section, shape_section = _sections(stdout)
    meta_fields = meta.split(FIELD_SEPARATOR)
    _verify_opened(meta_fields, staged_path)
    paragraph_count = _integer_or(_field(meta_fields, 1), 0)
    table_count = _integer_or(_field(meta_fields, 2), 0)
    shape_count = _integer_or(_field(meta_fields, 3), 0)
    inline_shape_count = _integer_or(_field(meta_fields, 4), 0)

    cells, cell_text = [], {}
    for fields in _records(cell_section):
        table = _integer(_field(fields, 0))
        row = _integer(_field(fields, 1))
        column = _integer(_field(fields, 2))
        start = _number(_field(fields, 3))
        end = _number(_field(fields, 4))
        if None in (table, row, column, start, end):
            continue
        cells.append({"table": table, "row": row, "column": column, "start": start, "end": end})
        cell_text[(table, row, column)] = native_text(_field(fields, 5), PREVIEW_CHARS)

    paragraphs, paragraph_starts = [], []
    for fields in _records(paragraph_section):
        index = _integer(_field(fields, 0))
        if index is None:
            continue
        start = _number(_field(fields, 1))
        paragraph_starts.append(start)
        entry = {
            "index": index,
            "text": native_text(_field(fields, 5)),
            "alignment": _field(fields, 3) or None,
            "font_size": _number(_field(fields, 4)),
            "table": None,
            "row": None,
            "column": None,
        }
        for cell in cells:
            if start is not None and cell["start"] <= start < cell["end"]:
                entry.update(table=cell["table"], row=cell["row"], column=cell["column"])
        paragraphs.append(entry)

    tables = []
    cells_by_table = {}
    for cell in cells:
        cells_by_table.setdefault(cell["table"], []).append(cell)
    for index, (row_count, column_count) in enumerate(
        (
            (_integer_or(_field(fields, 1), 0), _integer_or(_field(fields, 2), 0))
            for fields in _records(table_section)
        ),
        start=1,
    ):
        table_cells = cells_by_table.get(index, [])
        preview = [
            [cell_text.get((index, row, column), "") for column in range(1, column_count + 1)]
            for row in range(1, min(row_count, PREVIEW_ROWS) + 1)
        ]
        cell_total = row_count * column_count
        tables.append(
            {
                "index": index,
                "rows": row_count,
                "columns": column_count,
                "preview": preview,
                "cells_truncated": _capped(cell_total, max_items),
                "cells_unreadable": _unreadable(cell_total, max_items, table_cells, 0),
            }
        )

    shapes = []
    for fields in _records(shape_section):
        index = _integer(_field(fields, 0))
        if index is None:
            continue
        shapes.append(
            {
                "index": index,
                "name": _field(fields, 1) or None,
                "type": _field(fields, 2) or None,
                "left_pt": _number(_field(fields, 3)),
                "top_pt": _number(_field(fields, 4)),
                "width_pt": _number(_field(fields, 5)),
                "height_pt": _number(_field(fields, 6)),
                "text": native_text(_field(fields, 7), SHAPE_TEXT_CHARS),
            }
        )

    page_setup = None
    keys = ("page_width", "page_height", "top_margin", "bottom_margin", "left_margin", "right_margin")
    values = [_number(_field(meta_fields, 5 + offset)) for offset in range(len(keys))]
    if any(value is not None for value in values):
        page_setup = dict(zip(keys, values))

    return {
        "inspector": "applescript_fallback",
        "paragraph_count": paragraph_count,
        "paragraphs_truncated": _capped(paragraph_count, max_items),
        "paragraphs_unreadable": _unreadable(
            paragraph_count, max_items, paragraphs, sum(1 for start in paragraph_starts if start is None)
        ),
        "paragraphs": paragraphs[:max_items],
        "table_count": table_count,
        "tables_truncated": _capped(table_count, max_items),
        "tables": tables[:max_items],
        "shape_count": shape_count,
        "shapes_truncated": _capped(shape_count, max_items),
        "shapes_unreadable": _unreadable(
            shape_count, max_items, shapes, sum(1 for shape in shapes if _degraded(shape))
        ),
        "shapes": shapes[:max_items],
        "inline_shape_count": inline_shape_count,
        "page_setup": page_setup,
    }


def parse_excel(stdout, max_items, staged_path):
    meta, sheet_section, cell_section, _, _ = _sections(stdout)
    meta_fields = meta.split(FIELD_SEPARATOR)
    _verify_opened(meta_fields, staged_path)
    sheet_count = _integer_or(_field(meta_fields, 1), 0)
    cells = []
    for fields in _records(cell_section):
        sheet = _integer(_field(fields, 0))
        row = _integer(_field(fields, 1))
        column = _integer(_field(fields, 2))
        if None in (sheet, row, column):
            continue
        cells.append((sheet, row, column, native_text(_field(fields, 3), PREVIEW_CHARS)))
    worksheets = []
    for index, fields in enumerate(_records(sheet_section), start=1):
        used_range = (_field(fields, 1) or "").replace("$", "") or None
        rows = _integer_or(_field(fields, 2), 0)
        columns = _integer_or(_field(fields, 3), 0)
        preview_rows = min(rows, SHEET_PREVIEW_ROWS)
        preview_columns = min(columns, SHEET_PREVIEW_COLUMNS)
        grid = {}
        for sheet, row, column, text in cells:
            if sheet == index and row <= preview_rows and column <= preview_columns:
                grid[(row, column)] = text
        preview = [
            [grid.get((row, column), "") for column in range(1, preview_columns + 1)]
            for row in range(1, preview_rows + 1)
        ]
        worksheets.append(
            {
                "name": _field(fields, 0),
                "used_range": used_range,
                "preview": preview,
                "truncated": preview_rows < rows or preview_columns < columns,
            }
        )
    return {
        "inspector": "applescript_fallback",
        "sheet_count": sheet_count,
        "sheets_truncated": _capped(sheet_count, max_items),
        "worksheets": worksheets[:max_items],
    }


def parse_powerpoint(stdout, max_items, staged_path):
    meta, slide_section, shape_section, _, _ = _sections(stdout)
    meta_fields = meta.split(FIELD_SEPARATOR)
    _verify_opened(meta_fields, staged_path)
    slide_count = _integer_or(_field(meta_fields, 1), 0)
    shapes_by_slide = {}
    for fields in _records(shape_section):
        slide = _integer(_field(fields, 0))
        index = _integer(_field(fields, 1))
        if slide is None or index is None:
            continue
        shapes_by_slide.setdefault(slide, []).append(
            {
                "index": index,
                "id": None,
                "name": _field(fields, 2) or None,
                "kind": "sp",
                "left_pt": _number(_field(fields, 3)),
                "top_pt": _number(_field(fields, 4)),
                "width_pt": _number(_field(fields, 5)),
                "height_pt": _number(_field(fields, 6)),
                "text": native_text(_field(fields, 7), SHAPE_TEXT_CHARS),
            }
        )
    slides = []
    for fields in _records(slide_section):
        index = _integer(_field(fields, 0))
        if index is None:
            continue
        shapes = shapes_by_slide.get(index, [])
        shape_count = _integer_or(_field(fields, 1), len(shapes))
        slides.append(
            {
                "index": index,
                "shape_count": shape_count,
                "shapes_truncated": _capped(shape_count, max_items),
                "shapes_unreadable": _unreadable(
                    shape_count, max_items, shapes, sum(1 for shape in shapes if _degraded(shape))
                ),
                "shapes": shapes[:max_items],
            }
        )
    return {
        "inspector": "applescript_fallback",
        "slide_count": slide_count,
        "slides_truncated": _capped(slide_count, max_items),
        "slides": slides[:max_items],
    }


PARSERS = {"docx": parse_word, "xlsx": parse_excel, "pptx": parse_powerpoint}


class AppInspectError(RuntimeError):
    """The native-app inspection failed; ``hint`` becomes the result's ``error_hint``."""

    def __init__(self, message, hint=None):
        super().__init__(message)
        self.hint = hint


def inspect_via_app(path, kind, max_items, timeout):
    """Inspect ``path`` with the owning Office app; raises AppInspectError on failure."""
    app = APP_FOR_KIND[kind]
    timeout = max(MIN_INSPECT_TIMEOUT, timeout)
    with StageDir(label=os.path.basename(path)) as stage:
        staged = stage.stage_uniquely(path, staged_extension(path, kind))
        script = (
            SCRIPTS[kind]
            .replace("@MAX@", str(max_items))
            .replace("@OPEN_FAILED@", OPEN_FAILED)
            .replace("@STAGED@", encode_string(staged))
            .replace("@NAME@", encode_string(os.path.basename(staged)))
        )
        run = run_in_app(app, script, timeout, staged_paths=[staged])
    if run.completed is None:
        message = f"{APP_NAMES[app]} timed out after {timeout}s" if run.timed_out else run.error
        raise AppInspectError(message, run.error_hint)
    completed = run.completed
    if completed.returncode != 0:
        raise AppInspectError((completed.stderr or completed.stdout).strip())
    if (completed.stdout or "").strip() == OPEN_FAILED:
        raise AppInspectError(
            f"{APP_NAMES[app]} could not open the file (it ignored the request); the format may "
            "be blocked by policy or the extension may not match the container"
        )
    return PARSERS[kind](completed.stdout, max_items, staged), run.focus_restored

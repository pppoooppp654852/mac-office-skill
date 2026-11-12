#!/usr/bin/env python3
"""Rule-based grader: score one finished task from the files it left and its transcript.

Every check reads the deliverable through office_runner.py itself (--inspect, --render, or a
read-only --app probe for Excel formulas), so a score is reproducible and never depends on an
LLM. Each failed check carries a semantic ``feedback`` sentence (what is wrong with the
document, never the checker's thresholds or patterns): that sentence is all the optimizer
sees of the grader.

hard = 1 only when every check passes; soft = fraction of checks passed.
"""

import os
import tempfile

from runner_cli import run_runner
from task_catalog import before_inspect, before_render, md5

LAYOUT_TOLERANCE_PT = 0.5
FIELD = "\x1f"
RECORD = "\x1e"


def norm(text):
    return " ".join(str(text if text is not None else "").split())


class Deliverable:
    """Lazily inspected / rendered view of the graded file."""

    def __init__(self, path, work_dir, excel_cells=()):
        self.path = path
        self.work_dir = work_dir
        self.excel_cells = list(excel_cells)
        self._inspect = None
        self._render = None
        self._excel = None

    @property
    def exists(self):
        return bool(self.path) and os.path.isfile(self.path)

    def inspect(self):
        if self._inspect is None:
            self._inspect = run_runner(["--inspect", self.path, "--max-items", "400"])
        return self._inspect

    def render(self):
        if self._render is None:
            self._render = run_runner(
                ["--render", self.path, "--output-dir", os.path.join(self.work_dir, "render"), "--dpi", "60"]
            )
        return self._render

    def excel(self):
        """``({(sheet, cell): {"value", "formula"}}, error)`` for every cell the checks name."""
        if self._excel is None:
            self._excel = probe_excel(self.path, self.excel_cells)
        return self._excel


def all_texts(inspect):
    texts = [p.get("text") for p in inspect.get("paragraphs", [])]
    for table in inspect.get("tables", []):
        for row in table.get("preview", []):
            texts.extend(row)
    texts.extend(shape.get("text") for shape in inspect.get("shapes", []))
    for slide in inspect.get("slides", []):
        texts.extend(shape.get("text") for shape in slide.get("shapes", []))
    return [norm(text) for text in texts if text]


def short(text, limit=60):
    text = norm(text)
    return text if len(text) <= limit else text[: limit - 1] + "…"


# --- content checks: each returns (passed, feedback) -------------------------------------

def check_paragraph_count(check, deliverable):
    actual = deliverable.inspect().get("paragraph_count")
    if actual == check["equals"]:
        return True, ""
    return False, (
        f"The document should still have {check['equals']} paragraphs but has {actual}: "
        "paragraphs were merged, split, or an empty paragraph was added."
    )


def _paragraph(inspect, index):
    for paragraph in inspect.get("paragraphs", []):
        if paragraph.get("index") == index:
            return paragraph
    return None


def check_paragraph_text(check, deliverable):
    paragraph = _paragraph(deliverable.inspect(), check["index"])
    actual = paragraph.get("text") if paragraph else None
    if norm(actual) == norm(check["equals"]):
        return True, ""
    return False, f"Paragraph {check['index']} should read “{check['equals']}” but reads “{short(actual)}”."


def check_paragraph_alignment(check, deliverable):
    paragraph = _paragraph(deliverable.inspect(), check["index"])
    actual = (paragraph or {}).get("alignment")
    if actual == check["equals"]:
        return True, ""
    return False, (
        f"Paragraph {check['index']} (“{short((paragraph or {}).get('text'))}”) should keep its "
        f"original alignment ({check['equals']}) but is now {actual}."
    )


def _table(inspect, index):
    tables = inspect.get("tables", [])
    return tables[index - 1] if 0 < index <= len(tables) else None


def check_cell_text(check, deliverable):
    table = _table(deliverable.inspect(), check["table"])
    actual = None
    if table and check["row"] <= len(table["preview"]) and check["column"] <= len(table["preview"][0]):
        actual = table["preview"][check["row"] - 1][check["column"] - 1]
    if norm(actual) == norm(check["equals"]):
        return True, ""
    return False, (
        f"Table {check['table']} cell (row {check['row']}, column {check['column']}) should read "
        f"“{check['equals']}” but reads “{short(actual)}”."
    )


def check_table_shape(check, deliverable):
    table = _table(deliverable.inspect(), check["table"])
    actual = (table["rows"], table["columns"]) if table else None
    if actual == (check["rows"], check["columns"]):
        return True, ""
    return False, (
        f"Table {check['table']} should keep {check['rows']} rows × {check['columns']} columns "
        f"but has {actual[0]} × {actual[1]}." if actual else f"Table {check['table']} is missing."
    )


def check_shape_text(check, deliverable):
    shapes = {shape.get("name"): shape for shape in deliverable.inspect().get("shapes", [])}
    actual = (shapes.get(check["name"]) or {}).get("text")
    if norm(actual) == norm(check["equals"]):
        return True, ""
    return False, f"The text box “{check['name']}” should read “{check['equals']}” but reads “{short(actual)}”."


def check_picture_count(check, deliverable):
    inspect = deliverable.inspect()
    floating = sum(1 for shape in inspect.get("shapes", []) if "picture" in (shape.get("type") or ""))
    actual = floating + (inspect.get("inline_shape_count") or 0)
    if actual == check["equals"]:
        return True, ""
    return False, f"The document should contain exactly {check['equals']} picture(s) but contains {actual}."


def check_shape_count(check, deliverable):
    actual = deliverable.inspect().get("shape_count")
    if actual == check["equals"]:
        return True, ""
    return False, f"The document should keep {check['equals']} floating shapes but has {actual}."


def check_text_present(check, deliverable):
    texts = all_texts(deliverable.inspect())
    count = sum(text.count(norm(check["text"])) for text in texts)
    if count >= check.get("min", 1):
        return True, ""
    return False, f"The document should contain “{check['text']}” at least {check.get('min', 1)} time(s); found {count}."


def check_text_absent(check, deliverable):
    texts = all_texts(deliverable.inspect())
    hits = [text for text in texts if norm(check["text"]) in text]
    if not hits:
        return True, ""
    return False, f"“{check['text']}” should no longer appear, but is still in: “{short(hits[0])}”."


def check_layout_preserved(check, deliverable, task):
    before = before_render(task)
    after = deliverable.render()
    if not after.get("success"):
        return False, f"The deliverable could not be rendered: {after.get('error')}"
    if before["pages"] != after["pages"]:
        return False, f"The page count changed from {before['pages']} to {after['pages']}."
    ignore = [norm(token) for token in check.get("ignore_lines_containing", [])]
    remaining = [dict(line, used=False) for line in after["lines"]]
    moved = []
    for line in before["lines"]:
        text = norm(line["text"])
        if any(token in text for token in ignore):
            continue
        match = next(
            (
                candidate
                for candidate in remaining
                if not candidate["used"]
                and candidate["page"] == line["page"]
                and norm(candidate["text"]) == text
                and abs(candidate["x"] - line["x"]) <= LAYOUT_TOLERANCE_PT
                and abs(candidate["y"] - line["y"]) <= LAYOUT_TOLERANCE_PT
            ),
            None,
        )
        if match:
            match["used"] = True
        else:
            moved.append(text)
    if not moved:
        return True, ""
    sample = "; ".join(f"“{short(text, 40)}”" for text in moved[:3])
    return False, f"Content the user did not ask to change moved or disappeared on the page: {sample}."


def check_page_count(check, deliverable, task):
    before = before_render(task)
    after = deliverable.render()
    if after.get("success") and after["pages"] == before["pages"]:
        return True, ""
    return False, f"The page count changed from {before['pages']} to {after.get('pages')}."


def check_pdf_text(check, docs_dir, work_dir):
    path = os.path.join(docs_dir, check["file"])
    if not os.path.isfile(path):
        return False, f"The requested PDF {check['file']} was not produced."
    render = run_runner(["--render", path, "--output-dir", os.path.join(work_dir, "pdf"), "--dpi", "40"])
    if not render.get("success"):
        return False, f"The PDF {check['file']} cannot be read: {render.get('error')}"
    text = norm(" ".join(line["text"] for line in render["lines"]))
    squeezed = text.replace(" ", "")
    missing = [item for item in check.get("contains", []) if norm(item).replace(" ", "") not in squeezed]
    present = [item for item in check.get("absent", []) if norm(item).replace(" ", "") in squeezed]
    if "pages" in check and render["pages"] != check["pages"]:
        return False, f"The PDF should have {check['pages']} page(s) but has {render['pages']}."
    if missing:
        return False, f"The PDF does not show the expected content “{missing[0]}”."
    if present:
        return False, f"The PDF still shows outdated content “{present[0]}”."
    return True, ""


def check_slide_count(check, deliverable):
    actual = deliverable.inspect().get("slide_count")
    if actual == check["equals"]:
        return True, ""
    return False, f"The presentation should keep {check['equals']} slides but has {actual}."


def check_slide_shape_text(check, deliverable):
    slides = deliverable.inspect().get("slides", [])
    actual = None
    if check["slide"] <= len(slides):
        shapes = slides[check["slide"] - 1]["shapes"]
        if check["shape"] <= len(shapes):
            actual = shapes[check["shape"] - 1].get("text")
    if norm(actual) == norm(check["equals"]):
        return True, ""
    return False, (
        f"On slide {check['slide']}, shape {check['shape']} should read “{check['equals']}” "
        f"but reads “{short(actual)}”."
    )


EXCEL_PROBE = """set fSep to ASCII character 31
set rSep to ASCII character 30
set p to system attribute "MAC_OFFICE_TARGET_FILE"
set AppleScript's text item delimiters to "/"
set bookName to last text item of p
set AppleScript's text item delimiters to ""
set out to {}
tell application "Microsoft Excel"
	open workbook workbook file name p read only true add to mru false
	set wb to workbook bookName
	repeat with spec in {@SPECS@}
		set r to range (item 2 of spec) of worksheet (item 1 of spec) of wb
		set v to ""
		set f to ""
		try
			set v to (value of r) as text
		end try
		try
			set f to (formula of r) as text
		end try
		set end of out to (item 1 of spec) & fSep & (item 2 of spec) & fSep & v & fSep & f
	end repeat
	close wb saving no
end tell
set AppleScript's text item delimiters to rSep
return out as text
"""


def _quote(text):
    return '"' + text.replace("\\", "\\\\").replace('"', '\\"') + '"'


def probe_excel(path, cells):
    specs = ", ".join("{" + _quote(sheet) + ", " + _quote(cell) + "}" for sheet, cell in cells)
    result = run_runner(
        ["--app", "excel", "--target-file", path, "--timeout", "60"], stdin_text=EXCEL_PROBE.replace("@SPECS@", specs)
    )
    if not result.get("success"):
        return None, result.get("error")
    values = {}
    for record in (result.get("output") or "").split(RECORD):
        fields = record.split(FIELD)
        if len(fields) == 4:
            values[(fields[0], fields[1])] = {"value": fields[2], "formula": fields[3]}
    return values, None


def _same_value(actual, expected):
    if isinstance(expected, (int, float)):
        try:
            return abs(float(actual) - float(expected)) < 1e-6
        except ValueError:
            return False
    return norm(actual) == norm(expected)


def check_excel_cells(check, deliverable):
    expectations = check["cells"]
    values, error = deliverable.excel()
    if values is None:
        return False, f"The workbook could not be read back: {error}"
    for spec in expectations:
        got = values.get((spec["sheet"], spec["cell"]))
        if got is None:
            return False, f"Cell {spec['sheet']}!{spec['cell']} could not be read."
        if "value" in spec and not _same_value(got["value"], spec["value"]):
            return False, f"Cell {spec['sheet']}!{spec['cell']} should be {spec['value']} but is “{short(got['value'])}”."
        if spec.get("formula") and not got["formula"].startswith("="):
            return False, (
                f"Cell {spec['sheet']}!{spec['cell']} holds a typed value instead of a formula; the user "
                "asked for a live calculation."
            )
    return True, ""


def check_answer_contains(check, final_text):
    squeezed = norm(final_text).replace(" ", "").replace(",", "")
    for alternatives in check["all"]:
        options = alternatives if isinstance(alternatives, list) else [alternatives]
        if not any(norm(option).replace(" ", "").replace(",", "") in squeezed for option in options):
            return False, f"The reply to the user does not state the requested fact ({check['describe']})."
    return True, ""


INSPECT_CHECKS = {
    "paragraph_count": check_paragraph_count,
    "paragraph_text": check_paragraph_text,
    "paragraph_alignment": check_paragraph_alignment,
    "cell_text": check_cell_text,
    "table_shape": check_table_shape,
    "shape_text": check_shape_text,
    "shape_count": check_shape_count,
    "picture_count": check_picture_count,
    "text_present": check_text_present,
    "text_absent": check_text_absent,
    "slide_count": check_slide_count,
    "slide_shape_text": check_slide_shape_text,
    "excel_cells": check_excel_cells,
}
RENDER_CHECKS = {"layout_preserved": check_layout_preserved, "page_count": check_page_count}


def hygiene_checks(task, docs_dir, original_hashes, staging_leftovers, transcript):
    results = []
    present = sorted(name for name in os.listdir(docs_dir) if not name.startswith("."))
    expected = set(task["expected_files"])
    missing = sorted(expected - set(present))
    extra = sorted(set(present) - expected)
    results.append(
        ("deliverables_exist", not missing, f"Expected deliverable(s) were not produced: {', '.join(missing)}.")
    )
    results.append(
        (
            "no_extra_files",
            not extra,
            f"The user's folder now holds files nobody asked for: {', '.join(extra)}.",
        )
    )
    changed = [
        name
        for name, digest in original_hashes.items()
        if name not in task.get("modified_in_place", [])
        and (not os.path.isfile(os.path.join(docs_dir, name)) or md5(os.path.join(docs_dir, name)) != digest)
    ]
    results.append(
        (
            "originals_untouched",
            not changed,
            f"The user's original file(s) were modified or removed although no in-place edit was requested: {', '.join(changed)}.",
        )
    )
    results.append(
        (
            "staging_clean",
            not staging_leftovers,
            "Temporary files were left behind in the Office staging folder.",
        )
    )
    results.append(
        (
            "focus_not_stolen",
            transcript.get("focus_restored_count", 0) == 0,
            "An Office app grabbed the user's keyboard focus during the task.",
        )
    )
    return [{"name": name, "passed": passed, "feedback": "" if passed else feedback} for name, passed, feedback in results]


def judge(task, docs_dir, original_hashes, staging_leftovers, transcript):
    """Score a finished task. ``transcript`` = {"final_text", "focus_restored_count"}."""
    results = hygiene_checks(task, docs_dir, original_hashes, staging_leftovers, transcript)
    with tempfile.TemporaryDirectory(prefix="mac-office-judge-") as work_dir:
        graded = os.path.join(docs_dir, task["graded_file"]) if task.get("graded_file") else None
        excel_cells = [
            (cell["sheet"], cell["cell"]) for check in task["checks"] if check["type"] == "excel_cells" for cell in check["cells"]
        ]
        deliverable = Deliverable(graded, work_dir, excel_cells)
        for check in task["checks"]:
            kind = check["type"]
            if check.get("equals") == "unchanged":
                check = dict(check, equals=before_inspect(task)["paragraph_count"])
            if kind == "answer_contains":
                passed, feedback = check_answer_contains(check, transcript.get("final_text") or "")
            elif kind == "pdf_text":
                passed, feedback = check_pdf_text(check, docs_dir, work_dir)
            elif not deliverable.exists:
                passed, feedback = False, f"The deliverable {task['graded_file']} does not exist, so it cannot be checked."
            elif kind in RENDER_CHECKS:
                passed, feedback = RENDER_CHECKS[kind](check, deliverable, task)
            else:
                inspect = deliverable.inspect()
                if not inspect.get("success"):
                    passed, feedback = False, f"The deliverable cannot be opened: {inspect.get('error')}"
                else:
                    passed, feedback = INSPECT_CHECKS[kind](check, deliverable)
            results.append(
                {"name": check.get("name", kind), "passed": passed, "feedback": "" if passed else feedback}
            )
    passed = sum(1 for result in results if result["passed"])
    return {
        "hard": 1 if passed == len(results) else 0,
        "soft": round(passed / len(results), 4),
        "checks": results,
    }

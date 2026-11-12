---
name: mac-office
description: "macOS-only skill to read, inspect, create, edit, format, and export Microsoft Word (.docx), Excel (.xlsx), and PowerPoint (.pptx) files natively using local Microsoft Office for Mac and headless OpenXML inspection. Trigger whenever the user mentions Word, Excel, PowerPoint, .docx, .xlsx, .pptx, spreadsheet, slide deck, or Office document formatting/PDF export on macOS."
---

# macOS Native Office Skill (`mac-office`)

Drive local Microsoft Word, Excel, and PowerPoint on macOS via AppleScript for 100% layout fidelity, combined with fast headless XML inspection (with automatic native-app fallback for corporate MIP/RMS-encrypted files).

> **Note:** All `scripts/` and `references/` paths below are relative to this skill's directory. **Platform:** macOS only (Apple Events + Microsoft Office for Mac 16.x).

## 1. Progressive Disclosure Routing Protocol (CRITICAL)

To prevent context pollution, **NEVER** read all reference files at once. Follow this strict 2-step routing flow:

| Step | Condition / Target | Action |
| :--- | :--- | :--- |
| **Step 1: Inspect** | Reading content OR discovering structure (paragraph indices, table count, sheet names, slide shape IDs) before editing | Run `python3 scripts/office_runner.py --inspect "/abs/path/to/file"`. **Stop here** if pure extraction answers the user's request. |
| **Step 2A: Word** | Creating, editing, styling, or exporting `.docx` | Read **ONLY** `references/word.md` |
| **Step 2B: Excel** | Creating, editing, formulas, styling, or exporting `.xlsx` | Read **ONLY** `references/excel.md` |
| **Step 2C: PowerPoint** | Creating, editing, styling, or exporting `.pptx` | Read **ONLY** `references/ppt.md` |

*(If a task bridges two apps, e.g. Excel data to Word report, inspect/extract the source first, then load ONLY the target app's reference.)*

### Routing on the inspect `container` field

| `container` | Meaning | Action |
| :--- | :--- | :--- |
| `openxml` | Plain OOXML zip | Headless inspection is authoritative; edits go through `--app` |
| `ole2` | Legacy OLE2 container (Word 97 binary; also what Office for Mac writes on MIP-managed machines) | Headless XML inspection is impossible — the JSON comes from the AppleScript fallback |
| `ole2_mip` | OLE2 holding `EncryptedPackage` / `DRMEncryptedTransform` (M365 Sensitivity Label) | Same as `ole2`; see `hint` |

For `container: ole2` or `ole2_mip`: **NEVER** attempt unzip/`zipfile`, LibreOffice, `textutil`, `olefile`, or any XML rewriting — none of them can read or write the content. Only the native app can, so do the work through `--app <word|excel|powerpoint>`. Such files stay OLE2 containers after saving, whatever extension the save uses. The fallback still reports structure (Word: `paragraphs` with `table`/`row`/`column`, `shapes`, `inline_shape_count`, `page_setup`); an individual item that cannot be read degrades to `null` instead of aborting the inspect.

Each collection reports `<collection>_count` plus `<collection>_truncated` (e.g. `paragraph_count` / `paragraphs_truncated`, `sheet_count` / `sheets_truncated`, `slide_count` / `slides_truncated`; nested lists carry their own flag such as `shapes_truncated`, `cells_truncated`, or a worksheet preview's `truncated`). `true` always means the returned list was capped by `--max-items` (default 200; raise it to see more), in both the headless and the fallback path. The fallback also reports items it could not read apart from truncation: `paragraphs_unreadable` / `shapes_unreadable` (Word), `cells_unreadable` per table, `shapes_unreadable` per slide. Paragraph text is whitespace-collapsed, so count blank runs with a render, not from inspect text.

## 2. Universal Execution Rules

- **Always execute AppleScript via `scripts/office_runner.py`** (never run raw `osascript -e` in shell):
  ```bash
  python3 scripts/office_runner.py --app [word|excel|powerpoint] \
      [--file <script.applescript>] \
      [--target-file "/abs/existing.ext"] \
      [--output-file "/abs/new_or_export.ext"] \
      [--asset "/abs/extra/input.png" ...] \
      [--in-place] [--timeout 60]
  ```
- **Sandbox-Safe File Paths**: Office for Mac runs inside the macOS App Sandbox and blocks direct saves to arbitrary folders with a modal `Grant File Access` popup. `office_runner.py` transparently stages files inside the shared Office Group Container (`~/Library/Group Containers/UBF8T346G9.Office/mac-office-staging`), so **never hardcode user paths inside AppleScript**. Read them from `system attribute` instead:
  - `MAC_OFFICE_TARGET_FILE` (`--target-file`), `MAC_OFFICE_OUTPUT_FILE` (`--output-file`).
  - `MAC_OFFICE_ASSET_1` … `MAC_OFFICE_ASSET_N` and `MAC_OFFICE_ASSET_COUNT` (`--asset`, repeatable) — for images, logos, or any other file the script must insert. The runner copies each asset into the run's stage dir and deletes it afterwards. **Never** hand-copy inputs into the Group Container yourself.
  - `MAC_OFFICE_STAGE_DIR` — this run's own staging subdirectory (unique per invocation, removed after the run whether it succeeded or failed). Put any file the script has to create there and delete it before finishing.
- **Lazy Write-Back (no eager `_revised`)**: nothing is copied up front; the target itself is staged. After a successful script, the runner writes `<stem>_revised<ext>` (or the original, with `--in-place`) **only if** the staged target was actually modified, also when an `--output-file` was given (edit, then export). The original is never touched before that write-back, so a failed run leaves it as it was.
- **Never pass `--in-place` on the user's file unless the user explicitly asked to overwrite it.** The deliverable is `<stem>_revised<ext>`; the original must stay byte-identical. Every run against the original rebuilds `_revised` from the original, discarding earlier edits, so for a follow-up edit pass the `_revised` file as `--target-file` **with** `--in-place` (that edits the copy, never the original). Report the `working_file` from the JSON and check it ends in `_revised` before reporting success.
- **Read the JSON result** (do not guess from the script text):
  - `target_file` — original absolute path of `--target-file`, or null; `target_modified` — whether the script changed the staged target.
  - `working_file` — the `_revised`/original path actually written back, or null; `output_file` — as passed, or null.
  - `output_exists` / `output_modified` — describe the deliverable (`--output-file` if given, else `working_file`); both false when nothing was written.
  - `focus_restored` — true when an Office app took the focus during the run and the runner gave it back to the user's app.
- **AppleScript Error Handling**: AppleScript does **NOT** have `try ... finally`. Use `try ... on error errMsg number errNum ... end try`.
- **Never call `activate` or `launch`**: `activate` steals window focus from the user. The runner starts a stopped app hidden (`open -g -j`), keeps an app it runs hidden hidden, never hides an app the user has visible, restores the focus if an Office app grabs it, and serializes calls per app. An app the runner launched is quit on a later call after 10 minutes unused (only while still hidden and empty); an app the user opened is never quit.
- **Timeouts**: Default 30 s; pass `--timeout 60` for PDF export or large documents (ceiling 180 s). The native inspect fallback always gets at least 60 s. On a timeout read `error_hint`: a hidden app may be waiting on a modal dialog (for example Word's unreadable-content repair prompt), which the user has to answer after bringing the app forward from the Dock.

## 3. `--render`: look at the result and compare

```bash
python3 scripts/office_runner.py --render "/abs/report.docx" [--dpi 110] [--output-dir DIR]
```

Accepts `.docx`/`.xlsx`/`.pptx` (staged copy opened read-only and exported to PDF by the owning app; the original is never touched and no `_revised` is created) or `.pdf` (rasterize only). Returns `{pages, pdf, cached, png:[...], lines:[{page,text,x,y,w,h}], focus_restored, error, error_hint, execution_time_sec}`; coordinates are PDF points with the origin at the top-left. The PDF is kept in the output dir for diffing as `<name>.pdf` (`<name>-2.pdf`, … when that name is taken; an existing PDF is never overwritten), and pages as `<pdf name>-page-N.png`. Rendering an unchanged file again reuses the exported PDF (`cached: true`) without opening the app.

**Verification loop — mandatory after any layout-affecting edit:**

1. Render the edited file, then render the original (or an earlier render of it).
2. Compare the page count and the line coordinates of content you did not touch. Identical page count and coordinates ⇒ layout preserved.
3. Any extra page, shifted coordinate, or overflowing picture means the edit broke the layout (usually a lost paragraph mark, a re-flowed table row, or an oversized image). Fix it; never report success on a broken render.

## 4. Tests

Run from this skill's directory. **Default verification is the unit suite only** (fake `osascript`/`open`, no Office app is started or shown):

```bash
python3 -m unittest discover -s tests -v
```

Run the integration suite only when a change touches an AppleScript string (runner preambles, inspect/export scripts) or when explicitly asked. It drives real Word/Excel/PowerPoint and fails if an Office app takes the focus during a test class:

```bash
MAC_OFFICE_INTEGRATION=1 python3 -m unittest discover -s tests -v
```

---
name: mac-office
description: "macOS-only skill to read, inspect, create, edit, format, and export Microsoft Word (.docx), Excel (.xlsx), and PowerPoint (.pptx) files natively using local Microsoft Office for Mac and headless OpenXML inspection. Trigger whenever the user mentions Word, Excel, PowerPoint, .docx, .xlsx, .pptx, spreadsheet, slide deck, or Office document formatting/PDF export on macOS."
---

# macOS Native Office Skill (`mac-office`)

Drive local Microsoft Word, Excel, and PowerPoint on macOS via AppleScript for 100% layout fidelity, combined with fast headless XML inspection (with automatic fallback for corporate MIP/RMS encrypted files).

> **Note:** All `scripts/` and `references/` paths below are relative to this skill's directory.
> **Platform:** macOS only (Apple Events + Microsoft Office for Mac 16.x).

## 1. Progressive Disclosure Routing Protocol (CRITICAL)

To prevent context pollution, **NEVER** read all reference files at once. Follow this strict 2-step routing flow:

| Step | Condition / Target | Action |
| :--- | :--- | :--- |
| **Step 1: Inspect** | Reading content OR discovering structure (paragraph indices, table count, sheet names, slide shape IDs) before editing | Run `python3 scripts/office_runner.py --inspect "/abs/path/to/file"`. **Stop here** if pure extraction answers the user's request. |
| **Step 2A: Word** | Creating, editing, styling, or exporting `.docx` | Read **ONLY** `references/word.md` |
| **Step 2B: Excel** | Creating, editing, formulas, styling, or exporting `.xlsx` | Read **ONLY** `references/excel.md` |
| **Step 2C: PowerPoint** | Creating, editing, styling, or exporting `.pptx` | Read **ONLY** `references/ppt.md` |

*(If a task bridges two apps, e.g. Excel data to Word report, inspect/extract the source first, then load ONLY the target app's reference.)*

## 2. Universal Execution Rules

- **Always execute AppleScript via `scripts/office_runner.py`** (never run raw `osascript -e` in shell):
  ```bash
  python3 scripts/office_runner.py --app [word|excel|powerpoint] \
      [--file <script.applescript>] \
      [--target-file "/abs/existing.ext"] \
      [--output-file "/abs/new_or_export.ext"] \
      [--in-place] [--timeout 60]
  ```
- **Sandbox-Safe File Paths (`MAC_OFFICE_TARGET_FILE` / `MAC_OFFICE_OUTPUT_FILE`)**:
  - Office for Mac runs inside macOS App Sandbox and blocks direct saves to arbitrary folders with a modal `Grant File Access` popup.
  - `office_runner.py` transparently stages files inside the shared Office Group Container (`~/Library/Group Containers/UBF8T346G9.Office/mac-office-staging`).
  - Always pass `--target-file "/abs/input"` (when editing an existing file) and/or `--output-file "/abs/output"` (when creating a new file or exporting PDF/CSV), and read `system attribute "MAC_OFFICE_TARGET_FILE"` / `system attribute "MAC_OFFICE_OUTPUT_FILE"` inside AppleScript. Never hardcode user paths inside AppleScript!
- **Non-Destructive Default**: `office_runner.py` copies `--target-file` to `<name>_revised.<ext>` unless `--in-place` is passed. With `--in-place`, a `.bak` backup is kept during execution and automatically restored on failure.
- **Verify the write**: Check `output_exists` and `output_modified` in the JSON response.
- **AppleScript Error Handling**: AppleScript does **NOT** have `try ... finally`. Use `try ... on error errMsg number errNum ... end try`.
- **Never call `activate`**: It steals window focus from the user. `launch` is automatically injected by the runner.
- **Timeouts**: Default 30 s; pass `--timeout 60` for PDF export or large documents (ceiling 180 s).

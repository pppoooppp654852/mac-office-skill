# mac-office-skill

⚠️ **macOS only.** This skill drives Microsoft Word, Excel, and PowerPoint for Mac through
AppleScript (`osascript`). It does not work on Linux or Windows, and it needs a licensed
Microsoft Office for Mac 16.x installation (Word, Excel, PowerPoint in `/Applications`).

A single-folder Agent Skill that lets an AI agent read, create, edit, format, and export
`.docx`, `.xlsx`, and `.pptx` files using the real Microsoft layout engine on your Mac,
plus a headless OpenXML inspector (with automatic AppleScript fallback for corporate
Sensitivity-Label / RMS-encrypted files).

## Why not XML manipulation?

Most cross-platform Office automation unpacks the OOXML package, edits XML, and repacks it
(or shells out to LibreOffice/Pandoc). That breaks on anything the XML does not model:
pagination, table column auto-width, East Asian font substitution, formula evaluation, and
corporate M365 Sensitivity-Label encryption (`EncryptedPackage`). Repacking also risks
corrupting relationship IDs and triggering "document is damaged" prompts on open.

This skill takes the opposite route:

- **Native engine first.** Creating, editing, formatting, and exporting runs through the
  user's own Word/Excel/PowerPoint, so the output is 100% faithful to Microsoft Office.
- **Headless reads + corporate MIP fallback.** Pure inspection (paragraph indices, table
  shapes, sheet names, slide shape coordinates) runs via Python's `zipfile` +
  `xml.etree.ElementTree` in milliseconds without launching a GUI app. The container is
  sniffed first (OLE2 magic **before** any zip parsing), so Word-97 binaries and
  MIP/RMS-wrapped saves — which may embed a perfectly readable theme zip — are routed to the
  native AppleScript inspector instead of being misreported as damaged. The JSON reports
  `container: openxml | ole2 | ole2_mip`, and for the OLE2 variants no external converter
  (LibreOffice, `textutil`, `olefile`) can read or write the content: only the owning app can.
- **Sandbox-safe staging.** Files are automatically staged inside the shared Microsoft
  Office Group Container (`~/Library/Group Containers/UBF8T346G9.Office/mac-office-staging`)
  so Office never blocks execution with macOS `Grant File Access` modal dialogs. Each run
  gets its own subdirectory (`MAC_OFFICE_STAGE_DIR`), removed afterwards.
- **Verification by rendering.** `--render` exports a staged copy to PDF through the owning
  app, rasterizes the pages, and returns per-line text coordinates, so an agent can prove a
  layout change did not reflow the document instead of guessing.
- **Lazy write-back.** Nothing is copied up front. `<name>_revised.<ext>` is created only if
  the script actually modified the document (or the original is written with `--in-place`),
  and the JSON result says exactly what landed on disk. The original is never touched before
  a successful run's write-back.
- **Out of the user's way.** A stopped Office app is started hidden (`open -g -j`) and kept
  hidden; an app the user has visible is never hidden. If an Office app grabs the focus the
  runner gives it back (`focus_restored`), calls to one app are serialized with a file lock,
  and an app the runner launched is quit on a later call once it sat unused, hidden and empty
  for 10 minutes. Apps the user opened are never quit.
- **Progressive disclosure.** `SKILL.md` stays a lightweight router. Only the reference for
  the target application is loaded, so the agent's context never mixes Word, Excel, and
  PowerPoint AppleScript rules.
- **Zero third-party dependencies.** macOS `osascript` + Python 3 standard library only.

## Layout

```text
mac-office-skill/
├── README.md
├── LICENSE
└── skills/
    └── mac-office/            # <- the only folder you install
        ├── SKILL.md           # router: inspect vs. edit vs. render, which reference to load
        ├── scripts/
        │   └── office_runner.py   # headless inspector, --render, AppleScript sandbox
        ├── references/
        │   ├── word.md        # verified Word AppleScript dictionary + pitfalls
        │   ├── excel.md       # verified Excel AppleScript dictionary
        │   └── ppt.md         # verified PowerPoint AppleScript dictionary
        └── tests/             # stdlib unittest regression suite
            ├── support.py     # fake `osascript` shim + shared fixtures
            ├── test_inspect_openxml.py
            ├── test_inspect_applescript_fallback.py
            ├── test_execute.py
            ├── test_render.py
            ├── test_app_session.py   # hidden launch, focus guard, lock, idle quit
            └── test_integration_{word,excel,ppt}.py   # MAC_OFFICE_INTEGRATION=1 only
```

## Install

The skill is harness-agnostic: anything that reads a `SKILL.md` directory works. Clone and
copy or symlink `skills/mac-office` into your agent harness's skill folder:

```bash
git clone https://github.com/pppoooppp654852/mac-office-skill.git
ln -s "$(pwd)/mac-office-skill/skills/mac-office" ~/.agent/skills/mac-office
# or copy instead of linking:
cp -r mac-office-skill/skills/mac-office ~/.agent/skills/mac-office
```

### Ask your agent to install it

```text
Clone https://github.com/pppoooppp654852/mac-office-skill into ~/dev, then symlink its
skills/mac-office folder into my agent skill directory and verify that
`python3 <skill>/scripts/office_runner.py --inspect <some .docx>` returns JSON.
```

## First run: macOS Automation permission

The first time your terminal/IDE sends an Apple Event to Word, Excel, PowerPoint, or System
Events (used to keep the apps hidden), macOS asks whether it may control the app. Allow it, or
enable it manually in **System Settings → Privacy & Security → Automation**. If a call fails
with `-1743`, the runner returns this guidance in its JSON `error_hint`. Detecting a modal
dialog after a timeout reads the app's windows and needs **Accessibility** permission; without
it the hint only says the dialog state could not be read.

## Usage

Inspection (`--max-items` caps every collection, default 200; each collection reports
`<collection>_count` and `<collection>_truncated`, e.g. `paragraph_count` / `paragraphs_truncated`,
where `_truncated` only ever means "cut by `--max-items`"; the native-app fallback adds
`<collection>_unreadable` counts for items it could not read):

```bash
python3 skills/mac-office/scripts/office_runner.py --inspect "/abs/report.docx"
python3 skills/mac-office/scripts/office_runner.py --inspect "/abs/book.xlsx" --max-items 20
python3 skills/mac-office/scripts/office_runner.py --inspect "/abs/deck.pptx"
```

AppleScript execution through the sandbox:

```bash
# Edit an existing file. Lazy write-back: <name>_revised.<ext> appears only if the
# script really modified it; pass --in-place to overwrite the original instead.
# The original is never touched unless the run succeeded.
python3 skills/mac-office/scripts/office_runner.py --app word \
    --file /tmp/edit.applescript \
    --target-file "/abs/report.docx" \
    --timeout 60

# Create a new file or export to PDF/CSV (the output is the deliverable; if the script
# also modified the target, <name>_revised.<ext> is written as well)
python3 skills/mac-office/scripts/office_runner.py --app word \
    --file /tmp/export.applescript \
    --target-file "/abs/report.docx" \
    --output-file "/abs/report.pdf" \
    --timeout 60

# Bring extra input files (images, logos, snippets) into the sandbox: --asset is repeatable
python3 skills/mac-office/scripts/office_runner.py --app word \
    --file /tmp/insert_logo.applescript \
    --target-file "/abs/report.docx" \
    --asset "/abs/logo.png" \
    --timeout 60
```

- Inside the AppleScript, read the staged paths from `system attribute`:
  - `MAC_OFFICE_TARGET_FILE` (`--target-file`) and `MAC_OFFICE_OUTPUT_FILE` (`--output-file`).
  - `MAC_OFFICE_ASSET_1` … `MAC_OFFICE_ASSET_N` plus `MAC_OFFICE_ASSET_COUNT` (`--asset`); the
    copies are deleted when the run ends.
  - `MAC_OFFICE_STAGE_DIR` — this run's private staging subdirectory, unique per invocation and
    removed afterwards either way.

Render for verification (staged copy opened read-only and exported to PDF by the owning app,
then rasterized; the original is never touched and no `_revised` is created; `.pdf` input is
rasterized as is). The PDF lands as `<name>.pdf` (`<name>-2.pdf`, … instead of overwriting an
existing file) next to `<pdf name>-page-N.png`; an unchanged file is rendered again from the
cached export (`cached: true`) without opening the app:

```bash
python3 skills/mac-office/scripts/office_runner.py --render "/abs/report.docx" --dpi 110
python3 skills/mac-office/scripts/office_runner.py --render "/abs/book.xlsx" --output-dir /tmp/render-xlsx
```

The JSON carries `pages`, the kept `pdf`, `png` paths, and `lines` (`{page,text,x,y,w,h}` in PDF
points, origin top-left) — diff those against a render of the original to prove an edit did not
reflow the document.

### What the JSON says about the result

| Field | Meaning |
| :--- | :--- |
| `target_file` | Original absolute path of `--target-file`, or `null` |
| `working_file` | The `_revised` (or original, `--in-place`) path written back, or `null` |
| `output_file` | `--output-file` as passed, or `null` |
| `target_modified` | Whether the script changed the staged target |
| `output_exists` / `output_modified` | Describe the deliverable (`--output-file` if given, else `working_file`); both `false` when nothing was written |
| `focus_restored` | An Office app took the focus during the run and the runner gave it back |

## Testing

The regression suite is stdlib `unittest` only; unit tests use fake `osascript` / `open` shims and
never launch Office. They are the default verification:

```bash
cd skills/mac-office && python3 -m unittest discover -s tests -v
```

Integration tests drive the real apps and are opt-in: run them when a change touches an
AppleScript string, or on request. A background sampler fails a test class if an Office app
takes the focus while it runs:

```bash
cd skills/mac-office && MAC_OFFICE_INTEGRATION=1 python3 -m unittest discover -s tests -v
```

## Requirements

- macOS (Apple Silicon or Intel)
- Microsoft Word, Excel, PowerPoint for Mac 16.x+ (licensed installs in `/Applications`)
- Python 3.9+ (standard library only: no Homebrew, no LibreOffice, no Pandoc)

## License

MIT — see [LICENSE](LICENSE).

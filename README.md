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
  `xml.etree.ElementTree` in milliseconds without launching a GUI app, and automatically
  falls back to native Office AppleScript inspection if the file is wrapped in an OLE2
  `EncryptedPackage` container.
- **Sandbox-safe staging.** Files are automatically staged inside the shared Microsoft
  Office Group Container (`~/Library/Group Containers/UBF8T346G9.Office/mac-office-staging`)
  so Office never blocks execution with macOS `Grant File Access` modal dialogs.
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
        ├── SKILL.md           # router: inspect vs. edit, which reference to load
        ├── scripts/
        │   └── office_runner.py   # headless inspector + AppleScript sandbox
        └── references/
            ├── word.md        # verified Word AppleScript dictionary (per .sdef)
            ├── excel.md       # verified Excel AppleScript dictionary
            └── ppt.md         # verified PowerPoint AppleScript dictionary
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

The first time your terminal/IDE sends an Apple Event to Word, Excel, or PowerPoint, macOS
asks whether it may control the app. Allow it, or enable it manually in **System Settings →
Privacy & Security → Automation**. If a call fails with `-1743`, the runner returns this
guidance in its JSON `error_hint`.

## Usage

Headless inspection:

```bash
python3 skills/mac-office/scripts/office_runner.py --inspect "/abs/report.docx"
python3 skills/mac-office/scripts/office_runner.py --inspect "/abs/book.xlsx" --max-items 20
python3 skills/mac-office/scripts/office_runner.py --inspect "/abs/deck.pptx"
```

AppleScript execution through the sandbox:

```bash
# Edit an existing file (writes <name>_revised.<ext> by default, or use --in-place)
python3 skills/mac-office/scripts/office_runner.py --app word \
    --file /tmp/edit.applescript \
    --target-file "/abs/report.docx" \
    --timeout 60

# Create a new file or export to PDF/CSV
python3 skills/mac-office/scripts/office_runner.py --app word \
    --file /tmp/export.applescript \
    --target-file "/abs/report.docx" \
    --output-file "/abs/report.pdf" \
    --timeout 60
```

- Inside the AppleScript, read `system attribute "MAC_OFFICE_TARGET_FILE"` (for `--target-file`)
  and `system attribute "MAC_OFFICE_OUTPUT_FILE"` (for `--output-file`).

## Requirements

- macOS (Apple Silicon or Intel)
- Microsoft Word, Excel, PowerPoint for Mac 16.x+ (licensed installs in `/Applications`)
- Python 3.9+ (standard library only: no Homebrew, no LibreOffice, no Pandoc)

## License

MIT — see [LICENSE](LICENSE).

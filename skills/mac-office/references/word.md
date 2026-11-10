# Microsoft Word (`references/word.md`)

Level 3 reference. Load only for `.docx` work. All syntax below is verified against
`/Applications/Microsoft Word.app/Contents/Resources/Word.sdef` (Office 16.x).

## 0. Ground Rules

- **Sandbox Staging**: Never hardcode user file paths in AppleScript (doing so triggers a blocking macOS `Grant File Access` dialog). Pass `--target-file "/abs/in.docx"` and/or `--output-file "/abs/out.docx"` to `office_runner.py`, and read `system attribute "MAC_OFFICE_TARGET_FILE"` / `system attribute "MAC_OFFICE_OUTPUT_FILE"`.
- **Variable Coercion Inside `tell` Blocks**: When coercing a string variable `p` to a file reference inside `tell application "Microsoft Word"`, always write `(p as POSIX file)` (never `POSIX file p`, which evaluates to `missing value` on variables inside `tell` blocks). For `save as ... file name`, pass `outPath` directly or `((outPath as POSIX file) as text)`.
- **Silence**: The runner injects `set display alerts to alerts none` (`WdAlertLevel`). Never write `set display alerts to false`.
- **No `try ... finally`**: AppleScript only supports `try ... on error errMsg number errNum ... end try`.
- **Never `activate`**: Steals focus. `launch` is already in the runner preamble.
- **Always target `active document`**: After `open (targetPath as POSIX file)` or `make new document`, address `active document` (never pass a `theDoc` variable to `save as`, and do not nest `at end of active document` inside `tell active document`).

## 1. Open, Create, Save

Editing an existing document (`--target-file "/abs/report.docx"`):

```applescript
tell application "Microsoft Word"
	set targetPath to system attribute "MAC_OFFICE_TARGET_FILE"
	open (targetPath as POSIX file)
	set content of text object of paragraph 1 of active document to "報告標題"
	save active document
end tell
```

Creating a new document from scratch (`--output-file "/abs/out.docx"`):

```applescript
tell application "Microsoft Word"
	set outPath to system attribute "MAC_OFFICE_OUTPUT_FILE"
	make new document
	set content of text object of paragraph 1 of active document to "報告標題"
	save as active document file name outPath file format format document
	close active document saving no
end tell
```

- Formats (`WdSaveFormat`): `format document` (`.docx`), `format PDF` (`.pdf`).

## 2. Text, Font, and Appending Paragraphs

Text lives on `text object of paragraph N of active document`; character formatting lives on its `font object`. Colors are `{R, G, B}` integers `0..255`.

```applescript
tell application "Microsoft Word"
	tell text object of paragraph 1 of active document
		set content to "主標題文字"
		tell font object
			set name to "PingFang TC"
			set font size to 16.0
			set bold to true
			set italic to false
			set underline to underline single
			set color to {0, 51, 102}
		end tell
	end tell
end tell
```

Appending a paragraph (always call `make new paragraph at end of active document` at the `tell application "Microsoft Word"` level, then address `last paragraph of active document`):

```applescript
tell application "Microsoft Word"
	make new paragraph at end of active document
	set content of text object of last paragraph of active document to "第二段內文"
end tell
```

## 3. Paragraph Formatting

`paragraph format` is on the paragraph object:

```applescript
tell application "Microsoft Word"
	tell paragraph format of paragraph 1 of active document
		set alignment to align paragraph center
		set space before to 0.0
		set space after to 12.0
		set line spacing to 18.0
		set first line indent to 24.0
	end tell
end tell
```

- Alignment (`WdParagraphAlignment`): `align paragraph left`, `align paragraph center`, `align paragraph right`, `align paragraph justify`.
- Built-in styles (`WdBuiltinStyle`): `set style of paragraph 2 of active document to style heading1` (`style heading2`, `style normal`).

## 4. Tables

To append a table without overwriting existing paragraphs, first append a new paragraph at the end of `active document`, then create the table `at text object of last paragraph of active document`:

```applescript
tell application "Microsoft Word"
	make new paragraph at end of active document
	set theTable to make new table at text object of last paragraph of active document with properties {number of rows:3, number of columns:2}
	set content of text object of cell 1 of row 1 of theTable to "欄位"
	set content of text object of cell 2 of row 1 of theTable to "數值"
	set width of column 1 of theTable to 120.0
	set background pattern color index of shading of cell 1 of row 1 of theTable to gray25
	merge cell cell 1 of row 3 of theTable with cell 2 of row 3 of theTable
end tell
```

- Cell shading: `background pattern color index` (`WdColorIndex`, e.g. `gray25`, `yellow`) or `background pattern color` (`{R, G, B}`).
- Borders: `get border cell 1 of row 1 of theTable which border border top` then set `line style` / `line width` / `color`.

## 5. Find and Replace

The property on `text range` is **`find object`** (not `find`):

```applescript
tell application "Microsoft Word"
	execute find (find object of text object of active document) find text "舊字串" replace with "新字串" replace replace all
end tell
```

- `replace` (`WdReplace`): `replace one`, `replace all`, `replace none`.

## 6. PDF Export

Pass `--target-file "/abs/in.docx" --output-file "/abs/out.pdf" --timeout 60`:

```applescript
tell application "Microsoft Word"
	set targetPath to system attribute "MAC_OFFICE_TARGET_FILE"
	set outPath to system attribute "MAC_OFFICE_OUTPUT_FILE"
	open (targetPath as POSIX file)
	save as active document file name outPath file format format PDF
	close active document saving no
end tell
```

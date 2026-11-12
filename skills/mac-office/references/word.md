# Microsoft Word (`references/word.md`)

Level 3 reference. Load only for `.docx` work. All syntax below is verified against
`/Applications/Microsoft Word.app/Contents/Resources/Word.sdef` (Office 16.x).

## 0. Ground Rules

- **Sandbox Staging**: Never hardcode user file paths in AppleScript (doing so triggers a blocking macOS `Grant File Access` dialog). Pass `--target-file "/abs/in.docx"` and/or `--output-file "/abs/out.docx"` to `office_runner.py`, and read `system attribute "MAC_OFFICE_TARGET_FILE"` / `system attribute "MAC_OFFICE_OUTPUT_FILE"`.
- **Extra inputs**: pass images/logos/snippets with `--asset "/abs/logo.png"` (repeatable) and read `system attribute "MAC_OFFICE_ASSET_1"` … `"MAC_OFFICE_ASSET_n"` (count in `"MAC_OFFICE_ASSET_COUNT"`). Never copy files into the Group Container by hand — the runner stages and cleans them up.
- **Per-run stage dir**: `system attribute "MAC_OFFICE_STAGE_DIR"` is this invocation's own staging subdirectory (unique per run, deleted afterwards, success or failure). Write any scratch file the script needs there and remove it before finishing; do not reuse another run's directory.
- **Variable Coercion Inside `tell` Blocks**: When coercing a string variable `p` to a file reference inside `tell application "Microsoft Word"`, always write `(p as POSIX file)` (never `POSIX file p`, which evaluates to `missing value` on variables inside `tell` blocks). For `save as ... file name`, pass `outPath` directly or `((outPath as POSIX file) as text)`.
- **Silence**: The runner injects `set display alerts to alerts none` (`WdAlertLevel`). Never write `set display alerts to false`.
- **No `try ... finally`**: AppleScript only supports `try ... on error errMsg number errNum ... end try`.
- **Never `activate` or `launch`**: `activate` steals focus. The runner starts Word hidden (and keeps it hidden) before your script runs.
- **Always target `active document`**: After `open (targetPath as POSIX file)` or `make new document`, address `active document` (never pass a `theDoc` variable to `save as`, and do not nest `at end of active document` inside `tell active document`).
- **Keep the paragraph mark**: a content write that replaces a paragraph's text consumes its paragraph mark. For any paragraph that has a following paragraph, end the string with `& return`, or it merges with the next one (see §8.1). Do **not** append `& return` to the document's last paragraph or to a table cell's text object: their final mark is never consumed, so `& return` adds a stray empty paragraph.

## 1. Open, Create, Save

Editing an existing document (`--target-file "/abs/report.docx"`):

```applescript
tell application "Microsoft Word"
	set targetPath to system attribute "MAC_OFFICE_TARGET_FILE"
	open (targetPath as POSIX file)
	set content of text object of paragraph 1 of active document to "報告標題" & return
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
- Write-back is lazy: the file at `MAC_OFFICE_TARGET_FILE` is the staged copy. If you edited it
  and passed no `--output-file`, the runner writes `<stem>_revised<ext>`; report that path from
  the JSON `working_file`, not from an assumed filename.

## 2. Text, Font, and Appending Paragraphs

Text lives on `text object of paragraph N of active document`; character formatting lives on its `font object`. Colors are `{R, G, B}` integers `0..255`.

```applescript
tell application "Microsoft Word"
	tell text object of paragraph 1 of active document
		set content to "主標題文字" & return
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

Character formatting set *after* the content write survives; **paragraph** formatting does not
(content replacement resets it — see §3 and §8.2).

Appending a paragraph (always call `make new paragraph at end of active document` at the `tell application "Microsoft Word"` level, then address `last paragraph of active document`):

```applescript
tell application "Microsoft Word"
	make new paragraph at end of active document
	set content of text object of last paragraph of active document to "第二段內文"
end tell
```

## 3. Paragraph Formatting

`paragraph format` is reachable from the paragraph object, but after a content write the reliable
form addresses it through the text object of a document reference:

```applescript
tell application "Microsoft Word"
	set d to active document
	set alignment of paragraph format of text object of paragraph 1 of d to align paragraph center
	set space before of paragraph format of text object of paragraph 1 of d to 0.0
	set space after of paragraph format of text object of paragraph 1 of d to 12.0
	set line spacing of paragraph format of text object of paragraph 1 of d to 18.0
	set first line indent of paragraph format of text object of paragraph 1 of d to 24.0
	set paragraph format left indent of paragraph format of text object of paragraph 1 of d to 0.0
	set paragraph format right indent of paragraph format of text object of paragraph 1 of d to 0.0
end tell
```

- **Order matters**: write text first, then the paragraph format. Any later content write resets it.
- Alignment (`WdParagraphAlignment`): `align paragraph left`, `align paragraph center`, `align paragraph right`, `align paragraph justify`.
- Indents (`paragraph format` properties, sdef `paragraph format left indent` / `paragraph format right indent`) are in points, not characters.
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

- **Read cells via `content of text object of cell c of row r of theTable`.** `content of cell` and
  `paragraph N of cell` frequently return `missing value` (§8.3).
- Table-cell paragraphs are part of the document's paragraph numbering: indices read headlessly
  will not line up with `paragraph N of active document` after a table (§8.3).
- Cell shading: `background pattern color index` (`WdColorIndex`, e.g. `gray25`, `yellow`) or `background pattern color` (`{R, G, B}`).
- Borders: `get border cell 1 of row 1 of theTable which border border top` then set `line style` / `line width` / `color`.

## 5. Pictures

Pictures must come from a staged asset: `--asset "/abs/logo.png"` → read
`system attribute "MAC_OFFICE_ASSET_1"`. `file name` is read-only in the dictionary
(`inline picture`, code `WilP`; `picture`, code `cD04`), so it must be supplied in `with properties`
at creation time.

**Inline picture inside a table cell** (the cell grows to fit; use when the row may grow):

```applescript
tell application "Microsoft Word"
	set p to system attribute "MAC_OFFICE_ASSET_1"
	set theTable to table 1 of active document
	set thePic to make new inline picture at beginning of text object of cell 2 of row 1 of theTable ¬
		with properties {file name:p}
	set lock aspect ratio of thePic to true
	set width of thePic to 120.0
end tell
```

Always insert `at beginning of text object of cell …` (empty or not; the picture goes before any
text). Never use `at end of text object of cell` (lands in the next cell, or adds a column after a
row's last cell), `at text object of cell` (splits the cell) or `at start of …` (errors) — §8.4.

**Floating picture** (does not disturb row height; wrap settings decide overlap):

```applescript
tell application "Microsoft Word"
	set p to system attribute "MAC_OFFICE_ASSET_1"
	set pic to make new picture at text object of paragraph 1 of active document with properties {file name:p}
	set wrap type of wrap format of pic to wrap front
	set relative horizontal position of pic to relative horizontal position page
	set relative vertical position of pic to relative vertical position page
	set lock aspect ratio of pic to false
	set left position of pic to 380.0
	set top of pic to 120.0
	set width of pic to 200.0
	set height of pic to 150.0
end tell
```

- Coordinates are points from the top-left of the page reference you chose; `wrap front` overlays text, `wrap behind` sits underneath, `wrap square`/`wrap top bottom` re-flow text.
- `lock aspect ratio` must be false before setting `width` and `height` independently.
- Capture the returned object from `make new` and use it. `shape k` indices count pre-existing shapes, so a positional index is unreliable after the first insertion (§8.6).
- Verify the result with `--render` (page count + line coordinates), not by eye.

## 6. Find and Replace

The property on `text range` is **`find object`** (not `find`):

```applescript
tell application "Microsoft Word"
	execute find (find object of text object of active document) find text "舊字串" replace with "新字串" replace replace all
end tell
```

- `replace` (`WdReplace`): `replace one`, `replace all`, `replace none`.
- Symbol-font glyphs (`w:sym`, e.g. checkbox squares) are invisible to both `content` and find —
  they cannot be replaced; rewrite the paragraph instead (§8.9).

## 7. PDF Export

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

For a throwaway visual check (PDF kept for diffing, original untouched, no `_revised`) prefer
`python3 scripts/office_runner.py --render "/abs/in.docx"` — see `SKILL.md` §3.

<!-- PROTECTED_START: verified against Word.sdef and real sessions; edit by hand only -->
## 8. Verified pitfalls

All ten come from real sessions against Office 16.x (and are why the recipes above look the way
they do). Each is a hard failure, not a preference.

1. **Content write eats the paragraph mark.** `set content of text object of paragraph N of d to s`
   replaces the mark too, merging N with N+1 and shifting every later index. Append `& return` when
   N has a following paragraph; omit it for the document's last paragraph and for table cells
   (there it adds an empty paragraph — verified: `"A" & return` in a cell yields `A¶¶`).
2. **Content replacement resets paragraph format.** With `& return` the rewritten paragraph takes the
   format of the following paragraph's mark (a left-aligned line next to a centred one comes back
   centred). Set alignment/indents/spacing *after* writing, then confirm with `--render`.
   Working form: `set alignment of paragraph format of text object of paragraph N of d to align paragraph left`.
3. **Paragraph indices include table-cell paragraphs.** Read cell text as
   `content of text object of cell c of row r of t`; `content of cell` and `paragraph N of cell`
   often yield `missing value`.
4. **Inline picture in a cell:** `make new inline picture at beginning of text object of cell c of row r of t with properties {file name:p}`
   puts the picture in that cell, before any text it holds, whether the cell is empty or not
   (verified Word 16.114). `at end of text object of cell …` writes past the cell's end mark: the
   picture lands in the next cell, and after a row's last cell Word adds a column (one more
   paragraph, table one column wider). `at text object of cell` splits the cell (adds a column);
   `at start of …` raises an error.
5. **Floating picture:** `set pic to make new picture at text object of paragraph 1 of d with properties {file name:p}`,
   then `set wrap type of wrap format of pic to wrap front`, `relative horizontal/vertical position page`,
   and `left position` / `top` / `width` / `height` (set `lock aspect ratio` false first). Use it when a
   table row must not grow.
6. **Use the reference returned by `make new`.** `shape k` indices include pre-existing shapes, so
   counting shapes after an insertion addresses the wrong object. The returned reference works for
   `make new picture` and for `make new inline picture at beginning of …`; the one returned by
   `make new inline picture at end of text object of cell …` is dead (`-1728`/`-10006` on the next
   property access), one more reason never to insert there (verified Word 16.114).
7. **Text boxes hold a text range, not a text object:** read with
   `content of text range of text frame of shape k`. Writing there replaces any drawn content
   inside the box as well.
8. **Reserved words and term quirks:** `log` and `is` cannot be used as AppleScript variable names;
   indents are `paragraph format left indent` / `paragraph format right indent`; a
   `character id N` reference inside a Word `tell` block needs `my (character id N)`;
   `ASCII number` returns 63 for any non-ASCII character; coerce enums in two steps
   (`set v to alignment of pf` then `v as text`).
9. **Symbol-font glyphs are unreachable through AppleScript.** `w:sym` characters (e.g. ☐/☑ in
   documents built from forms) are absent from `content` and cannot be found or replaced — rebuild
   the paragraph with real text instead.
10. **MIP-labelled sources stay OLE2.** On sensitivity-label-managed machines, `save as … format document default`
    still produces an OLE2 container; detect it via inspect `container: ole2_mip` and never unzip it.
<!-- PROTECTED_END -->

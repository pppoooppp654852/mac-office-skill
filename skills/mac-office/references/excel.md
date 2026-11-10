# Microsoft Excel (`references/excel.md`)

Level 3 reference. Load only for `.xlsx` work. All syntax below is verified against
`/Applications/Microsoft Excel.app/Contents/Resources/Excel.sdef` (Office 16.x).

## 0. Ground Rules

- **Sandbox Staging**: Never hardcode user file paths in AppleScript (doing so triggers a blocking macOS `Grant File Access` dialog). Pass `--target-file "/abs/in.xlsx"` and/or `--output-file "/abs/out.xlsx"` to `office_runner.py`, and read `system attribute "MAC_OFFICE_TARGET_FILE"` / `system attribute "MAC_OFFICE_OUTPUT_FILE"`.
- **Variable Coercion**: When opening a path stored in a variable `targetPath` inside `tell application "Microsoft Excel"`, write `open (targetPath as POSIX file)` (or `open targetPath`).
- **Silence**: The runner injects `set display alerts to false` (boolean).
- **Workbook reference after `save workbook as`**: `save workbook as` renames the workbook in memory, invalidating any earlier `set wb to make new workbook` variable (`-1728`). Always address `active workbook` after `save workbook as`.
- **IPC ceiling**: One AppleEvent carrying more than ~500 rows of a 2D list stalls `osascript`. Chunk large data (≤ 500 rows per `set value`), never write cell by cell in a loop.
- **No `try ... finally`**: Use `try ... on error errMsg number errNum ... end try`.

## 1. Open, Create, Save

Editing an existing workbook (`--target-file "/abs/book.xlsx"`):

```applescript
tell application "Microsoft Excel"
	set targetPath to system attribute "MAC_OFFICE_TARGET_FILE"
	open (targetPath as POSIX file)
	tell active sheet of active workbook
		set value of range "A1:B1" to {{"標題A", "標題B"}}
	end tell
	save active workbook
end tell
```

Creating a new workbook from scratch (`--output-file "/abs/out.xlsx"`):

```applescript
tell application "Microsoft Excel"
	set outPath to system attribute "MAC_OFFICE_OUTPUT_FILE"
	set theWorkbook to make new workbook
	set name of active sheet of theWorkbook to "報表"
	save workbook as theWorkbook filename outPath file format Excel XML file format
	close active workbook saving no
end tell
```

- Formats (`XlFileFormat`): `Excel XML file format` (`.xlsx`), `PDF file format` (`.pdf`), `CSV file format` (`.csv`).

## 2. Batch Read and Write (2D Arrays, ≤ 500 Rows per Call)

One `set value` call transfers an entire 2D block (`{{row1_col1, row1_col2}, {row2_col1, row2_col2}}`):

```applescript
tell application "Microsoft Excel"
	tell active sheet of active workbook
		set value of range "A1:C2" to {{"欄位1", "欄位2", "欄位3"}, {"值A", 100, 200}}
	end tell
end tell
```

Reading a block back:

```applescript
tell application "Microsoft Excel"
	set block to value of range "A1:C2" of active sheet of active workbook
end tell
```

## 3. Formulas and Recalculation

```applescript
tell application "Microsoft Excel"
	tell active sheet of active workbook
		set formula of range "D2" to "=SUM(B2:C2)"
		calculate
	end tell
end tell
```

## 4. Formatting

Font properties live on `font object of range ...` (note: property name is **`font size`**, NOT `size`!). Fill color lives on `interior object of range ...`.

```applescript
tell application "Microsoft Excel"
	tell active sheet of active workbook
		tell font object of range "A1:C1"
			set bold to true
			set name to "PingFang TC"
			set font size to 12.0
			set color to {255, 255, 255}
		end tell
		set color of interior object of range "A1:C1" to {0, 51, 102}
		set number format of range "B2:D2" to "#,##0"
		set horizontal alignment of range "A1:C1" to horizontal align center
		border around range "A1:C1" line style continuous weight border weight medium color {0, 0, 0}
		autofit (entire column of range "A1:D2")
	end tell
end tell
```

- Alignment (`XlHAlign`): `horizontal align center`, `horizontal align left`, `horizontal align right`.

## 5. Worksheet Lifecycle

```applescript
tell application "Microsoft Excel"
	tell active workbook
		make new worksheet at end
		set name of active sheet to "明細"
		move sheet "明細" to before sheet 1
		delete worksheet "暫存"
	end tell
end tell
```

## 6. PDF / CSV Export

Pass `--target-file "/abs/in.xlsx" --output-file "/abs/out.pdf" --timeout 60`:

```applescript
tell application "Microsoft Excel"
	set targetPath to system attribute "MAC_OFFICE_TARGET_FILE"
	set outPath to system attribute "MAC_OFFICE_OUTPUT_FILE"
	open (targetPath as POSIX file)
	save workbook as active workbook filename outPath file format PDF file format
	close active workbook saving no
end tell
```

For CSV export, replace `PDF file format` with `CSV file format`.

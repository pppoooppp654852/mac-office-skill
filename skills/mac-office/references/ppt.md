# Microsoft PowerPoint (`references/ppt.md`)

Level 3 reference. Load only for `.pptx` work. All syntax below is verified against
`/Applications/Microsoft PowerPoint.app/Contents/Resources/PowerPoint.sdef` (Office 16.x).

## 0. Ground Rules

- **Sandbox Staging**: Never hardcode user file paths in AppleScript. Pass `--target-file "/abs/in.pptx"` and/or `--output-file "/abs/out.pptx"` to `office_runner.py`, and read `system attribute "MAC_OFFICE_TARGET_FILE"` / `system attribute "MAC_OFFICE_OUTPUT_FILE"`.
- **Variable Coercion**: Inside `tell application "Microsoft PowerPoint"`, coerce path variables with `(targetPath as POSIX file)` / `(outPath as POSIX file)` (never `POSIX file targetPath`, which fails on variables inside `tell` blocks).
- **No alert suppression**: `PowerPoint.sdef` has no `display alerts` property. Never inject it.
- **Units are points**: `72 pt = 1 inch`. Coordinates are `left position`, `top`, `width`, `height` (`left position` is the exact property name — `left` does not exist).
- **`shape type` is read-only**: Create rectangles with `auto shape type:autoshape rectangle` (`MsoAutoShapeType`). Never use VBA's `msoShapeRectangle`.
- **Text & Font**: Content lives on `content of text range of text frame of <shape>`; font properties live on `font of <textRange>` (bind to a variable `set theFont to font of theTextRange`, never nest `tell font` inside `tell text range`).
- **New presentations start empty**: `make new presentation` has **zero** slides; call `make new slide at end of thePres` first.
- **Reference after `save`**: Saving renames the presentation in memory; always use `active presentation` after `save`.

## 1. Open, Create, Save

Editing an existing deck (`--target-file "/abs/deck.pptx"`):

```applescript
tell application "Microsoft PowerPoint"
	set targetPath to system attribute "MAC_OFFICE_TARGET_FILE"
	open (targetPath as POSIX file)
	repeat 25 times
		if (count of presentations) > 0 then exit repeat
		delay 0.1
	end repeat
	set thePres to active presentation
	set theSlide to slide 1 of thePres
	save active presentation
end tell
```

Creating a new deck from scratch (`--output-file "/abs/out.pptx"`):

```applescript
tell application "Microsoft PowerPoint"
	set outPath to system attribute "MAC_OFFICE_OUTPUT_FILE"
	set thePres to make new presentation
	set theSlide to make new slide at end of thePres
	set layout of theSlide to slide layout title only
	save thePres in (outPath as POSIX file) as save as Open XML presentation
	close active presentation saving no
end tell
```

- Formats (`EPPSaveAsFileType`): `save as Open XML presentation` (`.pptx`), `save as PDF` (`.pdf`), `save as PNG`, `save as JPG`.
- **Corporate M365 IRM Note**: On corporate M365 Macs where saving as `.pptx` automatically applies a Sensitivity Label (`is protected: true`), if you need both `.pdf` and `.pptx` from a newly created presentation, run `save thePres in (pdfPath as POSIX file) as save as PDF` **before** saving as `Open XML presentation`.

## 2. Slides

```applescript
tell application "Microsoft PowerPoint"
	set newSlide to make new slide at end of active presentation
	set layout of newSlide to slide layout title only
	move slide 2 of active presentation to before slide 1 of active presentation
end tell
```

- Layouts (`EPPSlideLayout`): `slide layout title slide`, `slide layout title only`, `slide layout blank`, `slide layout text slide`, `slide layout two column text`.

## 3. Shapes, Text Boxes, and Fill/Line Formatting

```applescript
tell application "Microsoft PowerPoint"
	set theSlide to slide 1 of active presentation
	set newBox to make new shape at end of theSlide with properties {auto shape type:autoshape rectangle, top:100, left position:50, width:400, height:60}
	set theFill to fill format of newBox
	set fore color of theFill to {0, 51, 102}
	set theLine to line format of newBox
	set fore color of theLine to {0, 51, 102}
	set theTextRange to text range of text frame of newBox
	set content of theTextRange to "立德盃賽程規劃"
	set theFont to font of theTextRange
	set font name of theFont to "PingFang TC"
	set font size of theFont to 24.0
	set bold of theFont to true
	set font color of theFont to {255, 255, 255}
end tell
```

- **Note on `line format`**: `fill format` has a `visible` property (`boolean`), whereas `line format` does **NOT** have `visible` in `PowerPoint.sdef` (use `fore color`, `line weight`, or `transparency` on `line format`).

Move or resize an existing shape:

```applescript
tell application "Microsoft PowerPoint"
	tell shape 1 of slide 1 of active presentation
		set left position to 60.0
		set top to 200.0
		set width to 360.0
		set height to 80.0
	end tell
end tell
```

## 4. Pictures

```applescript
tell application "Microsoft PowerPoint"
	set theSlide to slide 1 of active presentation
	set newPicture to make new picture at end of theSlide with properties {file name:"/abs/photo.png"}
	set left position of newPicture to 40.0
	set top of newPicture to 320.0
	set width of newPicture to 240.0
end tell
```

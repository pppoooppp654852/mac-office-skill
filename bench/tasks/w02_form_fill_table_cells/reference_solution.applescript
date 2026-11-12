tell application "Microsoft Word"
	open ((system attribute "MAC_OFFICE_TARGET_FILE") as POSIX file)
	set d to active document
	set tb to table 1 of d
	set content of text object of cell 2 of row 1 of tb to "陳志明"
	set content of text object of cell 2 of row 2 of tb to "研發部"
	set content of text object of cell 2 of row 3 of tb to "特休"
	set content of text object of cell 2 of row 4 of tb to "2026/12/01–2026/12/03"
	save active document
end tell

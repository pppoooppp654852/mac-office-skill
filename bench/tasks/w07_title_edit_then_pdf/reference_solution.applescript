tell application "Microsoft Word"
	open ((system attribute "MAC_OFFICE_TARGET_FILE") as POSIX file)
	set d to active document
	execute find (find object of text object of paragraph 1 of d) find text "（草稿）" replace with "" replace replace one
	save active document
	save as active document file name (system attribute "MAC_OFFICE_OUTPUT_FILE") file format format PDF
end tell

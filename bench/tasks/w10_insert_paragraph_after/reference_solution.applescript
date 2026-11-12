tell application "Microsoft Word"
	open ((system attribute "MAC_OFFICE_TARGET_FILE") as POSIX file)
	set d to active document
	insert text "地點：第二會議室" & return at beginning of text object of paragraph 4 of d
	save active document
end tell

tell application "Microsoft Word"
	open ((system attribute "MAC_OFFICE_TARGET_FILE") as POSIX file)
	set d to active document
	execute find (find object of text object of d) find text "大同貿易" replace with "大同國際貿易" replace replace all
	save active document
end tell

tell application "Microsoft Word"
	open ((system attribute "MAC_OFFICE_TARGET_FILE") as POSIX file)
	set d to active document
	set content of text object of paragraph 1 of d to "下午四點改在第三會議室開會"
	save active document
end tell

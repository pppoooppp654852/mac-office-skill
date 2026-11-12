tell application "Microsoft Word"
	open ((system attribute "MAC_OFFICE_TARGET_FILE") as POSIX file)
	set d to active document
	set content of text range of text frame of shape 2 of d to "狀態：已核准"
	save as active document file name (system attribute "MAC_OFFICE_OUTPUT_FILE") file format format document
	close active document saving no
end tell

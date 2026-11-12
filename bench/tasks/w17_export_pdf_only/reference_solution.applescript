tell application "Microsoft Word"
	open ((system attribute "MAC_OFFICE_TARGET_FILE") as POSIX file)
	set d to active document
	save as active document file name (system attribute "MAC_OFFICE_OUTPUT_FILE") file format format PDF
	close active document saving no
end tell

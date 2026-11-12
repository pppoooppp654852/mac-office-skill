tell application "Microsoft Word"
	set outPath to system attribute "MAC_OFFICE_OUTPUT_FILE"
	set d to make new document
	set content of text object of d to "下午三點開會"
	save as active document file name outPath file format format document
	close active document saving no
end tell

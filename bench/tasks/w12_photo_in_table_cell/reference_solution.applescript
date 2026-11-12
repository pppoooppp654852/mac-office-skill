tell application "Microsoft Word"
	open ((system attribute "MAC_OFFICE_TARGET_FILE") as POSIX file)
	set d to active document
	set p to system attribute "MAC_OFFICE_ASSET_1"
	make new inline picture at beginning of text object of cell 3 of row 2 of table 1 of d with properties {file name:p}
	set pic to inline shape 1 of d
	set lock aspect ratio of pic to true
	set width of pic to 71.0
	save active document
end tell

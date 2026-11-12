tell application "Microsoft Word"
	open ((system attribute "MAC_OFFICE_TARGET_FILE") as POSIX file)
	set d to active document
	set p to system attribute "MAC_OFFICE_ASSET_1"
	set pic to make new picture at text object of cell 1 of row 1 of table 1 of d with properties {file name:p}
	set wrap type of wrap format of pic to wrap front
	set lock aspect ratio of pic to true
	set height of pic to 14.0
	save active document
end tell

tell application "Microsoft Excel"
	open ((system attribute "MAC_OFFICE_TARGET_FILE") as POSIX file)
	tell worksheet "報帳" of active workbook
		set value of range "D3" to "需主管核准"
		set value of range "D5" to "需主管核准"
	end tell
	save active workbook
end tell

tell application "Microsoft Excel"
	open ((system attribute "MAC_OFFICE_TARGET_FILE") as POSIX file)
	tell worksheet "明細" of active workbook
		set formula of range "D2:D6" to "=B2*C2"
		set formula of range "D7" to "=SUM(D2:D6)"
	end tell
	save active workbook
end tell

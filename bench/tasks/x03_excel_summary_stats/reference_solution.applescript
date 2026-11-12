tell application "Microsoft Excel"
	open ((system attribute "MAC_OFFICE_TARGET_FILE") as POSIX file)
	tell worksheet "摘要" of active workbook
		set formula of range "B1" to "=AVERAGE(成績!B2:B7)"
		set formula of range "B2" to "=MAX(成績!B2:B7)"
		set formula of range "B3" to "=COUNTIF(成績!B2:B7,\">=70\")"
	end tell
	save active workbook
end tell

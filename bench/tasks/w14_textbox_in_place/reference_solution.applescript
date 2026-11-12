tell application "Microsoft Word"
	open ((system attribute "MAC_OFFICE_TARGET_FILE") as POSIX file)
	set d to active document
	set content of text range of text frame of shape 1 of d to "聯絡人：周小姐 分機 5678"
	save active document
end tell

tell application "Microsoft Word"
	open ((system attribute "MAC_OFFICE_TARGET_FILE") as POSIX file)
	set d to active document
	set content of text object of paragraph 2 of d to "姓名：吳佩珊" & return
	set content of text object of paragraph 3 of d to "單位：資訊處" & return
	set content of text object of paragraph 4 of d to "備註：素食" & return
	set alignment of paragraph format of text object of paragraph 4 of d to align paragraph left
	save active document
end tell

tell application "Microsoft Word"
	open ((system attribute "MAC_OFFICE_TARGET_FILE") as POSIX file)
	set d to active document
	set content of text object of paragraph 2 of d to "申請人：蔡宗翰" & return
	set content of text object of paragraph 3 of d to "出差事由：客戶系統導入" & return
	set alignment of paragraph format of text object of paragraph 3 of d to align paragraph left
	set content of text object of cell 2 of row 1 of table 1 of d to "高雄"
	set content of text object of cell 2 of row 2 of table 1 of d to "3"
	save active document
end tell

tell application "Microsoft Word"
	set outPath to system attribute "MAC_OFFICE_OUTPUT_FILE"
	set d to make new document
	set content of text object of d to "會議紀錄" & return & "時間：11 月 10 日 14:00" & return & "出席：王經理、李工程師" & return & "決議：下週完成測試。" & return & "紀錄：陳秘書"
	set alignment of paragraph format of text object of paragraph 1 of d to align paragraph center
	set alignment of paragraph format of text object of paragraph 5 of d to align paragraph right
	save as active document file name outPath file format format document
	close active document saving no
end tell

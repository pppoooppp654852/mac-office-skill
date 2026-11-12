tell application "Microsoft Word"
	set outPath to system attribute "MAC_OFFICE_OUTPUT_FILE"
	set d to make new document
	set content of text object of d to "會員申請表" & return & "姓名：" & return & "電話：" & return & "地址：" & return & "以上資料僅供會員管理使用。"
	set alignment of paragraph format of text object of paragraph 1 of d to align paragraph center
	set alignment of paragraph format of text object of paragraph 5 of d to align paragraph center
	save as active document file name outPath file format format document
	close active document saving no
end tell

tell application "Microsoft Word"
	set outPath to system attribute "MAC_OFFICE_OUTPUT_FILE"
	set d to make new document
	set content of text object of d to "秋季健行活動" & return & "日期：11 月 22 日（六）" & return & "集合地點：捷運劍潭站 1 號出口"
	set alignment of paragraph format of text object of paragraph 1 of d to align paragraph center
	set tb1 to make new text box at d with properties {left position:150.0, top:400.0, width:240.0, height:40.0}
	set content of text range of text frame of tb1 to "聯絡人：周小姐 分機 1234"
	save as active document file name outPath file format format document
	close active document saving no
end tell

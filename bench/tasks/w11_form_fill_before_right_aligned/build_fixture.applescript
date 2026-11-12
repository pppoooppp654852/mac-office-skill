tell application "Microsoft Word"
	set outPath to system attribute "MAC_OFFICE_OUTPUT_FILE"
	set d to make new document
	set content of text object of d to "研習報名表" & return & "姓名：" & return & "單位：" & return & "備註：" & return & "報名日期：2026/11/11"
	set alignment of paragraph format of text object of paragraph 1 of d to align paragraph center
	set alignment of paragraph format of text object of paragraph 5 of d to align paragraph right
	save as active document file name outPath file format format document
	close active document saving no
end tell

tell application "Microsoft Word"
	set outPath to system attribute "MAC_OFFICE_OUTPUT_FILE"
	set d to make new document
	set content of text object of d to "出差申請單" & return & "申請人：" & return & "出差事由：" & return & "" & return & "核准主管："
	set alignment of paragraph format of text object of paragraph 1 of d to align paragraph center
	set alignment of paragraph format of text object of paragraph 5 of d to align paragraph right
	set t to make new table at text object of paragraph 4 of d with properties {number of rows:2, number of columns:2}
	set content of text object of cell 1 of row 1 of t to "地點"
	set content of text object of cell 1 of row 2 of t to "天數"
	save as active document file name outPath file format format document
	close active document saving no
end tell

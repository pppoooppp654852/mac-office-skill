tell application "Microsoft Word"
	set outPath to system attribute "MAC_OFFICE_OUTPUT_FILE"
	set d to make new document
	set content of text object of d to "請假單" & return & "" & return & "主管簽核：" & return & "備註：請於三日前提出。"
	set alignment of paragraph format of text object of paragraph 1 of d to align paragraph center
	set t to make new table at text object of paragraph 2 of d with properties {number of rows:4, number of columns:2}
	set content of text object of cell 1 of row 1 of t to "姓名"
	set content of text object of cell 1 of row 2 of t to "部門"
	set content of text object of cell 1 of row 3 of t to "假別"
	set content of text object of cell 1 of row 4 of t to "日期"
	save as active document file name outPath file format format document
	close active document saving no
end tell

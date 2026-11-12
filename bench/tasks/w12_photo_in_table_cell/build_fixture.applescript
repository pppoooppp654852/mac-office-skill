tell application "Microsoft Word"
	set outPath to system attribute "MAC_OFFICE_OUTPUT_FILE"
	set d to make new document
	set content of text object of d to "員工資料卡" & return & "" & return & "資料更新日期：2026/11/01"
	set alignment of paragraph format of text object of paragraph 1 of d to align paragraph center
	set t to make new table at text object of paragraph 2 of d with properties {number of rows:3, number of columns:3}
	set content of text object of cell 1 of row 1 of t to "姓名"
	set content of text object of cell 2 of row 1 of t to "李俊傑"
	set content of text object of cell 3 of row 1 of t to "照片"
	set content of text object of cell 1 of row 2 of t to "部門"
	set content of text object of cell 2 of row 2 of t to "業務部"
	set content of text object of cell 1 of row 3 of t to "分機"
	set content of text object of cell 2 of row 3 of t to "2301"
	save as active document file name outPath file format format document
	close active document saving no
end tell

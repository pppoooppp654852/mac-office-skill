tell application "Microsoft Word"
	set outPath to system attribute "MAC_OFFICE_OUTPUT_FILE"
	set d to make new document
	set content of text object of d to "報價單" & return & "" & return & "總金額：NT$ 128,000" & return & "付款條件：月結 30 天"
	set alignment of paragraph format of text object of paragraph 1 of d to align paragraph center
	set t to make new table at text object of paragraph 2 of d with properties {number of rows:4, number of columns:3}
	set content of text object of cell 1 of row 1 of t to "品項"
	set content of text object of cell 2 of row 1 of t to "數量"
	set content of text object of cell 3 of row 1 of t to "金額"
	set content of text object of cell 1 of row 2 of t to "伺服器"
	set content of text object of cell 2 of row 2 of t to "2"
	set content of text object of cell 3 of row 2 of t to "NT$ 96,000"
	set content of text object of cell 1 of row 3 of t to "安裝服務"
	set content of text object of cell 2 of row 3 of t to "1"
	set content of text object of cell 3 of row 3 of t to "NT$ 20,000"
	set content of text object of cell 1 of row 4 of t to "保固延長"
	set content of text object of cell 2 of row 4 of t to "1"
	set content of text object of cell 3 of row 4 of t to "NT$ 12,000"
	save as active document file name outPath file format format document
	close active document saving no
end tell

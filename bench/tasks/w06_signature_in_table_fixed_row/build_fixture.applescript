tell application "Microsoft Word"
	set outPath to system attribute "MAC_OFFICE_OUTPUT_FILE"
	set d to make new document
	set content of text object of d to "驗收單" & return & "" & return & "本單一式兩份。" & return & "總務處 製表"
	set alignment of paragraph format of text object of paragraph 1 of d to align paragraph center
	set alignment of paragraph format of text object of paragraph 4 of d to align paragraph right
	set t to make new table at text object of paragraph 2 of d with properties {number of rows:3, number of columns:2}
	set content of text object of cell 1 of row 1 of t to "驗收項目"
	set content of text object of cell 2 of row 1 of t to "辦公桌椅 20 組"
	set content of text object of cell 1 of row 2 of t to "驗收結果"
	set content of text object of cell 2 of row 2 of t to "合格"
	set content of text object of cell 1 of row 3 of t to "驗收人簽名"
	save as active document file name outPath file format format document
	close active document saving no
end tell

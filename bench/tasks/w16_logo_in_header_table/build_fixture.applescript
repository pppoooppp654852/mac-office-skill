tell application "Microsoft Word"
	set outPath to system attribute "MAC_OFFICE_OUTPUT_FILE"
	set d to make new document
	set content of text object of d to "" & return & "客戶：遠東物流" & return & "報價日期：2026/11/11" & return & "總金額：NT$ 128,000" & return & "有效期限：30 天"
	set t to make new table at text object of paragraph 1 of d with properties {number of rows:1, number of columns:2}
	set content of text object of cell 2 of row 1 of t to "宏遠科技有限公司 報價單"
	save as active document file name outPath file format format document
	close active document saving no
end tell

tell application "Microsoft Word"
	set outPath to system attribute "MAC_OFFICE_OUTPUT_FILE"
	set d to make new document
	set content of text object of d to "公文簽辦單" & return & "承辦人：" & return & "主旨：採購辦公用品。" & return & "說明：依年度預算辦理。"
	set alignment of paragraph format of text object of paragraph 1 of d to align paragraph center
	save as active document file name outPath file format format document97
	close active document saving no
end tell

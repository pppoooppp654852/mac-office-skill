tell application "Microsoft Word"
	set outPath to system attribute "MAC_OFFICE_OUTPUT_FILE"
	set d to make new document
	set content of text object of d to "2026 第三季營運報告（草稿）" & return & "一、營收概況" & return & "本季營收較上季成長 12%。" & return & "二、展望" & return & "預計第四季維持穩定成長。"
	set alignment of paragraph format of text object of paragraph 1 of d to align paragraph center
	save as active document file name outPath file format format document
	close active document saving no
end tell

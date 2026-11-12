tell application "Microsoft Word"
	set outPath to system attribute "MAC_OFFICE_OUTPUT_FILE"
	set d to make new document
	set content of text object of d to "切結書" & return & "本人保證所提供之資料均屬實，如有不實願負法律責任。" & return & "立書人：" & return & "日期："
	set alignment of paragraph format of text object of paragraph 1 of d to align paragraph center
	set alignment of paragraph format of text object of paragraph 4 of d to align paragraph right
	save as active document file name outPath file format format document
	close active document saving no
end tell

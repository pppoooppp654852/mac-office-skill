tell application "Microsoft Word"
	set outPath to system attribute "MAC_OFFICE_OUTPUT_FILE"
	set d to make new document
	set content of text object of d to "合作備忘錄" & return & "立約人：大同貿易股份有限公司（以下簡稱甲方）" & return & "甲方同意提供大同貿易之產品目錄予乙方。" & return & "本備忘錄由大同貿易法務部保管。" & return & "中華民國一一五年十一月十一日"
	set alignment of paragraph format of text object of paragraph 1 of d to align paragraph center
	set alignment of paragraph format of text object of paragraph 5 of d to align paragraph right
	set bold of font object of text object of paragraph 1 of d to true
	save as active document file name outPath file format format document
	close active document saving no
end tell

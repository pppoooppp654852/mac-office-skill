tell application "Microsoft Word"
	set outPath to system attribute "MAC_OFFICE_OUTPUT_FILE"
	set d to make new document
	set content of text object of d to "專案週報" & return & "本週完成登入模組。" & return & "下週預計整合測試。"
	set alignment of paragraph format of text object of paragraph 1 of d to align paragraph center
	set tb1 to make new text box at d with properties {left position:72.0, top:300.0, width:160.0, height:40.0}
	set content of text range of text frame of tb1 to "機密等級：內部"
	set tb2 to make new text box at d with properties {left position:380.0, top:300.0, width:160.0, height:40.0}
	set content of text range of text frame of tb2 to "狀態：草稿"
	save as active document file name outPath file format format document
	close active document saving no
end tell

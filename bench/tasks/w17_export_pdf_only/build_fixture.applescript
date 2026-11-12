tell application "Microsoft Word"
	set outPath to system attribute "MAC_OFFICE_OUTPUT_FILE"
	set d to make new document
	set content of text object of d to "2027 年度工作計畫" & return & "一、目標" & return & "提升客戶滿意度至 90%。" & return & "二、重點項目" & return & "導入新版客服系統。" & return & "三、預算" & return & "總預算新台幣 500 萬元。"
	set alignment of paragraph format of text object of paragraph 1 of d to align paragraph center
	save as active document file name outPath file format format document
	close active document saving no
end tell

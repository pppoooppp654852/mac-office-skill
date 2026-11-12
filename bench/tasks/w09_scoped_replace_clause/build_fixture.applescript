tell application "Microsoft Word"
	set outPath to system attribute "MAC_OFFICE_OUTPUT_FILE"
	set d to make new document
	set content of text object of d to "委託契約書" & return & "第一條　甲方委託乙方辦理系統維護。" & return & "甲方應於每月五日前支付費用。" & return & "第二條　乙方應每季向甲方提出報告。" & return & "甲方得隨時查核乙方工作紀錄。" & return & "第三條　本契約自簽訂日起生效。"
	set alignment of paragraph format of text object of paragraph 1 of d to align paragraph center
	save as active document file name outPath file format format document
	close active document saving no
end tell

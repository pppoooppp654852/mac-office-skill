tell application "Microsoft Excel"
	set outPath to system attribute "MAC_OFFICE_OUTPUT_FILE"
	set wb to make new workbook
	set ws to worksheet 1 of wb
	set name of ws to "報帳"
	set value of range "A1:D1" of ws to {{"日期", "項目", "金額", "審核"}}
	set value of range "A2:C6" of ws to {{"11/01", "計程車", 1200}, {"11/03", "研討會報名", 15800}, {"11/04", "文具", 860}, {"11/06", "機票", 23000}, {"11/08", "餐敘", 9999}}
	save workbook as wb filename outPath file format Excel XML file format
	close active workbook saving no
end tell

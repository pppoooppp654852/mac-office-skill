tell application "Microsoft Excel"
	set outPath to system attribute "MAC_OFFICE_OUTPUT_FILE"
	set wb to make new workbook
	set ws to worksheet 1 of wb
	set name of ws to "明細"
	set value of range "A1:D1" of ws to {{"品項", "數量", "單價", "小計"}}
	set value of range "A2:C6" of ws to {{"筆記型電腦", 5, 32000}, {"螢幕", 10, 6500}, {"鍵盤", 20, 890}, {"滑鼠", 20, 450}, {"擴充座", 5, 3200}}
	set value of range "A7" of ws to "合計"
	save workbook as wb filename outPath file format Excel XML file format
	close active workbook saving no
end tell

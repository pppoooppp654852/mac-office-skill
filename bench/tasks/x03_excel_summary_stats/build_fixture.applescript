tell application "Microsoft Excel"
	set outPath to system attribute "MAC_OFFICE_OUTPUT_FILE"
	set wb to make new workbook
	set ws to worksheet 1 of wb
	set name of ws to "成績"
	set value of range "A1:B1" of ws to {{"姓名", "分數"}}
	set value of range "A2:B7" of ws to {{"甲生", 78}, {"乙生", 92}, {"丙生", 65}, {"丁生", 88}, {"戊生", 71}, {"己生", 95}}
	set ws to make new worksheet at end of wb
	set name of ws to "摘要"
	set value of range "A1:A3" of ws to {{"平均"}, {"最高"}, {"70分以上人數"}}
	save workbook as wb filename outPath file format Excel XML file format
	close active workbook saving no
end tell

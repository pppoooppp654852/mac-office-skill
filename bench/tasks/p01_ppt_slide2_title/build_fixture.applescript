tell application "Microsoft PowerPoint"
	set outPath to system attribute "MAC_OFFICE_OUTPUT_FILE"
	set thePres to make new presentation
	set s to make new slide at end of thePres with properties {layout:slide layout title slide}
	set content of text range of text frame of shape 1 of slide 1 of thePres to "新產品發表"
	set content of text range of text frame of shape 2 of slide 1 of thePres to "2026 年 11 月"
	set s to make new slide at end of thePres with properties {layout:slide layout text slide}
	set content of text range of text frame of shape 1 of slide 2 of thePres to "市場分析（初稿）"
	set content of text range of text frame of shape 2 of slide 2 of thePres to "目標客群：中小企業"
	set s to make new slide at end of thePres with properties {layout:slide layout text slide}
	set content of text range of text frame of shape 1 of slide 3 of thePres to "下一步"
	set content of text range of text frame of shape 2 of slide 3 of thePres to "完成定價策略"
	save thePres in (outPath as POSIX file) as save as Open XML presentation
	close active presentation saving no
end tell

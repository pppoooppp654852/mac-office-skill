tell application "Microsoft PowerPoint"
	set outPath to system attribute "MAC_OFFICE_OUTPUT_FILE"
	set thePres to make new presentation
	set s to make new slide at end of thePres with properties {layout:slide layout title slide}
	set content of text range of text frame of shape 1 of slide 1 of thePres to "第三季回顧"
	set content of text range of text frame of shape 2 of slide 1 of thePres to "業務部"
	set s to make new slide at end of thePres with properties {layout:slide layout text slide}
	set content of text range of text frame of shape 1 of slide 2 of thePres to "重點成果"
	set content of text range of text frame of shape 2 of slide 2 of thePres to "簽約客戶 12 家"
	set s to make new slide at end of thePres with properties {layout:slide layout text slide}
	set content of text range of text frame of shape 1 of slide 3 of thePres to "待改善"
	set content of text range of text frame of shape 2 of slide 3 of thePres to "交期延誤 3 件"
	save thePres in (outPath as POSIX file) as save as Open XML presentation
	close active presentation saving no
end tell

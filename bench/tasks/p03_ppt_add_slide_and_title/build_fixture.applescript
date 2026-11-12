tell application "Microsoft PowerPoint"
	set outPath to system attribute "MAC_OFFICE_OUTPUT_FILE"
	set thePres to make new presentation
	set s to make new slide at end of thePres with properties {layout:slide layout title slide}
	set content of text range of text frame of shape 1 of slide 1 of thePres to "資安教育訓練（草案）"
	set content of text range of text frame of shape 2 of slide 1 of thePres to "資訊處"
	set s to make new slide at end of thePres with properties {layout:slide layout text slide}
	set content of text range of text frame of shape 1 of slide 2 of thePres to "課程大綱"
	set content of text range of text frame of shape 2 of slide 2 of thePres to "密碼管理"
	save thePres in (outPath as POSIX file) as save as Open XML presentation
	close active presentation saving no
end tell

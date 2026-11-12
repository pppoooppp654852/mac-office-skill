tell application "Microsoft Word"
	open ((system attribute "MAC_OFFICE_TARGET_FILE") as POSIX file)
	set d to active document
	set content of text object of paragraph 3 of d to "立書人：黃建國" & return
	set alignment of paragraph format of text object of paragraph 3 of d to align paragraph left
	set content of text object of paragraph 4 of d to "日期：2026 年 11 月 11 日"
	set alignment of paragraph format of text object of paragraph 4 of d to align paragraph right
	save active document
end tell

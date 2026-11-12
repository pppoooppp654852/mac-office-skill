tell application "Microsoft Word"
	open ((system attribute "MAC_OFFICE_TARGET_FILE") as POSIX file)
	set d to active document
	set content of text object of paragraph 2 of d to "姓名：林雅婷" & return
	set content of text object of paragraph 3 of d to "電話：0912-345-678" & return
	set content of text object of paragraph 4 of d to "地址：台北市信義區松仁路 100 號" & return
	set alignment of paragraph format of text object of paragraph 4 of d to align paragraph left
	save active document
end tell

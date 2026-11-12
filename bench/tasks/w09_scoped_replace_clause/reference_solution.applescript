tell application "Microsoft Word"
	open ((system attribute "MAC_OFFICE_TARGET_FILE") as POSIX file)
	set d to active document
	set content of text object of paragraph 4 of d to "第二條　乙方應每季向委託人提出報告。" & return
	set content of text object of paragraph 5 of d to "委託人得隨時查核乙方工作紀錄。" & return
	save active document
end tell

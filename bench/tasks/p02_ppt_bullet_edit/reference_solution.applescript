tell application "Microsoft PowerPoint"
	open ((system attribute "MAC_OFFICE_TARGET_FILE") as POSIX file)
	repeat 25 times
		if (count of presentations) > 0 then exit repeat
		delay 0.1
	end repeat
	set content of text range of text frame of shape 2 of slide 3 of active presentation to "交期延誤 1 件"
	save active presentation in ((system attribute "MAC_OFFICE_OUTPUT_FILE") as POSIX file) as save as Open XML presentation
end tell

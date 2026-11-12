tell application "Microsoft PowerPoint"
	open ((system attribute "MAC_OFFICE_TARGET_FILE") as POSIX file)
	repeat 25 times
		if (count of presentations) > 0 then exit repeat
		delay 0.1
	end repeat
	set content of text range of text frame of shape 1 of slide 2 of active presentation to "市場分析"
	save active presentation
end tell

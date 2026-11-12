tell application "Microsoft PowerPoint"
	open ((system attribute "MAC_OFFICE_TARGET_FILE") as POSIX file)
	repeat 25 times
		if (count of presentations) > 0 then exit repeat
		delay 0.1
	end repeat
	set thePres to active presentation
	set content of text range of text frame of shape 1 of slide 1 of thePres to "資安教育訓練"
	set newSlide to make new slide at end of thePres with properties {layout:slide layout text slide}
	set content of text range of text frame of shape 1 of slide 3 of thePres to "課後測驗"
	set content of text range of text frame of shape 2 of slide 3 of thePres to "線上測驗 11/30 截止"
	save active presentation
end tell

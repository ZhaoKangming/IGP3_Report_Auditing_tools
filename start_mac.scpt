tell application "Terminal"
	set newTab to do script "cd /Users/liqun/Desktop/【李群】赋能起航二期审核工具/programs/"
	do script "python3 ./window_report_checker.py" in front window
end tell
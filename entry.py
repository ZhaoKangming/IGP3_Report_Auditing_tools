import subprocess

CREATE_NO_WINDOW = 0x08000000
subprocess.call('python window_report_checker.py', creationflags=CREATE_NO_WINDOW)
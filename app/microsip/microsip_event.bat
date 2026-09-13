@echo off

set "PYTHON=C:\Users\Asko\Desktop\ZIK\.venv\Scripts\python.exe"
set "RECEIVER=C:\Users\Asko\Desktop\ZIK\app\microsip\event_receiver.py"

"%PYTHON%" "%RECEIVER%" %*

exit /b 0
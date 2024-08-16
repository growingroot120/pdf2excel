@echo off
REM Change directory to the script location
cd /d %~dp0

REM If using a virtual environment, activate it
REM Uncomment the next line and modify the path if you use a virtual environment
REM 
call .\pdf2excel\Scripts\activate

REM Run the Python script
python excel_merge.py

REM Pause the script to see the output
pause

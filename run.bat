@echo off
REM Double-click to run the ASX odd securities scan on this PC.
REM The report is saved in the "output" folder. Your open Excel files are not touched.
cd /d "%~dp0"
echo Installing / checking the packages this needs (first run takes a minute)...
python -m pip install --quiet --user -r requirements.txt
echo.
echo Running the scan...
python scan.py %*
echo.
echo Done. Opening the output folder...
start "" "%~dp0output"
pause

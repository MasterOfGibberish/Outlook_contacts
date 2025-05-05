@echo off
echo Starting Outlook Contact Exporter...
echo.

:: Try to run with python command
python main.py
if %ERRORLEVEL% EQU 0 goto end

:: If that fails, try with py command
py main.py
if %ERRORLEVEL% EQU 0 goto end

:: If both fail, try with pythonw
pythonw main.py
if %ERRORLEVEL% EQU 0 goto end

:: If all commands fail, show error
echo.
echo ERROR: Could not run the Outlook Contact Exporter.
echo Please make sure Python is installed on your system.
echo.
echo Try installing Python from: https://www.python.org/downloads/
echo.
pause

:end 
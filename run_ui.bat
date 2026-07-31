@echo off
REM Double-click this file to launch the Web UI with an elevated (admin)
REM token automatically - see run_ui.ps1 for why this is needed and what
REM it does. You will see one Windows "User Account Control" prompt; click
REM Yes. No other steps outside the browser UI are required afterward.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0run_ui.ps1"

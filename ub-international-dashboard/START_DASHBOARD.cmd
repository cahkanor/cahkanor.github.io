@echo off
setlocal
where python >nul 2>&1
if %errorlevel%==0 goto start_server
where py >nul 2>&1
if %errorlevel%==0 goto start_with_py
echo Python was not found.
echo Install Python 3, then run this launcher again.
pause
exit /b 1

:start_server
start "UB Dashboard Server" /min /D "%~dp0dist" python -m http.server 4173
goto open_browser

:start_with_py
start "UB Dashboard Server" /min /D "%~dp0dist" py -m http.server 4173

:open_browser
timeout /t 2 /nobreak >nul
start "" "http://localhost:4173/"
echo UB Dashboard started at http://localhost:4173/
echo Close the minimized "UB Dashboard Server" window to stop it.
timeout /t 3 /nobreak >nul
endlocal

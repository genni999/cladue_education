@echo off
REM ===================================================================
REM  MGB Downloader launcher (Windows)
REM  Double-click this file, or run "run.bat" in a terminal.
REM  On first run it creates a virtual environment and installs deps.
REM
REM  NOTE: messages are kept in plain ASCII on purpose. cmd.exe reads
REM  .bat files using the OEM codepage, so non-ASCII text here gets
REM  garbled and can break the script's control flow.
REM ===================================================================

setlocal
set "HERE=%~dp0"
set "VENV_PY=%HERE%.venv\Scripts\python.exe"

if not exist "%VENV_PY%" (
    echo [setup] First run - creating virtual environment. This may take a few minutes.
    py -m venv "%HERE%.venv"
    if errorlevel 1 goto nopython
    "%VENV_PY%" -m pip install --quiet --disable-pip-version-check -r "%HERE%requirements.txt"
    if errorlevel 1 goto installfailed
    echo [setup] Done.
)

echo.
echo Starting MGB Downloader...
echo Opening  http://localhost:8501  in your browser.
echo Press Ctrl+C in this window to stop.
echo.
REM The app runs headless (see .streamlit\config.toml), so open the browser here.
REM "ping" is used instead of "timeout" as a delay: timeout fails when stdin
REM is redirected (e.g. when launched from another script).
start "" /b cmd /c "ping -n 7 127.0.0.1 >nul & start "" http://localhost:8501"
"%VENV_PY%" -m streamlit run "%HERE%app.py"
goto end

:nopython
echo.
echo [ERROR] Python was not found.
echo Install it from https://www.python.org/downloads/
echo and be sure to tick "Add python.exe to PATH" during setup.
pause
goto end

:installfailed
echo.
echo [ERROR] Dependency install failed. See the messages above.
pause

:end
endlocal

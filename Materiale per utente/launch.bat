@echo off
REM ===========================================================================
REM Sunrise + Astra Input Builder — launcher
REM Runs entirely from this folder. Only requirement: internet access at the
REM very first launch (to download Python + dependencies). No admin rights,
REM no Python installation, no compiled executable.
REM
REM Design note: every step that would previously live inside a parenthesised
REM IF block has been moved to a labelled block reached via GOTO. This
REM sidesteps a whole class of cmd.exe quoting/expansion bugs that were
REM causing pip install to silently fail on some Windows configurations
REM even though the exact same command typed by hand worked fine.
REM
REM All setup output (in particular pip install output) is also mirrored to
REM install.log next to launch.bat, so any residual failure can be
REM diagnosed after the fact.
REM ===========================================================================

setlocal EnableDelayedExpansion
cd /d "%~dp0"

set "PY_VER=3.11.9"
set "PY_ZIP_URL=https://www.python.org/ftp/python/%PY_VER%/python-%PY_VER%-embed-amd64.zip"
set "GET_PIP_URL=https://bootstrap.pypa.io/get-pip.py"
set "PY_DIR=python"
set "PY_EXE=%PY_DIR%\python.exe"
set "PTH_FILE=%PY_DIR%\python311._pth"
set "MARKER=%PY_DIR%\.deps-installed"
set "LOG=install.log"

REM --- Step 1: download and unpack Python embeddable ------------------------
if exist "%PY_EXE%" goto step_pth
echo First-time setup — downloading Python %PY_VER% embeddable...
if not exist "%PY_DIR%" mkdir "%PY_DIR%"
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "$ErrorActionPreference='Stop'; try { Invoke-WebRequest -UseBasicParsing -Uri '%PY_ZIP_URL%' -OutFile 'python.zip'; Expand-Archive -Force 'python.zip' -DestinationPath '%PY_DIR%'; Remove-Item 'python.zip' } catch { Write-Host $_; exit 1 }"
if errorlevel 1 goto err_python_download

REM --- Step 2: enable pip / site-packages loader ---------------------------
:step_pth
REM Python embeddable ships with a *._pth file that disables site.py. Pip
REM needs `import site` uncommented to install into Lib\site-packages.
if not exist "%PTH_FILE%" goto step_pip_bootstrap
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "$p='%PTH_FILE%'; $t=Get-Content $p -Raw; if ($t -notmatch '(?m)^import site') { $t=[Regex]::Replace($t,'(?m)^#\s*import site','import site'); if ($t -notmatch '(?m)^import site') { $t += \"`nimport site`n\" }; Set-Content $p $t -NoNewline }"

REM --- Step 3: bootstrap pip -----------------------------------------------
:step_pip_bootstrap
if exist "%PY_DIR%\Scripts\pip.exe" goto step_deps
echo Bootstrapping pip...
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "$ErrorActionPreference='Stop'; try { Invoke-WebRequest -UseBasicParsing -Uri '%GET_PIP_URL%' -OutFile 'get-pip.py' } catch { Write-Host $_; exit 1 }"
if errorlevel 1 goto err_getpip_download
"%PY_EXE%" get-pip.py --no-warn-script-location
if errorlevel 1 goto err_pip_bootstrap
del get-pip.py

REM --- Step 4: install runtime dependencies (one-time) ---------------------
REM Moved out of parenthesised IF blocks: some cmd.exe versions mangled the
REM quoted pip install line when it was nested inside one, causing pip to
REM be launched with the wrong arguments (or not launched at all).
:step_deps
if not exist "%MARKER%" goto do_install
"%PY_EXE%" -c "import streamlit, openpyxl, pandas" >nul 2>nul
if not errorlevel 1 goto step_credentials
echo Marker present but dependencies broken — reinstalling.
del "%MARKER%"

:do_install
echo Installing dependencies (one-time, may take a couple of minutes)...
echo Detailed output will be written to %LOG%.
"%PY_EXE%" -m pip install --no-warn-script-location "streamlit>=1.30" "openpyxl>=3.1" "pandas>=2.0" > "%LOG%" 2>&1
if errorlevel 1 goto err_pip_install
"%PY_EXE%" -c "import streamlit, openpyxl, pandas" >nul 2>nul
if errorlevel 1 goto err_pip_import
echo done > "%MARKER%"

REM --- Step 5: pre-populate Streamlit user credentials ---------------------
:step_credentials
REM Some Streamlit versions prompt for an email address the very first
REM time they run for a given Windows user, blocking the terminal until
REM something is typed. Writing an empty credentials.toml in advance
REM suppresses the prompt entirely.
if not exist "%USERPROFILE%\.streamlit" mkdir "%USERPROFILE%\.streamlit" 2>nul
if not exist "%USERPROFILE%\.streamlit\credentials.toml" (
    > "%USERPROFILE%\.streamlit\credentials.toml" echo [general]
    >> "%USERPROFILE%\.streamlit\credentials.toml" echo email = ""
)

REM --- Step 6: open the default browser once Streamlit is ready ------------
start "" /B powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "Start-Sleep -Seconds 4; Start-Process 'http://localhost:8501'"

REM --- Step 7: launch Streamlit --------------------------------------------
REM No CLI flags: .streamlit\config.toml shipped in this folder already sets
REM `server.headless = true` and `browser.gatherUsageStats = false`. Env
REM vars are kept as a belt-and-suspenders fallback in case someone deletes
REM the config file.
set "STREAMLIT_SERVER_HEADLESS=true"
set "STREAMLIT_BROWSER_GATHER_USAGE_STATS=false"
echo.
echo Starting Sunrise + Astra Input Builder...
echo (Close this window to stop the app.)
echo.
"%PY_EXE%" -m streamlit run app.py
set "STREAMLIT_EXIT=!ERRORLEVEL!"
goto end

REM --- Error handlers -------------------------------------------------------
:err_python_download
echo.
echo ================================================================
echo ERROR: Python download failed. Check internet connection and rerun.
echo ================================================================
pause
exit /b 1

:err_getpip_download
echo.
echo ================================================================
echo ERROR: get-pip.py download failed. Check internet connection and rerun.
echo ================================================================
pause
exit /b 1

:err_pip_bootstrap
echo.
echo ================================================================
echo ERROR: pip bootstrap (python get-pip.py) failed.
echo ================================================================
pause
exit /b 1

:err_pip_install
echo.
echo ================================================================
echo ERROR: pip install failed. Full log written to %LOG%.
echo Last 30 lines of the log:
echo ----------------------------------------------------------------
powershell -NoProfile -Command "Get-Content '%LOG%' -Tail 30"
echo ================================================================
pause
exit /b 1

:err_pip_import
echo.
echo ================================================================
echo ERROR: pip install completed but Python still cannot import
echo        streamlit / openpyxl / pandas. This usually means the
echo        embeddable Python's path configuration is broken.
echo        Try deleting the "python" folder and rerun this launcher.
echo Log at %LOG%.
echo ================================================================
pause
exit /b 1

REM --- Final pause (also on clean exit) ------------------------------------
:end
echo.
echo ================================================================
if not "!STREAMLIT_EXIT!"=="0" (
    echo Streamlit exited with error code !STREAMLIT_EXIT!.
    echo If you don't see a traceback above, likely causes:
    echo  - antivirus blocked python.exe from listening on port 8501;
    echo  - port 8501 already in use by another process;
    echo  - a dependency install completed only partially.
    echo Try:
    echo  1) close this window;
    echo  2) delete the "python\.deps-installed" file;
    echo  3) run launch.bat again.
) else (
    echo Streamlit has exited cleanly. Close this window whenever you like.
)
echo ================================================================
pause

endlocal

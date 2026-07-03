@echo off
REM ===========================================================================
REM Sunrise + Astra Input Builder — launcher
REM Runs entirely from this folder. Only requirement: internet access at the
REM very first launch (to download Python + dependencies). No admin rights,
REM no Python installation, no compiled executable.
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

REM --- Step 1: download and unpack Python embeddable ------------------------
if not exist "%PY_EXE%" (
    echo First-time setup — downloading Python %PY_VER% embeddable.
    if not exist "%PY_DIR%" mkdir "%PY_DIR%"
    powershell -NoProfile -ExecutionPolicy Bypass -Command ^
        "$ErrorActionPreference='Stop'; try { Invoke-WebRequest -UseBasicParsing -Uri '%PY_ZIP_URL%' -OutFile 'python.zip'; Expand-Archive -Force 'python.zip' -DestinationPath '%PY_DIR%'; Remove-Item 'python.zip' } catch { Write-Host $_; exit 1 }"
    if !ERRORLEVEL! NEQ 0 (
        echo.
        echo ERROR: Python download failed. Check your internet connection and rerun.
        pause
        exit /b 1
    )
)

REM --- Step 2: enable pip / site-packages loader ---------------------------
REM Python embeddable ships with a *._pth file that disables site.py. Pip
REM needs `import site` uncommented to install packages into Lib\site-packages.
if exist "%PTH_FILE%" (
    powershell -NoProfile -ExecutionPolicy Bypass -Command ^
        "$p='%PTH_FILE%'; $t=Get-Content $p -Raw; if ($t -notmatch '(?m)^import site') { $t=[Regex]::Replace($t,'(?m)^#\s*import site','import site'); if ($t -notmatch '(?m)^import site') { $t += \"`nimport site`n\" }; Set-Content $p $t -NoNewline }"
)

REM --- Step 3: bootstrap pip -----------------------------------------------
if not exist "%PY_DIR%\Scripts\pip.exe" (
    echo Bootstrapping pip...
    powershell -NoProfile -ExecutionPolicy Bypass -Command ^
        "$ErrorActionPreference='Stop'; try { Invoke-WebRequest -UseBasicParsing -Uri '%GET_PIP_URL%' -OutFile 'get-pip.py' } catch { Write-Host $_; exit 1 }"
    if !ERRORLEVEL! NEQ 0 (
        echo ERROR: get-pip.py download failed.
        pause
        exit /b 1
    )
    "%PY_EXE%" get-pip.py --no-warn-script-location
    if !ERRORLEVEL! NEQ 0 (
        echo ERROR: pip install failed.
        pause
        exit /b 1
    )
    del get-pip.py
)

REM --- Step 4: install runtime dependencies (one-time) ---------------------
REM The marker is written ONLY after we verify that streamlit is actually
REM importable — that way a partially-failed install doesn't fool us at the
REM next launch.
if exist "%MARKER%" (
    "%PY_EXE%" -c "import streamlit, openpyxl, pandas" 1>nul 2>nul
    if !ERRORLEVEL! NEQ 0 (
        echo Marker present but dependencies broken — reinstalling.
        del "%MARKER%"
    )
)

if not exist "%MARKER%" (
    echo Installing dependencies (one-time, may take a couple of minutes)...
    "%PY_EXE%" -m pip install --no-warn-script-location streamlit^>=1.30 openpyxl^>=3.1 pandas^>=2.0
    if !ERRORLEVEL! NEQ 0 (
        echo.
        echo ERROR: Dependency installation failed. Check your internet connection.
        pause
        exit /b 1
    )
    REM Verify the install really worked before writing the marker.
    "%PY_EXE%" -c "import streamlit, openpyxl, pandas" 1>nul 2>nul
    if !ERRORLEVEL! NEQ 0 (
        echo.
        echo ERROR: Dependencies installed but Python cannot import them.
        echo        Try deleting the "python" folder and rerun this launcher.
        pause
        exit /b 1
    )
    echo done > "%MARKER%"
)

REM --- Step 5: open the default browser once Streamlit is ready ------------
start "" /B powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "Start-Sleep -Seconds 4; Start-Process 'http://localhost:8501'"

REM --- Step 6: launch Streamlit --------------------------------------------
echo.
echo Starting Sunrise + Astra Input Builder...
echo (Close this window to stop the app.)
echo.
"%PY_EXE%" -m streamlit run app.py --server.headless=true
set "STREAMLIT_EXIT=!ERRORLEVEL!"

REM Keep the window open on non-zero exit so any traceback stays visible.
if !STREAMLIT_EXIT! NEQ 0 (
    echo.
    echo ================================================================
    echo Streamlit exited with error code !STREAMLIT_EXIT!.
    echo If you don't see a traceback above, likely causes:
    echo  - antivirus blocked python.exe from listening on port 8501;
    echo  - port 8501 already in use by another process;
    echo  - a dependency install completed only partially.
    echo Try:
    echo  1) close this window;
    echo  2) delete the "python\.deps-installed" file;
    echo  3) run launch.bat again.
    echo ================================================================
    pause
)

endlocal

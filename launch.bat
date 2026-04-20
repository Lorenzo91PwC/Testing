@echo off
REM ===========================================================================
REM Excel Pipeline launcher for Windows
REM Checks for uv, pulls latest logic from GitHub, syncs deps, starts Streamlit.
REM ===========================================================================

cd /d "%~dp0"

REM --- Check for uv ---------------------------------------------------------
where uv >nul 2>nul
if %ERRORLEVEL% NEQ 0 (
    echo Installing uv ^(one-time^)...
    powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
    if %ERRORLEVEL% NEQ 0 (
        echo.
        echo ERROR: uv install failed. Check your internet connection.
        pause
        exit /b 1
    )
    REM Add uv to PATH for this session
    set "PATH=%USERPROFILE%\.local\bin;%PATH%"
)

REM --- Pull latest logic from GitHub ---------------------------------------
echo Checking for updates...
git pull --ff-only
if %ERRORLEVEL% NEQ 0 (
    echo WARNING: Could not pull updates. Continuing with local version.
)

REM --- Sync dependencies ---------------------------------------------------
echo Syncing dependencies...
uv sync
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ERROR: Dependency sync failed.
    pause
    exit /b 1
)

REM --- Check for .env ------------------------------------------------------
if not exist .env (
    echo.
    echo ==========================================================
    echo   .env file not found.
    echo   Copy .env.example to .env and add your ANTHROPIC_API_KEY
    echo   Get a key at https://console.anthropic.com
    echo ==========================================================
    echo.
    pause
    exit /b 1
)

REM --- Launch --------------------------------------------------------------
echo Starting Excel Pipeline...
echo (Close this window to stop the app.)
uv run streamlit run app.py

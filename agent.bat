@echo off
setlocal EnableExtensions EnableDelayedExpansion

rem =============================================================
rem AI Agent - Installer and launcher
rem -------------------------------------------------------------
rem 1) Check the repository version and update if needed.
rem 2) Ensure Python and dependencies are installed/up-to-date.
rem 3) Launch the GUI application.
rem =============================================================

rem -------- Paths and configuration --------
set "BASE_DIR=%~dp0"
if "%BASE_DIR:~-1%"=="\\" set "BASE_DIR=%BASE_DIR:~0,-1%"
set "INSTALL_ROOT=%BASE_DIR%\ai_agent"
set "SRC_DIR=%INSTALL_ROOT%\source"
set "PY_HOME=%INSTALL_ROOT%\.python"
set "VENV_DIR=%INSTALL_ROOT%\.venv"
set "VENV_PY=%VENV_DIR%\Scripts\python.exe"
set "APP_ENTRY=%SRC_DIR%\app\main.py"
set "REQ_FILE=%SRC_DIR%\requirements.txt"
set "VERSION_FILE=%INSTALL_ROOT%\.source_version"
set "GIT_REMOTE=https://github.com/NIAENGD/AI_Agent.git"
set "GIT_BRANCH=main"
set "REMOTE_ZIP=https://github.com/NIAENGD/AI_Agent/archive/refs/heads/main.zip"
set "PY_VERSION=3.11.9"
set "PY_EMBED_ZIP_URL=https://www.python.org/ftp/python/%PY_VERSION%/python-%PY_VERSION%-embed-amd64.zip"

set "SELF_PATH=%~f0"
set "NEED_RESTART="
set "UPDATED="

call :print_banner
call :ensure_directories || goto :fail
call :sync_source || goto :fail
if defined NEED_RESTART (
    call :restart_self
    endlocal
    exit /b %errorlevel%
)
call :prepare_python || goto :fail
call :install_dependencies || goto :fail
call :launch_app || goto :fail

endlocal
exit /b 0

rem =============================================================
:print_banner
echo ==================================================
echo  AI Agent installer and launcher
echo  Install root: %INSTALL_ROOT%
echo ==================================================
exit /b 0

rem =============================================================
:ensure_directories
if not exist "%INSTALL_ROOT%" mkdir "%INSTALL_ROOT%" >nul 2>nul
if not exist "%INSTALL_ROOT%" (
    echo Failed to create installation directory at %INSTALL_ROOT%.
    exit /b 1
)
exit /b 0

rem =============================================================
:sync_source
set "GIT_AVAILABLE="
where git >nul 2>nul && set "GIT_AVAILABLE=1"

call :get_remote_sha
set "REMOTE_SHA=%RETURN_VALUE%"
call :get_local_sha
set "LOCAL_SHA=%RETURN_VALUE%"

if defined GIT_AVAILABLE if exist "%SRC_DIR%\.git" (
    call :git_update_repo || exit /b 1
    goto :post_sync
)

if defined GIT_AVAILABLE if not exist "%SRC_DIR%\.git" (
    call :git_clone_repo && goto :post_sync
)

call :zip_sync_repo || exit /b 1

:post_sync
call :refresh_self
if defined UPDATED echo Update detected.
exit /b 0

rem =============================================================
:git_update_repo
pushd "%SRC_DIR%" >nul 2>nul
for /f "delims=" %%h in ('git rev-parse HEAD 2^>nul') do set "OLD_HEAD=%%h"
git fetch --all --prune
git pull --ff-only origin %GIT_BRANCH%
set "PULL_RC=%errorlevel%"
for /f "delims=" %%h in ('git rev-parse HEAD 2^>nul') do set "NEW_HEAD=%%h"
popd >nul 2>nul
if "%PULL_RC%"=="0" if not "%OLD_HEAD%"=="%NEW_HEAD%" set "UPDATED=1"
if "%PULL_RC%"=="0" exit /b 0
echo Git pull failed, continuing with existing checkout.
exit /b 0

rem =============================================================
:git_clone_repo
if exist "%SRC_DIR%" if not exist "%SRC_DIR%\.git" call :wipe_existing_source
if exist "%SRC_DIR%\.git" exit /b 0
echo Cloning repository into %SRC_DIR% ...
git clone --branch %GIT_BRANCH% "%GIT_REMOTE%" "%SRC_DIR%"
if errorlevel 1 (
    echo Clone failed. Falling back to zip download.
    exit /b 1
)
set "UPDATED=1"
exit /b 0

rem =============================================================
:zip_sync_repo
echo Using zip download flow.
if defined REMOTE_SHA if exist "%VERSION_FILE%" (
    set /p "CACHED_SHA=" < "%VERSION_FILE%"
    if /i "%CACHED_SHA%"=="%REMOTE_SHA%" (
        if exist "%SRC_DIR%" (
            echo Existing files match latest known version. Skipping download.
            exit /b 0
        )
    )
)
set "ZIP_PATH=%TEMP%\ai_agent_latest.zip"
set "EXTRACT_DIR=%TEMP%\ai_agent_extract"

powershell -NoLogo -NoProfile -Command "try { $ProgressPreference='SilentlyContinue'; Invoke-WebRequest -UseBasicParsing -Headers @{ 'User-Agent'='ai-agent-bootstrapper' } -Uri '%REMOTE_ZIP%' -OutFile '%ZIP_PATH%' } catch { exit 1 }" >nul
if errorlevel 1 (
    echo Failed to download source archive.
    exit /b 1
)

if exist "%EXTRACT_DIR%" rd /s /q "%EXTRACT_DIR%" >nul 2>nul
powershell -NoLogo -NoProfile -Command "try { Expand-Archive -Path '%ZIP_PATH%' -DestinationPath '%EXTRACT_DIR%' -Force } catch { exit 1 }" >nul
if errorlevel 1 (
    echo Failed to extract archive.
    exit /b 1
)
for /d %%d in ("%EXTRACT_DIR%\*") do set "UNPACKED_DIR=%%d"
if not defined UNPACKED_DIR (
    echo Extracted archive is empty.
    exit /b 1
)
call :wipe_existing_source || exit /b 1
xcopy "%UNPACKED_DIR%" "%SRC_DIR%" /E /I /Y >nul
if errorlevel 1 (
    echo Failed to copy extracted files.
    exit /b 1
)
if exist "%ZIP_PATH%" del "%ZIP_PATH%" >nul 2>nul
if exist "%EXTRACT_DIR%" rd /s /q "%EXTRACT_DIR%" >nul 2>nul
set "UPDATED=1"
if defined REMOTE_SHA echo %REMOTE_SHA%>"%VERSION_FILE%"
exit /b 0

rem =============================================================
:get_remote_sha
set "RETURN_VALUE="
if defined GIT_AVAILABLE (
    for /f "tokens=1" %%s in ('git ls-remote "%GIT_REMOTE%" %GIT_BRANCH% 2^>nul') do set "RETURN_VALUE=%%s"
)
if not defined RETURN_VALUE (
    powershell -NoLogo -NoProfile -Command "try { $ProgressPreference='SilentlyContinue'; $resp=Invoke-RestMethod -UseBasicParsing -Headers @{ 'User-Agent'='ai-agent-bootstrapper' } 'https://api.github.com/repos/NIAENGD/AI_Agent/commits/%GIT_BRANCH%'; if($resp.sha){ [Console]::WriteLine($resp.sha) } } catch { }" >"%TEMP%\ai_agent_sha.txt"
    if exist "%TEMP%\ai_agent_sha.txt" (
        set /p "RETURN_VALUE=" < "%TEMP%\ai_agent_sha.txt"
        del "%TEMP%\ai_agent_sha.txt" >nul 2>nul
    )
)
exit /b 0

rem =============================================================
:get_local_sha
set "RETURN_VALUE="
if exist "%SRC_DIR%\.git" (
    for /f "delims=" %%h in ('git -C "%SRC_DIR%" rev-parse HEAD 2^>nul') do set "RETURN_VALUE=%%h"
) else if exist "%VERSION_FILE%" (
    set /p "RETURN_VALUE=" < "%VERSION_FILE%"
)
exit /b 0

rem =============================================================
:wipe_existing_source
if not exist "%SRC_DIR%" exit /b 0
rd /s /q "%SRC_DIR%" >nul 2>nul
if exist "%SRC_DIR%" exit /b 1
exit /b 0

rem =============================================================
:refresh_self
if not exist "%SRC_DIR%\agent.bat" exit /b 0
fc /b "%SELF_PATH%" "%SRC_DIR%\agent.bat" >nul 2>nul
if errorlevel 1 (
    copy /Y "%SRC_DIR%\agent.bat" "%SELF_PATH%" >nul
    set "NEED_RESTART=1"
)
exit /b 0

rem =============================================================
:restart_self
echo Restarting with latest launcher...
call "%SELF_PATH%" /relaunch
exit /b %errorlevel%

rem =============================================================
:prepare_python
set "PYTHON_CMD="
for %%p in ("%PY_HOME%\python.exe" "%PY_HOME%\python3.exe" "%PY_HOME%\pythonw.exe") do if exist "%%~p" set "PYTHON_CMD=%%~fp"

if not defined PYTHON_CMD (
    rem Check for system Python before downloading our own
    for %%p in (python.exe python3.exe) do (
        for /f "delims=" %%q in ('where %%p 2^>nul') do if not defined PYTHON_CMD set "PYTHON_CMD=%%q"
    )
)

if defined PYTHON_CMD (
    "%PYTHON_CMD%" -V >nul 2>nul
    if not errorlevel 1 exit /b 0
    echo Existing Python appears broken. Reinstalling...
    set "PYTHON_CMD="
)

if not exist "%PY_HOME%" mkdir "%PY_HOME%" >nul 2>nul
call :install_embeddable_python || exit /b 1
for %%p in ("%PY_HOME%\python.exe" "%PY_HOME%\python3.exe" "%PY_HOME%\pythonw.exe") do if exist "%%~p" set "PYTHON_CMD=%%~fp"
if not defined PYTHON_CMD (
    echo No usable Python interpreter found after installation.
    exit /b 1
)
exit /b 0

rem =============================================================
:install_embeddable_python
echo Downloading embeddable Python %PY_VERSION% ...
set "PY_EMBED_ZIP=%TEMP%\python-%PY_VERSION%-embed-amd64.zip"
powershell -NoLogo -NoProfile -Command "try { $ProgressPreference='SilentlyContinue'; Invoke-WebRequest -UseBasicParsing -Uri '%PY_EMBED_ZIP_URL%' -OutFile '%PY_EMBED_ZIP%' } catch { exit 1 }" >nul
if errorlevel 1 (
    echo Failed to download embeddable Python package.
    exit /b 1
)
if not exist "%PY_HOME%" mkdir "%PY_HOME%" >nul 2>nul
echo Extracting embeddable runtime to %PY_HOME% ...
powershell -NoLogo -NoProfile -Command "try { Expand-Archive -Path '%PY_EMBED_ZIP%' -DestinationPath '%PY_HOME%' -Force } catch { exit 1 }" >nul
if errorlevel 1 (
    echo Failed to extract embeddable Python runtime.
    if exist "%PY_EMBED_ZIP%" del "%PY_EMBED_ZIP%" >nul 2>nul
    exit /b 1
)
if exist "%PY_EMBED_ZIP%" del "%PY_EMBED_ZIP%" >nul 2>nul

set "PY_TAG="
for /f "tokens=1-2 delims=." %%a in ("%PY_VERSION%") do set "PY_TAG=%%a%%b"
if defined PY_TAG if exist "%PY_HOME%\python%PY_TAG%._pth" (
    (
        echo python%PY_TAG%.zip
        echo .
        echo Lib
        echo Lib\site-packages
        echo import site
    )> "%PY_HOME%\python%PY_TAG%._pth"
)
exit /b 0

rem =============================================================
:install_dependencies
if not exist "%SRC_DIR%" (
    echo Source directory missing at %SRC_DIR%.
    exit /b 1
)
set "USE_VENV=1"
"%PYTHON_CMD%" -c "import venv" >nul 2>nul || set "USE_VENV="

if defined USE_VENV (
    if not exist "%VENV_DIR%" (
        echo Creating virtual environment...
        "%PYTHON_CMD%" -m venv "%VENV_DIR%" >nul 2>nul
    )
    if not exist "%VENV_PY%" (
        echo Virtual environment looks broken. Recreating...
        rd /s /q "%VENV_DIR%" >nul 2>nul
        "%PYTHON_CMD%" -m venv "%VENV_DIR%" >nul 2>nul
    )
    set "APP_PY=%VENV_PY%"
    ) else (
        set "APP_PY=%PYTHON_CMD%"
    )

    call :ensure_pip "%APP_PY%" || exit /b 1
    "%APP_PY%" -m pip install --upgrade pip >nul 2>nul
    if exist "%REQ_FILE%" (
        echo Installing Python dependencies...
        "%APP_PY%" -m pip install --upgrade -r "%REQ_FILE%"
    )
    exit /b %errorlevel%

rem =============================================================
:ensure_pip
set "TARGET_PY=%~1"
if not defined TARGET_PY exit /b 1
"%TARGET_PY%" -m pip --version >nul 2>nul && exit /b 0
echo Bootstrapping pip...
"%TARGET_PY%" -m ensurepip --upgrade >nul 2>nul
if not errorlevel 1 "%TARGET_PY%" -m pip --version >nul 2>nul && exit /b 0
set "GETPIP_SCRIPT=%TEMP%\get-pip.py"
powershell -NoLogo -NoProfile -Command "try { $ProgressPreference='SilentlyContinue'; Invoke-WebRequest -UseBasicParsing -Uri 'https://bootstrap.pypa.io/get-pip.py' -OutFile '%GETPIP_SCRIPT%' } catch { exit 1 }" >nul
if errorlevel 1 (
    echo Failed to download get-pip.py.
    exit /b 1
)
"%TARGET_PY%" "%GETPIP_SCRIPT%" >nul 2>nul
set "PIP_RC=%errorlevel%"
if exist "%GETPIP_SCRIPT%" del "%GETPIP_SCRIPT%" >nul 2>nul
exit /b %PIP_RC%

rem =============================================================
:launch_app
if not exist "%APP_ENTRY%" (
    echo Application entry point missing: %APP_ENTRY%
    exit /b 1
)
if defined USE_VENV (
    set "PATH=%VENV_DIR%\Scripts;%PATH%"
) else (
    set "PATH=%PY_HOME%;%PY_HOME%\Scripts;%PATH%"
)
echo Launching AI Agent...
"%APP_PY%" "%APP_ENTRY%"
set "APP_RC=%errorlevel%"
if not "%APP_RC%"=="0" (
    echo Application exited with code %APP_RC%.
    exit /b %APP_RC%
)
exit /b 0

rem =============================================================
:fail
endlocal
exit /b 1

@echo off
rem Ensure command extensions stay enabled so label calls and conditional operators
rem like "||" work reliably even if the host shell disables them by default.
setlocal EnableExtensions EnableDelayedExpansion

rem =============================================================
rem AI Agent - Single entry installer/launcher
rem -------------------------------------------------------------
rem * Drop this file anywhere and run it.
rem * Installs everything into subfolders next to this file.
rem * Handles code download/update, Python runtime, dependencies,
rem   Tesseract OCR, and launches the app automatically.
rem * If an update is applied, the script restarts itself to
rem   finish with the latest version.
rem =============================================================

if "%~1"=="/relaunch" (
    set "AGENT_RELAUNCHED=1"
) else (
    set "AGENT_RELAUNCHED="
)

set "BASE_DIR=%~dp0"
if "%BASE_DIR:~-1%"=="\\" set "BASE_DIR=%BASE_DIR:~0,-1%"
set "INSTALL_ROOT=%BASE_DIR%\ai_agent"
set "SRC_DIR=%INSTALL_ROOT%\source"
set "PY_HOME=%INSTALL_ROOT%\.python"
set "VENV_DIR=%INSTALL_ROOT%\.venv"
set "VENV_PY=%VENV_DIR%\Scripts\python.exe"
set "APP_PY="
set "APP_ENTRY=%SRC_DIR%\app\main.py"
set "REQ_FILE=%SRC_DIR%\requirements.txt"
set "GIT_REMOTE=https://github.com/NIAENGD/AI_Agent.git"
set "GIT_BRANCH=main"
set "REMOTE_ZIP=https://github.com/NIAENGD/AI_Agent/archive/refs/heads/main.zip"
set "VERSION_FILE=%INSTALL_ROOT%\.source_version"
set "SELF_PATH=%~f0"
set "PY_VERSION=3.11.9"
set "PY_INSTALLER_URL=https://www.python.org/ftp/python/%PY_VERSION%/python-%PY_VERSION%-amd64.exe"
set "PY_EMBED_ZIP_URL=https://www.python.org/ftp/python/%PY_VERSION%/python-%PY_VERSION%-embed-amd64.zip"
set "TESS_HOME=%INSTALL_ROOT%\.tesseract"
set "TESS_EXE=%TESS_HOME%\tesseract.exe"
set "TESS_DEFAULT_64=%ProgramFiles%\Tesseract-OCR\tesseract.exe"
set "TESS_DEFAULT_32=%ProgramFiles(x86)%\Tesseract-OCR\tesseract.exe"
set "TESS_DL_URL=https://github.com/UB-Mannheim/tesseract/releases/latest/download/tesseract-ocr-w64-setup.exe"
set "TESS_DL_FALLBACK=https://digi.bib.uni-mannheim.de/tesseract/tesseract-ocr-w64-setup.exe"

call :print_banner
call :ensure_directories || goto :fail
call :sync_source || goto :fail
if defined NEED_RESTART (
    call :restart_if_needed
    endlocal
    exit /b 0
)
call :bootstrap_python || goto :fail
call :bootstrap_dependencies || goto :fail
call :run_app
endlocal
exit /b %errorlevel%

:print_banner
echo ==================================================
echo  AI Agent installer and launcher
echo  Install root: %INSTALL_ROOT%
echo ==================================================
exit /b 0

:ensure_directories
if not exist "%INSTALL_ROOT%" mkdir "%INSTALL_ROOT%" >nul 2>nul
if not exist "%INSTALL_ROOT%" (
    echo Failed to create installation directory at %INSTALL_ROOT%.
    exit /b 1
)
exit /b 0

:sync_source
set "UPDATED="
set "NEED_RESTART="
set "GIT_AVAILABLE="
where git >nul 2>nul && set "GIT_AVAILABLE=1"

call :get_remote_sha
set "REMOTE_SHA=%RETURN_VALUE%"
call :get_local_sha
set "LOCAL_SHA=%RETURN_VALUE%"

if defined GIT_AVAILABLE if exist "%SRC_DIR%\.git" (
    call :git_update_repo
    goto :post_sync
)

if defined GIT_AVAILABLE if not exist "%SRC_DIR%\.git" (
    call :git_clone_repo
    goto :post_sync
)

call :zip_sync_repo

:post_sync
call :refresh_self
if defined UPDATED echo Update detected.
exit /b 0

:git_update_repo
pushd "%SRC_DIR%" >nul 2>nul
for /f "delims=" %%h in ('git rev-parse HEAD 2^>nul') do set "OLD_HEAD=%%h"
git fetch --all --prune
if errorlevel 1 (
    echo Git fetch failed. Trying to continue...
)
git pull --ff-only origin %GIT_BRANCH%
set "PULL_RC=%errorlevel%"
for /f "delims=" %%h in ('git rev-parse HEAD 2^>nul') do set "NEW_HEAD=%%h"
popd >nul 2>nul
if "%PULL_RC%"=="0" if not "%OLD_HEAD%"=="%NEW_HEAD%" set "UPDATED=1"
if "%PULL_RC%"=="0" exit /b 0
if not defined NEW_HEAD (
    echo Git pull failed and repository is unusable.
    exit /b 1
)
echo Git pull failed, continuing with existing checkout.
exit /b 0

:git_clone_repo
echo Cloning repository into %SRC_DIR% ...
git clone --branch %GIT_BRANCH% "%GIT_REMOTE%" "%SRC_DIR%"
if errorlevel 1 (
    echo Clone failed. Falling back to zip download.
    call :zip_sync_repo
    exit /b %errorlevel%
)
set "UPDATED=1"
exit /b 0

:zip_sync_repo
echo Using zip download flow.
if not defined REMOTE_SHA (
    echo Unable to determine remote version. Proceeding with download.
)
if exist "%SRC_DIR%" (
    if defined REMOTE_SHA if exist "%VERSION_FILE%" (
        set /p "CACHED_SHA=" < "%VERSION_FILE%"
        if /i "%CACHED_SHA%"=="%REMOTE_SHA%" (
            echo Existing files match latest known version. Skipping download.
            exit /b 0
        )
    )
)
set "ZIP_PATH=%TEMP%\ai_agent_latest.zip"
echo Downloading latest source archive...
powershell -NoLogo -NoProfile -Command "try { $ProgressPreference='SilentlyContinue'; Invoke-WebRequest -UseBasicParsing -Headers @{ 'User-Agent'='ai-agent-bootstrapper' } -Uri '%REMOTE_ZIP%' -OutFile '%ZIP_PATH%' } catch { exit 1 }" >nul
if errorlevel 1 (
    echo Failed to download source archive.
    exit /b 1
)
echo Extracting files...
set "EXTRACT_DIR=%TEMP%\ai_agent_extract"
if exist "%EXTRACT_DIR%" rd /s /q "%EXTRACT_DIR%" >nul 2>nul
powershell -NoLogo -NoProfile -Command "try { Expand-Archive -Path '%ZIP_PATH%' -DestinationPath '%EXTRACT_DIR%' -Force } catch { exit 1 }" >nul
if errorlevel 1 (
    echo Failed to extract archive.
    exit /b 1
)
set "UNPACKED_DIR="
for /d %%d in ("%EXTRACT_DIR%\*") do set "UNPACKED_DIR=%%d"
if not defined UNPACKED_DIR (
    echo Extracted archive is empty.
    exit /b 1
)
if exist "%SRC_DIR%" rd /s /q "%SRC_DIR%" >nul 2>nul
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

:get_remote_sha
set "RETURN_VALUE="
if defined GIT_AVAILABLE (
    for /f "delims=" %%s in ('git ls-remote "%GIT_REMOTE%" %GIT_BRANCH% 2^>nul') do (
        for /f "tokens=1" %%h in ("%%s") do set "RETURN_VALUE=%%h"
    )
)
if not defined RETURN_VALUE (
    powershell -NoLogo -NoProfile -Command "try { $ProgressPreference='SilentlyContinue'; $resp=Invoke-RestMethod -UseBasicParsing -Headers @{ 'User-Agent'='ai-agent-bootstrapper' } 'https://api.github.com/repos/NIAENGD/AI_Agent/commits/%GIT_BRANCH%'; if($resp.sha){ [Console]::WriteLine($resp.sha) } } catch { }" >"%TEMP%\ai_agent_sha.txt"
    if exist "%TEMP%\ai_agent_sha.txt" (
        set /p "RETURN_VALUE=" < "%TEMP%\ai_agent_sha.txt"
        del "%TEMP%\ai_agent_sha.txt" >nul 2>nul
    )
)
exit /b 0

:get_local_sha
set "RETURN_VALUE="
if exist "%SRC_DIR%\.git" (
    for /f "delims=" %%h in ('git -C "%SRC_DIR%" rev-parse HEAD 2^>nul') do set "RETURN_VALUE=%%h"
) else if exist "%VERSION_FILE%" (
    set /p "RETURN_VALUE=" < "%VERSION_FILE%"
)
exit /b 0

:refresh_self
if not exist "%SRC_DIR%\agent.bat" exit /b 0
fc /b "%SELF_PATH%" "%SRC_DIR%\agent.bat" >nul 2>nul
if errorlevel 1 (
    copy /Y "%SRC_DIR%\agent.bat" "%SELF_PATH%" >nul
    set "NEED_RESTART=1"
)
exit /b 0

:restart_if_needed
if defined AGENT_RELAUNCHED exit /b 0
echo Restarting with latest launcher...
start "AI Agent" "%SELF_PATH%" /relaunch
exit /b 0

:bootstrap_python
call :locate_python_exe
if defined PYTHON_CMD (
    "%PYTHON_CMD%" -V >nul 2>nul
    if not errorlevel 1 exit /b 0
    echo Existing Python appears broken. Reinstalling...
    rd /s /q "%PY_HOME%" >nul 2>nul
    set "PYTHON_CMD="
)
if not exist "%PY_HOME%" mkdir "%PY_HOME%" >nul 2>nul
set "PY_INSTALLER=%TEMP%\python-%PY_VERSION%-amd64.exe"
echo Downloading Python %PY_VERSION% ...
powershell -NoLogo -NoProfile -Command "try { $ProgressPreference='SilentlyContinue'; Invoke-WebRequest -UseBasicParsing -Uri '%PY_INSTALLER_URL%' -OutFile '%PY_INSTALLER%' } catch { exit 1 }" >nul
if errorlevel 1 (
    echo Failed to download Python installer.
    exit /b 1
)
echo Installing private Python runtime...
"%PY_INSTALLER%" /quiet InstallAllUsers=0 PrependPath=0 Include_pip=1 TargetDir="%PY_HOME%"
    if errorlevel 1 (
        echo Python installer returned an error.
        exit /b 1
    )
    if exist "%PY_INSTALLER%" del "%PY_INSTALLER%" >nul 2>nul
    call :locate_python_exe
    if not defined PYTHON_CMD (
        echo Python installation did not produce python.exe. Attempting embeddable runtime...
        call :install_embeddable_python || exit /b 1
        call :locate_python_exe
    )
    if not defined PYTHON_CMD (
        echo No usable Python interpreter found after installation.
        exit /b 1
    )
    call :ensure_pip "%PYTHON_CMD%" || exit /b 1
    "%PYTHON_CMD%" -m pip config set global.no-cache-dir true >nul 2>nul
    exit /b 0

:install_embeddable_python
echo Downloading embeddable Python %PY_VERSION% ...
set "PY_EMBED_ZIP=%TEMP%\python-%PY_VERSION%-embed-amd64.zip"
powershell -NoLogo -NoProfile -Command "Invoke-WebRequest -UseBasicParsing -Uri '%PY_EMBED_ZIP_URL%' -OutFile '%PY_EMBED_ZIP%'" >nul
if errorlevel 1 (
    echo Failed to download embeddable Python package.
    exit /b 1
)

if not exist "%PY_HOME%" mkdir "%PY_HOME%" >nul 2>nul
echo Extracting embeddable runtime to %PY_HOME% ...
powershell -NoLogo -NoProfile -Command "Expand-Archive -Path '%PY_EMBED_ZIP%' -DestinationPath '%PY_HOME%' -Force" >nul
if errorlevel 1 (
    echo Failed to extract embeddable Python runtime.
    if exist "%PY_EMBED_ZIP%" del "%PY_EMBED_ZIP%" >nul 2>nul
    exit /b 1
)
if exist "%PY_EMBED_ZIP%" del "%PY_EMBED_ZIP%" >nul 2>nul

if not exist "%PY_HOME%\python.exe" (
    echo Embeddable runtime did not provide python.exe.
    exit /b 1
)

set "PY_TAG="
for /f "tokens=1-2 delims=." %%a in ("%PY_VERSION%") do (
    set "PY_TAG=%%a%%b"
)
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

:locate_python_exe
set "PYTHON_CMD="
for %%p in ("%PY_HOME%\python.exe" "%PY_HOME%\python3.exe" "%PY_HOME%\pythonw.exe") do (
    if exist %%~p if not defined PYTHON_CMD set "PYTHON_CMD=%%~fp"
)
exit /b 0

:bootstrap_dependencies
    if not exist "%SRC_DIR%" (
        echo Source directory missing at %SRC_DIR%.
        exit /b 1
    )
    if not defined PYTHON_CMD call :locate_python_exe
    if not defined PYTHON_CMD (
        echo Python interpreter unavailable; bootstrap failed.
        exit /b 1
    )
    set "USE_VENV=1"
    "%PYTHON_CMD%" -c "import venv" >nul 2>nul || set "USE_VENV="

if defined USE_VENV if not exist "%VENV_DIR%" (
    echo Creating virtual environment...
    "%PY_HOME%\python.exe" -m venv "%VENV_DIR%" >nul 2>nul
)
if defined USE_VENV if not exist "%VENV_PY%" (
    echo Virtual environment looks broken.
    rd /s /q "%VENV_DIR%" >nul 2>nul
    "%PY_HOME%\python.exe" -m venv "%VENV_DIR%" >nul 2>nul
)

if defined USE_VENV (
    set "APP_PY=%VENV_PY%"
    "%APP_PY%" -m pip install --upgrade pip >nul
    if errorlevel 1 (
        echo Failed to upgrade pip inside virtual environment.
        exit /b 1
    )
) else (
        set "APP_PY=%PYTHON_CMD%"
        call :ensure_pip "%APP_PY%" || exit /b 1
    )
if exist "%REQ_FILE%" (
    echo Installing Python dependencies...
    "%APP_PY%" -m pip install -r "%REQ_FILE%"
    if errorlevel 1 (
        echo Dependency installation failed.
        exit /b 1
    )
)
call :ensure_tesseract || exit /b 1
exit /b 0

:ensure_tesseract
set "TESS_CMD="
if exist "%TESS_EXE%" set "TESS_CMD=%TESS_EXE%"
if not defined TESS_CMD if exist "%TESS_DEFAULT_64%" set "TESS_CMD=%TESS_DEFAULT_64%"
if not defined TESS_CMD if exist "%TESS_DEFAULT_32%" set "TESS_CMD=%TESS_DEFAULT_32%"
if not defined TESS_CMD (
    for /f "delims=" %%t in ('where tesseract 2^>nul') do if not defined TESS_CMD set "TESS_CMD=%%t"
)
if defined TESS_CMD (
    for %%p in ("%TESS_CMD%") do set "TESS_HOME=%%~dpf"
    exit /b 0
)
echo Installing Tesseract OCR locally...
set "TESS_INSTALLER=%TEMP%\tesseract-ocr-w64-setup.exe"
powershell -NoLogo -NoProfile -Command "try { $ProgressPreference='SilentlyContinue'; Invoke-WebRequest -UseBasicParsing -Headers @{ 'User-Agent'='ai-agent-bootstrapper' } -Uri '%TESS_DL_URL%' -OutFile '%TESS_INSTALLER%' } catch { exit 1 }" >nul
if errorlevel 1 (
    powershell -NoLogo -NoProfile -Command "try { $ProgressPreference='SilentlyContinue'; Invoke-WebRequest -UseBasicParsing -Headers @{ 'User-Agent'='ai-agent-bootstrapper' } -Uri '%TESS_DL_FALLBACK%' -OutFile '%TESS_INSTALLER%' } catch { exit 1 }" >nul
)
if errorlevel 1 (
    echo Failed to download Tesseract installer.
    exit /b 1
)
if not exist "%TESS_HOME%" mkdir "%TESS_HOME%" >nul 2>nul
"%TESS_INSTALLER%" /quiet /DIR="%TESS_HOME%"
if errorlevel 1 (
    echo Tesseract installer reported an error.
    exit /b 1
)
if exist "%TESS_INSTALLER%" del "%TESS_INSTALLER%" >nul 2>nul
if not exist "%TESS_EXE%" (
    echo Tesseract installation completed but executable not found.
    exit /b 1
)
set "TESS_CMD=%TESS_EXE%"
exit /b 0

:run_app
if not exist "%APP_ENTRY%" (
    echo Application entry point missing: %APP_ENTRY%
    exit /b 1
)
    if defined USE_VENV (
        set "PATH=%TESS_HOME%;%VENV_DIR%\Scripts;%PATH%"
    ) else (
        set "PATH=%TESS_HOME%;%PY_HOME%;%PY_HOME%\Scripts;%PATH%"
    )
echo Launching AI Agent...
"%APP_PY%" "%APP_ENTRY%"
exit /b %errorlevel%

:fail
endlocal
exit /b 1

:ensure_pip
set "ENSURE_PY=%~1"
if not defined ENSURE_PY exit /b 1
"%ENSURE_PY%" -m pip --version >nul 2>nul && exit /b 0
set "GET_PIP=%TEMP%\get-pip.py"
powershell -NoLogo -NoProfile -Command "try { $ProgressPreference='SilentlyContinue'; Invoke-WebRequest -UseBasicParsing -Uri 'https://bootstrap.pypa.io/get-pip.py' -OutFile '%GET_PIP%' } catch { exit 1 }" >nul
if errorlevel 1 (
    echo Failed to download get-pip.py.
    if exist "%GET_PIP%" del "%GET_PIP%" >nul 2>nul
    exit /b 1
)
"%ENSURE_PY%" "%GET_PIP%" >nul
set "RC=%errorlevel%"
if exist "%GET_PIP%" del "%GET_PIP%" >nul 2>nul
if not "%RC%"=="0" (
    echo Failed to install pip.
    exit /b 1
)
exit /b 0

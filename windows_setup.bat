@echo off
setlocal EnableDelayedExpansion

rem =============================================================
rem AI Agent - Windows helper script (all-in-one bootstrapper)
rem -------------------------------------------------------------
rem * Installs a private Python runtime into .\.python (no admin)
rem * Creates/refreshes a virtual environment and installs deps
rem * Installs Tesseract OCR into .\.tesseract (no admin)
rem * Can pull the latest code and run the app
rem -------------------------------------------------------------
rem Recommended: Run this script from an elevated prompt for the
rem best installation experience. It will still attempt user
rem installations when admin rights are unavailable.
rem =============================================================

rem Resolve project root without a trailing backslash to avoid escaping issues
set "PROJECT_ROOT=%~dp0"
if "%PROJECT_ROOT:~-1%"=="\\" set "PROJECT_ROOT=%PROJECT_ROOT:~0,-1%"
pushd "%PROJECT_ROOT%"

set "PY_VERSION=3.11.9"
set "PY_INSTALLER_URL=https://www.python.org/ftp/python/%PY_VERSION%/python-%PY_VERSION%-amd64.exe"
set "PYTHON_HOME=%PROJECT_ROOT%\.python"
set "VENV_DIR=.venv"
set "VENV_PYTHON=%VENV_DIR%\Scripts\python.exe"
set "REQUIREMENTS=requirements.txt"
set "APP_ENTRY=app\main.py"
set "GIT_REMOTE=https://github.com/NIAENGD/AI_Agent.git"
set "TESSERACT_INSTALLER_URL=https://github.com/UB-Mannheim/tesseract/releases/latest/download/tesseract-ocr-w64-setup.exe"
set "TESSERACT_INSTALLER_URL_FALLBACK=https://digi.bib.uni-mannheim.de/tesseract/tesseract-ocr-w64-setup.exe"
set "TESSERACT_HOME=%PROJECT_ROOT%\.tesseract"
set "TESSERACT_EXE=%TESSERACT_HOME%\tesseract.exe"
set "TESSERACT_DEFAULT_64=%ProgramFiles%\Tesseract-OCR\tesseract.exe"
set "TESSERACT_DEFAULT_32=%ProgramFiles(x86)%\Tesseract-OCR\tesseract.exe"

call :print_header

:menu
echo.
echo [1] Install local Python runtime into .\\.python
echo [2] Create or refresh virtual environment
echo [3] Install/Update Python dependencies
echo [4] Install/Verify Tesseract OCR
echo [5] Update project from Git (pull)
echo [6] Run AI Agent
echo [7] Full setup (Python + Tesseract + venv + deps + run)
echo [8] Clean __pycache__ folders
echo [0] Exit
set /p choice="Select an option: "
if "%choice%"=="1" call :ensure_local_python & goto menu
if "%choice%"=="2" call :setup_venv & goto menu
if "%choice%"=="3" call :install_requirements & goto menu
if "%choice%"=="4" call :install_tesseract & goto menu
if "%choice%"=="5" call :update_project & goto menu
if "%choice%"=="6" call :run_app & goto menu
if "%choice%"=="7" call :full_setup & goto menu
if "%choice%"=="8" call :clean_pycache & goto menu
if "%choice%"=="0" goto end

echo.
echo Invalid option. Please try again.
goto menu

:print_header
echo ==================================================
echo  AI Agent - Windows helper script (all-in-one)
echo  Location: %PROJECT_ROOT%
echo  Python target: %PYTHON_HOME%
echo ==================================================
exit /b 0

:detect_python
set "PYTHON_CMD="
if exist "%PYTHON_HOME%\python.exe" set "PYTHON_CMD=%PYTHON_HOME%\python.exe"
if not defined PYTHON_CMD (
    where python >nul 2>nul && for /f "delims=" %%p in ('where python') do if not defined PYTHON_CMD set "PYTHON_CMD=%%p"
)
if not defined PYTHON_CMD exit /b 1
exit /b 0

:ensure_local_python
if exist "%PYTHON_HOME%\python.exe" (
    set "PYTHON_CMD=%PYTHON_HOME%\python.exe"
    echo Found project-local Python at %PYTHON_CMD%.
    exit /b 0
)

echo Downloading Python %PY_VERSION% installer...
set "PY_INSTALLER=%TEMP%\python-%PY_VERSION%-amd64.exe"
powershell -NoLogo -NoProfile -Command "Invoke-WebRequest -Uri '%PY_INSTALLER_URL%' -OutFile '%PY_INSTALLER%'" >nul
if errorlevel 1 (
    echo Failed to download Python installer.
    exit /b 1
)

echo Installing private Python to %PYTHON_HOME% ...
"%PY_INSTALLER%" /quiet InstallAllUsers=0 PrependPath=0 Include_pip=1 TargetDir="%PYTHON_HOME%"
if errorlevel 1 (
    echo Python installation failed.
    exit /b 1
)

del "%PY_INSTALLER%" >nul 2>nul
set "PYTHON_CMD=%PYTHON_HOME%\python.exe"
echo Installed Python to %PYTHON_CMD%
%PYTHON_CMD% -m pip config set global.no-cache-dir true >nul 2>nul
exit /b 0

:ensure_python_available
call :ensure_local_python
exit /b %errorlevel%

:setup_venv
call :ensure_python_available || exit /b 1
if not exist "%VENV_DIR%\Scripts\activate.bat" (
    echo Creating virtual environment at "%VENV_DIR%"...
    "%PYTHON_CMD%" -m venv "%VENV_DIR%"
    if errorlevel 1 (
        echo Failed to create virtual environment.
        exit /b 1
    )
) else (
    echo Virtual environment already present at "%VENV_DIR%".
)
echo To activate manually run: "%VENV_DIR%\Scripts\activate.bat"
exit /b 0

:install_requirements
call :setup_venv || exit /b 1
call "%VENV_DIR%\Scripts\activate.bat"
"%VENV_PYTHON%" -m pip install --upgrade pip
if exist "%REQUIREMENTS%" (
    "%VENV_PYTHON%" -m pip install -r "%REQUIREMENTS%"
    if errorlevel 1 exit /b 1
) else (
    echo %REQUIREMENTS% not found. Cannot install dependencies.
    exit /b 1
)
exit /b 0

:install_tesseract
echo Checking for Tesseract OCR...
set "TESSERACT_CMD="
call :detect_tesseract
if not errorlevel 1 (
    echo Found Tesseract at "%TESSERACT_CMD%".
    exit /b 0
)

echo Attempting to install Tesseract automatically (no user input required)...
call :install_tesseract_winget
if not errorlevel 1 (
    call :detect_tesseract
    if not errorlevel 1 (
        echo Installed Tesseract using winget at "%TESSERACT_CMD%".
        exit /b 0
    )
)

call :install_tesseract_portable || exit /b 1
call :detect_tesseract
if not errorlevel 1 (
    echo Installed project-local Tesseract at "%TESSERACT_CMD%".
    exit /b 0
)

echo Tesseract installation completed, but executable was not found.
exit /b 1

:detect_tesseract
set "TESSERACT_CMD="
if exist "%TESSERACT_EXE%" set "TESSERACT_CMD=%TESSERACT_EXE%"
if not defined TESSERACT_CMD (
    if exist "%TESSERACT_DEFAULT_64%" set "TESSERACT_CMD=%TESSERACT_DEFAULT_64%"
)
if not defined TESSERACT_CMD (
    if exist "%TESSERACT_DEFAULT_32%" set "TESSERACT_CMD=%TESSERACT_DEFAULT_32%"
)
if not defined TESSERACT_CMD (
    for /f "delims=" %%t in ('where tesseract 2^>nul') do if not defined TESSERACT_CMD set "TESSERACT_CMD=%%t"
)
if not defined TESSERACT_CMD exit /b 1
for %%p in ("%TESSERACT_CMD%") do set "TESSERACT_HOME=%%~dpf"
exit /b 0

:install_tesseract_winget
winget --version >nul 2>nul
if errorlevel 1 exit /b 1
echo Winget detected. Installing Tesseract from official source...
winget install -e --id Tesseract-OCR.Tesseract --silent --accept-package-agreements --accept-source-agreements
exit /b %errorlevel%

:install_tesseract_portable
echo Downloading portable Tesseract installer (UB Mannheim build)...
set "TESS_INSTALLER=%TEMP%\tesseract-ocr-w64-setup.exe"
powershell -NoLogo -NoProfile -Command "try { $ProgressPreference='SilentlyContinue'; $release=Invoke-RestMethod -UseBasicParsing -Headers @{ 'User-Agent'='AI-Agent-setup' } 'https://api.github.com/repos/UB-Mannheim/tesseract/releases/latest'; $asset=$release.assets | Where-Object { $_.name -like 'tesseract-ocr-w64-setup-*.exe' } | Sort-Object -Property name -Descending | Select-Object -First 1; $url=$asset.browser_download_url; if(-not $url){ $url='%TESSERACT_INSTALLER_URL%' }; Invoke-WebRequest -UseBasicParsing -Headers @{ 'User-Agent'='AI-Agent-setup' } -Uri $url -OutFile '%TESS_INSTALLER%' } catch { exit 1 }" >nul
if errorlevel 1 (
    echo Primary download failed. Trying fallback mirror...
    powershell -NoLogo -NoProfile -Command "try { $ProgressPreference='SilentlyContinue'; Invoke-WebRequest -UseBasicParsing -Headers @{ 'User-Agent'='AI-Agent-setup' } -Uri '%TESSERACT_INSTALLER_URL_FALLBACK%' -OutFile '%TESS_INSTALLER%' } catch { exit 1 }" >nul
)

if errorlevel 1 (
    echo Failed to download Tesseract installer from both primary and fallback URLs.
    exit /b 1
)

if not exist "%TESSERACT_HOME%" mkdir "%TESSERACT_HOME%" >nul 2>nul
echo Running installer silently into %TESSERACT_HOME% ...
"%TESS_INSTALLER%" /quiet /DIR="%TESSERACT_HOME%"
if errorlevel 1 (
    echo Tesseract installation failed.
    exit /b 1
)
del "%TESS_INSTALLER%" >nul 2>nul
exit /b 0

:update_project
git --version >nul 2>nul
if errorlevel 1 (
    echo Git is not available. Install Git to pull updates.
    exit /b 1
)

if not exist "%PROJECT_ROOT%\\.git" (
    echo No git repository found in %PROJECT_ROOT%.
    echo Run windows_install.bat to create a clean installation with update support.
    exit /b 1
)

pushd "%PROJECT_ROOT%" >nul
for /f "delims=" %%s in ('git status --porcelain') do (
    set "GIT_DIRTY=1"
)
if defined GIT_DIRTY (
    echo Working tree is not clean. Skip pulling to avoid overwriting local files.
    echo Please commit, stash, or move the files before retrying the update.
    popd >nul
    exit /b 1
)

git rev-parse --abbrev-ref --symbolic-full-name @{u} >nul 2>nul
if errorlevel 1 (
    echo No upstream configured for current branch. Defaulting to origin/main...
    git pull --ff-only origin main
) else (
    git pull --ff-only
)
set "GIT_STATUS=%errorlevel%"
popd >nul
exit /b %GIT_STATUS%

:run_app
if not exist "%APP_ENTRY%" (
    echo Could not find application entry at %APP_ENTRY%.
    exit /b 1
)
call :install_requirements || exit /b 1
call :install_tesseract || exit /b 1
call "%VENV_DIR%\Scripts\activate.bat"
for %%p in ("%TESSERACT_CMD%") do set "PATH=%%~dpf;%PATH%"
"%VENV_PYTHON%" "%APP_ENTRY%"
exit /b %errorlevel%

:full_setup
echo Running full setup...
call :ensure_local_python || exit /b 1
call :install_tesseract || exit /b 1
call :install_requirements || exit /b 1
call :run_app
exit /b %errorlevel%

:clean_pycache
echo Removing __pycache__ folders...
for /d /r %%d in ("__pycache__") do if exist "%%d" rd /s /q "%%d"
echo Done.
exit /b 0

:end
popd
endlocal
exit /b 0

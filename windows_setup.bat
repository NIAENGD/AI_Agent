@echo off
setlocal EnableDelayedExpansion

rem =============================================================
rem AI Agent - Windows helper script (all-in-one bootstrapper)
rem -------------------------------------------------------------
rem * Installs a private Python runtime into .\.python (no admin)
rem * Creates/refreshes a virtual environment and installs deps
rem * Installs Tesseract OCR (winget/choco/installer fallback)
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
set "REQUIREMENTS=requirements.txt"
set "APP_ENTRY=app\main.py"
set "GIT_REMOTE=https://github.com/NIAENGD/AI_Agent.git"
set "TESSERACT_INSTALLER_URL=https://digi.bib.uni-mannheim.de/tesseract/tesseract-ocr-w64-setup-5.3.0.20230401.exe"
set "TESSERACT_EXE=%ProgramFiles%\Tesseract-OCR\tesseract.exe"
set "TESSERACT_EXE_X86=%ProgramFiles(x86)%\Tesseract-OCR\tesseract.exe"

call :print_header
call :enforce_startup_update || goto end

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

:check_internet
powershell -NoLogo -NoProfile -Command "if (Test-Connection -Quiet -Count 1 8.8.8.8) { exit 0 } else { exit 1 }" >nul
exit /b %errorlevel%

:enforce_startup_update
echo Verifying internet connectivity for mandatory update...
call :check_internet
if errorlevel 1 (
    echo No internet connection detected. Exiting to enforce mandatory updates.
    exit /b 1
)
echo Internet connection detected. Pulling latest changes...
call :update_project
if errorlevel 1 (
    echo Git update failed. Resolve the issue and rerun the script.
    exit /b 1
)
echo Repository is up to date.
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
call :detect_python
if errorlevel 1 (
    echo Python not found. Attempting local installation...
    call :ensure_local_python || exit /b 1
)
exit /b 0

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
"%PYTHON_CMD%" -m pip install --upgrade pip
if exist "%REQUIREMENTS%" (
    "%PYTHON_CMD%" -m pip install -r "%REQUIREMENTS%"
    if errorlevel 1 exit /b 1
) else (
    echo %REQUIREMENTS% not found. Cannot install dependencies.
    exit /b 1
)
exit /b 0

:install_tesseract
echo Checking for Tesseract OCR...
if exist "%TESSERACT_EXE%" (
    echo Found at "%TESSERACT_EXE%".
    exit /b 0
)
if exist "%TESSERACT_EXE_X86%" (
    echo Found at "%TESSERACT_EXE_X86%".
    exit /b 0
)

where tesseract >nul 2>nul
if not errorlevel 1 (
    echo Tesseract already available on PATH.
    exit /b 0
)

echo Attempting installation via winget...
winget --version >nul 2>nul
if not errorlevel 1 (
    winget install -e --id UB-Mannheim.TesseractOCR --accept-package-agreements --accept-source-agreements
    if not errorlevel 1 exit /b 0
)

echo Attempting installation via Chocolatey...
choco -v >nul 2>nul
if not errorlevel 1 (
    choco install tesseract --yes
    if not errorlevel 1 exit /b 0
)

echo Downloading Tesseract installer (UB Mannheim build)...
set "TESS_INSTALLER=%TEMP%\tesseract-ocr-w64-setup.exe"
powershell -NoLogo -NoProfile -Command "Invoke-WebRequest -Uri '%TESSERACT_INSTALLER_URL%' -OutFile '%TESS_INSTALLER%'" >nul
if errorlevel 1 (
    echo Failed to download Tesseract installer.
    exit /b 1
)

echo Running installer silently...
"%TESS_INSTALLER%" /quiet
if errorlevel 1 (
    echo Tesseract installation failed.
    exit /b 1
)
del "%TESS_INSTALLER%" >nul 2>nul

if exist "%TESSERACT_EXE%" (
    echo Installed Tesseract at "%TESSERACT_EXE%".
    exit /b 0
)
if exist "%TESSERACT_EXE_X86%" (
    echo Installed Tesseract at "%TESSERACT_EXE_X86%".
    exit /b 0
)

echo Tesseract installation completed. You may need to restart the terminal to refresh PATH.
exit /b 0

:update_project
git --version >nul 2>nul
if errorlevel 1 (
    echo Git is not available. Install Git to pull updates.
    exit /b 1
)

if not exist "%PROJECT_ROOT%\\.git" (
    call :workdir_ready_for_clone
    if errorlevel 1 (
        echo No git repository found and this folder already has files.
        echo For auto-updates, place ONLY windows_setup.bat in an empty folder and rerun.
        echo The script will clone the latest project automatically.
        exit /b 1
    )

    echo No git repository detected. Bootstrapping from %GIT_REMOTE% ...
    pushd "%PROJECT_ROOT%" >nul
    git init
    git remote add origin "%GIT_REMOTE%" >nul 2>nul
    git fetch origin
    git checkout -B main origin/main
    set "GIT_STATUS=%errorlevel%"
    popd >nul
    exit /b %GIT_STATUS%
)

pushd "%PROJECT_ROOT%" >nul
git rev-parse --abbrev-ref --symbolic-full-name @{u} >nul 2>nul
if errorlevel 1 (
    echo No upstream configured for current branch. Defaulting to origin/main...
    git pull origin main
) else (
    git pull
)
set "GIT_STATUS=%errorlevel%"
popd >nul
exit /b %GIT_STATUS%

:workdir_ready_for_clone
set "HAS_FOREIGN_FILES=0"
for /f "delims=" %%f in ('dir /b') do (
    if /i not "%%f"=="windows_setup.bat" if /i not "%%f"=="." if /i not "%%f"==".." (
        set "HAS_FOREIGN_FILES=1"
    )
)
if %HAS_FOREIGN_FILES%==0 ( exit /b 0 ) else ( exit /b 1 )

:run_app
if not exist "%APP_ENTRY%" (
    echo Could not find application entry at %APP_ENTRY%.
    exit /b 1
)
call :install_requirements || exit /b 1
call "%VENV_DIR%\Scripts\activate.bat"
"%PYTHON_CMD%" "%APP_ENTRY%"
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

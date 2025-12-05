@echo off
setlocal EnableDelayedExpansion

set "PROJECT_ROOT=%~dp0"
pushd "%PROJECT_ROOT%"

set "VENV_DIR=.venv"
set "REQUIREMENTS=requirements.txt"
set "PYTHON_CMD=python"
set "APP_ENTRY=app\main.py"

call :print_header

:menu
echo.
echo [1] Create or refresh virtual environment
echo [2] Install/Update Python dependencies
echo [3] Update project from Git (pull)
echo [4] Run AI Agent
echo [5] Full setup (venv + deps + run)
echo [6] Clean __pycache__ folders
echo [0] Exit
set /p choice="Select an option: "
if "%choice%"=="1" call :setup_venv & goto menu
if "%choice%"=="2" call :install_requirements & goto menu
if "%choice%"=="3" call :update_project & goto menu
if "%choice%"=="4" call :run_app & goto menu
if "%choice%"=="5" call :install_requirements && call :run_app & goto menu
if "%choice%"=="6" call :clean_pycache & goto menu
if "%choice%"=="0" goto end

echo.
echo Invalid option. Please try again.
goto menu

:print_header
echo ==================================================
echo  AI Agent - Windows helper script (Phase 2)
echo  Location: %PROJECT_ROOT%
echo ==================================================
exit /b 0

:ensure_python
where %PYTHON_CMD% >nul 2>nul
if errorlevel 1 (
    echo Python 3.10+ is required but not found on PATH.
    echo Install Python from https://www.python.org/downloads/ and try again.
    exit /b 1
)
for /f "tokens=2 delims= " %%v in ('%PYTHON_CMD% -V 2^>^&1') do set "PYTHON_VERSION=%%v"
for /f "tokens=1,2 delims=." %%m in ("!PYTHON_VERSION!") do (
    set "PY_MAJOR=%%m"
    set "PY_MINOR=%%n"
)
if not defined PY_MAJOR exit /b 0
if !PY_MAJOR! LSS 3 (
    echo Python 3.10+ is required. Found version !PYTHON_VERSION!.
    exit /b 1
)
if !PY_MAJOR! EQU 3 if !PY_MINOR! LSS 10 (
    echo Python 3.10+ is required. Found version !PYTHON_VERSION!.
    exit /b 1
)
exit /b 0

:setup_venv
call :ensure_python || exit /b 1
if not exist "%VENV_DIR%\Scripts\activate.bat" (
    echo Creating virtual environment at "%VENV_DIR%"...
    %PYTHON_CMD% -m venv "%VENV_DIR%"
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
%PYTHON_CMD% -m pip install --upgrade pip
if exist "%REQUIREMENTS%" (
    %PYTHON_CMD% -m pip install -r "%REQUIREMENTS%"
    if errorlevel 1 exit /b 1
) else (
    echo %REQUIREMENTS% not found. Cannot install dependencies.
    exit /b 1
)
exit /b 0

:update_project
git --version >nul 2>nul
if errorlevel 1 (
    echo Git is not available. Install Git to pull updates.
    exit /b 1
)
git -C "%PROJECT_ROOT%" pull
exit /b %errorlevel%

:run_app
if not exist "%APP_ENTRY%" (
    echo Could not find application entry at %APP_ENTRY%.
    exit /b 1
)
call :setup_venv || exit /b 1
call "%VENV_DIR%\Scripts\activate.bat"
if not exist "%REQUIREMENTS%" (
    echo %REQUIREMENTS% missing. Please ensure dependencies list is present.
    exit /b 1
)
for /f "tokens=1" %%i in ('%PYTHON_CMD% -m pip show PyQt5 2^>^&1 ^| find /c "Name"') do set "HAS_PYQT=%%i"
if "!HAS_PYQT!"=="0" (
    echo Installing dependencies before launch...
    call :install_requirements || exit /b 1
)
%PYTHON_CMD% "%APP_ENTRY%"
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

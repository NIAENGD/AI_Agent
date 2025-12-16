@echo off
setlocal EnableDelayedExpansion

rem =============================================================
rem AI Agent - Windows installer
rem -------------------------------------------------------------
rem * Can be run from any folder
rem * Prompts for an installation directory via a folder picker
rem * Clones or refreshes the repository in the chosen location
rem * Leaves the windows_setup.bat inside the install folder for
rem   updates, dependency management, and app launch
rem =============================================================

set "GIT_REMOTE=https://github.com/NIAENGD/AI_Agent.git"
set "INSTALL_DIR="

for /f "usebackq delims=" %%i in (`powershell -NoLogo -NoProfile -Sta -Command "Add-Type -AssemblyName System.Windows.Forms; [System.Windows.Forms.Application]::EnableVisualStyles(); $dlg=New-Object System.Windows.Forms.FolderBrowserDialog; $dlg.Description='Choose where AI Agent should be installed'; $dlg.ShowNewFolderButton=$true; $dlg.SelectedPath=$env:USERPROFILE; if($dlg.ShowDialog() -eq 'OK'){[Console]::WriteLine($dlg.SelectedPath)}"`) do set "BASE_DIR=%%i"

if not defined BASE_DIR (
    echo No folder selected. Installation cancelled.
    exit /b 1
)

set "INSTALL_DIR=%BASE_DIR%\AI_Agent"
echo.
echo Target installation: %INSTALL_DIR%

git --version >nul 2>nul
if errorlevel 1 (
    echo Git is not available. Please install Git and rerun this installer.
    exit /b 1
)

if exist "%INSTALL_DIR%\.git" (
    echo Existing AI Agent installation detected. Checking for updates...
    pushd "%INSTALL_DIR%" >nul
    for /f "delims=" %%s in ('git status --porcelain') do (
        set "GIT_DIRTY=1"
    )
    if defined GIT_DIRTY (
        echo Repository has local changes. Skipping pull to avoid overwriting them.
        echo Use windows_setup.bat inside the installation folder to manage updates once clean.
        popd >nul
        exit /b 0
    )

    git pull --ff-only
    if errorlevel 1 (
        echo Git pull failed. Please resolve the repository state and try again.
        popd >nul
        exit /b 1
    )
    popd >nul
    echo Installation refreshed successfully.
    goto finished
)

if exist "%INSTALL_DIR%" (
    echo The target folder already exists but is not a Git checkout.
    echo Please choose an empty folder or remove the existing files.
    exit /b 1
)

echo Cloning AI Agent into %INSTALL_DIR% ...
git clone "%GIT_REMOTE%" "%INSTALL_DIR%"
if errorlevel 1 (
    echo Clone failed. Check your network connection and try again.
    exit /b 1
)

echo Clone completed.

echo.
echo Next steps:
echo  - Open "%INSTALL_DIR%\windows_setup.bat" for updates, dependencies, and running the app.
echo  - The maintenance script will avoid pulling if local files would be overwritten.

goto finished

:finished
echo.
echo Installer finished.
endlocal
exit /b 0

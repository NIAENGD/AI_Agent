param(
    [switch]$Relaunch
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# =============================================================
# AI Agent - PowerShell installer/launcher
# -------------------------------------------------------------
# * Drop this file anywhere and run it.
# * Installs everything into subfolders next to this file.
# * Handles code download/update, Python runtime, dependencies,
#   Tesseract OCR, and launches the app automatically.
# * If an update is applied, the script restarts itself to
#   finish with the latest version.
# =============================================================

$BaseDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
$InstallRoot = Join-Path $BaseDir 'ai_agent'
$SourceDir = Join-Path $InstallRoot 'source'
$PythonHome = Join-Path $InstallRoot '.python'
$VenvDir = Join-Path $InstallRoot '.venv'
$VenvPython = Join-Path $VenvDir 'Scripts\python.exe'
$AppEntry = Join-Path $SourceDir 'app/main.py'
$Requirements = Join-Path $SourceDir 'requirements.txt'
$GitRemote = 'https://github.com/NIAENGD/AI_Agent.git'
$GitBranch = 'main'
$RemoteZip = 'https://github.com/NIAENGD/AI_Agent/archive/refs/heads/main.zip'
$VersionFile = Join-Path $InstallRoot '.source_version'
$SelfPath = $MyInvocation.MyCommand.Definition
$PythonVersion = '3.11.9'
$PythonInstallerUrl = "https://www.python.org/ftp/python/$PythonVersion/python-$PythonVersion-amd64.exe"
$PythonEmbedUrl = "https://www.python.org/ftp/python/$PythonVersion/python-$PythonVersion-embed-amd64.zip"
$TessHome = Join-Path $InstallRoot '.tesseract'
$TessExe = Join-Path $TessHome 'tesseract.exe'
$TessDefault64 = Join-Path ${env:ProgramFiles} 'Tesseract-OCR/tesseract.exe'
$TessDefault32 = Join-Path ${env:'ProgramFiles(x86)'} 'Tesseract-OCR/tesseract.exe'
$TessDlUrl = 'https://github.com/UB-Mannheim/tesseract/releases/latest/download/tesseract-ocr-w64-setup.exe'
$TessDlFallback = 'https://digi.bib.uni-mannheim.de/tesseract/tesseract-ocr-w64-setup.exe'

$global:Updated = $false
$global:NeedRestart = $false
$global:PythonCmd = $null

Write-Host '=================================================='
Write-Host ' AI Agent installer and launcher'
Write-Host " Install root: $InstallRoot"
Write-Host '=================================================='

function Ensure-Directory {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Failed to create directory at $Path."
    }
}

function Get-GitAvailable {
    return [bool](Get-Command git -ErrorAction SilentlyContinue)
}

function Get-RemoteSha {
    param([bool]$GitAvailable)
    $remoteSha = $null
    if ($GitAvailable) {
        try {
            $remoteSha = (git ls-remote $GitRemote $GitBranch 2>$null | Select-Object -First 1).Split()[0]
        } catch {}
    }
    if (-not $remoteSha) {
        try {
            $resp = Invoke-RestMethod -Headers @{ 'User-Agent'='ai-agent-bootstrapper' } "https://api.github.com/repos/NIAENGD/AI_Agent/commits/$GitBranch"
            if ($resp.sha) { $remoteSha = $resp.sha }
        } catch {}
    }
    return $remoteSha
}

function Get-LocalSha {
    if (Test-Path -LiteralPath (Join-Path $SourceDir '.git')) {
        try { return (git -C $SourceDir rev-parse HEAD 2>$null) } catch { return $null }
    }
    if (Test-Path -LiteralPath $VersionFile) {
        return (Get-Content -LiteralPath $VersionFile -ErrorAction SilentlyContinue | Select-Object -First 1)
    }
    return $null
}

function Wipe-ExistingSource {
    if (Test-Path -LiteralPath $SourceDir) {
        Remove-Item -LiteralPath $SourceDir -Recurse -Force -ErrorAction SilentlyContinue
        if (Test-Path -LiteralPath $SourceDir) {
            throw 'Failed to clear existing source directory.'
        }
    }
}

function Git-UpdateRepo {
    $oldHead = git -C $SourceDir rev-parse HEAD 2>$null
    git -C $SourceDir fetch --all --prune
    $pull = git -C $SourceDir pull --ff-only origin $GitBranch
    if ($LASTEXITCODE -ne 0 -and -not (git -C $SourceDir rev-parse HEAD 2>$null)) {
        throw 'Git pull failed and repository is unusable.'
    }
    $newHead = git -C $SourceDir rev-parse HEAD 2>$null
    if ($LASTEXITCODE -eq 0 -and $oldHead -ne $newHead) { $global:Updated = $true }
}

function Git-CloneRepo {
    if (Test-Path -LiteralPath $SourceDir -and -not (Test-Path -LiteralPath (Join-Path $SourceDir '.git'))) {
        Wipe-ExistingSource
    }
    Write-Host "Cloning repository into $SourceDir ..."
    git clone --branch $GitBranch $GitRemote $SourceDir
    if ($LASTEXITCODE -ne 0) {
        Write-Warning 'Clone failed. Falling back to zip download.'
        Zip-SyncRepo
        return
    }
    $global:Updated = $true
}

function Zip-SyncRepo {
    Write-Host 'Using zip download flow.'
    $remoteSha = Get-RemoteSha -GitAvailable:(Get-GitAvailable)
    if ($remoteSha -and (Test-Path -LiteralPath $VersionFile)) {
        $cached = Get-Content -LiteralPath $VersionFile -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($cached -and ($cached -ieq $remoteSha)) {
            Write-Host 'Existing files match latest known version. Skipping download.'
            return
        }
    }

    $zipPath = Join-Path $env:TEMP 'ai_agent_latest.zip'
    Write-Host 'Downloading latest source archive...'
    Invoke-WebRequest -Headers @{ 'User-Agent'='ai-agent-bootstrapper' } -Uri $RemoteZip -OutFile $zipPath -UseBasicParsing

    $extractDir = Join-Path $env:TEMP 'ai_agent_extract'
    if (Test-Path -LiteralPath $extractDir) { Remove-Item -LiteralPath $extractDir -Recurse -Force -ErrorAction SilentlyContinue }
    Write-Host 'Extracting files...'
    Expand-Archive -Path $zipPath -DestinationPath $extractDir -Force

    $unpackedDir = Get-ChildItem -LiteralPath $extractDir | Select-Object -First 1
    if (-not $unpackedDir) { throw 'Extracted archive is empty.' }

    Wipe-ExistingSource
    Copy-Item -LiteralPath $unpackedDir.FullName -Destination $SourceDir -Recurse -Force

    Remove-Item -LiteralPath $zipPath -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $extractDir -Recurse -Force -ErrorAction SilentlyContinue

    $global:Updated = $true
    if ($remoteSha) { Set-Content -LiteralPath $VersionFile -Value $remoteSha }
}

function Sync-Source {
    $gitAvailable = Get-GitAvailable
    $remoteSha = Get-RemoteSha -GitAvailable:$gitAvailable
    $localSha = Get-LocalSha

    if ($gitAvailable -and (Test-Path -LiteralPath (Join-Path $SourceDir '.git'))) {
        try { Git-UpdateRepo } catch { Write-Warning $_ }
    } elseif ($gitAvailable) {
        Git-CloneRepo
    } else {
        Zip-SyncRepo
    }

    Refresh-Self
    if ($global:Updated) { Write-Host 'Update detected.' }
}

function Refresh-Self {
    $sourceScript = Join-Path $SourceDir 'agent.ps1'
    if (-not (Test-Path -LiteralPath $sourceScript)) { return }
    try {
        $currentHash = Get-FileHash -LiteralPath $SelfPath -Algorithm SHA256
        $sourceHash = Get-FileHash -LiteralPath $sourceScript -Algorithm SHA256
        if ($currentHash.Hash -ne $sourceHash.Hash) {
            Copy-Item -LiteralPath $sourceScript -Destination $SelfPath -Force
            $global:NeedRestart = $true
        }
    } catch {
        # If hashing fails, perform a best-effort overwrite.
        Copy-Item -LiteralPath $sourceScript -Destination $SelfPath -Force -ErrorAction SilentlyContinue
        $global:NeedRestart = $true
    }
}

function Restart-IfNeeded {
    if ($Relaunch) { return }
    if (-not $global:NeedRestart) { return }
    Write-Host 'Restarting with latest launcher...'
    & powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File $SelfPath -Relaunch
    exit $LASTEXITCODE
}

function Locate-PythonExe {
    $candidates = @(
        Join-Path $PythonHome 'python.exe'
        Join-Path $PythonHome 'python3.exe'
        Join-Path $PythonHome 'pythonw.exe'
    )
    foreach ($c in $candidates) {
        if (Test-Path -LiteralPath $c) { return $c }
    }
    return $null
}

function Install-EmbeddablePython {
    Write-Host "Downloading embeddable Python $PythonVersion ..."
    $embedZip = Join-Path $env:TEMP "python-$PythonVersion-embed-amd64.zip"
    Invoke-WebRequest -Uri $PythonEmbedUrl -OutFile $embedZip -UseBasicParsing
    Ensure-Directory -Path $PythonHome
    Write-Host "Extracting embeddable runtime to $PythonHome ..."
    Expand-Archive -Path $embedZip -DestinationPath $PythonHome -Force
    Remove-Item -LiteralPath $embedZip -Force -ErrorAction SilentlyContinue

    $pyExe = Join-Path $PythonHome 'python.exe'
    if (-not (Test-Path -LiteralPath $pyExe)) { throw 'Embeddable runtime did not provide python.exe.' }

    $tag = ($PythonVersion -split '\.')[0..1] -join ''
    $pthPath = Join-Path $PythonHome "python$tag._pth"
    if (Test-Path -LiteralPath $pthPath) {
        @(
            "python$tag.zip"
            '.'
            'Lib'
            'Lib\\site-packages'
            'import site'
        ) | Set-Content -LiteralPath $pthPath
    }
}

function Ensure-Pip {
    param([string]$Python)
    try {
        & $Python -m pip --version | Out-Null
        return
    } catch {}

    $getPip = Join-Path $env:TEMP 'get-pip.py'
    Write-Host 'Installing pip...'
    Invoke-WebRequest -Uri 'https://bootstrap.pypa.io/get-pip.py' -OutFile $getPip -UseBasicParsing
    & $Python $getPip
    Remove-Item -LiteralPath $getPip -Force -ErrorAction SilentlyContinue
}

function Bootstrap-Python {
    $global:PythonCmd = Locate-PythonExe
    if ($global:PythonCmd) {
        try { & $global:PythonCmd -V | Out-Null; return } catch {}
        Write-Warning 'Existing Python appears broken. Reinstalling...'
        Remove-Item -LiteralPath $PythonHome -Recurse -Force -ErrorAction SilentlyContinue
        $global:PythonCmd = $null
    }

    Ensure-Directory -Path $PythonHome
    $installerPath = Join-Path $env:TEMP "python-$PythonVersion-amd64.exe"
    Write-Host "Downloading Python $PythonVersion ..."
    Invoke-WebRequest -Uri $PythonInstallerUrl -OutFile $installerPath -UseBasicParsing
    Write-Host 'Installing private Python runtime...'
    & $installerPath /quiet InstallAllUsers=0 PrependPath=0 Include_pip=1 TargetDir="$PythonHome"
    $exitCode = $LASTEXITCODE
    Remove-Item -LiteralPath $installerPath -Force -ErrorAction SilentlyContinue

    $global:PythonCmd = Locate-PythonExe
    if (-not $global:PythonCmd -or $exitCode -ne 0) {
        Write-Warning 'Standard installer failed, attempting embeddable runtime.'
        Install-EmbeddablePython
        $global:PythonCmd = Locate-PythonExe
    }

    if (-not $global:PythonCmd) { throw 'No usable Python interpreter found after installation.' }
    Ensure-Pip -Python $global:PythonCmd
    & $global:PythonCmd -m pip config set global.no-cache-dir true *> $null
}

function Bootstrap-Dependencies {
    if (-not (Test-Path -LiteralPath $SourceDir)) { throw "Source directory missing at $SourceDir." }
    if (-not $global:PythonCmd) { $global:PythonCmd = Locate-PythonExe }
    if (-not $global:PythonCmd) { throw 'Python interpreter unavailable; bootstrap failed.' }

    $useVenv = $true
    try { & $global:PythonCmd -c 'import venv' | Out-Null } catch { $useVenv = $false }

    $appPython = $global:PythonCmd
    if ($useVenv) {
        if (-not (Test-Path -LiteralPath $VenvDir)) {
            Write-Host 'Creating virtual environment...'
            & $global:PythonCmd -m venv $VenvDir *> $null
        }
        if (-not (Test-Path -LiteralPath $VenvPython)) {
            Write-Host 'Virtual environment looks broken. Recreating...'
            Remove-Item -LiteralPath $VenvDir -Recurse -Force -ErrorAction SilentlyContinue
            & $global:PythonCmd -m venv $VenvDir *> $null
        }
        $appPython = $VenvPython
        & $appPython -m pip install --upgrade pip
    } else {
        Ensure-Pip -Python $appPython
    }

    if (Test-Path -LiteralPath $Requirements) {
        Write-Host 'Installing Python dependencies...'
        & $appPython -m pip install -r $Requirements
    }

    $depCheck = "import importlib,sys;mods=['PyQt5','pygetwindow','pyautogui','pytesseract','PIL'];missing=[m for m in mods if importlib.util.find_spec(m) is None];print('Missing modules:'+','.join(missing) if missing else '');sys.exit(1 if missing else 0)"
    try {
        & $appPython -c $depCheck *> $null
    } catch {
        Write-Host 'Dependency check failed; reinstalling requirements...'
        & $appPython -m pip install -r $Requirements
    }

    Ensure-Tesseract
    return $appPython
}

function Ensure-Tesseract {
    $tessCmd = $null
    foreach ($candidate in @($TessExe, $TessDefault64, $TessDefault32)) {
        if ($candidate -and (Test-Path -LiteralPath $candidate)) { $tessCmd = $candidate; break }
    }
    if (-not $tessCmd) {
        $which = Get-Command tesseract -ErrorAction SilentlyContinue
        if ($which) { $tessCmd = $which.Source }
    }
    if ($tessCmd) {
        $global:TessHome = Split-Path -Parent $tessCmd
        return
    }

    Write-Host 'Installing Tesseract OCR locally...'
    $installer = Join-Path $env:TEMP 'tesseract-ocr-w64-setup.exe'
    try {
        Invoke-WebRequest -Headers @{ 'User-Agent'='ai-agent-bootstrapper' } -Uri $TessDlUrl -OutFile $installer -UseBasicParsing
    } catch {
        Invoke-WebRequest -Headers @{ 'User-Agent'='ai-agent-bootstrapper' } -Uri $TessDlFallback -OutFile $installer -UseBasicParsing
    }

    Ensure-Directory -Path $TessHome
    & $installer /quiet /DIR="$TessHome"
    Remove-Item -LiteralPath $installer -Force -ErrorAction SilentlyContinue
    if (-not (Test-Path -LiteralPath $TessExe)) { throw 'Tesseract installation completed but executable not found.' }
}

function Run-App {
    param([string]$PythonExe)
    if (-not (Test-Path -LiteralPath $AppEntry)) { throw "Application entry point missing: $AppEntry" }

    if ($PythonExe -eq $VenvPython) {
        $env:PATH = "$TessHome;$VenvDir\\Scripts;$env:PATH"
    } else {
        $env:PATH = "$TessHome;$PythonHome;$PythonHome\\Scripts;$env:PATH"
    }

    Write-Host 'Launching AI Agent...'
    & $PythonExe $AppEntry
    if ($LASTEXITCODE -ne 0) { throw "Application exited with code $LASTEXITCODE." }
}

try {
    Ensure-Directory -Path $InstallRoot
    Sync-Source
    Restart-IfNeeded
    Bootstrap-Python
    $appPy = Bootstrap-Dependencies
    Run-App -PythonExe $appPy
} catch {
    Write-Error $_
    exit 1
}

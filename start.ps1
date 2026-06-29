param(
    [switch]$DryRun
)

$ErrorActionPreference = 'Stop'

$ProjectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$LogFile = Join-Path $ProjectDir 'startup.log'
$VenvDir = Join-Path $ProjectDir '.venv'
$VenvPython = Join-Path $VenvDir 'Scripts\python.exe'
$PythonInstaller = Join-Path $env:TEMP 'python-3.11.9-amd64.exe'
$RequirementsFile = Join-Path $ProjectDir 'requirements.txt'

$env:PYTHONUTF8 = '1'
$env:PYTHONIOENCODING = 'utf-8'
$env:PIP_DISABLE_PIP_VERSION_CHECK = '1'

function Write-Log {
    param([string]$Message)
    $line = "[{0}] {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message
    Write-Host $line
    Add-Content -Path $LogFile -Value $line -Encoding utf8
}

function New-CommandResult {
    param(
        [int]$ExitCode,
        [object[]]$Output
    )

    $lines = @($Output | Where-Object { $_ -ne $null } | ForEach-Object { $_.ToString() })
    [pscustomobject]@{
        ExitCode = $ExitCode
        Output = $lines
        Text = ($lines -join [Environment]::NewLine).Trim()
    }
}

function Invoke-Native {
    param(
        [Parameter(Mandatory = $true)][string]$FilePath,
        [string[]]$Arguments = @(),
        [switch]$AllowFailure,
        [switch]$SkipLog,
        [string]$Label = $FilePath
    )

    if (-not $FilePath) {
        return New-CommandResult -ExitCode 9009 -Output @('Empty command path.')
    }

    $output = @()
    try {
        $output = & $FilePath @Arguments 2>&1
        $exitCode = $LASTEXITCODE
        if ($null -eq $exitCode) {
            $exitCode = 0
        }
    } catch {
        $exitCode = 1
        $output = @($_.Exception.Message)
    }

    $result = New-CommandResult -ExitCode $exitCode -Output $output
    if (-not $SkipLog -and $result.Text) {
        Add-Content -Path $LogFile -Value $result.Text -Encoding utf8
    }
    if (-not $AllowFailure -and $result.ExitCode -ne 0) {
        throw "$Label failed with exit code $($result.ExitCode)."
    }
    return $result
}

function Test-RealPythonPath {
    param([string]$PythonPath)

    if (-not $PythonPath) { return $false }
    if ($PythonPath -match 'WindowsApps') { return $false }
    if (-not (Test-Path $PythonPath)) { return $false }
    if ((Split-Path $PythonPath -Leaf).ToLowerInvariant() -ne 'python.exe') { return $false }
    return $true
}

function Get-PythonInfo {
    param(
        [Parameter(Mandatory = $true)][string]$FilePath,
        [string[]]$Arguments = @()
    )

    $probeArgs = @()
    if ($Arguments) {
        $probeArgs += $Arguments
    }
    $probeArgs += '-c'
    $probeArgs += "import sys; print(f'{sys.version_info[0]}.{sys.version_info[1]}'); print(sys.executable)"

    $result = Invoke-Native -FilePath $FilePath -Arguments $probeArgs -AllowFailure -SkipLog -Label "Probe $FilePath"
    if ($result.ExitCode -ne 0) { return $null }

    $lines = @($result.Output | Where-Object { $_ -and $_.Trim() })
    if ($lines.Count -lt 2) { return $null }

    $versionText = $lines[-2].Trim()
    $exePath = $lines[-1].Trim()
    if (-not (Test-RealPythonPath $exePath)) { return $null }

    $parts = $versionText.Split('.')
    if ($parts.Count -lt 2) { return $null }

    try {
        $major = [int]$parts[0]
        $minor = [int]$parts[1]
    } catch {
        return $null
    }

    if ($major -lt 3 -or ($major -eq 3 -and $minor -lt 10)) {
        return $null
    }

    [pscustomobject]@{
        Executable = $exePath
        Major = $major
        Minor = $minor
        Score = ($major * 100) + $minor
    }
}

function Get-PythonVersion {
    param([string]$PythonPath)

    if (-not $PythonPath) { return $null }

    $result = Invoke-Native -FilePath $PythonPath -Arguments @(
        '-c',
        "import sys; print(f'{sys.version_info[0]}.{sys.version_info[1]}')"
    ) -AllowFailure -SkipLog -Label "Version check $PythonPath"
    if ($result.ExitCode -ne 0) { return $null }

    $line = @($result.Output | Where-Object { $_ -and $_.Trim() } | Select-Object -Last 1)
    if (-not $line) { return $null }

    $parts = $line[0].Trim().Split('.')
    if ($parts.Count -lt 2) { return $null }

    try {
        [pscustomobject]@{
            Major = [int]$parts[0]
            Minor = [int]$parts[1]
            Score = ([int]$parts[0] * 100) + [int]$parts[1]
        }
    } catch {
        return $null
    }
}

function Find-Python {
    $candidates = New-Object System.Collections.Generic.List[object]

    $pyLauncher = Get-Command py.exe -ErrorAction SilentlyContinue
    if ($pyLauncher) {
        foreach ($launcherArgs in @(@('-3.11'), @('-3'), @())) {
            $info = Get-PythonInfo -FilePath $pyLauncher.Source -Arguments $launcherArgs
            if ($info) { $candidates.Add($info) }
        }
    }

    $pythonCommand = Get-Command python.exe -ErrorAction SilentlyContinue
    if ($pythonCommand) {
        $info = Get-PythonInfo -FilePath $pythonCommand.Source
        if ($info) { $candidates.Add($info) }
    }

    foreach ($candidatePath in @(
        (Join-Path $env:LocalAppData 'Programs\Python\Python311\python.exe'),
        'C:\Program Files\Python311\python.exe',
        'C:\Python311\python.exe',
        'C:\Program Files\Python310\python.exe',
        'C:\Python310\python.exe'
    )) {
        if (Test-Path $candidatePath) {
            $info = Get-PythonInfo -FilePath $candidatePath
            if ($info) { $candidates.Add($info) }
        }
    }

    return $candidates |
        Sort-Object -Property Score -Descending |
        Select-Object -ExpandProperty Executable -First 1
}

function Install-Python {
    Write-Log 'No usable Python found. Trying to install Python 3.11.'

    $winget = Get-Command winget.exe -ErrorAction SilentlyContinue
    if ($winget) {
        Write-Log 'Trying winget first.'
        $result = Invoke-Native -FilePath $winget.Source -Arguments @(
            'install', '--id', 'Python.Python.3.11', '-e',
            '--accept-package-agreements', '--accept-source-agreements', '--scope', 'user'
        ) -AllowFailure -Label 'winget install Python'
        if ($result.ExitCode -eq 0) {
            $found = Find-Python
            if ($found) { return $found }
        }
        Write-Log "winget install failed with exit code $($result.ExitCode)."
    }

    Write-Log 'Downloading official Python installer.'
    Invoke-WebRequest -Uri 'https://www.python.org/ftp/python/3.11.9/python-3.11.9-amd64.exe' -OutFile $PythonInstaller
    $installResult = Invoke-Native -FilePath $PythonInstaller -Arguments @(
        '/quiet', 'InstallAllUsers=0', 'PrependPath=1', 'Include_test=0', 'Include_launcher=1'
    ) -AllowFailure -Label 'python installer'
    if ($installResult.ExitCode -ne 0) {
        throw 'Python auto install failed. Please install Python 3.10+ and retry.'
    }

    $foundAfterInstall = Find-Python
    if ($foundAfterInstall) { return $foundAfterInstall }
    throw 'Python installer finished but no usable python.exe was found.'
}

function Test-PythonExe {
    param([string]$PythonPath)
    if (-not $PythonPath) { return $false }
    if (-not (Test-Path $PythonPath)) { return $false }

    $version = Get-PythonVersion -PythonPath $PythonPath
    if ($null -eq $version) { return $false }
    return $version.Major -gt 3 -or ($version.Major -eq 3 -and $version.Minor -ge 10)
}

function Log-PythonValidationFailure {
    param([string]$PythonPath)

    if (-not $PythonPath) {
        Write-Log 'Python validation failed: empty python path.'
        return
    }

    if (-not (Test-Path $PythonPath)) {
        Write-Log "Python validation failed: file not found: $PythonPath"
        return
    }

    $result = Invoke-Native -FilePath $PythonPath -Arguments @(
        '-c',
        'import sys; print(sys.version); print(sys.executable)'
    ) -AllowFailure -Label "Validation details $PythonPath"
    Write-Log "Python validation exit code: $($result.ExitCode)"
}

function Remove-VenvDirectory {
    if (-not (Test-Path $VenvDir)) { return }
    try {
        Remove-Item -LiteralPath $VenvDir -Recurse -Force -ErrorAction Stop
    } catch {
        throw 'Old .venv could not be removed. Close any server windows and retry.'
    }
}

function Ensure-Venv {
    param([string]$PythonPath)

    $needsRebuild = $true
    if (Test-Path $VenvPython) {
        if (Test-PythonExe $VenvPython) {
            $needsRebuild = $false
        } else {
            Write-Log 'Existing .venv is not usable on this machine. Rebuilding.'
        }
    } elseif (Test-Path $VenvDir) {
        Write-Log 'Broken .venv directory found. Rebuilding.'
    }

    if ($needsRebuild) {
        Remove-VenvDirectory
        Write-Log 'Creating virtual environment.'
        $createResult = Invoke-Native -FilePath $PythonPath -Arguments @('-m', 'venv', $VenvDir) -AllowFailure -Label 'python -m venv'
        if ($createResult.ExitCode -ne 0 -or -not (Test-Path $VenvPython)) {
            throw 'Virtual environment creation failed.'
        }
    }

    if (-not (Test-PythonExe $VenvPython)) {
        Log-PythonValidationFailure -PythonPath $VenvPython
        Write-Log 'New .venv failed validation. Retrying once.'
        Remove-VenvDirectory
        $rebuildResult = Invoke-Native -FilePath $PythonPath -Arguments @('-m', 'venv', $VenvDir) -AllowFailure -Label 'python -m venv retry'
        if ($rebuildResult.ExitCode -ne 0 -or -not (Test-Path $VenvPython)) {
            throw 'Virtual environment is still invalid after rebuild.'
        }
        if (-not (Test-PythonExe $VenvPython)) {
            Log-PythonValidationFailure -PythonPath $VenvPython
            throw 'Virtual environment is still invalid after rebuild.'
        }
    }
}

function Ensure-Pip {
    $pipCheck = Invoke-Native -FilePath $VenvPython -Arguments @('-m', 'pip', '--version') -AllowFailure -SkipLog -Label 'pip --version'
    if ($pipCheck.ExitCode -eq 0) { return }

    Write-Log 'pip is missing. Running ensurepip.'
    $ensurePip = Invoke-Native -FilePath $VenvPython -Arguments @('-m', 'ensurepip', '--upgrade') -AllowFailure -Label 'ensurepip'
    if ($ensurePip.ExitCode -ne 0) {
        throw 'pip bootstrap failed.'
    }
}

function Ensure-Requirements {
    if (-not (Test-Path $RequirementsFile)) {
        throw 'requirements.txt is missing.'
    }

    $importCheck = Invoke-Native -FilePath $VenvPython -Arguments @(
        '-c',
        'import fastapi, uvicorn, httpx, multipart, openpyxl, pypdf, docx, cryptography, lark_oapi'
    ) -AllowFailure -SkipLog -Label 'dependency import check'
    if ($importCheck.ExitCode -eq 0) { return }

    Write-Log 'Dependencies are missing. Installing requirements.txt.'
    $upgradePip = Invoke-Native -FilePath $VenvPython -Arguments @('-m', 'pip', 'install', '--default-timeout', '120', '--upgrade', 'pip') -AllowFailure -Label 'pip upgrade'
    if ($upgradePip.ExitCode -ne 0) {
        throw 'pip upgrade failed.'
    }

    $installReq = Invoke-Native -FilePath $VenvPython -Arguments @('-m', 'pip', 'install', '--default-timeout', '120', '-r', $RequirementsFile) -AllowFailure -Label 'pip install requirements'
    if ($installReq.ExitCode -ne 0) {
        throw 'Dependency install failed. Check network access or pip config.'
    }

    $recheck = Invoke-Native -FilePath $VenvPython -Arguments @(
        '-c',
        'import fastapi, uvicorn, httpx, multipart, openpyxl, pypdf, docx, cryptography, lark_oapi'
    ) -AllowFailure -SkipLog -Label 'dependency recheck'
    if ($recheck.ExitCode -ne 0) {
        throw 'Dependencies still fail after install.'
    }
}

function Start-App {
    Write-Log 'Starting server window.'
    $appPath = Join-Path $ProjectDir 'app.py'
    $command = "& '$VenvPython' '$appPath'"
    Start-Process -FilePath 'powershell.exe' -ArgumentList @('-NoExit', '-ExecutionPolicy', 'Bypass', '-Command', $command) -WorkingDirectory $ProjectDir
    Start-Sleep -Seconds 3
    Start-Process 'http://127.0.0.1:8008'
    Write-Log 'Startup complete.'
}

if (Test-Path $LogFile) {
    try {
        Remove-Item -LiteralPath $LogFile -Force -ErrorAction Stop
    } catch {
        $backup = Join-Path $ProjectDir ("startup_{0}.log" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
        try { Move-Item -LiteralPath $LogFile -Destination $backup -Force -ErrorAction Stop } catch {}
    }
}

"==== $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ====" | Out-File -FilePath $LogFile -Encoding utf8

try {
    Write-Log "Project dir: $ProjectDir"

    $pythonExe = Find-Python
    if (-not $pythonExe) {
        $pythonExe = Install-Python
    }
    if (-not $pythonExe -or -not (Test-PythonExe $pythonExe)) {
        throw 'No usable Python 3.10+ interpreter was found.'
    }

    Write-Log "Python: $pythonExe"
    Ensure-Venv -PythonPath $pythonExe
    Ensure-Pip
    Ensure-Requirements

    if ($DryRun -or $env:DRY_RUN -eq '1') {
        Write-Log 'DRY_RUN passed.'
        exit 0
    }

    Start-App
    exit 0
} catch {
    Write-Log "Startup failed: $($_.Exception.Message)"
    exit 1
}

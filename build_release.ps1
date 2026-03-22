param(
    [switch]$SkipDriverBundle
)

$ErrorActionPreference = "Stop"
Set-Location -Path $PSScriptRoot

function Resolve-PythonLauncher {
    $candidates = @()

    $knownPython = "C:/Users/Administrator/AppData/Local/Microsoft/WindowsApps/python3.9.exe"
    if (Test-Path -LiteralPath $knownPython) {
        $candidates += @{
            Exe = $knownPython
            PrefixArgs = @()
        }
    }

    $pyCmd = Get-Command py -ErrorAction SilentlyContinue
    if ($pyCmd) {
        $candidates += @{
            Exe = $pyCmd.Source
            PrefixArgs = @("-3")
        }
    }

    $pythonCmd = Get-Command python -ErrorAction SilentlyContinue
    if ($pythonCmd) {
        $candidates += @{
            Exe = $pythonCmd.Source
            PrefixArgs = @()
        }
    }

    foreach ($candidate in $candidates) {
        try {
            & $candidate.Exe @($candidate.PrefixArgs) --version | Out-Null
            if ($LASTEXITCODE -eq 0) {
                return $candidate
            }
        }
        catch {
            continue
        }
    }

    throw "Python launcher not found or unavailable. Install Python 3.9+ first."
}

function Find-CachedEdgeDriver {
    $roots = @(
        (Join-Path $env:LOCALAPPDATA "SeleniumManager\msedgedriver\win64"),
        (Join-Path $env:LOCALAPPDATA "selenium\msedgedriver\win64"),
        (Join-Path $env:USERPROFILE ".cache\selenium\msedgedriver\win64")
    )

    $matches = @()
    foreach ($root in $roots) {
        if (-not (Test-Path -LiteralPath $root)) {
            continue
        }
        $found = Get-ChildItem -Path $root -Filter "msedgedriver.exe" -Recurse -File -ErrorAction SilentlyContinue
        if ($found) {
            $matches += $found
        }
    }

    if (-not $matches) {
        return $null
    }

    return ($matches | Sort-Object -Property LastWriteTime -Descending | Select-Object -First 1).FullName
}

$driverDir = Join-Path $PSScriptRoot "drivers"
$driverTarget = Join-Path $driverDir "msedgedriver.exe"
New-Item -Path $driverDir -ItemType Directory -Force | Out-Null

if (-not $SkipDriverBundle) {
    $cachedDriver = Find-CachedEdgeDriver
    if ($cachedDriver) {
        Copy-Item -Path $cachedDriver -Destination $driverTarget -Force
        Write-Host "[INFO] Bundled EdgeDriver: $cachedDriver"
    }
    elseif (Test-Path -LiteralPath $driverTarget) {
        Write-Host "[INFO] Reusing existing bundled EdgeDriver: $driverTarget"
    }
    else {
        Write-Warning "No cached EdgeDriver found. Build will rely on Selenium Manager at runtime."
    }
}

$python = Resolve-PythonLauncher
& $python.Exe @($python.PrefixArgs) -m PyInstaller -y NWAFUAutoLogin.spec
if ($LASTEXITCODE -ne 0) {
    throw "PyInstaller build failed with exit code $LASTEXITCODE."
}

$distExe = Join-Path $PSScriptRoot "dist\NWAFUAutoLogin.exe"
if (-not (Test-Path -LiteralPath $distExe)) {
    throw "Build succeeded but EXE not found: $distExe"
}

$releaseDir = Join-Path $PSScriptRoot "release"
New-Item -Path $releaseDir -ItemType Directory -Force | Out-Null

$releaseExe = Join-Path $releaseDir "NWAFUAutoLogin.exe"
Copy-Item -Path $distExe -Destination $releaseExe -Force

$releaseDriver = Join-Path $releaseDir "msedgedriver.exe"
if (Test-Path -LiteralPath $driverTarget) {
    Copy-Item -Path $driverTarget -Destination $releaseDriver -Force
}

$zipPath = Join-Path $releaseDir "NWAFUAutoLogin-win64.zip"
if (Test-Path -LiteralPath $zipPath) {
    Remove-Item -LiteralPath $zipPath -Force
}
if (Test-Path -LiteralPath $releaseDriver) {
    Compress-Archive -Path @($releaseExe, $releaseDriver) -DestinationPath $zipPath -Force
}
else {
    Compress-Archive -Path $releaseExe -DestinationPath $zipPath -Force
}

Write-Host "[OK] Release EXE : $releaseExe"
Write-Host "[OK] Release ZIP : $zipPath"

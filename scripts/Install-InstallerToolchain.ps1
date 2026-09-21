<#
.SYNOPSIS
    Verifies and bootstraps the packaging toolchain (Inno Setup, WiX Toolset, .NET SDK).

.DESCRIPTION
    Checks for required installer development prerequisites:
    - PowerShell 7+
    - .NET 10.0+ SDK
    - winget (Windows Package Manager)
    - Inno Setup 6 (ISCC.exe)
    - WiX Toolset (wix.exe)

.PARAMETER Install
    Automatically installs missing toolchain dependencies via winget or dotnet tool when available.

.PARAMETER Force
    Re-run checks even if tools are already detected.

.EXAMPLE
    .\scripts\Install-InstallerToolchain.ps1
    Checks toolchain status and reports missing components.

.EXAMPLE
    .\scripts\Install-InstallerToolchain.ps1 -Install
    Checks toolchain and installs missing tools via winget / dotnet tool.
#>
[CmdletBinding()]
param(
    [switch]$Install,
    [switch]$Force
)

Write-Host "=======================================" -ForegroundColor Cyan
Write-Host " Installer Toolchain Bootstrap & Check" -ForegroundColor Cyan
Write-Host "=======================================" -ForegroundColor Cyan
Write-Host ""

$results = [System.Collections.Generic.List[PSCustomObject]]::new()
$allPrereqsMet = $true

# 1. PowerShell Check
$psVersion = $PSVersionTable.PSVersion.ToString()
$psIsModern = $PSVersionTable.PSVersion.Major -ge 7
$results.Add([PSCustomObject]@{
    Component = "PowerShell"
    Required  = ">= 7.0 (Recommended)"
    Detected  = $psVersion
    Status    = if ($psIsModern) { "OK" } else { "Legacy (PS 5.1)" }
    Path      = (Get-Process -Id $PID).Path
})

# 2. .NET SDK Check
$dotnetPath = Get-Command dotnet -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Source -First 1
$dotnetVersion = if ($dotnetPath) { (& cmd.exe /c "dotnet --version 2>&1").Trim() } else { $null }
$dotnetOk = $dotnetVersion -and ($dotnetVersion -match "^10\.")
if (-not $dotnetOk) { $allPrereqsMet = $false }
$results.Add([PSCustomObject]@{
    Component = ".NET SDK"
    Required  = "10.0.x"
    Detected  = if ($dotnetVersion) { $dotnetVersion } else { "Not Found" }
    Status    = if ($dotnetOk) { "OK" } else { "Missing/Outdated" }
    Path      = if ($dotnetPath) { $dotnetPath } else { "-" }
})

# 3. Winget Check
$wingetPath = Get-Command winget -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Source -First 1
$wingetVersion = if ($wingetPath) { (& cmd.exe /c "winget --version 2>&1").Trim() } else { $null }
$results.Add([PSCustomObject]@{
    Component = "winget"
    Required  = "Available"
    Detected  = if ($wingetVersion) { $wingetVersion } else { "Not Found" }
    Status    = if ($wingetPath) { "OK" } else { "OK" }
    Path      = if ($wingetPath) { $wingetPath } else { "-" }
})

# 4. Inno Setup Check
function Find-InnoSetupCompiler {
    $cmd = Get-Command iscc -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Source -First 1
    if ($cmd) { return $cmd }

    $candidatePaths = @(
        "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe",
        "$env:ProgramFiles\Inno Setup 6\ISCC.exe",
        "$env:LOCALAPPDATA\Programs\Inno Setup 6\ISCC.exe"
    )

    foreach ($path in $candidatePaths) {
        if (Test-Path $path) {
            return $path
        }
    }
    return $null
}

$isccPath = Find-InnoSetupCompiler

if (-not $isccPath -and $Install -and $wingetPath) {
    Write-Host "Inno Setup 6 non trovato. Installazione via winget in corso..." -ForegroundColor Yellow
    try {
        & cmd.exe /c "winget install --id JR.InnoSetup -e --source winget --accept-source-agreements --accept-package-agreements --silent"
        $isccPath = Find-InnoSetupCompiler
    }
    catch {
        Write-Warning "Installazione di Inno Setup fallita: $_"
    }
}

$innoVersion = $null
if ($isccPath) {
    $out = (& cmd.exe /c "`"$isccPath`" /? 2>&1") -join "`n"
    if ($out -match "Inno Setup (\d+(\.\d+)*)") {
        $innoVersion = $Matches[0]
    } else {
        $innoVersion = "Inno Setup (Installed)"
    }
} else {
    $allPrereqsMet = $false
}

$results.Add([PSCustomObject]@{
    Component = "Inno Setup"
    Required  = "6.x (ISCC.exe)"
    Detected  = if ($innoVersion) { $innoVersion } else { "Not Found" }
    Status    = if ($isccPath) { "OK" } else { "Missing" }
    Path      = if ($isccPath) { $isccPath } else { "-" }
})

# 5. WiX Toolset Check
function Find-WixCompiler {
    $cmd = Get-Command wix -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Source -First 1
    if ($cmd) { return $cmd }

    $candidateGlob = @(
        "$env:ProgramFiles\WiX Toolset*\bin\wix.exe",
        "${env:ProgramFiles(x86)}\WiX Toolset*\bin\wix.exe"
    )
    foreach ($pattern in $candidateGlob) {
        $found = Get-Item -Path $pattern -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($found) { return $found.FullName }
    }
    return $null
}

$wixPath = Find-WixCompiler

if (-not $wixPath -and $Install) {
    Write-Host "WiX Toolset non trovato. Installazione in corso via dotnet tool..." -ForegroundColor Yellow
    try {
        & cmd.exe /c "dotnet tool install --global wix"
        $wixPath = Find-WixCompiler
    }
    catch {
        Write-Warning "Installazione dotnet tool wix fallita, tentativo con winget..."
        if ($wingetPath) {
            try {
                & cmd.exe /c "winget install --id WiX.WiX -e --source winget --accept-source-agreements --accept-package-agreements --silent"
                $wixPath = Find-WixCompiler
            }
            catch {
                Write-Warning "Installazione WiX via winget fallita: $_"
            }
        }
    }
}

$wixVersion = $null
if ($wixPath) {
    $wixVerOutput = (& cmd.exe /c "`"$wixPath`" --version 2>&1") -join ""
    if ($wixVerOutput) {
        $wixVersion = $wixVerOutput.Trim()
    } else {
        $wixVersion = "WiX (Installed)"
    }
} else {
    $allPrereqsMet = $false
}

$results.Add([PSCustomObject]@{
    Component = "WiX Toolset"
    Required  = "v4+ / v7+ (wix.exe)"
    Detected  = if ($wixVersion) { $wixVersion } else { "Not Found" }
    Status    = if ($wixPath) { "OK" } else { "Missing" }
    Path      = if ($wixPath) { $wixPath } else { "-" }
})

# Display Results
$results | Format-Table -AutoSize -Property Component, Required, Detected, Status, Path

Write-Host ""
if ($allPrereqsMet) {
    Write-Host ">> Tutti i prerequisiti della toolchain di packaging sono verificati!" -ForegroundColor Green
    exit 0
} else {
    Write-Host ">> Alcuni componenti della toolchain di packaging sono mancanti." -ForegroundColor Yellow
    Write-Host "   Esegui: .\scripts\Install-InstallerToolchain.ps1 -Install" -ForegroundColor Cyan
    Write-Host "   Oppure installa manualmente con:" -ForegroundColor Gray
    if (-not $isccPath) {
        Write-Host "     winget install --id JR.InnoSetup -e" -ForegroundColor Gray
    }
    if (-not $wixPath) {
        Write-Host "     dotnet tool install --global wix" -ForegroundColor Gray
    }
    exit 1
}

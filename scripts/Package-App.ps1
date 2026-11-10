<#
.SYNOPSIS
    Packages OnlyFirmaOutlook into distribution installers (Inno Setup EXE and/or WiX MSI).

.DESCRIPTION
    Automates the full packaging pipeline:
    1. Builds and publishes the application via Build-App.ps1 (unless -SkipBuild is specified)
    2. Builds the canonical Inno Setup EXE installer using packaging/innosetup/setup.iss
    3. Builds the canonical WiX MSI package using packaging/msi/OnlyFirmaOutlook.wxs

.PARAMETER Format
    Specifies which installer format to build: "All" (default), "InnoSetup", or "WiX".

.PARAMETER Configuration
    Build configuration: "Release" (default) or "Debug".

.PARAMETER Version
    Semantic version string for the package (default: "1.0.0").

.PARAMETER OutputDir
    Output directory for generated packages (default: "packaging/output").

.PARAMETER SkipBuild
    Skips the Build-App.ps1 compilation step and uses existing dist/ artifacts.

.EXAMPLE
    .\scripts\Package-App.ps1
    Builds the solution and creates both EXE and MSI installers.

.EXAMPLE
    .\scripts\Package-App.ps1 -Format InnoSetup -Version "1.1.0"
    Builds the solution and packages only the Inno Setup EXE.

.EXAMPLE
    .\scripts\Package-App.ps1 -Format WiX -SkipBuild
    Creates the WiX MSI from existing dist/ artifacts.
#>
[CmdletBinding()]
param(
    [ValidateSet("All", "InnoSetup", "WiX")]
    [string]$Format = "All",

    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Release",

    [string]$Version = "1.0.0",

    [string]$OutputDir = "packaging/output",

    [switch]$SkipBuild
)

$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$rootDir = Split-Path -Parent $scriptDir
$distDir = Join-Path $rootDir "dist"
$targetOutputDir = if ([System.IO.Path]::IsPathRooted($OutputDir)) { $OutputDir } else { Join-Path $rootDir $OutputDir }

Write-Host "=======================================" -ForegroundColor Cyan
Write-Host " OnlyFirmaOutlook Packaging Pipeline" -ForegroundColor Cyan
Write-Host "=======================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Root Directory:  $rootDir"
Write-Host "Target Format:   $Format"
Write-Host "Configuration:   $Configuration"
Write-Host "Version:         $Version"
Write-Host "Output Directory:$targetOutputDir"
Write-Host ""

# Ensure output directory exists
if (-not (Test-Path $targetOutputDir)) {
    New-Item -ItemType Directory -Path $targetOutputDir -Force | Out-Null
}

# 1. Step: Build app if required
if (-not $SkipBuild) {
    Write-Host ">> Compilazione applicazione con Build-App.ps1..." -ForegroundColor Yellow
    $buildScript = Join-Path $scriptDir "Build-App.ps1"
    & powershell -ExecutionPolicy Bypass -File $buildScript -Configuration $Configuration -PublishMode FrameworkDependent -SkipTests
    if ($LASTEXITCODE -ne 0) {
        throw "Compilazione fallita con codice di uscita $LASTEXITCODE"
    }
}

# Verify dist/ contains required artifacts
$launcherExe = Join-Path $distDir "OnlyFirmaOutlook.Launcher.exe"
if (-not (Test-Path $launcherExe)) {
    throw "File non trovato in dist: $launcherExe. Esegui la build prima di creare i pacchetti."
}

# Helper: Find Inno Setup ISCC
function Get-InnoCompiler {
    $cmd = Get-Command iscc -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Source -First 1
    if ($cmd) { return $cmd }
    $candidates = @(
        "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe",
        "$env:ProgramFiles\Inno Setup 6\ISCC.exe",
        "$env:LOCALAPPDATA\Programs\Inno Setup 6\ISCC.exe"
    )
    foreach ($p in $candidates) {
        if (Test-Path $p) { return $p }
    }
    return $null
}

# Helper: Find WiX wix.exe
function Get-WixCompiler {
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

$generatedPackages = [System.Collections.Generic.List[string]]::new()

# 2. Step: Inno Setup EXE Packaging
if ($Format -in @("All", "InnoSetup")) {
    Write-Host ""
    Write-Host ">> Creazione Installer Inno Setup (EXE)..." -ForegroundColor Yellow
    $iscc = Get-InnoCompiler
    if (-not $iscc) {
        Write-Warning "Inno Setup 6 non trovato. Esegui .\scripts\Install-InstallerToolchain.ps1 per installarlo."
        if ($Format -eq "InnoSetup") {
            throw "Compilazione Inno Setup interrotta: ISCC.exe non disponibile."
        }
    } else {
        $issFile = Join-Path $rootDir "packaging/innosetup/setup.iss"
        if (-not (Test-Path $issFile)) {
            throw "File di configurazione Inno Setup non trovato: $issFile"
        }

        $exeName = "OnlyFirmaOutlook-Setup-$Version"
        $cmdArgs = @(
            "`"$issFile`"",
            "/DAppVersion=$Version",
            "/O`"$targetOutputDir`"",
            "/F`"$exeName`""
        )
        Write-Host "Esecuzione: $iscc $($cmdArgs -join ' ')" -ForegroundColor Gray
        & cmd.exe /c "`"$iscc`" $issFile /DAppVersion=$Version /O`"$targetOutputDir`" /F`"$exeName`""
        if ($LASTEXITCODE -ne 0) {
            throw "Compilazione Inno Setup fallita con codice di uscita $LASTEXITCODE"
        }

        $expectedExe = Join-Path $targetOutputDir "$exeName.exe"
        if (Test-Path $expectedExe) {
            $generatedPackages.Add($expectedExe)
            Write-Host "   OK - $expectedExe" -ForegroundColor Green
        }
    }
}

# 3. Step: WiX MSI Packaging
if ($Format -in @("All", "WiX")) {
    Write-Host ""
    Write-Host ">> Creazione Pacchetto WiX (MSI)..." -ForegroundColor Yellow
    $wix = Get-WixCompiler
    if (-not $wix) {
        Write-Warning "WiX Toolset non trovato. Esegui .\scripts\Install-InstallerToolchain.ps1 per installarlo."
        if ($Format -eq "WiX") {
            throw "Compilazione WiX interrotta: wix.exe non disponibile."
        }
    } else {
        $wxsFile = Join-Path $rootDir "packaging/msi/OnlyFirmaOutlook.wxs"
        if (-not (Test-Path $wxsFile)) {
            throw "File di configurazione WiX non trovato: $wxsFile"
        }

        # Format version for MSI (must be x.x.x.x)
        $msiVersion = if ($Version -match "^\d+\.\d+\.\d+\.\d+$") { $Version } else { "$Version.0" }
        $msiPath = Join-Path $targetOutputDir "OnlyFirmaOutlook-$Version.msi"

        Write-Host "Esecuzione: $wix build $wxsFile -o $msiPath -d AppVersion=$msiVersion -b $rootDir --acceptEula wix7" -ForegroundColor Gray
        & cmd.exe /c "`"$wix`" build `"$wxsFile`" -o `"$msiPath`" -d AppVersion=$msiVersion -b `"$rootDir`" --acceptEula wix7"
        if ($LASTEXITCODE -ne 0) {
            throw "Compilazione WiX fallita con codice di uscita $LASTEXITCODE"
        }

        if (Test-Path $msiPath) {
            $generatedPackages.Add($msiPath)
            Write-Host "   OK - $msiPath" -ForegroundColor Green
        }
    }
}

Write-Host ""
Write-Host "=======================================" -ForegroundColor Cyan
Write-Host " Packaging Completato con Successo!" -ForegroundColor Green
Write-Host "=======================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Pacchetti generati:" -ForegroundColor White
foreach ($pkg in $generatedPackages) {
    $item = Get-Item $pkg
    $sizeMb = [math]::Round($item.Length / 1MB, 2)
    Write-Host "  - $($item.Name) ($sizeMb MB) in $($item.DirectoryName)" -ForegroundColor Green
}
Write-Host ""

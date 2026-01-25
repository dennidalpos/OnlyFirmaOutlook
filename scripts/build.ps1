<#
.SYNOPSIS
    Build e publish di OnlyFirmaOutlook.

.DESCRIPTION
    Questo script esegue:
    - Restore delle dipendenze NuGet
    - Build del progetto
    - Publish self-contained per win-x86 e win-x64
    - Publish del bootstrapper/launcher
    - Copia dei file nella cartella dist

.PARAMETER Configuration
    Configurazione di build (Debug o Release). Default: Release

.PARAMETER SkipRestore
    Salta il restore delle dipendenze

.PARAMETER SkipClean
    Salta la pulizia della cartella dist

.EXAMPLE
    .\build.ps1

.EXAMPLE
    .\build.ps1 -Configuration Debug

.EXAMPLE
    .\build.ps1 -SkipRestore
#>

param(
    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Release",

    [switch]$SkipRestore,

    [switch]$SkipClean
)

$ErrorActionPreference = "Stop"

# Percorsi
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$rootDir = Split-Path -Parent $scriptDir
$srcDir = Join-Path $rootDir "src"
$distDir = Join-Path $rootDir "dist"
$mainProjectDir = Join-Path $srcDir "OnlyFirmaOutlook"
$bootstrapperDir = Join-Path $srcDir "Bootstrapper"
$mediaSourceDir = Join-Path $mainProjectDir "media"

Write-Host "=======================================" -ForegroundColor Cyan
Write-Host " OnlyFirmaOutlook Build Script" -ForegroundColor Cyan
Write-Host "=======================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Root Directory: $rootDir"
Write-Host "Configuration:  $Configuration"
Write-Host ""

# Funzione per eseguire comandi e verificare errori
function Invoke-BuildCommand {
    param(
        [string]$Description,
        [scriptblock]$Command
    )

    Write-Host ">> $Description" -ForegroundColor Yellow
    & $Command

    if ($LASTEXITCODE -ne 0) {
        Write-Host "ERRORE: $Description fallito con codice $LASTEXITCODE" -ForegroundColor Red
        exit $LASTEXITCODE
    }

    Write-Host "   OK" -ForegroundColor Green
    Write-Host ""
}

# Pulizia cartella dist
if (-not $SkipClean) {
    Write-Host "Pulizia cartella dist..." -ForegroundColor Yellow
    if (Test-Path $distDir) {
        Remove-Item -Path $distDir -Recurse -Force
    }
    New-Item -ItemType Directory -Path $distDir -Force | Out-Null
    Write-Host "   OK" -ForegroundColor Green
    Write-Host ""
}

# Restore dipendenze
if (-not $SkipRestore) {
    Invoke-BuildCommand "Restore dipendenze" {
        dotnet restore "$rootDir\OnlyFirmaOutlook.sln"
    }
}

# Build soluzione
Invoke-BuildCommand "Build soluzione ($Configuration)" {
    dotnet build "$rootDir\OnlyFirmaOutlook.sln" -c $Configuration --no-restore
}

# Publish win-x86
Write-Host "Publish win-x86..." -ForegroundColor Yellow
$publishX86Dir = Join-Path $distDir "win-x86"
dotnet publish "$mainProjectDir\OnlyFirmaOutlook.csproj" `
    -c $Configuration `
    -r win-x86 `
    --self-contained true `
    -p:PublishSingleFile=false `
    -o $publishX86Dir

if ($LASTEXITCODE -ne 0) {
    Write-Host "ERRORE: Publish win-x86 fallito" -ForegroundColor Red
    exit $LASTEXITCODE
}
Write-Host "   OK - Output: $publishX86Dir" -ForegroundColor Green
Write-Host ""

# Publish win-x64
Write-Host "Publish win-x64..." -ForegroundColor Yellow
$publishX64Dir = Join-Path $distDir "win-x64"
dotnet publish "$mainProjectDir\OnlyFirmaOutlook.csproj" `
    -c $Configuration `
    -r win-x64 `
    --self-contained true `
    -p:PublishSingleFile=false `
    -o $publishX64Dir

if ($LASTEXITCODE -ne 0) {
    Write-Host "ERRORE: Publish win-x64 fallito" -ForegroundColor Red
    exit $LASTEXITCODE
}
Write-Host "   OK - Output: $publishX64Dir" -ForegroundColor Green
Write-Host ""

# Publish bootstrapper
Write-Host "Publish Bootstrapper..." -ForegroundColor Yellow
$bootstrapperOutput = Join-Path $distDir "OnlyFirmaOutlook.Launcher.exe"
dotnet publish "$bootstrapperDir\Bootstrapper.csproj" `
    -c $Configuration `
    -r win-x64 `
    --self-contained true `
    -p:PublishSingleFile=true `
    -o $distDir

if ($LASTEXITCODE -ne 0) {
    Write-Host "ERRORE: Publish Bootstrapper fallito" -ForegroundColor Red
    exit $LASTEXITCODE
}
Write-Host "   OK - Output: $distDir" -ForegroundColor Green
Write-Host ""

# Crea cartella media in entrambe le build se non esiste
$mediaX86 = Join-Path $publishX86Dir "media"
$mediaX64 = Join-Path $publishX64Dir "media"

if (-not (Test-Path $mediaX86)) {
    New-Item -ItemType Directory -Path $mediaX86 -Force | Out-Null
    Write-Host "Creata cartella media in win-x86" -ForegroundColor Gray
}

if (-not (Test-Path $mediaX64)) {
    New-Item -ItemType Directory -Path $mediaX64 -Force | Out-Null
    Write-Host "Creata cartella media in win-x64" -ForegroundColor Gray
}

# Copia eventuali file preset dalla cartella media sorgente
if (Test-Path $mediaSourceDir) {
    $presetFiles = Get-ChildItem -Path $mediaSourceDir -Filter "*.doc*" -File
    if ($presetFiles.Count -gt 0) {
        Write-Host "Copia file preset..." -ForegroundColor Yellow
        foreach ($file in $presetFiles) {
            Copy-Item -Path $file.FullName -Destination $mediaX86 -Force
            Copy-Item -Path $file.FullName -Destination $mediaX64 -Force
            Write-Host "   Copiato: $($file.Name)" -ForegroundColor Gray
        }
        Write-Host "   OK" -ForegroundColor Green
        Write-Host ""
    }
}

# Riepilogo
Write-Host "=======================================" -ForegroundColor Cyan
Write-Host " Build completata con successo!" -ForegroundColor Green
Write-Host "=======================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Output in: $distDir"
Write-Host ""
Write-Host "Contenuto:" -ForegroundColor White
Write-Host "  - OnlyFirmaOutlook.Launcher.exe  (Avvia la build corretta in base a Office)"
Write-Host "  - win-x86\                       (Build 32-bit per Office 32-bit)"
Write-Host "  - win-x64\                       (Build 64-bit per Office 64-bit)"
Write-Host ""
Write-Host "=======================================" -ForegroundColor Cyan
Write-Host " PRESET (Media)" -ForegroundColor Yellow
Write-Host "=======================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Per aggiungere preset (documenti Word predefiniti):" -ForegroundColor White
Write-Host ""
Write-Host "  1. Copia i file .doc/.docx in ENTRAMBE le cartelle:" -ForegroundColor Gray
Write-Host "     - $mediaX86" -ForegroundColor Gray
Write-Host "     - $mediaX64" -ForegroundColor Gray
Write-Host ""
Write-Host "  Oppure:" -ForegroundColor White
Write-Host ""
Write-Host "  2. Metti i file in src\OnlyFirmaOutlook\media\" -ForegroundColor Gray
Write-Host "     e riesegui questo script (verranno copiati automaticamente)" -ForegroundColor Gray
Write-Host ""
Write-Host "=======================================" -ForegroundColor Cyan
Write-Host " Distribuzione" -ForegroundColor Yellow
Write-Host "=======================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Per distribuire, copiare l'intera cartella 'dist' sulla share di rete."
Write-Host "Gli utenti dovranno eseguire OnlyFirmaOutlook.Launcher.exe"
Write-Host ""

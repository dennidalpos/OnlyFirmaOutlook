

param(
    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Release",

    [string]$OutputDir = "dist",

    [switch]$SkipRestore,

    [switch]$SkipClean,

    [string[]]$Runtimes = @("win-x86", "win-x64"),

    [switch]$SkipTests,

    [switch]$SkipPublish,

    [switch]$SkipBootstrapper,

    [switch]$SkipMediaCopy
)

$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$rootDir = Split-Path -Parent $scriptDir
$srcDir = Join-Path $rootDir "src"
$distDir = Join-Path $rootDir $OutputDir
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

if (-not $SkipClean) {
    Write-Host "Pulizia cartella dist..." -ForegroundColor Yellow
    if (Test-Path $distDir) {
        Remove-Item -Path $distDir -Recurse -Force
    }
    New-Item -ItemType Directory -Path $distDir -Force | Out-Null
    Write-Host "   OK" -ForegroundColor Green
    Write-Host ""
} elseif (-not (Test-Path $distDir)) {
    New-Item -ItemType Directory -Path $distDir -Force | Out-Null
}

if (-not $SkipRestore) {
    Invoke-BuildCommand "Restore dipendenze" {
        dotnet restore "$rootDir\OnlyFirmaOutlook.sln"
    }
}

Invoke-BuildCommand "Build soluzione ($Configuration)" {
    dotnet build "$rootDir\OnlyFirmaOutlook.sln" -c $Configuration --no-restore
}

if (-not $SkipTests) {
    Invoke-BuildCommand "Esecuzione test" {
        dotnet test "$rootDir\OnlyFirmaOutlook.sln" -c $Configuration --no-build
    }
}

if (-not $SkipPublish) {
    $normalizedRuntimes = $Runtimes | ForEach-Object { $_.ToLowerInvariant() } | Select-Object -Unique
    $publishDirs = @{}

    foreach ($runtime in $normalizedRuntimes) {
        Write-Host "Publish $runtime..." -ForegroundColor Yellow
        $publishDir = Join-Path $distDir $runtime
        $publishDirs[$runtime] = $publishDir

        dotnet publish "$mainProjectDir\OnlyFirmaOutlook.csproj" `
            -c $Configuration `
            -r $runtime `
            --self-contained true `
            -p:PublishSingleFile=false `
            -o $publishDir

        if ($LASTEXITCODE -ne 0) {
            Write-Host "ERRORE: Publish $runtime fallito" -ForegroundColor Red
            exit $LASTEXITCODE
        }
        Write-Host "   OK - Output: $publishDir" -ForegroundColor Green
        Write-Host ""
    }

    if (-not $SkipBootstrapper) {
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
    }

    foreach ($runtime in $publishDirs.Keys) {
        $mediaTarget = Join-Path $publishDirs[$runtime] "media"
        if (-not (Test-Path $mediaTarget)) {
            New-Item -ItemType Directory -Path $mediaTarget -Force | Out-Null
            Write-Host "Creata cartella media in $runtime" -ForegroundColor Gray
        }
    }

    if (-not $SkipMediaCopy) {
        if (Test-Path $mediaSourceDir) {
            $presetFiles = Get-ChildItem -Path $mediaSourceDir -Filter "*.doc*" -File
            if ($presetFiles.Count -gt 0) {
                Write-Host "Copia file preset..." -ForegroundColor Yellow
                foreach ($file in $presetFiles) {
                    foreach ($runtime in $publishDirs.Keys) {
                        Copy-Item -Path $file.FullName -Destination (Join-Path $publishDirs[$runtime] "media") -Force
                    }
                    Write-Host "   Copiato: $($file.Name)" -ForegroundColor Gray
                }
                Write-Host "   OK" -ForegroundColor Green
                Write-Host ""
            }
        }
    }
}

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
Write-Host "  1. Copia i file .doc/.docx nelle cartelle 'media' delle build pubblicate" -ForegroundColor Gray
Write-Host "     (es. dist\\win-x86\\media e dist\\win-x64\\media)" -ForegroundColor Gray
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

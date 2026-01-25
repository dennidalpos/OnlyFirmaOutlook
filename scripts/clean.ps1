

param(
    [switch]$All
)

$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$rootDir = Split-Path -Parent $scriptDir

Write-Host "=======================================" -ForegroundColor Cyan
Write-Host " OnlyFirmaOutlook Clean Script" -ForegroundColor Cyan
Write-Host "=======================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Root Directory: $rootDir"
Write-Host ""

$foldersToDelete = @(
    "src\OnlyFirmaOutlook\bin",
    "src\OnlyFirmaOutlook\obj",
    "src\Bootstrapper\bin",
    "src\Bootstrapper\obj",
    "dist"
)

if ($All) {
    $foldersToDelete += @(
        ".vs",
        "packages",
        "TestResults"
    )
}

$deletedCount = 0
$skippedCount = 0

foreach ($folder in $foldersToDelete) {
    $fullPath = Join-Path $rootDir $folder

    if (Test-Path $fullPath) {
        Write-Host "Eliminazione: $folder" -ForegroundColor Yellow
        try {
            Remove-Item -Path $fullPath -Recurse -Force
            Write-Host "   OK" -ForegroundColor Green
            $deletedCount++
        }
        catch {
            Write-Host "   ERRORE: $($_.Exception.Message)" -ForegroundColor Red
        }
    }
    else {
        Write-Host "Non trovato: $folder" -ForegroundColor Gray
        $skippedCount++
    }
}

$tempPatterns = @(
    "*.user",
    "*.suo",
    "*.cache"
)

if ($All) {
    Write-Host ""
    Write-Host "Ricerca file temporanei..." -ForegroundColor Yellow

    foreach ($pattern in $tempPatterns) {
        $files = Get-ChildItem -Path $rootDir -Filter $pattern -Recurse -Force -ErrorAction SilentlyContinue
        foreach ($file in $files) {
            try {
                Remove-Item -Path $file.FullName -Force
                Write-Host "   Eliminato: $($file.FullName)" -ForegroundColor Gray
                $deletedCount++
            }
            catch {
                Write-Host "   Impossibile eliminare: $($file.FullName)" -ForegroundColor Red
            }
        }
    }
}

Write-Host ""
Write-Host "=======================================" -ForegroundColor Cyan
Write-Host " Pulizia completata" -ForegroundColor Green
Write-Host "=======================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Elementi eliminati: $deletedCount"
Write-Host "Elementi non trovati: $skippedCount"
Write-Host ""

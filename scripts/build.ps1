<#
.SYNOPSIS
    Backward compatibility wrapper for Build-App.ps1.
#>
& (Join-Path $PSScriptRoot "Build-App.ps1") @args

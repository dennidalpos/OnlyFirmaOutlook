<#
.SYNOPSIS
    Backward compatibility wrapper for Clean-App.ps1.
#>
& (Join-Path $PSScriptRoot "Clean-App.ps1") @args

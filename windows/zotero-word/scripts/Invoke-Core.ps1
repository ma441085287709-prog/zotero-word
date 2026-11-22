#requires -Version 5.1
[CmdletBinding()]
param(
    [string]$Python,
    [Parameter(Mandatory = $true)][string[]]$CoreArguments
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
if ($env:OS -ne 'Windows_NT') { throw 'This launcher requires Windows Python.' }
$pythonPrefix = @()
if ($Python) {
    $pythonCommand = (Get-Command -Name $Python -CommandType Application -ErrorAction Stop | Select-Object -First 1).Source
} else {
    $launcher = Get-Command -Name 'py.exe' -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($launcher) {
        $pythonCommand = $launcher.Source
        $pythonPrefix = @('-3')
    } else {
        $pythonCommand = (Get-Command -Name 'python.exe' -CommandType Application -ErrorAction Stop | Select-Object -First 1).Source
    }
}
$coreScript = Join-Path (Split-Path -Path $PSScriptRoot -Parent) 'core\scripts\zotero_docx.py'
$previousUtf8 = $env:PYTHONUTF8
$previousEncoding = $env:PYTHONIOENCODING
try {
    $env:PYTHONUTF8 = '1'
    $env:PYTHONIOENCODING = 'utf-8'
    & $pythonCommand @pythonPrefix '-X' 'utf8' $coreScript @CoreArguments
    $coreExit = $LASTEXITCODE
} finally {
    $env:PYTHONUTF8 = $previousUtf8
    $env:PYTHONIOENCODING = $previousEncoding
}
exit $coreExit

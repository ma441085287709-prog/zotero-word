#requires -Version 5.1
[CmdletBinding()]
param([Parameter(Mandatory = $true)][string]$SkillsRoot)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
if ($env:OS -ne 'Windows_NT') { throw 'Run this installer on Windows.' }
if (-not [IO.Path]::IsPathRooted($SkillsRoot)) { throw 'SkillsRoot must be an absolute directory.' }
$sourceSkill = [IO.Path]::GetFullPath($PSScriptRoot).TrimEnd('\')
$installRoot = [IO.Path]::GetFullPath($SkillsRoot)
if ($installRoot.StartsWith('\\')) { throw 'Use a local skills directory, not a UNC share.' }
$destination = Join-Path $installRoot 'zotero-word'
$comparison = [StringComparison]::OrdinalIgnoreCase
if ($destination.Equals($sourceSkill, $comparison) -or
    $destination.StartsWith($sourceSkill + '\', $comparison)) {
    throw 'Choose a skills directory outside the extracted source package.'
}
if (Test-Path -LiteralPath $destination) {
    throw 'zotero-word already exists there. Use an empty skills root or preserve and move the older installation first.'
}
$scan = $installRoot
while ($scan) {
    if (Test-Path -LiteralPath $scan) {
        $item = Get-Item -LiteralPath $scan -Force
        if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
            throw 'Installation through a junction or symbolic link is unsupported.'
        }
    }
    $parent = Split-Path -Path $scan -Parent
    if ($parent -eq $scan) { break }
    $scan = $parent
}
$sourceItems = @(Get-ChildItem -LiteralPath $sourceSkill -Recurse -Force)
if (@($sourceItems | Where-Object { ($_.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0 }).Count -gt 0) {
    throw 'The source package must not contain junctions or symbolic links.'
}
[IO.Directory]::CreateDirectory($installRoot) | Out-Null
$staging = Join-Path $installRoot ('.zotero-word-install-' + [Guid]::NewGuid().ToString('N'))
try {
    [IO.Directory]::CreateDirectory($staging) | Out-Null
    Get-ChildItem -LiteralPath $sourceSkill -Force | ForEach-Object {
        Copy-Item -LiteralPath $_.FullName -Destination $staging -Recurse -ErrorAction Stop
    }
    # Directory.Move refuses an existing target; an earlier skill is never replaced.
    [IO.Directory]::Move($staging, $destination)
} finally {
    if (Test-Path -LiteralPath $staging) {
        Remove-Item -LiteralPath $staging -Recurse -Force
    }
}
[ordered]@{
    result = 'INSTALLED'
    skill = 'zotero-word'
    version = '2.4.4'
    build = 'direct-save.2-handoff.1-dispatch-fixed'
    variant = 'windows-unified'
    adapters = @('codex-windows', 'deepseek-harness-windows')
    installed_path = $destination
    test_status = 'NOT_RUN_AT_USER_REQUEST'
    dependency_installation = 'NOT_PERFORMED'
} | ConvertTo-Json -Depth 4

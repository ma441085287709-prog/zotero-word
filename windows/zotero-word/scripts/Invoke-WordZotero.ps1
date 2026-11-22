#requires -Version 5.1
[CmdletBinding(DefaultParameterSetName='Bridge')]
param(
    [Parameter(Mandatory, ParameterSetName='Bridge')][switch]$Bridge,
    [Parameter(Mandatory, ParameterSetName='Connect')][switch]$ConnectOnly,
    [Parameter(Mandatory, ParameterSetName='Resume')][string]$JobDirectory,
    [Parameter(ParameterSetName='Resume')][string]$ClientPython
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
$utf8 = [Text.UTF8Encoding]::new($false)
[Console]::InputEncoding = $utf8
[Console]::OutputEncoding = $utf8

if ($PSCmdlet.ParameterSetName -eq 'Resume') {
    $job = Get-Item -LiteralPath $JobDirectory
    if (!$job.PSIsContainer -or $job.Name -notmatch '^native-[0-9a-f]{32}$') { throw 'Use the run directory recorded in run.json.' }
    $index = Join-Path $job.Parent.FullName 'run.json'
    $state = [IO.File]::ReadAllText($index, $utf8) | ConvertFrom-Json
    $bound = foreach ($key in 'native_job','verification_job','library_job') {
        if ($state.PSObject.Properties[$key]) { $state.$key }
    }
    if ($job.FullName -notin @($bound)) { throw 'This native job is not bound to the active run.' }
    $receipt = Join-Path $job.Parent.FullName ('resume-' + [guid]::NewGuid().ToString('N') + '.json')
    & (Join-Path $PSScriptRoot 'Invoke-Core.ps1') -Python $ClientPython -CoreArguments @('run-next','--run',$job.Parent.FullName,'--output',$receipt)
    exit $LASTEXITCODE
}

$script:word = $null
$script:pageWord = $script:pageDoc = $null
$script:phase = 'connect'
$script:runSent = $false
$script:focusMethod = $null
. (Join-Path $PSScriptRoot 'ZoteroRunner.ps1')

function Connect-Word {
    if ($null -eq $script:word) {
        try { $script:word = [Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application') }
        catch {
            if ($_.Exception.GetBaseException().HResult -ne -2147221021) { throw }
            $script:word = New-Object -ComObject Word.Application
            $script:word.Visible = $true
        }
    }
}
function Find-WorkingDocument([string]$target) {
    foreach ($d in $script:word.Documents) {
        if ([string]::Equals($d.FullName, $target, [StringComparison]::OrdinalIgnoreCase)) { return ,$d }
    }
}
function Assert-ActiveDocument([string]$target) {
    $doc = $script:word.ActiveDocument
    if ($null -eq $doc -or ![string]::Equals($doc.FullName, $target, [StringComparison]::OrdinalIgnoreCase)) {
        throw 'Word selected a different document.'
    }
    if ($doc.ReadOnly) { throw 'Word opened the working copy read-only.' }
    return ,$doc
}
function Assert-IndependentCopy([string]$target, [string]$source, [string]$snapshot) {
    Initialize-DesktopTypes
    if ([IO.Path]::GetExtension($target) -ine '.docx' -or $target.StartsWith('\\') -or $target -ine [IO.Path]::GetFullPath($target)) {
        throw 'Word needs an absolute local working DOCX.'
    }
    $id = [WindowsDesktop]::FileId($target, $true)
    if ($id -ceq [WindowsDesktop]::FileId($source, $false) -or $id -ceq [WindowsDesktop]::FileId($snapshot, $true)) {
        throw 'The working document aliases its source or protected snapshot.'
    }
}
function Get-SelectedCode($doc, [string]$expected, [bool]$renderCopy, [string]$story) {
    $matchedField = $null
    $matchedCode = $null
    $ranges = [Collections.Generic.List[object]]::new()
    if (!$story -or $story -eq 'word/document.xml') { $ranges.Add($doc.Content) }
    else {
        foreach ($head in $doc.StoryRanges) {
            for ($part = $head; $null -ne $part; $part = $part.NextStoryRange) { $ranges.Add($part) }
        }
    }
    foreach ($range in $ranges) {
        foreach ($field in $range.Fields) {
            $code = [string]$field.Code.Text
            if ($code -notmatch '^\s*(?:ADDIN\s+)?ZOTERO_ITEM\s+CSL_CITATION\s*(\{[\s\S]+\})\s*$') { continue }
            if (($Matches[1] | ConvertFrom-Json).citationID -cne $expected) { continue }
            if ($null -ne $matchedField) { throw 'Word found duplicate matching citation IDs.' }
            $matchedField = $field
            $matchedCode = $code
        }
    }
    if ($null -eq $matchedField) { throw 'Word did not find the citation identified by Zotero.' }
    # The WinWord plugin version on this host does not position the cursor in select().
    # Position it in the exact field of the independently bound and active copy.
    if ($renderCopy -and $doc.TrackRevisions) { $doc.TrackRevisions = $false }
    $result = $matchedField.Result
    $result.Select()
    $sel = $script:word.Selection.Range
    if (!$sel.InStory($result) -or $sel.Start -ne $result.Start -or $sel.End -ne $result.End) {
        throw 'The selected range is not the requested Zotero field result.'
    }
    return $matchedCode
}
function Set-BibliographyAnchor($doc, [string]$paragraph) {
    $para = if ($paragraph -eq 'end') { $doc.Paragraphs.Last } else { $doc.Paragraphs.Item([int]$paragraph) }
    $range = $para.Range
    if ($range.StoryType -ne 1 -or [string]$range.Text -cne [string][char]13 -or $range.Information(12) -or
        $range.Fields.Count -or $range.Tables.Count -or $range.OMaths.Count) { throw 'The bibliography anchor is not an empty body paragraph.' }
    $start = $range.Start
    $range.Collapse(1)  # wdCollapseStart
    $range.Select()
    $selected = $script:word.Selection
    if ($selected.StoryType -ne 1 -or $selected.Start -ne $start -or $selected.End -ne $start) {
        throw 'Word moved the bibliography insertion point.'
    }
}
function Invoke-DocumentPreferences($doc) {
    # Only a positively identified Zotero tab/button in the bound Word window may be invoked.
    try {
        Add-Type -AssemblyName UIAutomationClient
        $element = [Windows.Automation.AutomationElement]
        $scope = [Windows.Automation.TreeScope]::Descendants
        $root = $element::FromHandle([IntPtr]($doc.ActiveWindow.Hwnd))
        $tabs = [Windows.Automation.PropertyCondition]::new($element::ControlTypeProperty,[Windows.Automation.ControlType]::TabItem)
        $tab = @($root.FindAll($scope,$tabs) | Where-Object { $_.Current.Name.Trim() -ieq 'Zotero' })
        if ($tab.Count -ne 1) { return 'unavailable' }
        $tab[0].GetCurrentPattern([Windows.Automation.SelectionItemPattern]::Pattern).Select()
        $buttons = [Windows.Automation.PropertyCondition]::new($element::ControlTypeProperty,[Windows.Automation.ControlType]::Button)
        $button = @($root.FindAll($scope,$buttons) | Where-Object {
            !$_.Current.IsOffscreen -and $_.Current.IsEnabled -and
            ($_.Current.AutomationId -match 'Zotero.*(SetDocPrefs|DocumentPreferences)' -or
             $_.Current.Name.Trim() -in @('Document Preferences','文档首选项','文档偏好设置','文件偏好設定','文件偏好設定值'))
        })
        if ($button.Count -ne 1) { return 'unavailable' }
        $invoke = $button[0].GetCurrentPattern([Windows.Automation.InvokePattern]::Pattern)
    } catch { return 'unavailable' }
    $null = Assert-ActiveDocument $doc.FullName
    try { $invoke.Invoke() } catch { return 'unknown' }
    return 'invoked'
}
function Invoke-WordAction([string]$action, [string[]]$values) {
    Connect-Word
    if ($action -eq 'connect') { return [int]$script:word.Documents.Count }
    if ($values.Count -lt 1) { throw 'Missing the bound document.' }
    $target = $values[0]
    $doc = Find-WorkingDocument $target
    if ($action -eq 'state') {
        if ($null -eq $doc) { return 'closed' }
        if ($doc.Saved) { return 'saved' }
        return 'unsaved'
    }
    if ($values.Count -lt 3) { throw 'Missing protected document paths.' }
    Assert-IndependentCopy $target $values[1] $values[2]
    if ($action -eq 'restore-tracking') {
        if ($null -eq $doc) { return 'closed' }
        if (!$doc.TrackRevisions) { $doc.TrackRevisions = $true }
        if (!$doc.TrackRevisions) { throw 'Word could not restore Track Changes on the working copy.' }
        return 'restored'
    }
    if ($action -in 'open','reopen') {
        if ($null -ne $doc) {
            if (!$doc.Saved) { throw 'The working copy has unsaved edits. Save or resolve them in Word before resuming.' }
            if ($action -eq 'reopen') {
                $doc.Close(-2)  # wdPromptToSaveChanges protects concurrent edits.
                if ($null -ne (Find-WorkingDocument $target)) { throw 'Word kept the document open.' }
                $doc = $null
            }
        }
        if ($null -eq $doc) { $doc = $script:word.Documents.Open($target, $false, $false, $false) }
        $doc.Activate()
        $doc = Assert-ActiveDocument $target
        $doc.Range(0, 0).Select()
        return 'open'
    }
    if ($action -in 'settle-save','settle-verify','close-saved','finish') {
        if ($null -eq $doc) {
            if ($action -in 'settle-save','settle-verify') { return 'closed' }
            throw 'The bound working copy was closed outside this operation.'
        }
        if (!$doc.Saved) {
            if ($action -in 'settle-verify','finish') { return 'left_open_unsaved' }
            throw 'The working copy has unsaved changes.'
        }
        $doc.Close(-2)
        if ($null -ne (Find-WorkingDocument $target)) { throw 'Word kept the document open.' }
        return 'closed'
    }
    if ($null -eq $doc) { throw 'The working copy is no longer open.' }
    if ($action -in 'focus','save','anchor','document-preferences') { $doc.Activate() }
    $doc = Assert-ActiveDocument $target
    switch ($action) {
        'document-preferences' { return Invoke-DocumentPreferences $doc }
        'tracking'  { return [bool]$doc.TrackRevisions }
        'selection' { return Get-SelectedCode $doc $values[3] ($values[4] -eq 'pause-tracking') $values[5] }
        'anchor'    { Set-BibliographyAnchor $doc $values[3] }
        'save'      { $doc.Save(); if (!$doc.Saved) { throw 'Word could not confirm the save.' } }
        'focus'     { }
        default     { throw 'Unknown Word operation.' }
    }
    $null = Assert-ActiveDocument $target
    return 'ok'
}

if ($ConnectOnly) {
    Connect-Word
    [Console]::WriteLine([int]$script:word.Documents.Count)
    Start-Zotero
    exit 0
}

while ($null -ne ($line = [Console]::ReadLine())) {
    $id = ''
    $script:runSent = $false
    $script:focusMethod = $null
    $script:focusRecovery = $null
    $script:phase = 'request'
    try {
        $request = $line | ConvertFrom-Json
        $id = [string]$request.id
        $action = [string]$request.action
        $values = [string[]]@($request.args)
        $script:phase = $action
        switch ($action) {
            'pages-open' {
                . (Join-Path $PSScriptRoot 'WordPages.ps1')
                Add-Type -AssemblyName System.Drawing
                $value = Open-PageReader $values[0]
            }
            'page-image' { $value = Write-PageImage ([int]$values[0]) $values[1] }
            'pages-close' { Close-PageReader; $value = 'closed' }
            'bootstrap'    { $value = Start-ZoteroRunner $values[0] $values[1] $values[2] }
            'launch-zotero' { Start-Zotero; $value = 'requested' }
            'zotero-alive' { $value = $null -ne (Get-ZoteroProcesses | Where-Object Id -eq ([int]$values[0]) | Select-Object -First 1) }
            default        { $value = Invoke-WordAction $action $values }
        }
        $reply = @{id=$id; ok=$true; value=$value}
    } catch {
        $cause = $_.Exception.GetBaseException()
        $reply = @{id=$id; ok=$false; error=$cause.Message; hresult=$cause.HResult.ToString('X8');
                   phase=$script:phase; runSent=$script:runSent; focusMethod=$script:focusMethod}
    }
    if ($script:focusRecovery) { $reply['focusRecovery'] = $script:focusRecovery }
    [Console]::WriteLine((ConvertTo-Json -InputObject $reply -Compress -Depth 5))
    [Console]::Out.Flush()
}
if ($null -ne $script:pageWord) { Close-PageReader }

# Windows PowerShell 5.1, on the bridge's STA. Only the newly opened Zotero window receives editor input.
$script:desktopSession = [Diagnostics.Process]::GetCurrentProcess().SessionId

function Initialize-DesktopTypes {
    if (!('WindowsDesktop' -as [type])) { Add-Type -Path (Join-Path $PSScriptRoot 'WindowsDesktop.cs') }
}
function Get-ZoteroProcesses {
    Get-Process -Name zotero -ErrorAction SilentlyContinue | Where-Object SessionId -eq $script:desktopSession
}
function Start-Zotero {
    if (Get-ZoteroProcesses) { return }
    $locations = @('HKCU:\Software\Microsoft\Windows\CurrentVersion\App Paths\zotero.exe',
                   'HKLM:\Software\Microsoft\Windows\CurrentVersion\App Paths\zotero.exe',
                   'HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\zotero.exe')
    $files = @(foreach ($key in $locations) {
        if (Test-Path -LiteralPath $key) { (Get-Item -LiteralPath $key).GetValue('') }
    })
    foreach ($base in @([Environment]::GetFolderPath('ProgramFiles'),[Environment]::GetFolderPath('ProgramFilesX86'),$env:LOCALAPPDATA)) {
        if ($base) { $files += (Join-Path $base 'Zotero\zotero.exe'), (Join-Path $base 'Programs\Zotero\zotero.exe') }
    }
    foreach ($file in $files) {
        if ($file -and [IO.File]::Exists($file)) { Start-Process -FilePath $file; return }
    }
    throw 'Cannot locate Zotero in the current desktop session. Open your installed Zotero and resume.'
}
function Get-ProcessWindows([int]$zid) {
    $filter = [Windows.Automation.PropertyCondition]::new([Windows.Automation.AutomationElement]::ProcessIdProperty,$zid)
    [Windows.Automation.AutomationElement]::RootElement.FindAll([Windows.Automation.TreeScope]::Children,$filter)
}
function Find-MenuItem($window, [string]$itemId, [string[]]$labels) {
    $element = [Windows.Automation.AutomationElement]
    $scope = [Windows.Automation.TreeScope]::Descendants
    $byId = [Windows.Automation.PropertyCondition]::new($element::AutomationIdProperty,$itemId)
    $item = $window.FindFirst($scope,$byId)
    if ($item) { return $item }
    $menuType = [Windows.Automation.PropertyCondition]::new($element::ControlTypeProperty,[Windows.Automation.ControlType]::MenuItem)
    foreach ($candidate in $window.FindAll($scope,$menuType)) {
        if (($candidate.Current.Name -replace '\s*\(&?.\)\s*$','' -replace '&','').Trim() -in $labels) { return $candidate }
    }
}
function Open-MenuItem($item, [bool]$opensRunner) {
    $pattern = $null
    if ($item.TryGetCurrentPattern([Windows.Automation.ExpandCollapsePattern]::Pattern,[ref]$pattern) -and
        $pattern.Current.ExpandCollapseState -ne [Windows.Automation.ExpandCollapseState]::LeafNode) {
        $invoke = { $pattern.Expand() }
    } elseif ($item.TryGetCurrentPattern([Windows.Automation.InvokePattern]::Pattern,[ref]$pattern)) {
        $invoke = { $pattern.Invoke() }
    } else {
        $item.SetFocus()
        $focused = [Windows.Automation.AutomationElement]::FocusedElement
        $front = [WindowsDesktop]::GetForegroundWindow()
        if (!$focused -or ![Windows.Automation.Automation]::Compare($focused,$item) -or
            [WindowsDesktop]::ProcessOf($front) -ne $item.Current.ProcessId) { throw 'Zotero did not focus its menu command.' }
        Assert-ZoteroFront $front $item.Current.ProcessId
        $invoke = { [Windows.Forms.SendKeys]::SendWait('{ENTER}') }
    }
    if ([WindowsDesktop]::ProcessOf([WindowsDesktop]::GetForegroundWindow()) -ne $item.Current.ProcessId) {
        throw 'The bound Zotero menu lost foreground focus.'
    }
    # After invoking Run JavaScript the command may already have opened its editor.
    if ($opensRunner) { $script:phase = 'window' }
    & $invoke
}
function Wait-ZoteroFront([IntPtr]$handle) {
    $limit = [Diagnostics.Stopwatch]::StartNew()
    while ($limit.ElapsedMilliseconds -lt 800 -and [WindowsDesktop]::GetForegroundWindow() -ne $handle) { Start-Sleep -Milliseconds 40 }
}
function Restore-ZoteroFront([IntPtr]$handle, [int]$zid) {
    if (!$script:canRecoverFocus -or $script:runSent) { throw 'The bound Zotero window lost foreground focus.' }
    $script:canRecoverFocus = $false
    $script:focusRecovery = @{ phase=$script:phase; restored=$false }
    [WindowsDesktop]::Activate($handle,$zid)
    Wait-ZoteroFront $handle
    if ([WindowsDesktop]::GetForegroundWindow() -ne $handle -and [WindowsDesktop]::ProcessOf($handle) -eq $zid) {
        $shell = New-Object -ComObject WScript.Shell
        $null = $shell.AppActivate($zid)
        Wait-ZoteroFront $handle
    }
    if ([WindowsDesktop]::GetForegroundWindow() -ne $handle -or [WindowsDesktop]::ProcessOf($handle) -ne $zid) {
        throw 'Windows did not return the bound Zotero window to the foreground.'
    }
    $script:focusRecovery.restored = $true
}
function Assert-ZoteroFront([IntPtr]$handle, [int]$zid) {
    if ($handle -eq [IntPtr]::Zero -or [WindowsDesktop]::ProcessOf($handle) -ne $zid) { throw 'The bound Zotero window changed.' }
    if ([WindowsDesktop]::GetForegroundWindow() -ne $handle) {
        if ($script:phase -eq 'menu') { throw 'The bound Zotero menu lost foreground focus.' }
        Restore-ZoteroFront $handle $zid
    }
}
function Activate-ZoteroWindow($window, [int]$zid) {
    $handle = [IntPtr]$window.Current.NativeWindowHandle
    [WindowsDesktop]::Activate($handle,$zid)
    if ([WindowsDesktop]::GetForegroundWindow() -ne $handle) {
        try { $window.SetFocus() } catch [System.InvalidOperationException] { }
        [WindowsDesktop]::Activate($handle,$zid)
    }
    Wait-ZoteroFront $handle
    Assert-ZoteroFront $handle $zid
    return $handle
}
function Send-EditorKey([IntPtr]$handle, [int]$zid, [string]$keys, [switch]$RunCommand) {
    Assert-ZoteroFront $handle $zid
    if ($RunCommand) { $script:runSent = $true }
    [Windows.Forms.SendKeys]::SendWait($keys)
}
function Get-RunnerWindow($windows, $existing) {
    foreach ($window in $windows) {
        $h = $window.Current.NativeWindowHandle
        if (!$h -or $existing.Contains([int]$h)) { continue }
        if ($window.Current.Name -match '(?:Run|Ex[eé]cuter|Ejecutar|执行|运行|執行|運行)\s*Java\s*Script|Java\s*Script\s+ausführen') { return $window }
        $idProperty = [Windows.Automation.AutomationElement]::AutomationIdProperty
        $isRunner = $true
        foreach ($id in 'run-as-async','help-button') {
            $condition = [Windows.Automation.PropertyCondition]::new($idProperty,$id)
            if (!$window.FindFirst([Windows.Automation.TreeScope]::Descendants,$condition)) { $isRunner = $false; break }
        }
        if ($isRunner) { return $window }
    }
}
function Test-OwnedEmptyEditor($window) {
    $handle = [IntPtr]$window.Current.NativeWindowHandle
    if (![WindowsDesktop]::WasFreshEditor($handle)) { return $false }
    if ([WindowsDesktop]::UntouchedEditor($handle)) { return $true }
    # The full editor document can prove emptiness. The hidden Ace input alone cannot.
    $id = [Windows.Automation.PropertyCondition]::new([Windows.Automation.AutomationElement]::AutomationIdProperty,'editor-code')
    $editor = $window.FindFirst([Windows.Automation.TreeScope]::Descendants,$id)
    $pattern = $null
    return ($null -ne $editor -and $editor.Current.ControlType -eq [Windows.Automation.ControlType]::Document -and
        $editor.TryGetCurrentPattern([Windows.Automation.TextPattern]::Pattern,[ref]$pattern) -and
        $pattern.DocumentRange.GetText(-1) -ceq '')
}
function Open-RunnerMenus($main, [IntPtr]$mainHandle, [int]$zid) {
    $wait = [Diagnostics.Stopwatch]::StartNew()
    $commands = @(
        @{id='toolsMenu'; labels=@('Tools','工具','Outils','Werkzeuge','Herramientas')},
        @{id='developer-menu'; labels=@('Developer','开发者','开发者工具','开发人员','開發者','Développeur','Entwickler')},
        @{id='menu_runJS'; labels=@('Run JavaScript','执行 JavaScript','运行 JavaScript','執行 JavaScript','執行 JavaScript 程式碼')}
    )
    $toolsKey = '%t'
    $toolsShortcutUsed = $false
    foreach ($command in $commands) {
        $wait.Restart()
        $item = $null
        $revealed = $false
        do {
            if ([WindowsDesktop]::ProcessOf([WindowsDesktop]::GetForegroundWindow()) -ne $zid) {
                throw 'The bound Zotero menu lost foreground focus.'
            }
            $item = Find-MenuItem $main $command.id $command.labels
            if (!$item -and $command.id -ne 'toolsMenu') {
                foreach ($popup in Get-ProcessWindows $zid) {
                    if ($popup.Current.ControlType -eq [Windows.Automation.ControlType]::Menu -or
                        $popup.Current.ClassName -match 'Popup|DropShadow') {
                        $item = Find-MenuItem $popup $command.id $command.labels
                        if ($item) { break }
                    }
                }
            }
            if (!$item -and !$revealed -and $command.id -eq 'toolsMenu') {
                Send-EditorKey $mainHandle $zid '{F10}'
                $revealed = $true
            }
            if (!$item -and !$toolsShortcutUsed -and $command.id -eq 'developer-menu' -and $wait.ElapsedMilliseconds -ge 450) {
                # The installed English and Chinese Tools menu has the T access key.
                # UI Automation can report Tools without showing its Gecko popup.
                Send-EditorKey $mainHandle $zid $toolsKey
                $toolsShortcutUsed = $true
            }
            if (!$item) { Start-Sleep -Milliseconds 100 }
        } while (!$item -and $wait.Elapsed.TotalSeconds -lt 3)
        if ($command.id -eq 'toolsMenu') {
            if (!$item) { Send-EditorKey $mainHandle $zid $toolsKey; $toolsShortcutUsed = $true; continue }
            if ($item.Current.AccessKey -match '^Alt\+([a-z0-9])$') { $toolsKey = '%' + $Matches[1].ToLowerInvariant() }
        }
        if (!$item) { throw ('Zotero did not expose menu item ' + $command.id + '.') }
        Open-MenuItem $item ($command.id -eq 'menu_runJS')
    }
}
function Open-Runner {
    $script:phase = 'process'
    Start-Zotero
    $wait = [Diagnostics.Stopwatch]::StartNew()
    $main = $null
    do {
        foreach ($p in Get-ZoteroProcesses) {
            if (!$p.MainWindowHandle.ToInt64()) { continue }
            $windows = @(Get-ProcessWindows $p.Id)
            foreach ($w in $windows) {
                if ($w.Current.NativeWindowHandle -and $w.Current.Name -match '(?i)\bZotero\b' -and
                    $w.Current.Name -notmatch 'Java\s*Script' -and
                    ($w.Current.NativeWindowHandle -eq $p.MainWindowHandle.ToInt64() -or $w.Current.ClassName -eq 'MozillaWindowClass')) {
                    $main = $w; $zid = $p.Id; break
                }
            }
            if ($main) { break }
        }
        if (!$main) { Start-Sleep -Milliseconds 150 }
    } while (!$main -and $wait.Elapsed.TotalSeconds -lt 14)
    if (!$main) { throw 'Zotero started without an available library window in this desktop session.' }
    $old = [Collections.Generic.HashSet[int]]::new()
    foreach ($w in $windows) {
        $null = $old.Add([int]$w.Current.NativeWindowHandle)
        if (Test-OwnedEmptyEditor $w) {
            $handle = Activate-ZoteroWindow $w $zid
            return @{handle=$handle; zid=$zid}
        }
    }
    $script:phase = 'menu'
    $mainHandle = [IntPtr]$main.Current.NativeWindowHandle
    try {
        $null = Activate-ZoteroWindow $main $zid
        Open-RunnerMenus $main $mainHandle $zid
    } catch {
        if (!$script:canRecoverFocus -or $script:runSent -or $script:phase -ne 'menu' -or
            [WindowsDesktop]::ProcessOf([WindowsDesktop]::GetForegroundWindow()) -eq $zid) { throw }
        Restore-ZoteroFront $mainHandle $zid
        Open-RunnerMenus $main $mainHandle $zid
    }
    $script:phase = 'window'
    $wait.Restart()
    do {
        $w = Get-RunnerWindow @(Get-ProcessWindows $zid) $old
        if ($w) {
            $handle = Activate-ZoteroWindow $w $zid
            [WindowsDesktop]::MarkFreshEditor($handle)
            return @{handle=$handle; zid=$zid}
        }
        Start-Sleep -Milliseconds 100
    } while ($wait.Elapsed.TotalSeconds -lt 6)
    throw 'Zotero did not create a new Run JavaScript window. Close an existing script editor after preserving its contents, then resume.'
}
function Start-ZoteroRunner([string]$loader, [string]$job, [string]$operation) {
    Initialize-DesktopTypes
    Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes, System.Windows.Forms
    if (!('ClipboardSnapshot' -as [type])) { Add-Type -Path (Join-Path $PSScriptRoot 'ClipboardSnapshot.cs') }
    $script:phase = 'clipboard'
    $script:canRecoverFocus = $operation -eq 'refresh'
    $control = [Windows.Forms.Control]::new()
    $backup = $null
    $owned = $null
    $jobName = [IO.Path]::GetFileName($job)
    try {
        $backup = [ClipboardSnapshot]::new($control.Handle)
        $owned = $backup.Stage($backup.Sequence,$loader,$jobName)
        $bound = Open-Runner
        $handle = $bound.handle
        $zid = $bound.zid
        $wanted = $loader.Replace([string][char]13,'')
        $echoed = $false
        foreach ($method in 'autofocus','left-pane','right-pane') {
            $script:phase = 'focus'
            $script:focusMethod = $method
            Assert-ZoteroFront $handle $zid
            if ($method -ne 'autofocus') {
                $fraction = if ($method -eq 'left-pane') { 0.25 } else { 0.75 }
                [WindowsDesktop]::ClickPane($handle,$zid,$fraction)
                Start-Sleep -Milliseconds 80
            }
            $script:phase = 'clipboard'
            Send-EditorKey $handle $zid '^a'
            if ($method -ne 'autofocus') { $owned = $backup.Stage($owned,$loader,$jobName) }
            $owned = $backup.ConfirmStage($loader,$jobName)
            [WindowsDesktop]::ReleaseEditor($handle)
            Send-EditorKey $handle $zid '^v'
            Start-Sleep -Milliseconds 100
            Send-EditorKey $handle $zid '^a'
            Send-EditorKey $handle $zid '^c'
            $wait = [Diagnostics.Stopwatch]::StartNew()
            do {
                $seq = [WindowsDesktop]::GetClipboardSequenceNumber()
                if ($seq -ne $owned -and ![ClipboardSnapshot]::IsClipboardFormatAvailable([ClipboardSnapshot]::Marker)) {
                    $owner = [WindowsDesktop]::GetClipboardOwner()
                    $echoed = [string]::Equals([Windows.Forms.Clipboard]::GetText().Replace([string][char]13,''),$wanted,[StringComparison]::Ordinal)
                    if (!$echoed -and ($owner -eq [IntPtr]::Zero -or [WindowsDesktop]::ProcessOf($owner) -ne $zid)) {
                        throw 'Another application changed the clipboard during editor readback.'
                    }
                    $settled = [WindowsDesktop]::GetClipboardSequenceNumber()
                    if ([WindowsDesktop]::GetClipboardOwner() -ne $owner -or
                        [WindowsDesktop]::GetClipboardSequenceNumber() -ne $settled -or
                        [ClipboardSnapshot]::IsClipboardFormatAvailable([ClipboardSnapshot]::Marker)) {
                        throw 'Clipboard changed while reading the editor echo.'
                    }
                    $owned = $settled
                    break
                }
                Start-Sleep -Milliseconds 40
            } while ($wait.ElapsedMilliseconds -lt 800)
            if ($echoed) { break }
        }
        if (!$echoed) { $script:phase = 'focus'; throw 'The new Zotero window did not return this job code when copied.' }
        $script:phase = 'dispatch'
        Send-EditorKey $handle $zid '^r' -RunCommand
    } finally {
        try { if ($null -ne $backup -and $null -ne $owned) { $backup.Restore($owned) } }
        finally { if ($null -ne $backup) { $backup.Dispose() }; $control.Dispose() }
    }
    return [int]$zid
}

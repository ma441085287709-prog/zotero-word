# The review reads the saved bytes in a separate read-only Word session, even if a customer's window has edits.
function Close-PageReader {
    try {
        if ($null -ne $script:pageDoc) { $script:pageDoc.Close(0) }
    } finally {
        $script:pageDoc = $null
        if ($null -ne $script:pageWord) {
            if ($script:pageWord.Documents.Count -eq 0) { $script:pageWord.Quit(0) }
            $script:pageWord = $null
        }
    }
}

function Open-PageReader([string]$target) {
    if ($null -ne $script:pageWord) { throw 'A page reader is already open.' }
    $private = New-Object -ComObject Word.Application
    if ($private.Documents.Count -ne 0) { throw 'Could not acquire an independent Word page reader.' }
    $script:pageWord = $private
    try {
        $private.AutomationSecurity = 3  # msoAutomationSecurityForceDisable in this private instance
        $script:pageDoc = $private.Documents.Open($target,$false,$true,$false)
        $doc = $script:pageDoc
        if (!$doc.ReadOnly -or ![string]::Equals($doc.FullName,$target,[StringComparison]::OrdinalIgnoreCase)) {
            throw 'The reader did not open the bound saved file read-only.'
        }
        $doc.ActiveWindow.View.Type = 3  # wdPrintView
        $doc.Repaginate()
        $pages = $doc.ActiveWindow.Panes.Item(1).Pages.Count
        if ($pages -lt 1) { throw 'Word returned no print-layout pages.' }
        $bibliographies = @($doc.Fields | Where-Object { [string]$_.Code.Text -match '^\s*(?:ADDIN\s+)?ZOTERO_BIBL\b' })
        if ($bibliographies.Count -ne 1) { throw 'Expected one bibliography in the accepted document.' }
        $field = $bibliographies[0].Result
        $rows = @(foreach ($para in $field.Paragraphs) {
            $r = $para.Range.Duplicate
            $start = [Math]::Max($field.Start,$r.Start)
            $end = [Math]::Min($field.End,$r.End)
            if ($end -le $start) { continue }
            $r.SetRange($start,$end)
            while ($r.End -gt $r.Start -and (@([string][char]13,[string][char]7) -contains [string]$r.Characters.Last.Text)) {
                $r.End = $r.End - 1
            }
            if ($r.Start -ge $r.End) { continue }
            $label = [regex]::Replace([string]$r.Text,'\s+',' ').Trim()
            if (!$label) { continue }
            $hash = [Security.Cryptography.SHA256]::Create()
            try { $key = [BitConverter]::ToString($hash.ComputeHash([Text.Encoding]::UTF8.GetBytes($label))).Replace('-','').ToLowerInvariant() }
            finally { $hash.Dispose() }
            $font = $r.Font
            $size = [double]$font.Size
            $format = $para.Format
            [ordered]@{text_sha256=$key; font=$font.Name; size_pt=$(if ($size -gt 0 -and $size -ne 9999999) { $size });
                first_character_size_pt=[double]$r.Characters.First.Font.Size;
                left_indent_pt=$format.LeftIndent; right_indent_pt=$format.RightIndent; first_line_indent_pt=$format.FirstLineIndent;
                space_before_pt=$format.SpaceBefore; space_after_pt=$format.SpaceAfter;
                line_spacing_rule=$format.LineSpacingRule; line_spacing_pt=$format.LineSpacing}
        })
        return @{page_count=[int]$pages; bibliography_paragraphs=$rows}
    } catch {
        Close-PageReader
        throw
    }
}

function Write-PageImage([int]$number,[string]$destination) {
    if ($null -eq $script:pageDoc -or !$script:pageDoc.ReadOnly) { throw 'Read-only page reader is not open.' }
    $page = $script:pageDoc.ActiveWindow.Panes.Item(1).Pages.Item($number)
    [byte[]]$bytes = $page.EnhMetaFileBits
    if (!$bytes.Length) { throw "Word returned an empty image for page $number." }
    $inputStream = [IO.MemoryStream]::new($bytes)
    $image = $bitmap = $graphics = $file = $null
    try {
        $image = [Drawing.Imaging.Metafile]::new($inputStream)
        $scale = 2000.0 / [Math]::Max($image.Width,$image.Height)
        $width = [Math]::Max(1,[int]($image.Width * $scale))
        $height = [Math]::Max(1,[int]($image.Height * $scale))
        $bitmap = [Drawing.Bitmap]::new($width,$height)
        $graphics = [Drawing.Graphics]::FromImage($bitmap)
        $graphics.Clear([Drawing.Color]::White)
        $graphics.DrawImage($image,[Drawing.Rectangle]::new(0,0,$width,$height))
        $file = [IO.File]::Open($destination,[IO.FileMode]::CreateNew,[IO.FileAccess]::Write,[IO.FileShare]::None)
        $bitmap.Save($file,[Drawing.Imaging.ImageFormat]::Png)
    } finally {
        foreach ($resource in @($file,$graphics,$bitmap,$image,$inputStream)) { if ($null -ne $resource) { $resource.Dispose() } }
    }
    return $number
}

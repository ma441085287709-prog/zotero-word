-- Save an interrupted, dirty working document under a new name; keep both versions.
on matchesFile(d, targetFile)
    tell application id "com.microsoft.Word"
        try
            if (path of d) is in {"", missing value} then return false
            set p to posix full name of d
        on error number -1728
            return false
        end try
    end tell
    if p is missing value then return false
    if p does not start with "/" then return false
    try
        return ((POSIX file p) as alias) = targetFile
    on error number -43
        return false
    end try
end matchesFile

on run argv
    set boundFile to (POSIX file (item 1 of argv)) as alias
    set outPath to item 2 of argv
    set newFile to POSIX file outPath
    try
        do shell script "/bin/test ! -e " & quoted form of outPath & " -a ! -L " & quoted form of outPath
    on error
        error "The recovery path already exists or cannot be checked safely."
    end try
    tell application id "com.microsoft.Word"
        set matched to missing value
        repeat with d in (get documents)
            if my matchesFile(d, boundFile) then
                set matched to contents of d
                exit repeat
            end if
        end repeat
        if matched is missing value then error "The bound Word working copy is no longer open."
        if read only of matched then error "The bound working document is read-only."
        if saved of matched then return "already-saved"
        set originalFormat to save format of matched
        save as matched file name outPath file format originalFormat add to recent files false
        if not (saved of matched) then error "The new recovery file has unconfirmed changes."
        set savedPath to posix full name of matched
    end tell
    if ((POSIX file savedPath) as alias) is not (newFile as alias) then error "Word saved to a different recovery path."
    return "saved"
end run

-- Read one saved bibliography without selecting or refreshing a field.
on isBibliographyCode(instruction)
    set k to offset of "ZOTERO_BIBL" in instruction
    if k is 0 then return false
    set prefixWords to {}
    if k > 1 then set prefixWords to words of (text 1 thru (k - 1) of instruction)
    return prefixWords = {} or prefixWords = {"ADDIN"}
end isBibliographyCode

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

on openCopy(targetFile)
    tell application id "com.microsoft.Word"
        repeat with d in (get documents)
            if my matchesFile(d, targetFile) then return contents of d
        end repeat
    end tell
    return missing value
end openCopy

on measure(d)
	tell application id "com.microsoft.Word"
		set instructions to {}
		repeat with f in (get fields of d)
			set end of instructions to get content of field code of f
		end repeat
		set bibIndex to 0
		repeat with i from 1 to (count instructions)
			set instruction to item i of instructions as text
            if my isBibliographyCode(instruction) then
					if bibIndex is not 0 then error "Word found multiple bibliography fields."
					set bibIndex to i as integer
			end if
		end repeat
		if bibIndex is 0 then error "Word did not find the saved bibliography."
		set bib to result range of field bibIndex of d
		set fieldText to content of bib
		set visibleEnd to end of content of bib
		repeat while fieldText ends with return or fieldText ends with (character id 7)
			if (length of fieldText) < 2 then error "Word returned an empty bibliography."
			set fieldText to text 1 thru -2 of fieldText
			set visibleEnd to visibleEnd - 1
		end repeat
		set fieldStart to start of content of bib
		set visibleRange to create range d start fieldStart end visibleEnd
		set measuredPt to font size of font object of visibleRange
		if measuredPt is missing value then error "Word could not resolve the bibliography font size."
		if measuredPt ≤ 0 or measuredPt ≥ 9999999 then error "Word reports mixed or undefined bibliography font sizes."
		-- Word stores text sizes in half points. Integers travel unchanged across locales.
		set halfPoints to (measuredPt * 2) as integer
		if halfPoints / 2 is not measuredPt then error "Word returned an unsupported font size."
		return (halfPoints as text) & linefeed & fieldText
	end tell
end measure

on run argv
	set targetFile to (POSIX file (item 1 of argv)) as alias
	set owned to false
	set d to missing value
	with timeout of 10 seconds
		tell application id "com.microsoft.Word"
			set d to my openCopy(targetFile)
			if d is not missing value then
				if not (saved of d) then return "OPEN_UNSAVED"
			end if
			try
				if d is missing value then
					open targetFile confirm conversions false read only true add to recent files false
					set d to my openCopy(targetFile)
					if d is missing value then error "Word did not open the bound file."
					if not (read only of d) then error "Word did not open the bound file read-only."
					set owned to true
				end if
				set reply to my measure(d)
				if not (saved of d) then error "The Word window changed while its bibliography was being read."
			on error msg number n
				if owned and d is not missing value then
					if saved of d then close d saving no
				end if
				error msg number n
			end try
			if owned then close d saving ask
			return reply
		end tell
	end timeout
end run

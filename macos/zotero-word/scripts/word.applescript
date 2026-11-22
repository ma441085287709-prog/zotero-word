-- Word supplies the path and selection independently of Zotero's MacWord handle.
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

on openCopy(targetFile, targetName)
	tell application id "com.microsoft.Word"
		try
			set d to get document targetName
		on error number -1728
			return missing value
		end try
		if not my matchesFile(d, targetFile) then error "The open Word document with this name has a different path."
		return d
	end tell
end openCopy

on boundDocument(targetFile, targetName, mayActivate)
	tell application id "com.microsoft.Word"
		set d to active document
		if not my matchesFile(d, targetFile) then
			if not mayActivate then error "Word's active document is not the bound working copy."
			set d to my openCopy(targetFile, targetName)
			if d is missing value then error "The bound working copy is no longer open."
			activate object d
			set d to active document
			if not my matchesFile(d, targetFile) then error "Word did not activate the requested working copy."
		end if
		if read only of d then error "Word opened the working copy read-only."
		return d
	end tell
end boundDocument

on selectedCode(s)
	tell application id "com.microsoft.Word"
		set lo to selection start of s
		set hi to selection end of s
		set area to text object of paragraph 1 of s
		repeat with f in (get fields of area)
			set valRange to result range of f
			set codeRange to field code of f
			if story type of valRange is story type of s then
				if lo ≥ (start of content of codeRange) - 1 and hi ≤ (end of content of valRange) + 1 and hi ≥ (start of content of valRange) then
					set instruction to content of codeRange
					if instruction contains "ZOTERO_ITEM" then return instruction
				end if
			end if
		end repeat
	end tell
	error "Word did not select a Zotero field in the bound document."
end selectedCode

on run argv
	set actionName to item 1 of argv
	if actionName = "connect" then
		-- Explicit legacy diagnostic; the manuscript workflow opens the supplied file directly.
		tell application id "com.microsoft.Word"
			launch
			set openDocs to get documents
			return count of openDocs
		end tell
	end if
	set targetFile to (POSIX file (item 2 of argv)) as alias
	set targetName to name of (info for targetFile)
	set paragraphMark to character id 13
	tell application id "com.microsoft.Word"
		if actionName = "state" then
			set d to my openCopy(targetFile, targetName)
			if d is missing value then return "closed"
			if saved of d then return "saved"
			return "unsaved"
		end if
		if actionName is in {"settle-save", "settle-verify"} then
			set d to my openCopy(targetFile, targetName)
			if d is missing value then return "closed"
			if not (saved of d) then
				if actionName is "settle-verify" then return "left_open_unsaved"
				error "The rendered working copy has unsaved changes."
			end if
			close d saving ask
			return "closed"
		end if
		if actionName is in {"open", "reopen"} then
			if actionName = "reopen" then
				set d to my openCopy(targetFile, targetName)
				if d is not missing value then
					if not (saved of d) then error "The bound working copy has unsaved edits. Preserve them before continuing."
					close d saving ask
				end if
			end if
			open targetFile confirm conversions false read only false add to recent files false revert false
			set d to my boundDocument(targetFile, targetName, false)
			if not (saved of d) then error "The bound working copy has unsaved edits. Preserve them before continuing."
			set beginningRange to create range d start 0 end 0
			«event miscslct» beginningRange
			return
		end if
		set mayActivate to (actionName is in {"focus", "save", "close-saved", "anchor"})
		set d to my boundDocument(targetFile, targetName, mayActivate)
		if actionName = "focus" then return
		if actionName = "selection" then
			return my selectedCode(selection)
		else if actionName = "anchor" then
			set idx to item 3 of argv
			if idx = "end" then
				set para to last paragraph of d
			else
				set para to paragraph (idx as integer) of d
			end if
			set r to text object of para
			if story type of r is not main text story or (content of r) is not paragraphMark then error "Bibliography anchor is not an empty body paragraph."
			if (count fields of r) > 0 or (count math objects of r) > 0 or (count tables of r) > 0 then error "Bibliography anchor contains a protected object."
			set inTable to (get range information r information type with in table) as text
			if inTable is not in {"false", "0", "no"} then error "Word cannot confirm an anchor outside tables."
			set pos to start of content of r
			set pointRange to collapse range r direction collapse start
			«event miscslct» pointRange
			set s to selection
			if story type of s is not main text story or selection start of s is not pos or selection end of s is not pos then error "Word moved the bibliography insertion point."
			my boundDocument(targetFile, targetName, false)
		else if actionName = "save" then
			save d
			if not (saved of d) then error "Word did not confirm saving the working copy."
			my boundDocument(targetFile, targetName, false)
		else if actionName = "close-saved" then
			if not (saved of d) then error "Working copy changed after its confirmed save."
			close d saving ask
		else if actionName = "finish" then
			if not (saved of d) then return "left_open_unsaved"
			close d saving ask
			return "closed"
		else
			error "Unknown Word operation."
		end if
	end tell
end run

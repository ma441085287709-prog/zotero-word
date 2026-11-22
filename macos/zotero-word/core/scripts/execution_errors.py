"""Stable recovery hints for HOLD receipts; never grant automatic retries."""
from __future__ import annotations

import errno
from zipfile import BadZipFile
import xml.etree.ElementTree as ET


def explain(error, *, command=None):
    explicit = getattr(error, "reason_code", None)
    if explicit:
        details = {"reason_code": explicit,
                   "next_action": getattr(error, "next_action", "RECHECK_CURRENT_STATE"),
                   "retry_allowed": getattr(error, "retry_allowed", False)}
        if getattr(error, "diagnostic", None):
            details["diagnostic"] = error.diagnostic
        return details
    if isinstance(error, PermissionError) or (isinstance(error, OSError) and error.errno in (errno.EPERM, errno.EACCES)):
        code, action = "ACCESS_DENIED", "CHECK_EXISTING_ACCESS_AND_AVAILABLE_FALLBACK"
    elif isinstance(error, ImportError):
        code, action = "CAPABILITY_MISSING", "DISCOVER_EXISTING_COMPATIBLE_RUNTIME"
    elif isinstance(error, FileNotFoundError):
        code, action = "INPUT_MISSING", "LOCATE_REQUIRED_INPUT"
    elif isinstance(error, (BadZipFile, ET.ParseError)):
        code, action = "DOCUMENT_INVALID", "INSPECT_OR_RESTORE_SAVED_DOCUMENT"
    elif isinstance(error, TimeoutError):
        code, action = "OUTCOME_UNKNOWN", "RECHECK_OPERATION_OUTCOME_BEFORE_RETRY"
    elif type(error).__name__ == "CacheInputChanged":
        code, action = "INPUT_CHANGED", "RECHECK_SAVED_STATE"
    else:
        # Legacy guards have human-readable messages. Preserve the actual HOLD
        # without guessing from its text that a write may safely be repeated.
        code, action = "EVIDENCE_INVALID_OR_INCOMPLETE", "RECHECK_REQUIRED_EVIDENCE"
    return {"reason_code": code, "next_action": action, "retry_allowed": False}

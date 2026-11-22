"""Recover an interrupted Mac job using its receipts and an independent Word save."""
import os
import re
import time
from zipfile import ZipFile


def recovery_path(job):
    return job / ("word-recovery-" + job.name.removeprefix("native-") + ".docx")


def _request(cfg, destination):
    return {"document": cfg["document"], "prepared_sha256": cfg["documentHash"], "source_sha256": cfg["sourceHash"],
            "recovery_copy": str(destination)}


def purpose(job, cfg):
    if cfg["reopenOnly"]:
        return "verification"
    if (job / "refreshed.json").is_file() or (job / "refresh-target.ok").is_file():
        return "refresh"
    return "before_refresh"


def check_request(job, cfg, destination, api):
    record = api.load(job / "recovery-request.json")
    expected = _request(cfg, destination)
    if (destination != recovery_path(job) or any(record.get(k) != v for k, v in expected.items())
            or record.get("purpose", purpose(job, cfg)) != purpose(job, cfg)):
        raise ValueError("The recovery request belongs to another file or operation: " + str(destination))
    return record


def await_dispatched(job, state, api):
    """The old controller's stop request alone cannot prove its Word command ended."""
    cfg = api.load(job / "job.json")
    verifying = cfg["reopenOnly"]
    terminal = "verified.json" if verifying else "refreshed.json"
    if not (job / "started.json").is_file():
        # No command on Word can begin before the controller acknowledges the document selection.
        (job / "abort").touch(exist_ok=True)
        return
    if (job / terminal).is_file() or (job / "error.json").is_file():
        return
    request = api.load(job / "request.json")
    session = request.get("session", "")
    if (not re.fullmatch(r"native-[0-9a-f]{32}", session) or request != api.load(job / "started.json")
            or request.get("job") != job.name or request.get("operation") != ("reopen_verify" if verifying else "refresh")
            or request.get("document") != cfg["document"]):
        raise ValueError("The interrupted native command does not match this working copy")
    last = state.get("zotero_connection", {})
    pid = state.get("native_sessions", {}).get(session) or (last.get("pid") if last.get("session") == session else None)
    deadline = time.monotonic() + 35
    labels = ("reattached",) if verifying else ("attached", "refresh-target", "bibliography-inserted")
    while not (job / terminal).is_file() and not (job / "error.json").is_file():
        if type(pid) is int and pid > 0:
            try:
                os.kill(pid, 0)
            except ProcessLookupError:
                return
        # A pending handshake depends on the controller that previously held the run lock.
        # Withdraw only that handoff; an already issued refresh must report its own outcome.
        for label in labels:
            if (job / (label + ".json")).is_file() and not (job / (label + ".ok")).is_file():
                (job / "abort").touch(exist_ok=True)
                break
        if time.monotonic() >= deadline:
            raise TimeoutError("Zotero has not confirmed the end of the previous command. The working copy and its run are preserved.")
        time.sleep(0.2)


def capture_recovery(job, cfg, api, guard, connection):
    """Caller owns the desktop lock. An unanswered Save As keeps its single destination."""
    bound = guard.independent(cfg["document"], cfg["source"], cfg["snapshot"])
    dest = guard.independent(recovery_path(job), bound, cfg["source"], cfg["snapshot"])
    if (guard.sha(cfg["source"]) != cfg["sourceHash"] or guard.sha(cfg["snapshot"]) != cfg["sourceHash"]
            or (cfg["reopenOnly"] and guard.sha(bound) != cfg["documentHash"])):
        raise ValueError("The recovery save is not bound to this document and protected original")
    request = job / "recovery-request.json"
    record = check_request(job, cfg, dest, api) if request.is_file() else None

    def saved_destination():
        if not request.is_file():
            raise ValueError("No bound Word save request for the existing file: " + str(dest))
        guard.independent(dest, bound, cfg["source"], cfg["snapshot"])
        with ZipFile(dest) as archive:
            archive.getinfo("word/document.xml")
        if connection.word("state", dest) == "unsaved" or connection.word("state", bound) == "unsaved":
            raise ValueError("Additional unsaved edits remain in a bound Word window. Retained copy: " + str(dest))
        return dest

    if dest.exists():
        return saved_destination()
    if record:
        window = connection.word("state", bound)
        if window == "saved" or (window == "closed" and
                (record.get("status") == "saved_in_place" or guard.sha(bound) != cfg["documentHash"])):
            return bound
        if record.get("status") == "awaiting_word":
            raise TimeoutError("Word has an unconfirmed Save As. Finish its open save or access prompt, then run-next: " + str(dest))
        if window == "closed":
            return bound
    record = {**_request(cfg, dest), "purpose": purpose(job, cfg), "status": "awaiting_word"}
    api.output(request, record)
    bridge = api._helper("mac_connection")
    try:
        result = bridge._osa("word-save-recovery.applescript", bound, dest, timeout=90)
    except bridge.DesktopError as exc:
        if dest.is_file():
            return saved_destination()
        if exc.reason_code == "WORD_TIMEOUT" or "-1712" in str(exc):
            raise TimeoutError("Word has not acknowledged its save. Resume this run after the access prompt: " + str(dest)) from exc
        api.output(request, {**record, "status": "failed"})
        raise bridge._script_error("word.applescript", "save-recovery", str(exc)) from exc
    if result == "already-saved":
        api.output(request, {**record, "status": "saved_in_place"})
        return bound
    if result != "saved" or not dest.is_file():
        raise ValueError("Word has not confirmed the recovery save: " + str(dest))
    return saved_destination()

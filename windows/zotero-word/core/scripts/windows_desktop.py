"""Coordinate Word COM with the WinWord integration in the user's Zotero."""
import json
import os
import re
import time
from xml.etree import ElementTree as ET
from zipfile import BadZipFile, ZipFile


def _bound_copy(cfg, guard):
    target = guard.independent(cfg["document"], cfg["source"], cfg["snapshot"])
    guard.independent(cfg["snapshot"], cfg["source"], target)
    if not target.is_file() or not os.access(target, os.W_OK):
        raise ValueError("Word requires the existing writable working copy.")
    return target


def _selected(job, label, target, api, connection, cfg, expected):
    request = connection.wait(job, label + ".json", 185 if label == "after-preferences" else 90)
    if request["citation"] != expected:
        raise ValueError("Zotero requested a different readback sample")
    story = next(row["story"] for row in cfg["citations"] if row["id"] == expected)
    suspend = label == "attached" and connection.word("tracking", target)
    if suspend:
        api.output(job / "tracking-paused.json", {"document": str(target), "before": cfg["documentHash"]})
    code = connection.word("selection", target, expected, "pause-tracking" if suspend else "select", story)
    match = re.fullmatch(r"\s*(?:ADDIN\s+)?ZOTERO_ITEM\s+CSL_CITATION\s*(\{.*\})\s*", code, re.I | re.S)
    if not match or json.loads(match[1]).get("citationID") != request["citation"]:
        raise ValueError("Word and Zotero selected different citations or documents.")
    count = request.get("bibliographyCount")
    if type(count) is not int or count not in (0, 1):
        raise ValueError("Zotero could not identify a unique bibliography in the bound document.")
    api.output(job / (label + ".ok"), True)


def restore_tracking(job, cfg, target, api, connection):
    marker = job / "tracking-paused.json"
    if marker.is_file() and not (job / "tracking-restored.json").exists():
        if cfg["reopenOnly"] or api.load(marker) != {"document": str(target), "before": cfg["documentHash"]}:
            raise ValueError("Tracking state belongs to another working copy")
        connection.bind_copy(target, cfg)
        window = connection.word("restore-tracking", target)
        if window == "closed":
            # A host may exit between restoring Track Changes and recording the restore.
            ns = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
            with ZipFile(target) as docx:
                settings = ET.fromstring(docx.read("word/settings.xml"))
            flag = settings.find(ns + "trackRevisions")
            if flag is None or flag.get(ns + "val", "true") in ("false", "0", "off"):
                raise ValueError("The working copy closed before its original Track Changes setting was restored")
        api.output(job / "tracking-restored.json", {"window": window, "document": str(target)})


def saved_fields(document, style, rows, api):
    report = api.inspect_docx(document, fields_only=True)
    if report["issues"] or ((report.get("preferences") or {}).get("style") or {}).get("id") != style:
        raise ValueError("Saved fields or document style differ from the run: " + ", ".join(report["issues"]))
    bibs = report["bibliographies"]
    if len(bibs) != 1 or bibs[0]["display"].strip() in ("", "Zotero bibliography"):
        raise ValueError("The saved copy needs one rendered bibliography")
    expected = {row["id"]: row["items"] for row in rows}
    if len(expected) != len(rows) or len(report["citations"]) != len(rows):
        raise ValueError("Saved citation coverage differs from the bound plan")
    def identities(uris):
        return {uri for value in uris if (uri := api.canonical_uri(value))}
    seen = set()
    for row in report["citations"]:
        cid = row["citationID"]
        wanted = expected.get(cid)
        items = row["citationItems"]
        if cid in seen or not wanted or len(items) != len(wanted) or any(
                not identities(item["uris"]).intersection(identities(aliases)) for item, aliases in zip(items, wanted)):
            raise ValueError("Saved citation identity or item order changed: " + cid)
        seen.add(cid)
    citations = [[row["citationID"], row["display"]] for row in report["citations"]]
    if not citations or any(text.strip() in ("", "{Citation}") for _, text in citations):
        raise ValueError("The saved copy has unrendered citations")
    return {"origin": "saved_copy", "document_sha256": report["document_sha256"],
            "citations": citations, "text": bibs[0]["display"]}


def _same_render(left, right):
    # Word's field text includes paragraph marks and tabs omitted by the DOCX reader.
    def visible(value):
        return "".join(value.split())
    def citations(render):
        return sorted((cid, visible(text)) for cid, text in render["citations"])
    return visible(left["text"]) == visible(right["text"]) and citations(left) == citations(right)


def await_dispatched(job, api, connection):
    """Wait for native completion or the actual worker to exit before reissuing anything."""
    issued = ("preferences-issued.json", "refresh-issued.json", "refresh-target.ok")
    if (job / "refreshed.json").is_file() or not any((job / name).is_file() for name in issued):
        return
    request = api.load(job / "request.json")
    session = request.get("session", "")
    if not re.fullmatch(r"native-[0-9a-f]{32}", session):
        raise ValueError("The previous Zotero session has no valid identity")
    folder = job.parent / session
    end = time.monotonic() + 20
    while not any((job / name).is_file() for name in ("refreshed.json", "error.json")):
        if (folder / "session-ended.json").is_file():
            break
        if time.monotonic() >= end:
            process = folder / "process.json"
            if process.is_file() and not connection.process_alive(api.load(process)["pid"]):
                break
            raise TimeoutError("The dispatched Zotero command has no terminal result yet. Preserve the working copy and resume this run.")
        time.sleep(0.2)


def _disk_matches_refresh(target, cfg, job, api):
    try:
        render = saved_fields(target, cfg["style"], cfg["citations"], api)
    except (ValueError, OSError, BadZipFile):
        return False
    return _same_render(render, api.load(job / "refreshed.json"))


def _save_and_close(job, cfg, target, api, guard, connection, *, resume=False):
    saving = time.monotonic()
    request = {"document": str(target), "before": cfg["documentHash"]}
    saved_file = job / "saved.json"
    restore_tracking(job, cfg, target, api, connection)
    if resume and (job / "saving.json").is_file() and api.load(job / "saving.json") != request:
        raise ValueError("The Word save request belongs to a different working copy")
    window = connection.word("state", target) if resume else "unsaved"
    saved = api.load(saved_file) if resume and saved_file.is_file() else None
    if saved is not None:
        digest = guard.sha(target)
        if saved.get("document") != str(target):
            raise ValueError("The Word save receipt belongs to another document")
        if window != "closed" and saved.get("sha256") != digest:
            raise ValueError("The working copy differs from the confirmed Word save")
    elif window == "unsaved":
        api.output(job / "saving.json", request)
        connection.word("save", target, timeout=90)
        _bound_copy(cfg, guard)
        saved = {"document": str(target), "sha256": guard.sha(target),
                 "save_seconds": round(time.monotonic() - saving, 3)}
        api.output(saved_file, saved)
    else:
        if not _disk_matches_refresh(target, cfg, job, api):
            if window == "closed" and guard.sha(target) == cfg["documentHash"]:
                return False
            raise ValueError("The disk fields differ from the confirmed Zotero refresh")
        saved = {"document": str(target), "sha256": guard.sha(target), "method": "disk_matches_native_refresh"}
        api.output(saved_file, saved)
    if window != "closed" and connection.word("settle-save", target) != "closed":
        raise ValueError("Word did not confirm closing the working copy")
    digest = guard.sha(target)
    if saved["sha256"] != digest:
        # Some Word versions write once more during Close. Bind the result actually left on disk.
        if not _disk_matches_refresh(target, cfg, job, api):
            raise ValueError("The closed file differs from the confirmed Zotero fields")
        api.output(saved_file, {**saved, "before_close_sha256": saved["sha256"], "sha256": digest})
    api.output(job / "closed.json", {"document": str(target), "sha256": digest,
               "save_close_seconds": round(time.monotonic() - saving, 3), "resumed": resume})
    return True


def settle(job, api, guard, connection, *, verification=False):
    """Resume the recorded Word operation without replaying a plugin command."""
    cfg = api.load(job / "job.json")
    with connection.desktop_lock(job):
        target = _bound_copy(cfg, guard)
        connection.bind_copy(target, cfg)
        if guard.sha(cfg["source"]) != cfg["sourceHash"] or guard.sha(cfg["snapshot"]) != cfg["sourceHash"]:
            raise ValueError("Protected source changed before finishing the recorded operation")
        if not verification:
            if not (job / "refreshed.json").is_file() or not (job / "attached.ok").is_file():
                raise ValueError("Missing bound native refresh")
            return _save_and_close(job, cfg, target, api, guard, connection, resume=True)
        digest = guard.sha(target)
        verified = api.load(job / "verified.json")
        if (not (job / "reattached.ok").is_file() or verified.get("sha256") != digest
                or verified.get("document") != str(target) or cfg["documentHash"] != digest):
            raise ValueError("Interrupted verification belongs to another saved file")
        window = connection.word("settle-verify", target)
        if guard.sha(target) != digest:
            raise ValueError("The saved artifact changed while settling its Word window")
        api.output(job / "finished.json", {"document": str(target), "sha256": digest,
                                           "word_window": window, "resumed": True})
        return True


def execute(job, api, guard, connection):
    cfg = api.load(job / "job.json")
    with connection.desktop_lock(job):
        opened = False
        try:
            if cfg["libraryOnly"]:
                connection.dispatch(job, cfg)
                return connection.wait(job, "library.json", 90)
            target = _bound_copy(cfg, guard)
            connection.bind_copy(target, cfg)
            if (guard.sha(target) != cfg["documentHash"] or guard.sha(cfg["source"]) != cfg["sourceHash"]
                    or guard.sha(cfg["snapshot"]) != cfg["sourceHash"]):
                raise ValueError("The bound original, snapshot or working copy has changed.")
            opened_at = time.monotonic()
            connection.word("reopen" if cfg["reopenOnly"] else "open", target)
            opened = True
            connection.dispatch(job, cfg)
            saved_hash = cfg["documentHash"]
            first = next((row for row in cfg["citations"] if row["story"] == "word/document.xml"), cfg["citations"][0])["id"]
            if not cfg["reopenOnly"]:
                _selected(job, "attached", target, api, connection, cfg, first)
                settings = connection.wait(job, "settings.json", 90)
                if (settings.get("document") != str(target) or settings.get("style") != cfg["style"]
                        or type(settings.get("needed")) is not bool):
                    raise ValueError("The settings request belongs to another document or style")
                if settings["needed"]:
                    api.output(job / "preferences-issued.json", {"document": str(target), "style": cfg["style"]})
                    result = connection.word("document-preferences", target, timeout=185)
                    api.output(job / "settings.ok", result)
                    _selected(job, "after-preferences", target, api, connection, cfg, first)
                request = connection.wait(job, "refresh-target.json", 90)
                bibliography_count = request.get("bibliographyCount")
                if type(bibliography_count) is not int or bibliography_count not in (0, 1):
                    raise ValueError("The bound working copy has an unexpected bibliography count")
                _bound_copy(cfg, guard)
                if bibliography_count == 0:
                    if cfg["anchor"] is None:
                        raise ValueError("Missing bibliography and no bound insertion point in the working copy.")
                    connection.word("anchor", target, "end" if cfg["anchorEnd"] else cfg["anchor"])
                else:
                    connection.word("focus", target)
                api.output(job / "refresh-target.ok", True)
                if bibliography_count == 0:
                    connection.wait(job, "bibliography-inserted.json", 90)
                    connection.word("focus", target)
                    api.output(job / "bibliography-inserted.ok", True)
                connection.wait(job, "refreshed.json", 150)
                _bound_copy(cfg, guard)
                _save_and_close(job, cfg, target, api, guard, connection)
                opened = False
                return
            _selected(job, "reattached", target, api, connection, cfg, first)
            for index, citation in enumerate(cfg["samples"]):
                _selected(job, "sample-" + str(index), target, api, connection, cfg, citation)
            connection.wait(job, "verified.json", 120)
            _bound_copy(cfg, guard)
            window = connection.word("finish", target)
            if guard.sha(target) != saved_hash or guard.sha(cfg["source"]) != cfg["sourceHash"]:
                raise ValueError("The saved document or original changed during reopen verification.")
            api.output(job / "finished.json", {"document": str(target), "sha256": saved_hash, "word_window": window,
                                               "reopen_verify_seconds": round(time.monotonic() - opened_at, 3)})
        except BaseException:
            (job / "abort").touch(exist_ok=True)
            if opened and not cfg["reopenOnly"]:
                try:
                    restore_tracking(job, cfg, target, api, connection)
                except (OSError, ValueError):
                    pass  # A disconnected host leaves its restoration receipt pending for the same run.
            if opened and (cfg.get("reopenOnly") or not (job / "attached.ok").exists()):
                try:
                    connection.word("finish", target)
                except (OSError, ValueError):
                    pass  # Preserve any unsaved document or changed active window.
            raise

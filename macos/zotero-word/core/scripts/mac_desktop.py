"""Coordinate Word Apple events with the MacWord integration in the user's Zotero."""
import json
import os
import re
import time
from zipfile import BadZipFile


def _bound_copy(cfg, guard):
    target = guard.independent(cfg["document"], cfg["source"], cfg["snapshot"])
    guard.independent(cfg["snapshot"], cfg["source"], target)
    if not target.is_file() or not os.access(target, os.W_OK):
        raise ValueError("Word requires the existing writable working copy.")
    return target


def _selected(job, label, target, api, connection):
    request = connection.wait(job, label + ".json", 90)
    code = connection.word("selection", target)
    match = re.fullmatch(r"\s*(?:ADDIN\s+)?ZOTERO_ITEM\s+CSL_CITATION\s*(\{.*\})\s*", code, re.I | re.S)
    if not match or json.loads(match[1]).get("citationID") != request["citation"]:
        raise ValueError("Word and Zotero selected different citations or documents.")
    count = request.get("bibliographyCount")
    if type(count) is not int or count not in (0, 1):
        raise ValueError("Zotero could not identify a unique bibliography in the bound document.")
    api.output(job / (label + ".ok"), True)


class IncompleteRender(ValueError):
    """A valid working document still needs its target style or field results."""


def saved_fields(document, style, rows, api):
    report = api.inspect_docx(document, fields_only=True)
    if report["issues"]:
        raise ValueError("The saved copy has field errors: " + ", ".join(report["issues"]))
    bibs = report["bibliographies"]
    if len(bibs) > 1:
        raise ValueError("The saved copy contains multiple bibliographies")
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
    if ((report.get("preferences") or {}).get("style") or {}).get("id") != style:
        raise IncompleteRender("The target document style has not been applied")
    if not bibs or bibs[0]["display"].strip() in ("", "Zotero bibliography"):
        raise IncompleteRender("The saved copy needs one rendered bibliography")
    if not citations or any(text.strip() in ("", "{Citation}") for _, text in citations):
        raise IncompleteRender("The saved copy has unrendered citations")
    return {"origin": "saved_copy", "document_sha256": report["document_sha256"],
            "citations": citations, "text": bibs[0]["display"]}


def _same_render(left, right):
    # Word's field text includes paragraph marks and tabs omitted by the DOCX reader.
    def visible(value):
        return "".join(value.split())
    def citations(render):
        return sorted((cid, visible(text)) for cid, text in render["citations"])
    return visible(left["text"]) == visible(right["text"]) and citations(left) == citations(right)


def _disk_matches_refresh(target, cfg, job, api):
    try:
        render = saved_fields(target, cfg["style"], cfg["citations"], api)
    except (ValueError, OSError, BadZipFile):
        return False
    return _same_render(render, api.load(job / "refreshed.json"))


def _save_and_close(job, cfg, target, api, guard, connection, *, resume=False):
    """Finish a confirmed refresh on the same Word document. Caller owns the desktop lock."""
    saving = time.monotonic()
    request = {"document": str(target), "before": cfg["documentHash"]}
    saved_path = job / "saved.json"
    window = None
    if resume:
        if (job / "saving.json").is_file() and api.load(job / "saving.json") != request:
            raise ValueError("Word save request belongs to another working copy")
        digest = guard.sha(target)
        if saved_path.is_file():
            saved = api.load(saved_path)
            if saved.get("document") != str(target) or saved.get("sha256") != digest:
                raise ValueError("The confirmed Word save differs from the working copy")
        else:
            window = connection.word("state", target)
            if window != "unsaved":
                if not _disk_matches_refresh(target, cfg, job, api):
                    if window == "closed" and digest == cfg["documentHash"]:
                        return False
                    raise ValueError("The saved fields differ from the confirmed Zotero refresh")
                api.output(saved_path, {"document": str(target), "sha256": digest, "method": "disk_matches_native_refresh"})
    if not resume or window == "unsaved":
        api.output(job / "saving.json", request)
        try:
            connection.word("save", target, timeout=90)
        except ValueError as exc:
            if getattr(exc, "reason_code", "") != "WORD_TIMEOUT" and "-1712" not in str(exc):
                raise
            if not _disk_matches_refresh(target, cfg, job, api) or connection.word("state", target) == "unsaved":
                raise
        _bound_copy(cfg, guard)
        digest = guard.sha(target)
        api.output(saved_path, {"document": str(target), "sha256": digest,
                                "save_seconds": round(time.monotonic() - saving, 3)})
    if window != "closed" and connection.word("settle-save", target) != "closed":
        raise ValueError("Word did not confirm closing the saved working copy")
    api.output(job / "closed.json", {"document": str(target), "sha256": digest,
                                     "save_close_seconds": round(time.monotonic() - saving, 3), "resumed": resume})
    return True


def settle(job, api, guard, connection, *, verification=False):
    """Resume the recorded Word operation without replaying a plugin command."""
    cfg = api.load(job / "job.json")
    with connection.desktop_lock(job):
        target = _bound_copy(cfg, guard)
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
            if (guard.sha(target) != cfg["documentHash"] or guard.sha(cfg["source"]) != cfg["sourceHash"]
                    or guard.sha(cfg["snapshot"]) != cfg["sourceHash"]):
                raise ValueError("The bound original, snapshot or working copy has changed.")
            opened_at = time.monotonic()
            connection.word("reopen" if cfg["reopenOnly"] else "open", target)
            opened = True
            connection.dispatch(job, cfg)
            saved_hash = cfg["documentHash"]
            if not cfg["reopenOnly"]:
                _selected(job, "attached", target, api, connection)
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
                return
            _selected(job, "reattached", target, api, connection)
            connection.wait(job, "verified.json", 120)
            _bound_copy(cfg, guard)
            window = connection.word("finish", target)
            if guard.sha(target) != saved_hash or guard.sha(cfg["source"]) != cfg["sourceHash"]:
                raise ValueError("The saved document or original changed during reopen verification.")
            api.output(job / "finished.json", {"document": str(target), "sha256": saved_hash, "word_window": window,
                                               "reopen_verify_seconds": round(time.monotonic() - opened_at, 3)})
        except BaseException:
            (job / "abort").touch(exist_ok=True)
            if opened and (cfg.get("reopenOnly") or not (job / "attached.ok").exists()):
                try:
                    connection.word("finish", target)
                except (OSError, ValueError):
                    pass  # Preserve any unsaved document or changed active window.
            raise

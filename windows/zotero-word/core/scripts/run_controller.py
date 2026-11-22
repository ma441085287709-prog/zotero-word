"""Deliver the working copy after the native WinWord refresh, save and close."""
from contextlib import closing
from datetime import datetime, timezone
import hashlib
from pathlib import Path
import re
import sys
import time
import uuid
from zipfile import BadZipFile

SCHEMA = "zotero-word.windows-run.v2"
DIRECT_SCOPE = "zotero_refresh_word_save_close"
PREFS_SCOPE = "zotero_preferences_update_word_save_close"
DEFAULTS = {"preparation": "preparation.json", "hybrid_plan": "plan.json",
            "library": "library.json",
            "snapshot": "source-snapshots/original.docx", "review": "matching.json"}


def _now():
    return datetime.now(timezone.utc).isoformat()


def artifact_paths(directory):
    # Keep independent Word runs distinguishable even when their source names coincide.
    token = hashlib.sha256(str(directory).casefold().encode("utf-8")).hexdigest()[:12]
    paths = {key: directory / name for key, name in DEFAULTS.items()}
    paths["citation_document"] = directory / f"citations-{token}.docx"
    paths["delivery_document"] = paths["citation_document"]
    paths["metadata"] = directory / "reference-metadata.json"
    paths["saved_copy"] = directory / "saved-copy.json"
    return paths


def _legacy_delivery(directory):
    working = artifact_paths(directory)["citation_document"]
    return working.with_name("verified-" + working.name)


def _manual_delivery(directory, sha):
    if not isinstance(sha, str) or not re.fullmatch("[0-9a-f]{64}", sha):
        raise ValueError("Invalid saved-copy fingerprint")
    base = _legacy_delivery(directory)
    return base.with_name(f"{base.stem}-{sha[:16]}.docx")


def _accepted_path(directory, state):
    path = Path(state["document"])
    if state.get("completion_scope") in (DIRECT_SCOPE, PREFS_SCOPE):
        working = Path(state["paths"]["citation_document"])
        origin = "refresh" if state["completion_scope"] == DIRECT_SCOPE else "preferences_update"
        return (path == working == Path(state["paths"]["delivery_document"])
                and working in (artifact_paths(directory)["citation_document"], directory / "citations.docx")
                and state.get("native_update_origin") == origin and state.get("word_window") == "closed"
                and state.get("native_refresh_observed") == (origin == "refresh")
                and state.get("document_sha256") == state.get("protection", {}).get("target_sha256"))
    if not state.get("native_validation_verified"):
        return False
    render_hash = state.get("protection", {}).get("rendered_sha256")
    if render_hash and path == _manual_delivery(directory, render_hash):
        return True
    if state.get("saved_copy_audited") and path == _manual_delivery(directory, state["document_sha256"]):
        return True  # A previously accepted 2.4.1 delivery.
    paths = artifact_paths(directory)
    return (path in (_legacy_delivery(directory), paths["citation_document"]) or
            (path == directory / "citations.docx" and Path(state["paths"]["citation_document"]) == path))


def _directory(value):
    path = Path(value).expanduser().absolute()
    if path != path.resolve() or path.drive.startswith("\\\\"):
        raise ValueError("Use an independent local Windows run directory")
    if path.suffix.lower() == ".json":
        if path.name != "run.json":
            raise ValueError("Use a run directory or its run.json")
        path = path.parent
    return path


def _job(directory, data, api):
    folder = directory / ("native-" + uuid.uuid4().hex)
    folder.mkdir(mode=0o700)
    api.output(folder / "job.json", data)
    return folder


def _previous_job(directory, state, key):
    if not state.get(key):
        return None
    job = Path(state[key])
    if job.parent != directory or job.resolve() != job or not re.fullmatch(r"native-[0-9a-f]{32}", job.name):
        raise ValueError("Native job path changed")
    return job


def _mappings(report, library, review, api, aliases, cached):
    """Keep every source position, including multiple cues consumed by a reviewed group."""
    cues = report["citation_cues"]
    rows = [dict(row) for row in review.get("entries", [])]
    ignored = review.get("non_citations", {})
    if not isinstance(ignored, dict) or any(not isinstance(v, str) or not v.strip() for v in ignored.values()):
        raise ValueError("Each non-citation needs its observed cue_id and reason")
    cue_ids = {c["cue_id"] for c in cues}
    if set(ignored) - cue_ids:
        raise ValueError("Review references a different source cue")
    choices = review.get("choices", {})
    if not isinstance(choices, dict) or set(choices) - cue_ids:
        raise ValueError("A selected citation position does not belong to this manuscript")
    lookup_cues = [c for c in cues if c["cue_id"] not in cached and c["cue_id"] not in choices]
    lookup = api._helper("execution_plan").prepare({**report, "citation_cues": lookup_cues}, library) if lookup_cues else {"query_groups": []}
    groups = {}
    for group in lookup["query_groups"]:
        for key in group["source_ids"]:
            groups.setdefault(key, []).append(group)
    blocks = {(b["story"], b["paragraph"]): b["text"] for b in report["blocks"]}
    doi_index = {}
    for uri, item in library.items():
        doi_index.setdefault(api._cues.normalize_doi(item.get("DOI")), []).append(uri)
    conflicts = {(c["story"], c["paragraph"], c["start"], c["end"]) for c in cues if c.get("link_target_conflict")}
    excluded = dict(ignored)
    pending, auto = [], []
    for cue in cues:
        cid, part, para = cue["cue_id"], cue["story"], cue["paragraph"]
        covering = [r for r in rows if r.get("story", "word/document.xml") == part and r["paragraph"] == para
                    and r["start"] <= cue["start"] and r["end"] >= cue["end"]]
        if len(covering) > 1 or (covering and cid in ignored):
            raise ValueError("Conflicting dispositions for " + cid)
        if covering or cid in ignored:
            continue
        conflict = (part, para, cue["start"], cue["end"]) in conflicts
        reason = api._helper("reference_import").non_citation(cue) if not conflict else None
        if reason:
            excluded[cid] = reason
            continue
        if cid in choices:
            selected = choices[cid]
            if not isinstance(selected, list) or not selected or any(uri not in library for uri in selected):
                raise ValueError("Select a real item from the bound library for " + cid)
            if conflict or part != "word/document.xml":
                raise ValueError("This location needs an explicit range review: " + cid)
            explicit_doi = api._cues.normalize_doi(cue.get("DOI"))
            if explicit_doi and (len(selected) != 1 or api._cues.normalize_doi(library[selected[0]].get("DOI")) != explicit_doi):
                raise ValueError("The selected Zotero item does not have the DOI at " + cid)
            auto.append({"story": part, "paragraph": para, "start": cue["start"], "end": cue["end"],
                         "id": cid, "uris": list(selected), "text": blocks[part, para][cue["start"]:cue["end"]],
                         "converted_link": cue.get("kind") in ("doi", "article_link") or bool(cue.get("target"))})
            continue
        if cid in cached:
            auto.append(dict(cached[cid]))
            continue
        matches = groups.get(cid, [])
        candidates = list(dict.fromkeys(uri for group in matches for uri in group["candidate_uris"]))
        if cue.get("target") in aliases:
            candidates = list(dict.fromkeys(candidates + doi_index.get(aliases[cue["target"]], [])))
        exact = len(matches) == 1 and matches[0]["query"]["kind"] in ("doi", "url")
        if not exact or len(candidates) != 1 or conflict or part != "word/document.xml":
            pending.append({**cue, "candidates": candidates, "link_target_conflict": conflict})
            continue
        text = blocks[part, para]
        start, end = cue["start"], cue["end"]
        auto.append({"story": part, "paragraph": para, "start": start, "end": end,
                     "id": cid, "uris": candidates, "text": text[start:end], "converted_link": True})
    resolved = {row["id"]: {**row, "uris": list(row["uris"])} for row in auto if row["id"] not in choices}
    # Only combine links in the same parenthesized/bracketed citation list.
    for row in sorted(auto, key=lambda r: (r["paragraph"], r["start"])):
        previous = rows[-1] if rows else None
        text = blocks[row["story"], row["paragraph"]]
        if previous and previous.get("_auto") and previous["paragraph"] == row["paragraph"]:
            gap = text[previous["end"]:row["start"]]
            prefix = text[:previous["start"]].rstrip()
            if (prefix.endswith(("(", "[", "（")) and previous["end"] <= row["start"]
                    and re.fullmatch(r"\s*[,;，；]\s*", gap)):
                previous["uris"].extend(uri for uri in row["uris"] if uri not in previous["uris"])
                previous["end"] = row["end"]
                previous["text"] = text[previous["start"]:row["end"]]
                continue
        rows.append(dict(row, uris=list(row["uris"]), _auto=True))
    for row in rows:
        if not row.pop("_auto", False):
            continue
        text = blocks[row["story"], row["paragraph"]]
        left, right = row["start"], row["end"]
        if left and right < len(text) and (text[left-1], text[right]) in (("(", ")"), ("[", "]"), ("（", "）")):
            row.update(start=left-1, end=right+1, text=text[left-1:right+1])
    return rows, pending, excluded, resolved


def _connection(index, state, api):
    def report(update):
        if "desktop_event" in update:
            counts = state.setdefault("desktop_counts", {"host_permission_prompts": None})
            event = update["desktop_event"]
            counts[event] = counts.get(event, 0) + 1
            api.output(index, state)
            return
        previous = state.get("zotero_connection", {})
        if update.get("state") == "library_read" and previous.get("job") == update["job"]:
            update = {**previous, **update}
        state["zotero_connection"] = update
        if update.get("operation") in ("library_read", "library_import"):
            state["library_job"] = update["job"]
            if update["state"] == "library_read":
                state["library_receipt"] = update["receipt"]
        api.output(index, state)
    return api._helper("windows_connection").Connection(index.parent, api, report)


def _set_hold(index, state, error, api):
    detail = api._helper("execution_errors").explain(error)
    if hasattr(error, "observation"):
        if detail["reason_code"] == "WORD_CONTROL_DENIED":
            state["word_control_error"] = error.observation
            state.pop("word_documents_open", None)
        diagnostic = detail.setdefault("diagnostic", str(index))
        error.diagnostic = diagnostic
        event = {"at": _now(), "stage": state["stage"], "reason": str(error),
                 "reason_code": detail["reason_code"], "diagnostic": diagnostic}
        if diagnostic == str(index):
            event["observation"] = error.observation
        state.setdefault("connection_failures", []).append(event)
    state.update(result="HOLD", native_validation_verified=False, reason=str(error), resume_from=state["stage"], **detail)


def _library_snapshot(job, export, api, guard):
    if not isinstance(export, dict) or not isinstance(export.get("library_id"), str) or not isinstance(export.get("items"), list):
        raise ValueError("Invalid native Zotero library receipt")
    return {"schema": "zotero-word.snapshot.v1", "kind": "library", **export, "filters_clear": True,
            "items_sha256": api.digest(export["items"]), "source_export_sha256": guard.sha(job / "library.json")}


def _matching_library(args, cache, previous, state, live_library, read_library, api):
    external = state.get("external_library", False)
    cfg = api.load(previous / "job.json") if previous else {}
    receipt = previous / "library.json" if previous else None
    uncertain_import = cfg.get("imports") and (previous / "dispatched").exists() and not receipt.is_file()
    if args.library:
        supplied = api.load(args.library)
        if not uncertain_import or not isinstance(supplied, dict) or supplied.get("library_id") != cfg["libraryID"]:
            return supplied, True
    if uncertain_import:
        return live_library(library_id=cfg["libraryID"]), False
    if args.reload_library:
        return live_library(), False
    if receipt and receipt.is_file() and (not external or state.get("library_receipt") != str(receipt)):
        snapshot = read_library(previous, api.load(receipt))
        external = False
    elif cache.is_file():
        snapshot = api.load(cache)
    else:
        return live_library(), False
    if not external:
        observed = datetime.fromisoformat(snapshot["observed_at"].replace("Z", "+00:00"))
        if (datetime.now(timezone.utc) - observed).total_seconds() >= 86400:
            return live_library(library_id=snapshot["library_id"]), False
    return snapshot, external


def _native_rows(report, plan):
    rows = []
    for c in report["citations"]:
        rows.append({"id": c["citationID"], "items": [i["uris"] for i in c["citationItems"]],
                     "story": c["context"]["story"], "location": c["context"]["location"]})
    locations = {(b["story"], b["paragraph"]): b["location"] for b in report["blocks"]}
    for e in plan["entries"]:
        rows.append({"id": e["citation_id"], "items": [[u] for u in e["uris"]],
                     "story": e["story"], "location": locations[e["story"], e["paragraph"]]})
    return rows


def _resume_refresh(job, state, target, candidate, api, guard, connection):
    prepared = state.get("resume_input_sha256", state["write"]["target_sha256"])
    cfg = api.load(job / "job.json")
    if (cfg.get("source") != state["source"]["path"] or cfg.get("sourceHash") != state["source"]["sha256"]
            or cfg.get("snapshot") != state["paths"]["snapshot"] or cfg.get("document") != str(target)
            or cfg.get("documentHash") != prepared or cfg.get("style") != state["style"]
            or cfg.get("planHash", state["write"]["plan_sha256"]) != state["write"]["plan_sha256"]
            or cfg.get("reopenOnly") is not False):
        raise ValueError("The previous native job belongs to another document or revision")
    desktop = api._helper("windows_desktop")
    desktop.await_dispatched(job, api, connection)
    if (job / "refreshed.json").is_file():
        return None
    if candidate:
        raise ValueError("This working copy has no confirmed Zotero update in this run")
    if (job / "refresh-issued.json").exists():
        raise ValueError("The WinWord refresh was issued but has no final receipt. Preserve this working copy: " + str(target))
    if "planHash" not in cfg and (job / "preferences-issued.json").is_file() and (job / "refresh-target.ok").is_file():
        # 2.4.2 could render in setDocPrefs without recording an explicit refresh.
        raise ValueError("The earlier WinWord update has no final receipt. Preserve its working copy: " + str(target))
    if (job / "attached.ok").is_file() or (job / "tracking-paused.json").is_file():
        with connection.desktop_lock(job):
            connection.bind_copy(target, cfg)
            desktop.restore_tracking(job, cfg, target, api, connection)
            if (job / "attached.ok").is_file():
                if connection.word("state", target) == "unsaved":
                    connection.word("save", target, timeout=90)
                prepared = guard.sha(target)
    (job / "abort").touch(exist_ok=True)
    return prepared


def _finish_working_copy(job, state, target, stage, api, guard, connection):
    if not (job / "attached.ok").is_file():
        raise ValueError("The refresh receipt has no bound Word selection")
    if not (job / "closed.json").is_file():
        stage("resume_word_save")
        if not api._helper("windows_desktop").settle(job, api, guard, connection):
            raise ValueError("Zotero refreshed this working copy, but Word closed without saving it: " + str(target))
    refreshed = api.load(job / "refreshed.json")
    saved, closed = api.load(job / "saved.json"), api.load(job / "closed.json")
    digest = guard.sha(target)
    if any(r.get("document") != str(target) or r.get("sha256") != digest for r in (saved, closed)):
        raise ValueError("The Word save and close receipts must match the current working copy")
    source_hash = state["source"]["sha256"]
    if guard.sha(state["source"]["path"]) != source_hash or guard.sha(state["paths"]["snapshot"]) != source_hash:
        raise ValueError("The original or its protected snapshot changed during the run")
    state["paths"]["delivery_document"] = str(target)
    for key in ("verified", "verification_job", "delivery_protection", "saved_copy_audited"):
        state.pop(key, None)
    # Older WinWord jobs could finish in Document Preferences with a verified native render.
    origin = refreshed.get("origin", "refresh")
    if origin not in ("refresh", "preferences_update"):
        raise ValueError("The native update receipt does not identify a WinWord operation")
    scope = DIRECT_SCOPE if origin == "refresh" else PREFS_SCOPE
    state.update(result="COMPLETE", completion_scope=scope, document=str(target), document_sha256=digest,
                 protection={"result": "SOURCE_UNCHANGED", "source_sha256": source_hash, "target_sha256": digest},
                 native_refresh_observed=origin == "refresh", native_update_origin=origin, word_window="closed",
                 native_validation_verified=False, sample_verification="not_run",
                 citation_count=len(refreshed["citations"]), bibliography_entries=refreshed["entryCount"],
                 native_seconds={"refresh": refreshed.get("refreshSeconds"), "save": saved.get("save_seconds"),
                                 "save_close": closed.get("save_close_seconds")}, finished_at=_now())
    state["elapsed_seconds"] = round((datetime.fromisoformat(state["finished_at"])
                                      - datetime.fromisoformat(state["started_at"])).total_seconds(), 3)
    stage("complete")
    return state


def run(args, api):
    directory = _directory(args.run)
    index = directory / "run.json"
    guard = api._helper("execution_guard")
    if args.command == "run-pages":
        with guard.exclusive(directory, api._helper("file_lock")):
            return _pages(directory, api.load(index), api, guard)
    if args.command == "run-verify-copy":
        with guard.exclusive(directory, api._helper("file_lock")):
            return verify_copy(directory, args.document, args.clean, api, guard)
    if args.command == "run-status":
        state = api.load(index)
        if state.get("schema") != SCHEMA:
            return {"result": "HISTORICAL_RUN", "run_index": str(index), "stage": state.get("stage")}
        if state.get("result") == "COMPLETE":
            try:
                roles = artifact_paths(directory)
                current = (_accepted_path(directory, state)
                           and guard.sha(state["source"]["path"]) == guard.sha(roles["snapshot"]) == state["source"]["sha256"]
                           and guard.sha(state["document"]) == state["document_sha256"] == state["protection"]["target_sha256"])
            except FileNotFoundError:
                current = False
            if not current:
                return {**state, "result": "STALE", "native_validation_verified": False, "sample_verification": "stale",
                        "reason": "The original or accepted file is missing or has changed"}
        return state
    directory.mkdir(parents=True, exist_ok=True, mode=0o700)
    with guard.exclusive(directory, api._helper("file_lock")):
        return _run_locked(directory, index, args, api, guard)


def _pages(directory, state, api, guard):
    if (state.get("schema") != SCHEMA or state.get("result") != "COMPLETE" or
            not _accepted_path(directory, state)):
        raise ValueError("A completed citation run is required to review its saved pages")
    target = guard.independent(state["document"], state["source"]["path"], state["paths"]["snapshot"])
    digest = state["document_sha256"]
    if guard.sha(target) != digest or digest != state["protection"]["target_sha256"]:
        raise ValueError("The accepted file differs from its receipt")
    root = directory / "visual-qa"
    if root.resolve() != root:
        raise ValueError("The page-review directory must be inside this run")
    root.mkdir(exist_ok=True)
    manifest = root / (digest + ".json")
    if manifest.is_file():
        cached = api.load(manifest)
        folder = Path(cached["image_directory"])
        images = cached["page_sha256"]
        if (cached.get("document") == str(target) and cached.get("document_sha256") == digest
                and folder.parent == root and folder.resolve() == folder and images
                and len(images) == cached["page_count"] and all(
                    (path := folder / f"page-{n:04}.png").is_file() and path.resolve() == path and guard.sha(path) == expected
                    for n, expected in enumerate(images, 1))):
            return cached
    folder = root / ("pages-" + uuid.uuid4().hex)
    folder.mkdir()
    bridge = api._helper("windows_connection")
    with closing(bridge.Connection(directory, api, lambda _: None)) as connection, connection.desktop_lock(folder):
        observed = connection.render_pages(target, folder)
    if guard.sha(target) != digest:
        raise ValueError("The saved document changed during page review")
    result = {**observed, "result": "PAGES_RENDERED", "document": str(target), "document_sha256": digest,
              "image_directory": str(folder), "visual_review": "pending", "renderer": "Word.Page.EnhMetaFileBits",
              "rendered_at": _now(), "page_sha256": [guard.sha(folder / f"page-{n:04}.png")
                  for n in range(1, observed["page_count"] + 1)]}
    api.output(manifest, result)
    return result


def _run_locked(directory, index, args, api, guard):
    expected_paths = artifact_paths(directory)
    if index.exists():
        state = api.load(index)
        if state.get("schema") != SCHEMA:
            raise ValueError("Use a new run directory for this release; historical evidence is preserved")
        legacy_target = directory / "citations.docx"
        if Path(state["paths"]["citation_document"]) == legacy_target:
            expected_paths["citation_document"] = expected_paths["delivery_document"] = legacy_target
        if state["paths"].get("delivery_document") == str(_legacy_delivery(directory)):
            expected_paths["delivery_document"] = _legacy_delivery(directory)
        if args.command == "run-start" and (Path(args.document).resolve() != Path(state["source"]["path"])
                or args.journal != state["journal"]):
            raise ValueError("The run belongs to another original or journal")
    else:
        if args.command != "run-start":
            raise ValueError("Start a run before resuming it")
        source = Path(args.document).resolve(strict=True)
        paths = {key: str(path) for key, path in expected_paths.items()}
        for dest in paths.values():
            if Path(dest).resolve() == source or Path(dest).exists():
                raise ValueError("Run directory already contains a reserved input or output: " + dest)
        guard.independent(paths["snapshot"], source)
        guard.independent(paths["citation_document"], source, paths["snapshot"])
        state = {"schema": SCHEMA, "run_index": str(index), "journal": args.journal, "source": {"path": str(source)},
                 "paths": paths, "result": "PENDING", "stage": "source", "started_at": _now(), "stage_seconds": {}}
        api.output(index, state)
    # Earlier v2 runs used citation_document for the Word render and the delivery.
    for key in ("delivery_document", "metadata", "saved_copy"):
        if key not in state["paths"]:
            if key != "delivery_document" and expected_paths[key].exists():
                raise ValueError("New run artifact path already in use: " + str(expected_paths[key]))
            state["paths"][key] = str(expected_paths[key])
    paths = {k: Path(v) for k, v in state["paths"].items()}
    for key, expected in expected_paths.items():
        if paths[key] != expected or paths[key].resolve() != expected:
            raise ValueError("Run artifact path changed: " + key)
    source = Path(state["source"]["path"])
    target = guard.independent(paths["citation_document"], source, paths["snapshot"])
    clock = time.monotonic()

    def stage(name):
        nonlocal clock
        elapsed = time.monotonic()
        times = state["stage_seconds"]
        old = state["stage"]
        times[old] = round(times.get(old, 0) + elapsed - clock, 3)
        clock = elapsed
        state["stage"] = name
        api.output(index, state)

    connection = _connection(index, state, api)
    desktop = api._helper("windows_desktop")

    def read_library(job, export):
        request = api.load(job / "request.json")
        local_read = isinstance(export, dict) and export.get("transport") == "local_api"
        if (not isinstance(request, dict) or request.get("job") != job.name
                or request.get("operation") not in ("library_read", "library_import")
                or (not local_read and request != api.load(job / "started.json"))):
            raise ValueError("The library export belongs to another native request")
        snapshot = _library_snapshot(job, export, api, guard)
        if local_read:
            observed = api.load(job / "read-connection.json")
            if (request["operation"] != "library_read" or request.get("session") != job.name
                    or observed.get("request") != request or observed.get("result") != "complete"
                    or observed.get("http_status") != 200 or observed.get("total") != len(export["items"])
                    or observed.get("library_id") != export["library_id"]
                    or observed.get("items_sha256") != snapshot["items_sha256"]):
                raise ValueError("Local library snapshot does not match its read receipt")
        receipt = str(job / "library.json")
        if state.get("library_receipt") != receipt:
            state["external_library"] = False
            connection.report({"state": "library_read", "session": request["session"], "operation": request["operation"],
                "job": str(job), "receipt": receipt, "library_id": export["library_id"], "items": len(export["items"])})
        return snapshot

    def live_library(records=(), library_id=None):
        stage("zotero_import" if records else "zotero_library")
        job = _job(directory, {"libraryOnly": True, "imports": records, "libraryID": library_id}, api)
        export = None
        if not records and connection.first is None:
            request = {"job": job.name, "session": job.name, "operation": "library_read", "document": None}
            api.output(job / "request.json", request)
            connection.report({**request, "job": str(job), "via": "local_api", "state": "reading"})
            export = api._helper("windows_library_read").read(job, request, library_id, api, connection.launch_zotero)
        if export is None:
            export = desktop.execute(job, api, guard, connection)
        snapshot = read_library(job, export)
        stage("matching")
        return snapshot

    try:
        candidate = getattr(args, "saved_copy", None)
        if candidate and not state.get("write"):
            raise ValueError("A saved copy needs this run's existing citation plan and working-copy checkpoint")
        if "sha256" in state["source"]:
            source_hash = state["source"]["sha256"]
            if guard.sha(source) != source_hash or guard.sha(paths["snapshot"]) != source_hash:
                raise ValueError("The protected original or snapshot changed")
        if state.get("result") == "COMPLETE":
            if not _accepted_path(directory, state) or guard.sha(state["document"]) != state["document_sha256"]:
                raise ValueError("Accepted file changed after completion")
            if candidate:
                try:
                    same = guard.sha(Path(candidate).expanduser()) in (
                        state["document_sha256"], state["protection"].get("rendered_sha256"))
                except OSError:
                    same = False
                if not same:
                    return {"result": "HOLD", "reason_code": "RUN_ALREADY_COMPLETE", "native_validation_verified": False,
                            "reason": "This run already has an accepted document: " + state["document"]}
            return state
        for key in ("reason", "reason_code", "retry_allowed", "resume_from", "next_action", "matching", "diagnostic", "word_control_error"):
            state.pop(key, None)
        state["result"] = "PENDING"
        previous = _previous_job(directory, state, "native_job")
        if candidate and (not previous or Path(candidate).expanduser().resolve() != target):
            raise ValueError("Resume this run using its own working copy and refresh receipt: " + str(target))
        if "sha256" not in state["source"]:
            state["source"]["sha256"] = guard.snapshot(source, paths["snapshot"])
            api.output(index, state)
        source_hash = state["source"]["sha256"]
        style = args.style_id or state.get("style")
        if not style:
            journal = state["journal"]
            if journal.startswith(("http://www.zotero.org/styles/", "https://www.zotero.org/styles/")):
                style = journal
            elif re.sub(r"[^a-z]", "", journal.lower()) in ("est", "environmentalsciencetechnology", "environmentalscienceandtechnology"):
                style = "http://www.zotero.org/styles/environmental-science-and-technology"
            elif re.sub(r"[^a-z]", "", journal.lower()) in ("cej", "chemicalengineeringjournal"):
                style = "http://www.zotero.org/styles/chemical-engineering-journal"
        if style and style.startswith("https://www.zotero.org/styles/"):
            style = "http://" + style[len("https://"):]
        if not style:
            raise ValueError("Supply the installed target CSL URI with --style-id")
        abbreviations = args.abbreviations if args.abbreviations is not None else state.get(
            "abbreviations", True if style.endswith(("/environmental-science-and-technology", "/chemical-engineering-journal")) else None)
        if target.exists() and (style != state.get("style") or abbreviations != state.get("abbreviations")):
            raise ValueError("Style settings changed after creating the working copy")
        state.update(style=style, abbreviations=abbreviations)
        if paths["saved_copy"].is_file():
            recorded = Path(api.load(paths["saved_copy"])["candidate"])
            if recorded != target or not previous or not (previous / "refreshed.json").is_file():
                raise ValueError("Preserve the saved candidate recorded by the earlier package: " + str(recorded))
        if not previous and state.get("delivery_protection"):
            published = state["delivery_protection"].get("document") or paths["delivery_document"]
            raise ValueError("The earlier published file has no bound WinWord update receipt: " + str(published))
        prepared = None
        if previous:
            prepared = _resume_refresh(previous, state, target, candidate, api, guard, connection)
            if prepared is None:
                return _finish_working_copy(previous, state, target, stage, api, guard, connection)
        previous_library = _previous_job(directory, state, "library_job")
        stage("source")
        if not paths["preparation"].exists():
            report = api.inspect_docx(source)
            api.output(paths["preparation"], report)
        else:
            report = api.load(paths["preparation"])
        if report["document_sha256"] != source_hash or report["issues"]:
            raise ValueError("Source inspection differs or has field errors: " + ", ".join(report["issues"]))
        writer = api._helper("hybrid_citations")
        if not state.get("write") and not paths["hybrid_plan"].exists():
            if target.exists():
                raise ValueError("A working file exists without its source-bound plan")
            stage("matching")
            data = api.load(args.entries) if args.entries else state.get("position_review", {})
            if isinstance(data, list):
                data = {"entries": data}
            if data.get("source_sha256", source_hash) != source_hash:
                raise ValueError("Position review belongs to another original")
            if args.resolve:
                selections = {}
                for pair in args.resolve:
                    cue_id, separator, uri = pair.partition("=")
                    if not separator or not cue_id or not uri:
                        raise ValueError("Use --resolve CUE_ID=ZOTERO_ITEM_URI")
                    selections.setdefault(cue_id, []).append(uri)
                data["choices"] = {**data.get("choices", {}), **selections}
            if args.entries or args.resolve:
                state["position_review"] = data
                api.output(index, state)
            lib = {}
            known = {}
            if report["citation_cues"] or data.get("entries") or data.get("choices") or args.library:
                lib, external = _matching_library(args, paths["library"], previous_library, state,
                    live_library, read_library, api)
                lib, known = api.validate_snapshot(lib, "library", 24)
                api.output(paths["library"], lib)
                state["external_library"] = external
                if external:
                    state.pop("library_job", None)
                if not external and lib.get("import_receipt"):
                    state["last_import"] = lib["import_receipt"]
            if args.import_missing is not None:
                state["import_missing"] = args.import_missing
                api.output(index, state)
            aliases = state.get("link_dois", {})
            cached = {}
            if paths["review"].is_file() and state.get("matching_cache_sha256"):
                review = api.load(paths["review"])
                if (api.digest(review) == state["matching_cache_sha256"] and review.get("source_sha256") == source_hash
                        and review.get("items_sha256") == lib.get("items_sha256")
                        and review.get("review_sha256") == api.digest(data.get("entries", []))):
                    cached = review.get("matched_cues", {})
            entries, pending, excluded, resolved_cues = _mappings(report, known, data, api, aliases, cached)
            problems = {}
            if pending:
                old_aliases = dict(aliases)
                records, aliases, problems = api._helper("reference_import").fetch_missing(
                    pending, known, paths["metadata"], source_hash, api, allow_import=bool(state.get("import_missing")))
                state["link_dois"] = aliases
                if records and state.get("import_missing"):
                    if guard.sha(source) != source_hash or guard.sha(paths["snapshot"]) != source_hash:
                        raise ValueError("Source changed before importing the reviewed identifiers")
                    lib = live_library(records, lib["library_id"])
                    lib, known = api.validate_snapshot(lib, "library", 24)
                    api.output(paths["library"], lib)
                    state["external_library"] = False
                    state["last_import"] = lib["import_receipt"]
                if records or aliases != old_aliases:
                    changed = {c["cue_id"] for c in report["citation_cues"]
                               if aliases.get(c.get("target")) != old_aliases.get(c.get("target"))}
                    retained = {} if records else {k: v for k, v in resolved_cues.items() if k not in changed}
                    entries, pending, excluded, resolved_cues = _mappings(report, known, data, api, aliases, retained)
            state["non_citations"] = excluded
            citation_markers = [p for p in report["unresolved_placeholders"]
                if re.search(r"\b(?:CITE|REFHERE|BIBHERE|ZOTERO_ANCHOR_\w+)\b|待引|参考文献待|\[\s*\?\s*\]",
                             p["text"], re.I)]
            placeholders = {p["cue_sha256"] for p in citation_markers}
            resolved = data.get("non_citation_placeholders", {})
            if not isinstance(resolved, dict) or any(not isinstance(v, str) or not v.strip() for v in resolved.values()):
                raise ValueError("Non-citation placeholders need their observed hash and reason")
            if pending or placeholders != set(resolved):
                review = {"source_sha256": source_hash, "items_sha256": lib.get("items_sha256"),
                    "review_sha256": api.digest(data.get("entries", [])), "matched_cues": resolved_cues,
                    "pending": pending, "placeholders": citation_markers, "matched": entries,
                    "non_citations": excluded, "metadata_errors": problems}
                api.output(paths["review"], review)
                state["matching_cache_sha256"] = api.digest(review)
                action = ("Confirm the unresolved item with run-next --resolve CUE_ID=ITEM_URI, or resume its metadata lookup" if problems
                          else "Use run-next --resolve CUE_ID=ITEM_URI for each choice, --entries for range reviews, or an updated library")
                state.update(result="NEEDS_MATCHING", next_action=action, matching=str(paths["review"]))
                stage("matching")
                return state
            config = {"style_id": style, "bibliography_paragraph": data.get("bibliography_paragraph"), "link_dois": aliases}
            plan = writer.build_plan(source, entries, lib or None, style_config=config)
            api.output(paths["hybrid_plan"], plan)
        plan = writer._validate_plan(source, api.load(paths["hybrid_plan"]))
        if not state.get("write"):
            stage("batch_write")
            state["write"] = writer.write_prepared(source, plan, target, resume=True, for_native_run=True)
            api.output(index, state)
        if state["write"]["plan_sha256"] != plan["plan_sha256"]:
            raise ValueError("The prepared plan changed")
        rows = _native_rows(report, plan)
        if not rows:
            raise ValueError("No Zotero citations to refresh")
        config = {"source": str(source), "sourceHash": source_hash, "snapshot": str(paths["snapshot"]),
            "citations": rows,
            "anchor": plan.get("bibliography_paragraph"), "anchorEnd": plan.get("bibliography_at_end", False),
            "style": style, "abbreviations": abbreviations, "libraryOnly": False}
        prepared = prepared or state["write"]["target_sha256"]
        if guard.sha(target) != prepared:
            raise ValueError("The prepared working copy changed before dispatch: " + str(target))
        stage("native_refresh_save")
        job = _job(directory, {**config, "document": str(target), "documentHash": prepared,
                               "planHash": plan["plan_sha256"], "reopenOnly": False}, api)
        state.update(native_job=str(job), resume_input_sha256=prepared)
        api.output(index, state)
        desktop.execute(job, api, guard, connection)
        return _finish_working_copy(job, state, target, stage, api, guard, connection)
    except (OSError, ValueError, KeyError, TypeError, ImportError, BadZipFile) as exc:
        _set_hold(index, state, exc, api)
        word_error = state.get("word_control_error", {})
        if (state["reason_code"] == "WORD_CONTROL_DENIED" and state["stage"] == "native_refresh_save"
                and word_error.get("operation", word_error.get("phase")) == "open"
                and not (Path(state["native_job"]) / "request.json").exists()):
            # PowerShell parses straight and typographic single quotes as delimiters.
            argv = (sys.executable, "-X", "utf8", str(Path(__file__).resolve().with_name("zotero_docx.py")),
                    "run-next", "--run", str(directory), "--output", str(Path(args.output).resolve()), "--brief")
            word_error["working_copy"] = str(target)
            quoted = [re.sub("['‘’‚‛]", lambda match: match[0] * 2, arg) for arg in argv]
            word_error["resume_command"] = "& " + " ".join(f"'{arg}'" for arg in quoted)
            state["next_action"] = ("工作副本已写入，刷新尚未发出。请在同一 Windows 登录桌面中，"
                                    "用可控制 Word 的宿主在 PowerShell 执行 word_control_error.resume_command。")
        stage(state["stage"])
        return state
    finally:
        connection.close()


def _reopen(directory, job, state, index, config, target, digest, refreshed, api, connection, guard):
    desktop = api._helper("windows_desktop")
    saved_copy = refreshed.get("origin") == "saved_copy"
    if job:
        bound = api.load(job / "job.json")
        if (bound.get("document") != str(target) or bound.get("documentHash") != digest
                or bool(bound.get("savedCopy")) != saved_copy or bound.get("sourceHash") != config["sourceHash"]
                or bound.get("executor") != connection.executor):
            job = None
    if job and (job / "verified.json").is_file() and not (job / "finished.json").is_file():
        desktop.settle(job, api, guard, connection, verification=True)
    if not job or not (job / "finished.json").is_file():
        job = _job(directory, {**config, "document": str(target), "documentHash": digest,
                               "reopenOnly": True, "savedCopy": saved_copy, "executor": connection.executor}, api)
        api.output(job / ("saved-copy.json" if saved_copy else "refreshed.json"), refreshed)
        state["verification_job"] = str(job)
        api.output(index, state)
        desktop.execute(job, api, guard, connection)
    observed, finished = api.load(job / "verified.json"), api.load(job / "finished.json")
    if (observed.get("executor") != connection.executor or finished["document"] != str(target) or observed.get("document") != str(target)
            or observed.get("sha256") != digest or finished["sha256"] != digest or digest != guard.sha(target)
            or not observed["bibliographyEntries"] or (not saved_copy and observed["bibliographyEntries"] != refreshed["entryCount"])
            or observed["citations"] != len(config["citations"]) or observed["style"] != config["style"]
            or [s["citation"] for s in observed["samples"]] != config["samples"]):
        raise ValueError("Native reopen evidence differs from the bound plan or saved file")
    planned = {row["id"]: row["items"] for row in config["citations"]}
    for index_in_job, sample in enumerate(observed["samples"]):
        label = "sample-" + str(index_in_job)
        if (api.load(job / (label + ".json")).get("citation") != sample["citation"] or
                api.load(job / (label + ".ok")) is not True):
            raise ValueError("The Word selection receipt does not match the sampled citation")
        aliases = planned[sample["citation"]]
        if len(sample["items"]) != len(aliases) or not all(
                type(item["id"]) is int and item["id"] > 0 and api.canonical_uri(item["uri"])
                and item.get("matchedURI") in choices
                for item, choices in zip(sample["items"], aliases)):
            raise ValueError("Edit Citation did not recover the expected real library items")
    return observed, finished


def verify_copy(directory, document, clean, api, guard, *, layout_protection=None):
    baseline = api.load(directory / "run.json")
    if baseline.get("schema") != SCHEMA or baseline.get("result") != "COMPLETE":
        raise ValueError("A completed Windows citation run is required to verify a derived copy")
    original = baseline["source"]["path"]
    paths = artifact_paths(directory)
    accepted = Path(baseline["document"])
    if not _accepted_path(directory, baseline):
        raise ValueError("Accepted citation path differs from the run")
    target = guard.independent(Path(document).expanduser(), original, paths["snapshot"], accepted, paths["citation_document"])
    if (guard.sha(original) != baseline["source"]["sha256"] or guard.sha(paths["snapshot"]) != baseline["source"]["sha256"]
            or guard.sha(accepted) != baseline["document_sha256"]):
        raise ValueError("Original, snapshot or accepted citation file changed")
    if layout_protection is None:
        source_info, target_info, math = api._copy_content(accepted, target, clean=clean)
        source_hash, digest = source_info["document_sha256"], target_info["document_sha256"]
    else:
        source_hash, digest = layout_protection["source_document_sha256"], layout_protection["target_document_sha256"]
        math = layout_protection["math_equivalence"]
    if source_hash != baseline["document_sha256"] or digest != guard.sha(target):
        raise ValueError("Derived-copy protection does not match the current files")
    if clean:
        api._helper("hybrid_citations").check_clean(target)
    folder = directory / "copies" / api.digest([str(target), digest])
    if folder.resolve() != folder:
        raise ValueError("Derived-copy evidence directory must stay inside the run")
    folder.mkdir(mode=0o700, parents=True, exist_ok=True)
    index = folder / "verification.json"
    state = api.load(index) if index.is_file() else {}
    previous = _previous_job(folder, state, "verification_job")
    plan = api._helper("hybrid_citations")._validate_plan(original, api.load(paths["hybrid_plan"]))
    rows = _native_rows(api.load(paths["preparation"]), plan)
    samples = {}
    for row in rows:
        samples.setdefault((len(row["items"]) > 1, row["location"]), row["id"])
    config = {"source": original, "sourceHash": baseline["source"]["sha256"], "snapshot": str(paths["snapshot"]),
              "citations": rows, "samples": list(samples.values()), "style": baseline["style"],
              "abbreviations": baseline["abbreviations"], "libraryOnly": False}
    for key in ("reason", "reason_code", "retry_allowed", "resume_from", "next_action", "diagnostic", "word_control_error"):
        state.pop(key, None)
    state.update(result="PENDING", native_validation_verified=False, stage="reopen_verify", document=str(target))
    try:
        with closing(_connection(index, state, api)) as connection:
            state["stage"] = "reopen_verify"
            if baseline.get("saved_copy_audited"):
                evidence = api.load(paths["saved_copy"])["render"]
            else:
                render = _previous_job(directory, baseline, "native_job")
                evidence = api.load(render / "refreshed.json")
            observed, finished = _reopen(folder, previous, state, index, config, target, digest,
                evidence, api, connection, guard)
        if (guard.sha(accepted) != baseline["document_sha256"] or guard.sha(target) != digest or
                guard.sha(original) != config["sourceHash"] or guard.sha(paths["snapshot"]) != config["sourceHash"]):
            raise ValueError("A bound file changed while verifying the derived copy")
    except (OSError, ValueError, KeyError, TypeError, ImportError, BadZipFile) as exc:
        _set_hold(index, state, exc, api)
        api.output(index, state)
        raise
    state.update(result="COMPLETE", completion_scope="derived_copy_content_and_dynamic_citations",
        stage="complete", document=str(target), document_sha256=digest, native_validation_verified=True, verified=observed,
        protection={"source_document": str(accepted), "source_sha256": source_hash,
                    "target_sha256": digest, "active_content_preserved": True,
                    "clean_copy_verified": bool(clean), "math_equivalence": math}, finish=finished)
    api.output(index, state)
    return state

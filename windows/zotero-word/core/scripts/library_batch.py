"""Read-only batch plans for lookup, imports and collection membership.

This module emits adapter work and rechecks observations.  It never calls a
Zotero API, writes a library, creates Word fields or certifies native execution.
The existing preparation and membership engines remain the authorities for
source positions, actual item identities and authorized collection differences.
"""
from __future__ import annotations

from collections import Counter
from datetime import datetime, timezone
import hashlib
import importlib.util
import json
from pathlib import Path
import re
import sys
from urllib.parse import urlsplit, urlunsplit


PLAN = "zotero-word.library-batch-plan.v1"
STATUS = "zotero-word.library-batch-status.v1"
OBSERVATIONS = "zotero-word.batch-observations.v1"


def _helper(name):
    spec = importlib.util.spec_from_file_location(
        "_zotero_word_batch_" + name, Path(__file__).with_name(name + ".py"))
    module = importlib.util.module_from_spec(spec)
    previous = sys.dont_write_bytecode
    try:
        sys.dont_write_bytecode = True
        spec.loader.exec_module(module)
    finally:
        sys.dont_write_bytecode = previous
    return module


def _require(ok, message):
    if not ok:
        raise ValueError(message)


def _digest(value):
    return hashlib.sha256(json.dumps(value, sort_keys=True, ensure_ascii=False,
                                    separators=(",", ":")).encode()).hexdigest()


def _text(value):
    return isinstance(value, str) and bool(value.strip())


def _now():
    return datetime.now(timezone.utc).isoformat()


def _url(value):
    if not isinstance(value, str):
        return ""
    parsed = urlsplit(value.strip())
    if parsed.scheme.lower() not in ("http", "https") or not parsed.netloc:
        return ""
    return urlunsplit((parsed.scheme.lower(), parsed.netloc.lower(),
                       parsed.path.rstrip("/"), parsed.query, ""))


def _identifier(value, core):
    identifier = core.doi(value.get("DOI") or value.get("URL"))
    if identifier:
        return {"kind": "doi", "DOI": identifier}
    url = _url(value.get("URL"))
    _require(url, "Batch imports require a verified DOI or absolute article URL; use the native reviewed path for other record types")
    return {"kind": "url", "URL": url}


def _hits(identifier, known, core):
    return sorted(uri for uri, item in known.items()
                  if (core.doi(item.get("DOI")) == identifier["DOI"]
                      if identifier["kind"] == "doi"
                      else _url(item.get("URL")) == identifier["URL"]))


def _batches(operations, size):
    # Keep read/import/membership calls separate so the adapter can discover a
    # real batch endpoint and respect its own smaller request-size limit.
    result = []
    for action in ("lookup", "import_item", "add_membership", "remove_membership"):
        chosen = [op for op in operations if op["action"] == action]
        for start in range(0, len(chosen), size):
            items = chosen[start:start + size]
            result.append({"batch_index": len(result) + 1, "action": action,
                           "operation_ids": [op["operation_id"] for op in items],
                           "count": len(items), "status": "PENDING_ADAPTER"})
    return result


def _plan_fingerprint(plan):
    return _digest({k: v for k, v in plan.items() if k != "plan_fingerprint"})


def _collection(core, library, known, snapshot):
    col, members = core.validate_snapshot(snapshot, "collection", library["max_age_hours"])
    _require(col["library_id"] == library["library_id"], "Batch library and collection identities differ")
    _require(all(uri in known and known[uri] == record for uri, record in members.items()),
             "Collection records differ from the library snapshot")
    return col, members


def build(preparation, library_snapshot, *, current_document_sha256, imports=None,
          membership_plan=None, collection_snapshot=None, batch_size=20,
          max_age_hours=24):
    """Build structured adapter batches without selecting ambiguous identities.

    ``preparation`` is a current ``prepare`` output, or None for a membership-
    only stage after native conversion changed the document bytes.  ``imports``
    are per-position, already identity-reviewed requests, not search hits.
    ``membership_plan`` must be the existing engine's authorized READY plan.
    """
    core = _helper("zotero_docx")
    _require(type(batch_size) is int and batch_size > 0, "Batch size must be a positive integer")
    _require(isinstance(current_document_sha256, str)
             and re.fullmatch(r"[a-f0-9]{64}", current_document_sha256), "Bind the current saved document hash")
    lib, known = core.validate_snapshot(library_snapshot, "library", max_age_hours)
    _require(preparation is not None or membership_plan is not None,
             "Provide preparation or a READY membership plan")
    groups = []
    if preparation is not None:
        _require(isinstance(preparation, dict)
                 and preparation.get("schema") == "zotero-word.execution-preparation.v1"
                 and preparation.get("source_document_sha256") == current_document_sha256,
                 "Preparation belongs to another saved document")
        _require(preparation.get("library_id", lib["library_id"]) == lib["library_id"],
                 "Preparation belongs to another library")
        if preparation.get("library_snapshot_sha256"):
            # Public prepare records the raw input-file hash; record content is
            # independently checked below instead of inventing JSON whitespace.
            _require(preparation.get("library_status") == "VALIDATED_LOCAL_SNAPSHOT",
                     "Preparation does not establish its observed local snapshot")
        groups = preparation.get("query_groups")
        _require(isinstance(groups, list), "Preparation has no per-position query groups")
        summaries = preparation.get("candidate_records", {})
        _require(isinstance(summaries, dict), "Invalid preparation candidate records")
        for uri, record in summaries.items():
            _require(uri in known and record.get("source_record_sha256") == core.digest(known[uri]),
                     "A prepared candidate changed; refresh preparation before batching")
    valid_sources, query_keys, source_queries = set(), set(), {}
    for group in groups:
        _require(isinstance(group, dict) and _text(group.get("key"))
                 and group["key"] not in query_keys and isinstance(group.get("query"), dict),
                 "Query groups require distinct keys and query objects")
        query_keys.add(group["key"])
        ids = group.get("source_ids")
        _require(isinstance(ids, list) and ids and len(ids) == len(set(ids))
                 and all(_text(sid) for sid in ids), "Query groups must retain distinct actual source positions")
        valid_sources.update(ids)
        for sid in ids:
            source_queries.setdefault(sid, []).append(group["query"])

    imports = [] if imports is None else imports
    _require(isinstance(imports, list), "Import requests must be a list")
    deduplicated = {}
    for request in imports:
        _require(isinstance(request, dict) and request.get("identity_verified") is True
                 and _text(request.get("evidence")), "Imports need actual identity verification and evidence")
        _require(request.get("metadata") is None or isinstance(request.get("metadata"), dict),
                 "Verified import metadata must be an object when provided")
        ids = request.get("source_ids")
        _require(isinstance(ids, list) and ids and len(ids) == len(set(ids))
                 and set(ids) <= valid_sources, "Import requests must bind exact prepared source positions")
        identifier = _identifier(request, core)
        resolved_source_urls = []
        # A source position is not a blank authorization to import any paper.
        # Exact DOI/URL clues remain binding even after semantic verification;
        # author/year queries may legitimately resolve to a newly found DOI.
        for sid in ids:
            for query in source_queries[sid]:
                if query.get("kind") == "doi":
                    _require(identifier["kind"] == "doi"
                             and identifier["DOI"] == core.doi(query.get("DOI")),
                             "Import DOI conflicts with its exact source clue; resolve the conflict instead of replacing the identifier")
                elif query.get("kind") in ("url", "article_url", "article_link"):
                    original_url = _url(query.get("URL"))
                    bindings = [request.get("URL")]
                    if identifier["kind"] == "doi":
                        # Publisher landing URLs often contain no DOI. A
                        # verified metadata lookup can resolve one while
                        # retaining an explicit binding to that original URL.
                        bindings.extend([request.get("resolved_from_URL"),
                                         (request.get("metadata") or {}).get("URL")])
                    _require(original_url and original_url in {_url(value) for value in bindings},
                             "Import URL conflicts with its exact source clue; retain the verified source URL when resolving its DOI")
                    resolved_source_urls.append(original_url)
        key = _digest(identifier)
        entry = deduplicated.setdefault(key, {"identifier": identifier, "source_ids": [],
                                             "evidence": [], "metadata": request.get("metadata"),
                                             "resolved_from_URLs": [],
                                             "selection": {k: request.get(k) for k in
                                                           ("selected_uri", "selection_confirmed", "decision_source")}})
        _require(entry["metadata"] == request.get("metadata")
                 and entry["selection"] == {k: request.get(k) for k in entry["selection"]},
                 "Duplicate import requests conflict; resolve metadata/selection before writing")
        entry["source_ids"] = sorted(set(entry["source_ids"]) | set(ids))
        entry["resolved_from_URLs"] = sorted(set(entry["resolved_from_URLs"]) | set(resolved_source_urls))
        entry["evidence"].append(request["evidence"])

    operations, reuse, reviews, import_rows = [], [], [], []
    reviewed_sources = {sid for row in deduplicated.values() for sid in row["source_ids"]}
    for key, entry in deduplicated.items():
        hits = _hits(entry["identifier"], known, core)
        selection = entry.pop("selection")
        selected = core.canonical_uri(selection.get("selected_uri")) if selection.get("selected_uri") else None
        if selected:
            _require(selected in hits and selection.get("selection_confirmed") is True
                     and _text(selection.get("decision_source")),
                     "An ambiguous existing identity needs the actual recorded selection decision")
        row = {**entry, "request_key": key, "candidate_uris": hits}
        if selected or len(hits) == 1:
            row.update(status="REUSE_EXISTING", uri=selected or hits[0],
                       selection_decision=selection if selected else None)
            reuse.append({"uri": row["uri"], "source_ids": row["source_ids"], "request_key": key})
        elif hits:
            row["status"] = "IDENTITY_REVIEW_REQUIRED"
            reviews.append({"kind": "duplicate_library_identity", **row})
        else:
            row["status"] = "IMPORT_READY"
            operations.append({"operation_id": "import:" + _digest([lib["library_id"], entry["identifier"]]),
                               "action": "import_item", "identifier": entry["identifier"],
                               "source_ids": entry["source_ids"], "metadata": entry["metadata"],
                               "resolved_from_URLs": entry["resolved_from_URLs"],
                               "evidence": entry["evidence"],
                               "preflight": "Recheck DOI/URL in this library immediately before import; reuse one proven existing identity, review multiple; never blindly replay an uncertain import",
                               "required_capability": "verified_zotero_import",
                               "idempotency_key": _digest([lib["library_id"], entry["identifier"]])})
        import_rows.append(row)

    for group in groups:
        pending_ids = [sid for sid in group["source_ids"] if sid not in reviewed_sources]
        if not pending_ids:
            continue
        candidates = group.get("candidate_uris", [])
        _require(isinstance(candidates, list) and all(uri in known for uri in candidates),
                 "A prepared candidate disappeared; refresh preparation")
        if candidates or group["query"].get("kind") in ("manual", "static_reference_mapping"):
            reviews.append({"kind": "per_context_semantic_review", "query_key": group["key"],
                            "source_ids": pending_ids, "candidate_uris": candidates,
                            "status": "REVIEW_REQUIRED"})
        else:
            operations.append({"operation_id": "lookup:" + _digest([current_document_sha256, group["key"]]),
                               "action": "lookup", "query_key": group["key"], "query": group["query"],
                               "source_ids": pending_ids, "positions": group.get("positions", []),
                               "required_capability": "public_metadata_lookup",
                               "policy": "Return candidates/evidence only; distinct contexts still require review; no automatic import or native insertion"})

    collection_binding = None
    if membership_plan is not None:
        _require(collection_snapshot is not None, "Membership batches require the observed collection snapshot")
        p = membership_plan
        _require(isinstance(p, dict) and p.get("schema") == core.PLAN and p.get("status") == "READY"
                 and p.get("document_sha256") == current_document_sha256,
                 "Only a current READY membership plan can supply membership operations")
        col, members = _collection(core, {**lib, "max_age_hours": max_age_hours}, known, collection_snapshot)
        identity = {k: col.get(k) for k in ("library_id", "collection_name", "collection_path", "collection_key")}
        _require(all(p.get(k) == value for k, value in identity.items()), "Membership plan targets a different collection")
        _require(p.get("library_record_hashes") == {u: core.digest(i) for u, i in known.items()}
                 and p.get("current_members") == sorted(members), "Library or collection changed after membership planning")
        desired = p.get("desired_members")
        _require(isinstance(desired, list) and len(desired) == len(set(desired)) and set(desired) <= set(known),
                 "Membership targets must be actual library identities")
        expected = {"add": set(desired) - set(members), "remove": set(members) - set(desired)}
        if expected["add"] or expected["remove"]:
            _require(p.get("update_decision") == "confirmed" and _text(p.get("decision_source")),
                     "Membership changes require the already recorded collection-update decision")
        for name, action in (("add", "add_membership"), ("remove", "remove_membership")):
            values = p.get(name)
            _require(isinstance(values, list) and all(isinstance(v, dict) for v in values), "Invalid membership difference rows")
            uris = [v.get("uri") for v in values]
            _require(len(uris) == len(set(uris)) and set(uris) == expected[name], "Membership difference was altered or contains a total-library deletion")
            for uri in uris:
                operations.append({"operation_id": action + ":" + _digest([identity, uri]),
                                   "action": action, "uri": uri, "target_collection": identity,
                                   "expected_record_sha256": core.digest(known[uri]),
                                   "expected_record_version": known[uri].get("version"),
                                   "required_capability": "zotero_collection_membership",
                                   "preflight": "Re-read this item and collection; use native version preconditions when available; preserve other memberships and total-library record",
                                   "policy": "Membership only; never delete/trash/merge the item or recreate the collection"})
        collection_binding = {**identity, "membership_plan_sha256": _digest(p),
                              "fingerprint": core.collection_fingerprint(col, members),
                              "current_members": sorted(members), "desired_members": sorted(desired)}
    else:
        _require(collection_snapshot is None, "A collection snapshot alone does not authorize membership work")

    plan = {"schema": PLAN, "result": "BATCH_PLANNED", "created_at": _now(),
            "source_document_sha256": current_document_sha256,
            "preparation_sha256": _digest(preparation) if preparation is not None else None,
            "library_id": lib["library_id"], "library_fingerprint": core.library_fingerprint(lib, known),
            "library_record_hashes": {u: core.digest(i) for u, i in known.items()},
            "library_record_preconditions": {u: {"content_sha256": core.digest({k: v for k, v in i.items() if k not in ("version", "collections")}),
                                                  "version": i.get("version"), "has_collections": "collections" in i,
                                                  "collections": i.get("collections")}
                                             for u, i in known.items()},
            "library_observed_at": lib["observed_at"], "collection": collection_binding,
            "batch_size": batch_size, "operations": operations, "batches": _batches(operations, batch_size),
            "import_requests": import_rows, "reuse_existing": reuse, "review_queue": reviews,
            "counts": {"query_groups": len(groups), "source_positions": len(valid_sources),
                       "unique_import_requests": len(import_rows), "reuse_existing": len(reuse),
                       "review_required": len(reviews), **dict(Counter(op["action"] for op in operations))},
            "policy": "Adapter tasks only. No MCP calls, imports, membership changes, merges, deletions or native Word actions have been performed. A plan/status is not citation completion evidence. Discover real host capabilities and honor provider batch limits."}
    plan["plan_fingerprint"] = _plan_fingerprint(plan)
    return plan


def status(plan, library_snapshot, *, current_document_sha256,
           collection_snapshot=None, observations=None, max_age_hours=24):
    """Recheck batch effects; do not replay unknown writes or trust success text."""
    core = _helper("zotero_docx")
    _require(isinstance(plan, dict) and plan.get("schema") == PLAN
             and plan.get("plan_fingerprint") == _plan_fingerprint(plan), "Batch plan changed; build a fresh plan")
    _require(plan.get("source_document_sha256") == current_document_sha256,
             "Saved document changed after the batch plan")
    lib, known = core.validate_snapshot(library_snapshot, "library", max_age_hours)
    _require(lib["library_id"] == plan["library_id"], "Batch status targets another library")
    operations = plan.get("operations")
    _require(isinstance(operations, list), "Batch plan has no operation list")
    by_id = {op.get("operation_id"): op for op in operations}
    _require(len(by_id) == len(operations) and all(_text(k) for k in by_id), "Batch operation IDs must be distinct")
    observations = {"schema": OBSERVATIONS, "plan_fingerprint": plan["plan_fingerprint"], "results": []} if observations is None else observations
    _require(isinstance(observations, dict) and observations.get("schema") == OBSERVATIONS
             and observations.get("plan_fingerprint") == plan["plan_fingerprint"], "Batch observations belong to another plan")
    entries = observations.get("results")
    _require(isinstance(entries, list), "Batch observations must list per-operation outcomes")
    observed = {}
    for entry in entries:
        oid = entry.get("operation_id") if isinstance(entry, dict) else None
        _require(oid in by_id and oid not in observed and _text(entry.get("evidence"))
                 and entry.get("outcome") in ("succeeded", "failed", "unknown", "not_applied"),
                 "Each actual batch outcome needs one known operation, status and evidence")
        when = core.iso(entry.get("observed_at", ""))
        _require(when >= core.iso(plan["created_at"]), "An outcome predates its batch plan")
        if by_id[oid]["action"] != "lookup":
            _require(core.iso(lib["observed_at"]) >= when,
                     "Re-export/re-observe the library after each completed or uncertain write; do not replay from old snapshots")
        observed[oid] = entry
    writes_observed = [entry for oid, entry in observed.items() if by_id[oid]["action"] != "lookup"]
    membership_observed = [entry for oid, entry in observed.items()
                           if by_id[oid]["action"] in ("add_membership", "remove_membership")]
    if writes_observed:
        _require(core.iso(lib["observed_at"]) >= core.iso(plan["created_at"]), "Batch recovery needs post-plan observations")
    # Imports may add records. Ordinary metadata must remain intact. A native
    # membership update can legitimately advance an item version; its narrow
    # exception is checked only after the current target collection is read.
    expected_hashes = plan["library_record_hashes"]
    drift = [uri for uri, fingerprint in expected_hashes.items()
             if uri not in known or core.digest(known[uri]) != fingerprint]
    binding = plan.get("collection")
    members = None
    if binding:
        _require(collection_snapshot is not None, "Membership status requires a fresh observed collection")
        col, members = _collection(core, {**lib, "max_age_hours": max_age_hours}, known, collection_snapshot)
        _require(all(col.get(k) == binding.get(k) for k in ("library_id", "collection_name", "collection_path", "collection_key")),
                 "Observed collection changed identity")
        if membership_observed:
            _require(core.iso(col["observed_at"]) >= max(core.iso(e["observed_at"]) for e in membership_observed),
                     "Collection observation predates the completed/uncertain batch")
        allowed = set(binding["current_members"]) | set(binding["desired_members"])
        _require(set(members) <= allowed, "Unplanned collection memberships appeared; rebuild after observing current state")
        retained = set(binding["current_members"]) & set(binding["desired_members"])
        _require(retained <= set(members), "Unplanned retained collection memberships disappeared; observe current state and rebuild")
    else:
        _require(collection_snapshot is None, "This batch has no collection target")

    verified_version_changes = []
    for uri in drift:
        precondition = plan.get("library_record_preconditions", {}).get(uri, {})
        member_ops = [op for op in operations if op["action"] in ("add_membership", "remove_membership") and op["uri"] == uri]
        safe = (uri in known and members is not None and len(member_ops) == 1
                and precondition.get("content_sha256") == core.digest({k: v for k, v in known[uri].items() if k not in ("version", "collections")})
                and core.iso(lib["observed_at"]) >= core.iso(plan["created_at"])
                and core.iso(col["observed_at"]) >= core.iso(plan["created_at"]))
        if safe:
            op = member_ops[0]
            safe = ((uri in members) == (op["action"] == "add_membership"))
            before_version, after_version = precondition.get("version"), known[uri].get("version")
            safe = safe and ((before_version is None and after_version is None)
                             or (type(before_version) is int and type(after_version) is int and after_version > before_version))
            before_collections, after_collections = precondition.get("collections"), known[uri].get("collections")
            if precondition.get("has_collections") or "collections" in known[uri]:
                target_key = binding.get("collection_key")
                safe = (safe and _text(target_key) and isinstance(before_collections, list)
                        and isinstance(after_collections, list)
                        and all(_text(k) for k in before_collections + after_collections)
                        and len(before_collections) == len(set(before_collections))
                        and len(after_collections) == len(set(after_collections)))
                if safe:
                    expected_collections = set(before_collections)
                    if op["action"] == "add_membership": expected_collections.add(target_key)
                    else: expected_collections.discard(target_key)
                    safe = set(after_collections) == expected_collections
        _require(safe, "Existing library records/versions changed or disappeared; preserve current state and rebuild the batch plan: " + uri)
        verified_version_changes.append({"uri": uri, "before_version": before_version,
                                         "after_version": after_version,
                                         "basis": "Planned membership effect observed; ordinary metadata and any exposed other memberships preserved"})

    results, ready = [], []
    for op in operations:
        oid = op["operation_id"]; entry = observed.get(oid)
        action = op["action"]; found = []
        state = "PENDING_ADAPTER"
        if action == "import_item":
            found = _hits(op["identifier"], known, core)
            if len(found) == 1:
                state = "VERIFIED_PRESENT"
                if entry and entry.get("uri"):
                    _require(core.canonical_uri(entry["uri"]) == found[0], "Reported imported identity conflicts with the actual library")
            elif len(found) > 1:
                state = "IDENTITY_REVIEW_REQUIRED"
            elif entry:
                state = "VERIFY_OUTCOME" if entry["outcome"] != "not_applied" else "PENDING_ADAPTER"
                if entry["outcome"] == "not_applied":
                    _require(entry.get("no_write_verified") is True,
                             "Retrying an import requires actual no-write verification, not only an error response")
        elif action in ("add_membership", "remove_membership"):
            _require(members is not None, "Membership task has no observed collection")
            present = op["uri"] in members
            if present == (action == "add_membership"):
                state = "VERIFIED_APPLIED"
            elif entry:
                state = "VERIFY_OUTCOME" if entry["outcome"] != "not_applied" else "PENDING_ADAPTER"
                if entry["outcome"] == "not_applied":
                    _require(entry.get("no_write_verified") is True,
                             "A membership retry needs actual no-write verification")
        elif action == "lookup":
            state = "LOOKUP_REVIEW_REQUIRED" if entry and entry["outcome"] == "succeeded" else "PENDING_ADAPTER"
        else:
            raise ValueError("Unsupported batch action; total-library deletion/merge is never a batch operation")
        if state == "PENDING_ADAPTER":
            ready.append(op)
        results.append({"operation_id": oid, "action": action, "status": state,
                        "candidate_uris": found, "observed_outcome": entry.get("outcome") if entry else None})
    counts = dict(Counter(row["status"] for row in results))
    pending_review = bool(plan["review_queue"] or any(row["status"] in ("IDENTITY_REVIEW_REQUIRED", "LOOKUP_REVIEW_REQUIRED") for row in results))
    unknown = any(row["status"] == "VERIFY_OUTCOME" for row in results)
    return {"schema": STATUS, "result": "BATCH_STATUS", "status": "RECHECK_REQUIRED" if unknown else "READY" if ready else "REVIEW_REQUIRED" if pending_review else "OBSERVED_BATCH_EFFECTS",
            "plan_fingerprint": plan["plan_fingerprint"], "source_document_sha256": current_document_sha256,
            "library_id": lib["library_id"], "library_after_fingerprint": core.library_fingerprint(lib, known),
            "verified_membership_version_changes": verified_version_changes,
            "operations": results, "ready_operations": ready, "ready_batches": _batches(ready, plan["batch_size"]),
            "review_queue": plan["review_queue"], "counts": counts,
            "policy": "Snapshot recheck only; not a new native observation or a citation/finalization PASS. Unknown writes must be inspected, not replayed. Final membership still requires the existing verify/finalize engine and real native evidence."}

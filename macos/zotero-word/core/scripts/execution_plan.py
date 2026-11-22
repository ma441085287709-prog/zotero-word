"""Batch local candidate enumeration without resolving papers or editing apps.

Only retrieval keys are deduplicated. Every source clue and native citation
keeps its own position and pending review, even when it cites the same paper.
"""
from __future__ import annotations

from collections import defaultdict
import hashlib
import importlib.util
import json
from pathlib import Path
import re
import sys
import unicodedata
from urllib.parse import urlsplit, urlunsplit


def _load_cues():
    spec = importlib.util.spec_from_file_location(
        "_zotero_word_execution_cues", Path(__file__).with_name("citation_cues.py"))
    module = importlib.util.module_from_spec(spec)
    previous = sys.dont_write_bytecode
    try:
        sys.dont_write_bytecode = True
        spec.loader.exec_module(module)
    finally:
        sys.dont_write_bytecode = previous
    return module


_CUES = _load_cues()
SCHEMA = "zotero-word.execution-preparation.v1"


def required_phases():
    """A pending dependency plan, never a claim that host tools were executed."""
    phases = [
        ("matching", [], ["read_local_document_and_library"], "Bind the independent copy and resolve its actual citation positions."),
        ("native_conversion", ["matching"], ["native_word_zotero"], "Write the bound fields in the copy and apply the requested style."),
        ("native_save_reopen", ["native_conversion"], ["native_word_zotero"], "Use the saved render, publish on the protected source and reopen the delivery."),
        ("current_file_checks", ["native_save_reopen"], ["local_joint_audit"], "Check protected content and actual editable Zotero items in the delivered file."),
    ]
    return [{"phase": name, "depends_on": dependencies, "required_capabilities": capabilities,
             "status": "PENDING", "action": action} for name, dependencies, capabilities, action in phases]


def _digest(value):
    return hashlib.sha256(json.dumps(value, ensure_ascii=False, sort_keys=True,
                                     separators=(",", ":")).encode()).hexdigest()


def _name(value):
    return unicodedata.normalize("NFKC", str(value or "")).casefold().strip()


def _url(value):
    """Normalize scheme/host and fragments, retaining article-defining queries."""
    if not isinstance(value, str):
        return ""
    try:
        p = urlsplit(value.strip())
    except ValueError:
        return ""
    if p.scheme.casefold() not in ("http", "https") or not p.netloc:
        return ""
    return urlunsplit((p.scheme.casefold(), p.netloc.casefold(), p.path.rstrip("/"), p.query, ""))


def _first_author(authors):
    # This is an exact local retrieval hint, not a general name parser. Missing
    # family-name/online-year aliases remain an external semantic-search task.
    authors = unicodedata.normalize("NFKC", str(authors or "")).strip()
    authors = re.sub(r"(?:\s+et\s+al\.?|\s*等)\s*$", "", authors, flags=re.I)
    authors = re.split(r"\s+(?:and|&)\s+", authors, maxsplit=1, flags=re.I)[0]
    return _name(authors)


def _summary(uri, item):
    return {"uri": uri, "DOI": item.get("DOI"), "title": item.get("title"),
            "author": item.get("author", []), "issued": item.get("issued"),
            "container-title": item.get("container-title"),
            "source_record_sha256": _digest(item)}


def query_batches(groups, batch_size=10):
    """Bound distinct queries per planning batch, without assuming a provider API.

    These batches are retrieval/review work, not permission to insert citations
    or confirmation that any result has been resolved. Repeated positions are
    retained; a provider may require smaller batches or separate requests.
    """
    if isinstance(batch_size, bool) or not isinstance(batch_size, int) or batch_size < 1:
        raise ValueError("Query batch size must be a positive integer")
    batches = []
    for offset in range(0, len(groups), batch_size):
        batch = groups[offset:offset + batch_size]
        source_ids = list(dict.fromkeys(sid for group in batch for sid in group["source_ids"]))
        batches.append({"batch_index": len(batches) + 1,
                        "query_keys": [group["key"] for group in batch],
                        "source_ids": source_ids,
                        "distinct_query_count": len(batch),
                        "source_position_count": len(source_ids),
                        "status": "PENDING_REVIEW"})
    return batches


def prepare(report, known, *, query_batch_size=10):
    """Prepare all local lookups in one pass; every result still needs review.

    ``known`` must already be validated by the caller against the current bound
    library snapshot. ``None`` means no library was provided, not an empty one.
    This function does not read a library, use a network or cache app state.
    ``query_batch_size`` only bounds planning batches; it is not a provider's
    supported request size. Queues contain unreviewed source positions and
    existing fields to retain by default, never completed application actions.
    """
    if not isinstance(report, dict) or not isinstance(report.get("citation_cues"), list):
        raise ValueError("Preparation requires a current inspection with citation_cues")
    if known is not None and not isinstance(known, dict):
        raise ValueError("Known records must be a validated URI-to-record mapping or None")
    source_hash = report.get("document_sha256")
    if not isinstance(source_hash, str) or not re.fullmatch(r"[a-f0-9]{64}", source_hash):
        raise ValueError("Preparation requires the actual saved source-document hash")
    if not isinstance(report.get("citations"), list):
        raise ValueError("Preparation requires the observed native citation list")
    # Validate even an empty report, where no batches would otherwise be built.
    if isinstance(query_batch_size, bool) or not isinstance(query_batch_size, int) or query_batch_size < 1:
        raise ValueError("Query batch size must be a positive integer")

    doi_index, url_index, author_year_index = defaultdict(list), defaultdict(list), defaultdict(list)
    if known is not None:
        for uri, item in known.items():
            if not isinstance(uri, str) or not isinstance(item, dict):
                raise ValueError("A validated library record must have a URI key and object value")
            identifier = _CUES.normalize_doi(item.get("DOI"))
            if identifier:
                doi_index[identifier].append(uri)
            url = _url(item.get("URL"))
            if url:
                url_index[url].append(uri)
            authors = item.get("author", [])
            parts = (item.get("issued", {}).get("date-parts") or [[]]) if isinstance(item.get("issued", {}), dict) else [[]]
            first = authors[0] if isinstance(authors, list) and authors and isinstance(authors[0], dict) else {}
            family = _name(first.get("family", first.get("literal", "")))
            year = str(parts[0][0]) if isinstance(parts, list) and parts and isinstance(parts[0], list) and parts[0] else ""
            if family and re.fullmatch(r"\d{4}", year):
                author_year_index[(family, year)].append(uri)

    groups, group_sources, order, pending, seen = {}, {}, [], [], set()
    cue_queue, native_queue = [], []
    manifest_entries = []
    def add_group(query, position, manual_status=None, year_label=None):
        key = "query:" + _digest(query)
        if key not in groups:
            candidate_uris = []
            if manual_status:
                status = manual_status
            elif known is None:
                status = "LIBRARY_NOT_PROVIDED"
            else:
                if query["kind"] == "doi":
                    candidate_uris = doi_index.get(query["DOI"], [])
                elif query["kind"] == "url":
                    candidate_uris = url_index.get(query["URL"], [])
                elif query["kind"] == "author_year":
                    candidate_uris = author_year_index.get((query["author"], query["year"]), [])
                status = "NEEDS_SEMANTIC_VERIFICATION" if candidate_uris else "NO_LOCAL_CANDIDATES"
            groups[key] = {"key": key, "query": query, "source_ids": [], "positions": [],
                           "candidate_uris": list(candidate_uris), "status": status,
                           "year_labels": [], "lookup_performed": known is not None and not manual_status}
            group_sources[key] = set()
            order.append(key)
        group = groups[key]
        if position["source_id"] not in group_sources[key]:
            group_sources[key].add(position["source_id"])
            group["source_ids"].append(position["source_id"])
            group["positions"].append(position)
        if year_label and year_label not in group["year_labels"]:
            group["year_labels"].append(year_label)

    for cue in report["citation_cues"]:
        if not isinstance(cue, dict) or not isinstance(cue.get("cue_id"), str) or not cue["cue_id"] or cue["cue_id"] in seen:
            raise ValueError("Every source cue requires a unique cue_id")
        sid = cue["cue_id"]; seen.add(sid); pending.append(sid)
        position = {"source_id": sid, **{k: cue[k] for k in ("story", "paragraph", "location", "start", "end", "text", "context_excerpt") if k in cue}}
        if cue.get("link_target_conflict"):
            position["link_target_conflict"] = True
        cue_queue.append(dict(position, status="PENDING_REVIEW"))
        manifest_entries.append({"source_id": sid, "disposition": "pending", "evidence": ""})
        identifier = _CUES.normalize_doi(cue.get("DOI") or cue.get("target") or "")
        if identifier:
            add_group({"kind": "doi", "DOI": identifier}, position)
        elif cue.get("kind") == "article_link" and _url(cue.get("target")):
            add_group({"kind": "url", "URL": _url(cue["target"])}, position)
        elif cue.get("kind") == "author_year":
            author = _first_author(cue.get("authors"))
            years = cue.get("years", [])
            valid_years = [str(v) for v in years if re.fullmatch(r"\d{4}[a-z]?", str(v))] if isinstance(years, list) else []
            if author and valid_years:
                for label in valid_years:
                    add_group({"kind": "author_year", "author": author, "year": label[:4]}, position, year_label=label)
            else:
                add_group({"kind": "manual", "source_id": sid}, position, "NEEDS_MANUAL_QUERY")
        elif cue.get("kind") == "numeric":
            # Identical numbers in different parts of a mixed draft can point
            # to different static bibliographies. Do not deduplicate or guess.
            add_group({"kind": "static_reference_mapping", "source_id": sid}, position, "NEEDS_REFERENCE_MAPPING")
        else:
            add_group({"kind": "manual", "source_id": sid}, position, "NEEDS_MANUAL_QUERY")

    native_ids = set()
    for citation in report["citations"]:
        cid = citation.get("citationID") if isinstance(citation, dict) else None
        if not isinstance(cid, str) or not cid or cid in native_ids:
            raise ValueError("Native citations require distinct observed citationID values")
        native_ids.add(cid)
        sid = "field:" + cid
        manifest_entries.append({"source_id": sid, "disposition": "pending", "evidence": ""})
        pending.append(sid)
        context = citation.get("context", {})
        items = citation.get("citationItems", [])
        observed_uris = sorted({uri for item in items if isinstance(item, dict)
                                for uri in (item.get("uris") if isinstance(item.get("uris"), list) else [])
                                if isinstance(uri, str)}) if isinstance(items, list) else []
        native_queue.append({"source_id": sid, "citation_id": cid,
                             "default_action": "retain_and_review",
                             "status": "PENDING_REVIEW", "observed_uris": observed_uris,
                             "position": {key: context[key] for key in ("story", "paragraph", "location", "start", "end")
                                          if isinstance(context, dict) and key in context}})

    records = {}
    for key in order:
        group = groups[key]
        for uri in group["candidate_uris"]:
            if uri not in records:
                records[uri] = _summary(uri, known[uri])
    result_groups = [groups[key] for key in order]
    batches = query_batches(result_groups, query_batch_size)
    source_query_keys = defaultdict(list)
    candidate_counts = {"zero": 0, "single": 0, "multiple": 0, "not_looked_up": 0}
    for group in result_groups:
        for sid in group["source_ids"]:
            source_query_keys[sid].append(group["key"])
        count = len(group["candidate_uris"])
        bucket = "not_looked_up" if not group["lookup_performed"] else "zero" if count == 0 else "single" if count == 1 else "multiple"
        candidate_counts[bucket] += 1
    for position in cue_queue:
        position["query_keys"] = source_query_keys[position["source_id"]]
    placeholders = report.get("unresolved_placeholders", [])
    if not isinstance(placeholders, list):
        raise ValueError("Placeholder diagnostics must be a per-position list")
    workload = {"source_cue_position_count": len(cue_queue),
                "placeholder_count": len(placeholders),
                "existing_native_field_count": len(native_queue),
                "unique_query_count": len(result_groups), "query_batch_count": len(batches),
                "query_batch_size": query_batch_size,
                "local_query_candidate_counts": candidate_counts,
                "unique_local_candidate_item_count": len(records),
                "library_index_pass_count": int(known is not None),
                "library_records_indexed_count": len(known) if known is not None else 0,
                "pending_ui_cue_position_count": len(cue_queue),
                "new_native_field_count": None, "estimated_seconds": None,
                "scope": "Counts describe pending source positions, not verified papers or required new fields. Existing native fields are retained by default. Semantic review may classify non-citations or map multiple co-located clues to one field. No UI throughput or ETA has been measured."}
    return {"schema": SCHEMA, "source_document_sha256": source_hash,
            "cue_count": len(report["citation_cues"]), "placeholder_count": len(placeholders), "unique_query_count": len(result_groups),
            "lookup_count": sum(g["lookup_performed"] for g in result_groups),
            "library_record_count": len(known) if known is not None else None,
            "library_status": "VALIDATED_LOCAL_SNAPSHOT" if known is not None else "LIBRARY_NOT_PROVIDED",
            "native_citation_count": len(native_ids), "query_groups": result_groups,
            "candidate_records": records, "unresolved_semantic_positions": pending,
            "workload": workload, "query_batches": batches,
            "execution_phases": required_phases(),
            "execution_policy": "Execute ready authorized phases automatically. A blocked capability stops only dependent work. This pending plan does not call host tools, certify an adapter or prove any phase completed.",
            "review_queues": {"source_document_sha256": source_hash,
                              "cue_positions": cue_queue,
                              "placeholder_diagnostics": placeholders,
                              "existing_native_fields": native_queue},
            "manifest_template": {"schema": "zotero-word.coverage-manifest.v1",
                "source_document_sha256": source_hash, "target_document_sha256": None,
                "entries": manifest_entries, "source_reference_blocks": [], "added_citations": [],
                "consolidations": [], "body_changes": []},
            "scope": "Local exact candidate enumeration only. Distinct citation positions and per-context claim support remain unresolved. A unique DOI match is not a merge, selection, import or authorization. Numeric markers require a verified static-reference mapping. Revalidate current app state and final saved-document/metadata evidence before writing or completion."}

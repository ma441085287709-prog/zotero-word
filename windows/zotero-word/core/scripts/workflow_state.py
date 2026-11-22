"""Bind an optional layout decision to an already completed citation artifact.

This module never opens applications or displays a pretend dialog. The host
must actually present the offered question and record the user's response.
It never changes the citation manuscript or its completed baseline.
"""
from __future__ import annotations

from datetime import datetime, timezone
import hashlib
import json
from pathlib import Path

OFFER = "zotero-word.layout-offer.v1"
DECISION = "zotero-word.layout-decision.v1"
CONSENT = "zotero-word.layout-consent.v1"


def require(ok, message):
    if not ok:
        raise ValueError(message)


def digest(value):
    raw = value if isinstance(value, bytes) else json.dumps(value, sort_keys=True,
        ensure_ascii=False, separators=(",", ":")).encode()
    return hashlib.sha256(raw).hexdigest()


def read(path):
    raw = Path(path).read_bytes()
    value = json.loads(raw)
    require(isinstance(value, dict), "Workflow evidence must be a JSON object")
    return raw, value


def now():
    return datetime.now(timezone.utc).isoformat()


def checkpoint(document, baseline_path, reference_profile=None):
    document, baseline_path = Path(document).resolve(), Path(baseline_path).resolve()
    raw, baseline = read(baseline_path)
    if baseline.get("schema") == "zotero-word.windows-run.v2":
        actual = digest(document.read_bytes())
        verified = baseline.get("verified", {})
        updated = ((baseline.get("completion_scope") == "zotero_refresh_word_save_close"
                    and baseline.get("native_refresh_observed") is True)
                   or (baseline.get("completion_scope") == "zotero_preferences_update_word_save_close"
                       and baseline.get("native_update_origin") == "preferences_update"))
        direct = (updated and baseline.get("word_window") == "closed"
                  and baseline.get("paths", {}).get("citation_document") == str(document)
                  and baseline.get("paths", {}).get("delivery_document") == str(document))
        reopened = (baseline.get("native_validation_verified") is True
                    and verified.get("samples") and verified.get("style") == baseline.get("style"))
        require(baseline.get("result") == "COMPLETE" and (direct or reopened)
                and Path(baseline["document"]).resolve() == document and baseline.get("document_sha256") == actual
                and baseline.get("protection", {}).get("target_sha256") == actual,
                "Optional layout requires the accepted and unchanged native working copy")
        require(reference_profile, "Provide the citation format profile for the separately requested layout")
        profile_raw, profile = read(reference_profile)
        require(profile.get("schema") == "zotero-word.reference-format-profile.v1"
                and profile.get("csl", {}).get("style_id") == baseline["style"] and profile.get("journal"),
                "Layout reference profile does not match the accepted citation style")
        return {"citation_document": str(document), "citation_document_sha256": actual,
                "citation_baseline": str(baseline_path), "citation_baseline_sha256": digest(raw),
                "reference_profile": str(Path(reference_profile).resolve()), "reference_profile_sha256": digest(profile),
                "reference_profile_file_sha256": digest(profile_raw), "journal": profile["journal"], "citation_stage_status": "COMPLETE"}
    audit = baseline.get("completion_audit", {})
    require(baseline.get("schema") == "zotero-word.baseline.v1"
            and isinstance(audit, dict) and audit.get("result") == "PASS"
            and audit.get("completion_scope") == "citations_and_references"
            and baseline.get("native_refresh_verified") is True,
            "Optional layout requires a completed citation/References baseline")
    actual = digest(document.read_bytes())
    require(baseline.get("document_sha256") == actual == audit.get("target_document_sha256")
            and Path(audit.get("target_document", "")).resolve() == document,
            "Citation checkpoint changed or belongs to another manuscript")
    require(audit.get("references_format_status") in ("PASS", "NOT_APPLICABLE"),
            "References must pass before offering layout")
    profile_path = audit.get("reference_profile")
    require(isinstance(profile_path, str) and profile_path,
            "The completed baseline must bind its reference profile; revalidate this checkpoint once")
    profile_raw, profile = read(profile_path)
    require(digest(profile) == audit.get("reference_profile_sha256")
            and profile.get("schema") == "zotero-word.reference-format-profile.v1",
            "The citation checkpoint's reference profile changed")
    journal = profile.get("journal")
    require(isinstance(journal, str) and journal.strip(), "The checkpoint has no target journal")
    return {"citation_document": str(document), "citation_document_sha256": actual,
            "citation_baseline": str(baseline_path), "citation_baseline_sha256": digest(raw),
            "reference_profile": str(Path(profile_path).resolve()),
            "reference_profile_sha256": digest(profile),
            "reference_profile_file_sha256": digest(profile_raw), "journal": journal,
            "citation_stage_status": "COMPLETE"}


def _same_binding(value, current):
    for name, expected in current.items():
        require(value.get(name) == expected, "Optional layout evidence is stale: " + name)


def offer(document, baseline_path, reference_profile=None):
    binding = checkpoint(document, baseline_path, reference_profile)
    return {"schema": OFFER, "result": "AWAITING_DECISION", **binding,
            "offered_at": now(), "layout_stage_status": "AWAITING_DECISION",
            "question": "引文与 References 已整理并核验完成。是否继续按目标期刊要求匹配全文格式？",
            "options": ["是，继续全文格式匹配", "否，保留引用完成稿"],
            "requires_explicit_yes": True,
            "next_action": "RECORD_EXPLICIT_LAYOUT_REQUEST_OR_CLARIFY",
            "scope": "Record the user’s explicit request for this manuscript and journal. An offer alone grants no consent."}


def decide(document, offer_path, choice, decision_source):
    require(choice in ("yes", "no", "pending"), "Layout choice must be yes, no or pending")
    offer_raw, offered = read(offer_path)
    require(offered.get("schema") == OFFER and offered.get("result") == "AWAITING_DECISION"
            and offered.get("requires_explicit_yes") is True,
            "A valid post-citation layout offer is required")
    current = checkpoint(document, offered.get("citation_baseline", ""), offered.get("reference_profile"))
    _same_binding(offered, current)
    require(isinstance(decision_source, str) and decision_source.strip(),
            "Record the actual host-dialog or chat response; silence is not approval")
    status = {"yes": "AUTHORIZED", "no": "DECLINED", "pending": "AWAITING_DECISION"}[choice]
    return {"schema": DECISION, "result": status, **current,
            "offer_path": str(Path(offer_path).resolve()), "offer_sha256": digest(offer_raw),
            "choice": choice, "decision_source": decision_source, "decided_at": now(),
            "layout_stage_status": status, "citation_checkpoint_retained": True,
            "next_action": "PREPARE_OPTIONAL_LAYOUT" if choice == "yes" else "KEEP_CITATION_CHECKPOINT"}


def authorize(decision_path, citation_document=None, baseline_path=None):
    decision_raw, decision = read(decision_path)
    require(decision.get("schema") == DECISION and decision.get("choice") == "yes"
            and decision.get("result") == "AUTHORIZED"
            and isinstance(decision.get("decision_source"), str) and decision["decision_source"].strip(),
            "Full layout requires an explicit yes after citation completion")
    offer_raw, offered = read(decision.get("offer_path", ""))
    require(digest(offer_raw) == decision.get("offer_sha256")
            and offered.get("schema") == OFFER and offered.get("requires_explicit_yes") is True
            and offered.get("result") == "AWAITING_DECISION",
            "The actual layout offer changed after the response")
    document = citation_document or decision.get("citation_document", "")
    baseline = baseline_path or decision.get("citation_baseline", "")
    current = checkpoint(document, baseline, offered.get("reference_profile"))
    _same_binding(offered, current)
    _same_binding(decision, current)
    return {"schema": CONSENT, **current,
            "decision_path": str(Path(decision_path).resolve()),
            "decision_file_sha256": digest(decision_raw), "decision_sha256": digest(decision_raw),
            "baseline_path": current["citation_baseline"],
            "baseline_sha256": current["citation_baseline_sha256"],
            "source_document_sha256": current["citation_document_sha256"],
            "choice": "yes"}


def validate_context(context):
    require(isinstance(context, dict) and context.get("schema") == CONSENT,
            "A bound layout consent context is required")
    current = authorize(context.get("decision_path", ""), context.get("citation_document"),
                        context.get("citation_baseline"))
    _same_binding(context, current)
    return current

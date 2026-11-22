"""Read-only coverage, ordinary-text and completion checks for saved manuscripts.

The receipts record an agent's evidence and decisions; they do not manufacture
native Word/Zotero observations or independently establish claim support.
"""
from __future__ import annotations

from collections import defaultdict
import hashlib
import json
import re
from urllib.parse import urlparse


MANIFEST = "zotero-word.coverage-manifest.v1"
QUALITY = "zotero-word.completion-review.v1"
AUDIT = "zotero-word.manuscript-audit.v1"
_URI = re.compile(r"^/(users/(?:local/[^/]+|[0-9]+)|groups/[0-9]+)/items/([A-Z0-9]{8})/?$")
_SEPARATORS = re.compile(r"^[\s,;:，；：（）()\[\]{}<>/\\|&—–-]*$")


def text_sha256(text):
    """The receipt hash of an exact, unnormalised block.text value."""
    return hashlib.sha256(text.encode("utf-8")).hexdigest()


def _digest(value):
    return hashlib.sha256(json.dumps(value, sort_keys=True, ensure_ascii=False,
                                     separators=(",", ":")).encode()).hexdigest()


def _nonempty(value):
    return isinstance(value, str) and bool(value.strip())


def _key(value):
    if not isinstance(value, dict):
        raise ValueError("A block/position must be an object")
    story, paragraph = value.get("story"), value.get("paragraph")
    if not _nonempty(story) or type(paragraph) is not int or paragraph < 1:
        raise ValueError("A block/position requires its exact story and positive paragraph")
    return story, paragraph


def _ordinary_text(block, ranges=(), limit=None):
    """Remove only observed native results and approved exact cue spans.

    Ordinary spaces, tabs, line breaks and punctuation remain exact. Native
    display lengths may change between styles; their masked padding is not
    manuscript text and must not be confused with real whitespace.
    """
    text = block["text"]
    end = len(text) if limit is None else limit
    if type(end) is not int or not 0 <= end <= len(text):
        raise ValueError("An ordinary-text prefix lies outside its exact block")
    excluded = [False] * len(text)
    segments = block.get("visible_segments")
    if not isinstance(segments, list):
        raise ValueError("Fresh visible_segments are required to preserve ordinary whitespace")
    for segment in segments:
        if segment.get("kind") == "zotero":
            start, stop = segment.get("start"), segment.get("end")
            if type(start) is int and start == stop and 0 <= start <= len(text):
                continue
            _mask(text, [(start, stop)])
            excluded[start:stop] = [True] * (stop - start)
    for start, stop in ranges:
        _mask(text, [(start, stop)])
        excluded[start:stop] = [True] * (stop - start)
    return "".join(char for index, char in enumerate(text[:end]) if not excluded[index])


def _mask(text, ranges):
    chars = list(text)
    for start, end in ranges:
        if type(start) is not int or type(end) is not int or not 0 <= start < end <= len(text):
            raise ValueError("A citation range is outside its exact paragraph")
        chars[start:end] = " " * (end - start)
    return "".join(chars)


def _uris(citation):
    items = citation.get("citationItems")
    if not isinstance(items, list) or not items:
        raise ValueError("An observed citation has no citationItems")
    found = set()
    for item in items:
        uris = item.get("uris") if isinstance(item, dict) else None
        if not isinstance(uris, list) or not uris:
            raise ValueError("An observed citation has no Zotero URI")
        for uri in uris:
            parsed = urlparse(uri) if isinstance(uri, str) else None
            match = _URI.fullmatch(parsed.path) if parsed else None
            if not parsed or parsed.scheme not in ("http", "https") or parsed.netloc.lower() not in ("zotero.org", "www.zotero.org") or parsed.query or parsed.fragment or not match:
                raise ValueError("An observed citation contains an invalid Zotero URI")
            found.add("http://zotero.org/" + match.group(1) + "/items/" + match.group(2))
    return found


def audit(source_report, target_report, manifest, quality, protected_equivalence=None):
    """Citation/body diagnostic; public completion also checks CSL/References.

    Never alter a manuscript, library, receipt or application. A limited PASS
    here alone does not certify native execution or References formatting.
    """
    issues = []

    def check(condition, message):
        if not condition:
            issues.append(message)
        return bool(condition)

    counts = {}
    alignment_receipt = []
    try:
        if not all(isinstance(v, dict) for v in (source_report, target_report, manifest, quality)):
            raise ValueError("Inspection reports, coverage manifest and completion review must be objects")
        source_hash, target_hash = source_report.get("document_sha256"), target_report.get("document_sha256")
        equivalent_math = protected_equivalence is not None
        if equivalent_math:
            evidence = protected_equivalence
            if not (isinstance(evidence, dict) and evidence.get("schema") == "zotero-word.math-equivalence-verification.v1"
                    and evidence.get("result") == "PASS" and not evidence.get("issues")
                    and evidence.get("source_document_sha256") == source_hash and evidence.get("target_document_sha256") == target_hash
                    and evidence.get("source_raw_equations_sha256") == source_report.get("protected_content_sha256", {}).get("equations")
                    and evidence.get("target_raw_equations_sha256") == target_report.get("protected_content_sha256", {}).get("equations")):
                raise ValueError("Math equivalence must be freshly computed for both original saved inspections")

        def protected_tokens(block, source):
            tokens = block["protected_tokens"]
            if not equivalent_math:
                return tokens
            mapping = protected_equivalence.get("source_equation_token_map" if source else "target_equation_token_map")
            if not isinstance(mapping, dict):
                raise ValueError("Math equivalence has no complete token mapping")
            result = []
            for token in tokens:
                if token.startswith("equation:"):
                    raw = token.removeprefix("equation:")
                    value = mapping.get(raw)
                    if not isinstance(value, str) or not re.fullmatch(r"[a-f0-9]{64}", value):
                        raise ValueError("An original equation anchor lacks computed equivalence evidence")
                    token = "equation:" + value
                result.append(token)
            return result
        check(bool(re.fullmatch(r"[a-f0-9]{64}", source_hash or "")), "Source inspection has no valid saved-document hash")
        check(bool(re.fullmatch(r"[a-f0-9]{64}", target_hash or "")), "Target inspection has no valid saved-document hash")
        check(not source_report.get("issues"), "Source contains unresolved structural issues")
        check(not target_report.get("issues"), "Target contains unresolved structural issues")
        check(manifest.get("schema") == MANIFEST, "Unsupported coverage-manifest schema")
        check(manifest.get("source_document_sha256") == source_hash and manifest.get("target_document_sha256") == target_hash,
              "Coverage manifest belongs to another source or saved target")
        check(quality.get("schema") == QUALITY, "Unsupported completion-review schema")
        check(quality.get("target_document_sha256") == target_hash, "Completion review belongs to another saved target")
        check(type(quality.get("unresolved_count")) is int and quality.get("unresolved_count") == 0 and quality.get("unresolved_issues") == [],
              "Completion review contains unresolved issues or omits the explicit empty issue list")
        checks = quality.get("checks", {})
        names = ["metadata", "native_refresh", "reopen_link", "visual_review"]
        if quality.get("claim_support_requested") or "claim_support" in checks:
            names.append("claim_support")
        for name in names:
            entry = checks.get(name, {}) if isinstance(checks, dict) else {}
            check(isinstance(entry, dict) and entry.get("passed") is True and _nonempty(entry.get("evidence")),
                  "Completion review lacks a passing, evidence-backed " + name + " check")

        def block_index(report, label):
            blocks = report.get("blocks")
            if not isinstance(blocks, list):
                raise ValueError(label + " inspection has no blocks")
            by_key, order = {}, []
            for block in blocks:
                key = _key(block)
                if key in by_key:
                    raise ValueError(label + " inspection contains duplicate block positions")
                if not isinstance(block.get("text"), str) or not isinstance(block.get("comparison_text"), str) or len(block["text"]) != len(block["comparison_text"]):
                    raise ValueError(label + " block requires an equally long comparison_text with only native Zotero results masked")
                if not isinstance(block.get("protected_tokens"), list):
                    raise ValueError(label + " block requires its protected-content anchor tokens")
                by_key[key] = block
                order.append(key)
            return by_key, order

        source_blocks, source_order = block_index(source_report, "Source")
        target_blocks, target_order = block_index(target_report, "Target")
        if not isinstance(source_report.get("citation_cues"), list) or not isinstance(target_report.get("citation_cues"), list):
            raise ValueError("Use fresh inspections containing per-position citation_cues")
        sources = {}
        for cue in source_report.get("citation_cues", []):
            sid = cue.get("cue_id") if isinstance(cue, dict) else None
            if not _nonempty(sid) or sid.startswith("field:") or sid in sources:
                raise ValueError("Source citation cues need distinct cue_id values")
            if _key(cue) not in source_blocks:
                raise ValueError("Source citation cue points to a missing paragraph")
            _mask(source_blocks[_key(cue)]["comparison_text"], [(cue.get("start"), cue.get("end"))])
            sources[sid] = dict(cue, source_type="cue")
        source_fields = {}
        for field in source_report.get("citations", []):
            cid = field.get("citationID")
            if not _nonempty(cid) or cid in source_fields:
                raise ValueError("Source contains duplicate or missing citationID values")
            pos = field.get("context", {})
            if _key(pos) not in source_blocks:
                raise ValueError("Source field points to a missing paragraph")
            source_fields[cid] = field
            sources["field:" + cid] = {**pos, "source_type": "field", "citation": field}
        targets = {}
        for field in target_report.get("citations", []):
            cid = field.get("citationID")
            if not _nonempty(cid) or cid in targets:
                raise ValueError("Target contains duplicate or missing citationID values")
            pos = field.get("context", {})
            if _key(pos) not in target_blocks:
                raise ValueError("Target field points to a missing paragraph")
            _mask(target_blocks[_key(pos)]["comparison_text"], [(pos.get("start"), pos.get("end"))])
            _uris(field)
            targets[cid] = field

        references = set()
        for entry in manifest.get("source_reference_blocks", []):
            key = _key(entry)
            if key not in source_blocks or key in references or not _nonempty(entry.get("reason")) or not _nonempty(entry.get("evidence")):
                raise ValueError("Each excluded source reference block needs a unique exact position, reason and evidence")
            references.add(key)
        entries = manifest.get("entries")
        if not isinstance(entries, list):
            raise ValueError("Coverage entries must list every source cue and activity field")
        by_source, mapped, masks, target_covered = {}, defaultdict(list), defaultdict(list), set()

        def expected_target(entry):
            cid = entry.get("target_citation_id")
            expected = entry.get("expected_uris")
            if cid not in targets or not isinstance(expected, list) or not expected:
                raise ValueError("A mapped/new citation needs an observed target_citation_id and expected_uris")
            # Use the same URI parser for receipts and actual fields.
            wanted = _uris({"citationItems": [{"uris": expected}]})
            if len(wanted) != len(expected) or wanted != _uris(targets[cid]):
                raise ValueError("Mapped expected Zotero identities differ from the actual target field")
            return cid

        for entry in entries:
            if not isinstance(entry, dict):
                raise ValueError("A coverage entry must be an object")
            sid = entry.get("source_id")
            if sid not in sources or sid in by_source:
                raise ValueError("Coverage includes an unknown or duplicate source_id")
            by_source[sid] = entry
            item = sources[sid]
            disposition = entry.get("disposition")
            if not _nonempty(entry.get("evidence")):
                raise ValueError("Every coverage disposition requires evidence")
            if disposition in ("converted", "retained"):
                cid = expected_target(entry)
                if disposition == "retained" and item["source_type"] == "field" and _uris(item["citation"]) != _uris(targets[cid]):
                    raise ValueError("A retained native citation changed its Zotero identities")
                mapped[cid].append(sid)
                target_covered.add(cid)
                if item["source_type"] == "cue":
                    masks[_key(item)].append((item["start"], item["end"]))
            elif disposition == "removed":
                if entry.get("user_confirmed") is not True or not _nonempty(entry.get("decision_source")) or not _nonempty(entry.get("reason")):
                    raise ValueError("Removing a source citation needs a recorded user decision and reason")
                if item["source_type"] == "cue":
                    masks[_key(item)].append((item["start"], item["end"]))
            elif disposition in ("non_citation", "reference_only"):
                if item["source_type"] != "cue" or not _nonempty(entry.get("reason")):
                    raise ValueError("Only a source cue with a reason can be marked non-citation/reference-only")
                if disposition == "reference_only" and _key(item) not in references:
                    raise ValueError("reference_only is outside an explicitly reviewed source reference block")
            else:
                raise ValueError("Unknown source-citation disposition")
        check(set(by_source) == set(sources), "Some source citation positions have no reviewed disposition")
        for sid, item in sources.items():
            if _key(item) in references:
                disposition = by_source.get(sid, {}).get("disposition")
                check(item["source_type"] == "cue" and disposition in ("reference_only", "non_citation"),
                      "A source reference-block exclusion overlaps an actual converted/retained manuscript citation")
        for entry in manifest.get("added_citations", []):
            if not isinstance(entry, dict) or not _nonempty(entry.get("evidence")) or not _nonempty(entry.get("decision_source")):
                raise ValueError("A new target citation needs evidence and an addition decision source")
            cid = expected_target(entry)
            if cid in target_covered:
                raise ValueError("A target citation is listed redundantly as a new citation")
            target_covered.add(cid)
        check(target_covered == set(targets), "Some target native citations have no source mapping or recorded addition")

        # Explicit multi-location consolidation must name exactly the source
        # positions being collapsed, rather than merely saying 'same paper'.
        consolidations = {}
        for item in manifest.get("consolidations", []):
            cid, ids = item.get("target_citation_id"), item.get("source_ids")
            if cid not in mapped or cid in consolidations or not isinstance(ids, list) or len(ids) != len(set(ids)) or set(ids) != set(mapped[cid]):
                raise ValueError("Consolidation must name exactly one mapped target and all its source positions")
            if item.get("user_confirmed") is not True or not _nonempty(item.get("decision_source")) or not _nonempty(item.get("evidence")):
                raise ValueError("Consolidating separate citation positions requires a user decision")
            consolidations[cid] = item
        for cid, ids in mapped.items():
            if len(ids) < 2 or cid in consolidations:
                continue
            items = [sources[sid] for sid in ids]
            if len({_key(item) for item in items}) != 1:
                check(False, "Separate source paragraphs were mapped to a single target citation without confirmation")
                continue
            items.sort(key=lambda item: item.get("start", -1))
            key = _key(items[0]); text = source_blocks[key]["comparison_text"]
            spans = [(item.get("start"), item.get("end")) for item in items]
            for start, end in spans:
                _mask(text, [(start, end)])
            end = spans[0][1]
            adjacent = True
            for start, next_end in spans[1:]:
                if start > end and not _SEPARATORS.fullmatch(text[end:start]):
                    adjacent = False
                end = max(end, next_end)
            if not adjacent:
                check(False, "Separate source citation positions were collapsed into one target field without confirmation")
            else:
                masks[key].append((spans[0][0], end))

        # Parentheses enclosing a combined author-year citation belong to the
        # citation only when every cue inside was resolved, and masking those
        # cues leaves punctuation/whitespace, never ordinary manuscript prose.
        containers, source_anchors = set(), {}
        for sid, item in sources.items():
            if item["source_type"] != "cue" or "container_start" not in item or "container_end" not in item:
                continue
            key = _key(item); start, end = item["container_start"], item["container_end"]
            text = source_blocks[key]["comparison_text"]
            _mask(text, [(start, end)])
            if not start <= item["start"] < item["end"] <= end:
                raise ValueError("Citation container does not enclose its cue")
            containers.add((key, start, end))
        for key, start, end in sorted(containers):
            text = source_blocks[key]["comparison_text"]
            if (text[start], text[end - 1]) not in (("(", ")"), ("（", "）")):
                raise ValueError("Author-year container must name its exact outer parentheses")
            enclosed, partial = [], False
            for sid, item in sources.items():
                if item["source_type"] != "cue" or _key(item) != key:
                    continue
                if item["start"] < end and start < item["end"]:
                    if not start <= item["start"] < item["end"] <= end:
                        partial = True
                    enclosed.append(sid)
            if partial or not enclosed or any(by_source.get(sid, {}).get("disposition") not in ("converted", "removed") for sid in enclosed):
                continue
            target_ids = {by_source[sid]["target_citation_id"] for sid in enclosed if by_source[sid]["disposition"] == "converted"}
            if len(target_ids) > 1:
                continue
            residual = _mask(text, [(sources[sid]["start"], sources[sid]["end"]) for sid in enclosed])[start:end]
            if _SEPARATORS.fullmatch(residual):
                masks[key].append((start, end))
                source_anchors.update({sid: start for sid in enclosed})

        # Body rewrites are narrow, exact-state receipts. They cannot approve a
        # changing whole manuscript with a single free-text 'reviewed' string.
        changed_source, changed_target = {}, {}
        for index, change in enumerate(manifest.get("body_changes", [])):
            if not isinstance(change, dict) or change.get("meaning_preserved") is not True or not _nonempty(change.get("reason")) or not _nonempty(change.get("evidence")):
                raise ValueError("A body change needs meaning_preserved, reason and evidence")
            for name, blocks, order, used in (("source_blocks", source_blocks, source_order, changed_source), ("target_blocks", target_blocks, target_order, changed_target)):
                values = change.get(name)
                if not isinstance(values, list) or not values:
                    raise ValueError("A body change must bind its exact source and target blocks")
                keys = []
                for item in values:
                    key = _key(item)
                    if key not in blocks or key in used or item.get("text_sha256") != text_sha256(blocks[key]["text"]):
                        raise ValueError("Body-change receipt is stale, points to a missing block, or reuses a block")
                    if key in references and name == "source_blocks":
                        raise ValueError("A source reference exclusion cannot also act as an ordinary-body rewrite")
                    used[key] = index
                    keys.append(key)
                positions = [order.index(key) for key in keys]
                if positions != list(range(positions[0], positions[0] + len(positions))):
                    raise ValueError("A body-change group must name contiguous blocks in stored order")
                if len({key[0] for key in keys}) != 1:
                    raise ValueError("A body-change group cannot merge different Word stories")

        source_plain = {key: _mask(block["comparison_text"], masks[key]) for key, block in source_blocks.items()}
        mapped_source_blocks = {_key(sources[sid]) for ids in mapped.values() for sid in ids}
        mapped_target_blocks = {_key(targets[cid]["context"]) for cid in mapped}

        def body_sequence(blocks, order, changed, source=False):
            result, emitted = [], set()
            for key in order:
                if source and key in references:
                    continue
                if key in changed:
                    index = changed[key]
                    if index not in emitted:
                        result.append((key, ("reviewed_body_change", key[0], index)))
                        emitted.add(index)
                    continue
                normal = _ordinary_text(blocks[key], masks[key] if source else ())
                tokens = protected_tokens(blocks[key], source)
                citation_block = key in (mapped_source_blocks if source else mapped_target_blocks)
                if normal or tokens or citation_block:
                    # Paragraph numbers are document-local: collapsing a static
                    # References section must not change the identity of later
                    # unchanged appendix blocks. Keep story, position type,
                    # ordered text and protected anchors as alignment evidence.
                    result.append((key, ("ordinary_block", key[0], blocks[key].get("location", ""),
                                         normal, tokens, bool(citation_block) if not normal and not tokens else False)))
            return result

        source_sequence = body_sequence(source_blocks, source_order, changed_source, True)
        target_sequence = body_sequence(target_blocks, target_order, changed_target)
        check([value for _, value in source_sequence] == [value for _, value in target_sequence],
              "Ordinary manuscript text or block placement changed outside explicitly reviewed citation/body ranges")
        block_alignment = {source_key: target_key for (source_key, left), (target_key, right) in zip(source_sequence, target_sequence) if left == right}
        alignment_receipt = [{"source": {"story": a[0], "paragraph": a[1]},
                              "target": {"story": b[0], "paragraph": b[1]}}
                             for a, b in block_alignment.items()]
        # An identifier match alone does not establish that a field was inserted
        # at the right sentence. Check the aligned block and ordinary prefix.
        for cid, ids in mapped.items():
            target_pos = targets[cid]["context"]; target_key = _key(target_pos)
            for sid in ids:
                item = sources[sid]; source_key = _key(item)
                if cid in consolidations:
                    continue
                if source_key in changed_source and target_key in changed_target and changed_source[source_key] == changed_target[target_key]:
                    continue
                if not check(block_alignment.get(source_key) == target_key, "A source citation moved to an unaligned paragraph without an exact reviewed body change"):
                    continue
                start = source_anchors.get(sid, item.get("start"))
                if type(start) is not int or not 0 <= start <= len(source_plain[source_key]):
                    check(False, "A source native-field position is missing or invalid")
                    continue
                before = _ordinary_text(source_blocks[source_key], masks[source_key], start)
                after = _ordinary_text(target_blocks[target_key], limit=target_pos["start"])
                check(before == after, "A target citation is at a different ordinary-text position from its source")
        protected_source, protected_target = source_report.get("protected_content_sha256"), target_report.get("protected_content_sha256")
        required_protected = {"media", "equations", "other_fields", "drawings"}
        check(isinstance(protected_source, dict) and isinstance(protected_target, dict) and required_protected <= set(protected_source) and required_protected <= set(protected_target),
              "Inspections lack media, drawing, equation or non-Zotero-field protection fingerprints")
        for value in (protected_source, protected_target):
            check(isinstance(value, dict) and all(isinstance(value.get(key), str) and re.fullmatch(r"[a-f0-9]{64}", value[key]) for key in required_protected),
                  "Protected-content fingerprints must be actual SHA256 strings")
        if equivalent_math:
            check({k: v for k, v in protected_source.items() if k != "equations"}
                  == {k: v for k, v in protected_target.items() if k != "equations"},
                  "Images, drawing anchors, non-Zotero fields or table structure changed")
        else:
            check(protected_source == protected_target, "Images, drawing anchors, equations or non-Zotero fields changed")
        check(source_report.get("media_sha256") == target_report.get("media_sha256"), "Embedded image payloads changed")

        zero_edition = []
        for cid, citation in targets.items():
            for item in citation["citationItems"]:
                embedded = item.get("itemData", {})
                if isinstance(embedded, dict) and str(embedded.get("edition", "")).strip() in ("0", "0.0"):
                    zero_edition.append(cid)
        for bibliography in target_report.get("bibliographies", []):
            display = bibliography.get("display", "")
            if re.search(r"\(\s*0(?:\.0)?\s+ed\.?\s*\)|第\s*0\s*版", display, re.I):
                zero_edition.append("bibliography")
        check(not zero_edition, "Known invalid zero-edition metadata/output remains; repair and recheck before completion")
        counts = {"source_cue_count": sum(item["source_type"] == "cue" for item in sources.values()),
                  "source_native_citation_count": len(source_fields), "source_position_count": len(sources),
                  "covered_source_position_count": len(by_source), "target_citation_count": len(targets),
                  "target_unique_reference_count": len(set().union(*(_uris(field) for field in targets.values()))) if targets else 0,
                  "approved_body_change_count": len(manifest.get("body_changes", []))}
    except (ValueError, TypeError, KeyError, AttributeError) as exc:
        issues.append(str(exc))
    return {"schema": AUDIT, "result": "HOLD" if issues else "PASS",
            "source_document_sha256": source_report.get("document_sha256") if isinstance(source_report, dict) else None,
            "target_document_sha256": target_report.get("document_sha256") if isinstance(target_report, dict) else None,
            "coverage_manifest_sha256": _digest(manifest), "completion_review_sha256": _digest(quality),
            "counts": counts, "issues": list(dict.fromkeys(issues)), "block_alignment": alignment_receipt,
            "scope": "Exact saved-state citation-position coverage, ordinary-text/protected-content changes and recorded completion evidence; not independent native execution or claim verification."}

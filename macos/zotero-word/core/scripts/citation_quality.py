"""Read-only citation boundary QA; never rewrite ordinary text or native fields.

The journal/CSL decision belongs to the evidence-backed profile. This narrow
check deliberately keeps whitespace significant, unlike prose alignment.
"""
from __future__ import annotations

import re

SCHEMA = "zotero-word.citation-quality.v1"
NUMERIC = re.compile(r"[0-9]+(?:\s*[,;\-\u2013\u2014]\s*[0-9]+)*")
HORIZONTAL = " \u00a0\u2009\u202f"


def scientific_end(text, end, raised):
    if end <= 0:
        return False
    char = text[end - 1]
    if char in "⁺⁻" or (raised and char in "+-−"):
        return True
    # Also recognize plain-text ion charges, for example Co3+ and SO4·−.
    return char in "+-−" and bool(re.search(
        r"(?<![A-Za-z])(?:[A-Z][a-z]?[0-9₀-₉⁰¹²³⁴⁵⁶⁷⁸⁹]*)+[·•]?[+\-−]$", text[max(0, end - 48):end]))


def check(report, profile):
    citations = report.get("citations", [])
    rule = profile.get("citation_boundary")
    issues, checked = [], 0
    result = {"schema": SCHEMA, "document_sha256": report.get("document_sha256"),
              "policy": rule.get("value") if isinstance(rule, dict) else None,
              "citation_count": len(citations), "checked_citation_count": 0,
              "issues": issues,
              "scope": "Citation boundary whitespace only; not metadata, claim, native refresh, full citation style or visual acceptance. No edits performed."}
    if not citations:
        result.update(result="PASS")
        return result
    if not isinstance(rule, dict) or set(rule) != {"value", "basis", "evidence"} or rule.get("value") not in ("superscript_no_gap", "style_managed") or rule.get("basis") not in ("required", "workflow_default") or not isinstance(rule.get("evidence"), str) or not rule["evidence"].strip():
        issues.append({"kind": "missing_or_invalid_citation_boundary_policy",
                       "action": "Verify the selected journal/CSL and declare an evidence-backed citation_boundary rule; do not guess or use style_managed to silence a known defect."})
    elif rule["value"] == "style_managed":
        # Author-date and other styles may require spaces. No blanket stripping.
        result["scope"] += " Profile declares style-managed spacing; this check does not assert that the selected style or its displayed spacing is correct."
    else:
        blocks = {(b["story"], b["paragraph"]): b for b in report.get("blocks", [])}
        for citation in citations:
            context = citation.get("context", {})
            cid = citation.get("citationID")
            location = {"citation_id": cid, **{k: context.get(k) for k in ("story", "paragraph", "start", "end")}}
            display = citation.get("display", "")
            if not isinstance(display, str) or not NUMERIC.fullmatch(display):
                issues.append({**location, "kind": "unsupported_superscript_display",
                               "action": "Review the real rendered citation and any prefix/locator in the native plugin; never strip whitespace inside fields."})
                continue
            block = blocks.get((context.get("story"), context.get("paragraph")))
            start, end = context.get("start"), context.get("end")
            text = block.get("text") if block else None
            if not isinstance(text, str) or type(start) is not int or type(end) is not int or not 0 <= start < end <= len(text) or text[start:end] != display:
                issues.append({**location, "kind": "unbound_citation_boundary",
                               "action": "Reinspect the actual saved document; a cross-paragraph or unbound field needs native review."})
                continue
            checked += 1
            left = start
            while left > 0 and text[left - 1].isspace():
                left -= 1
            preceding = next((s for s in block.get("visible_segments", []) if s["start"] < left <= s["end"]), {})
            if (preceding.get("kind") == "text" and scientific_end(text, left, preceding.get("superscript"))
                    and all(ch in HORIZONTAL for ch in text[left:start])
                    and all(s["kind"] == "text" and not s.get("control") for s in block.get("visible_segments", ())
                            if s["start"] < start and s["end"] > left)):
                if left == start:
                    issues.append({**location, "kind": "scientific_superscript_touches_citation",
                                   "action": "Compare with the source separator and review this location on the rendered page."})
                else:
                    result.setdefault("scientific_charge_separators", []).append(location)
                continue
            # Leading paragraph indentation is not a between-text/citation gap.
            if left == start or left == 0:
                continue
            gap = text[left:start]
            ordinary = all(any(s.get("kind") == "text" and not s.get("control") and s.get("start", -1) <= i < s.get("end", -1)
                               for s in block.get("visible_segments", [])) for i in range(left, start))
            issues.append({**location, "kind": "space_before_superscript_citation",
                           "gap_start": left, "gap_end": start, "gap_text": gap,
                           "context_excerpt": text[max(0, left - 55):min(len(text), end + 25)],
                           "ordinary_text_only": ordinary,
                           "action": "Check the exact boundary and preserve meaning. Repair confirmed extra spaces outside the native field on the working copy, then native Refresh/save/reinspect. Prefixes, tabs/line breaks or field-owned whitespace require separate review; no automatic blanket deletion."})
    result.update(result="HOLD" if issues else "PASS", checked_citation_count=checked)
    return result

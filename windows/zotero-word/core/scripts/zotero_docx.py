#!/usr/bin/env python3
"""Resume citation workflows; batch library and hybrid DOCX writes remain pending native Word validation."""
from __future__ import annotations

import argparse
from io import BytesIO
import importlib.util
from collections import Counter, defaultdict
from datetime import datetime, timezone
import hashlib
import json
import os
from pathlib import Path
import posixpath
import re
import sys
import tempfile
import time
import unicodedata
from urllib.parse import urlparse, unquote
import xml.etree.ElementTree as ET
from zipfile import ZipFile, BadZipFile

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
MC = "{http://schemas.openxmlformats.org/markup-compatibility/2006}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
SNAPSHOT = "zotero-word.snapshot.v1"
PLAN = "zotero-word.membership-plan.v1"
BASELINE = "zotero-word.baseline.v1"
URI_RE = re.compile(r"^/(users/(?:local/[^/]+|[0-9]+)|groups/[0-9]+)/items/([A-Z0-9]{8})/?$")
LIB_RE = re.compile(r"^(users/(?:local/[^/]+|[0-9]+)|groups/[0-9]+)$")


_HELPERS = {}
_PURE_MEMO = None


def _memo():
    global _PURE_MEMO
    if _PURE_MEMO is None:
        _PURE_MEMO = _helper("state_cache").Memo()
    return _PURE_MEMO


def _helper(name):
    """Resolve bundled helpers even when called through the legacy entrypoint."""
    if name in _HELPERS:
        return _HELPERS[name]
    spec = importlib.util.spec_from_file_location("_zotero_word_" + name, Path(__file__).with_name(name + ".py"))
    module = importlib.util.module_from_spec(spec)
    previous = sys.dont_write_bytecode
    try:
        sys.dont_write_bytecode = True
        spec.loader.exec_module(module)
    finally:
        sys.dont_write_bytecode = previous
    if name == "layout_format":
        compute = module._expected
        module._expected = lambda document, profile, role_map: _memo().layout_expected(document, profile, role_map, compute)
    _HELPERS[name] = module
    return module


_cues = _helper("citation_cues")
_audit = _helper("manuscript_audit")

INPUT_PATH_FIELDS = ("document", "saved_copy", "library", "collection", "export", "plan", "verification", "identity_map", "cue_review", "empty_review", "baseline", "previous", "source_document", "coverage_manifest", "completion_review", "format_source", "format_profile", "format_role_map", "format_final_role_map", "style_file", "parent_file", "zotero_archive", "abbreviation_override", "node_executable", "preparation", "batch_plan", "import_requests", "membership_plan", "observations", "offer", "decision", "citation_document", "layout_profile", "role_map", "target_document", "apply_receipt", "config", "artifact", "paths", "run", "previous_baseline", "entries", "preferences_seed", "style_config", "observation", "recovery_evidence", "evidence", "from_document", "host_context")


def digest(value):
    raw = value if isinstance(value, bytes) else json.dumps(value, sort_keys=True, ensure_ascii=False, separators=(",", ":")).encode()
    return hashlib.sha256(raw).hexdigest()


def load(path):
    return json.loads(Path(path).read_text(encoding="utf-8-sig"))


def output(path, value):
    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    name = None
    try:
        with tempfile.NamedTemporaryFile(mode="w", encoding="utf-8", dir=target.parent, prefix=".zotero-word-", suffix=".tmp", delete=False) as f:
            name = f.name
            json.dump(value, f, ensure_ascii=False, indent=2)
            f.write("\n"); f.flush(); os.fsync(f.fileno())
        os.replace(name, target)
    finally:
        if name and Path(name).exists():
            Path(name).unlink()


def protect_output(args):
    target = Path(args.output).resolve()
    require(target.suffix.lower() != ".docx", "A JSON receipt must never overwrite a Word document")
    for key in INPUT_PATH_FIELDS:
        value = getattr(args, key, None)
        if value:
            require(target != Path(value).resolve(), f"Output must not overwrite input {key}: {value}")
    # Follow only explicitly named linked evidence paths. A layout response
    # must never overwrite its A checkpoint, reference profile or bound CSL
    # resources merely because they were reached through another JSON file.
    linked = {"citation_document", "citation_baseline", "reference_profile", "offer_path", "decision_path",
              "source_document", "target_document", "baseline_path", "layout_profile", "role_map",
              "style_file", "parent_file", "zotero_archive", "abbreviation_override", "node_executable", "path", "host_context"}
    queue = [getattr(args, key, None) for key in ("format_profile", "baseline", "offer", "decision", "apply_receipt", "completion_review", "plan")]
    config = getattr(args, "config", None)
    if config:
        state_dir = load(config).get("state_dir")
        if isinstance(state_dir, str) and state_dir:
            require(not target.is_relative_to(Path(state_dir).resolve()), "Output must not overwrite durable execution state")
    run = getattr(args, "run", None)
    if run:
        run_path = Path(run).resolve()
        index = run_path / "run.json" if run_path.is_dir() or run_path.suffix.lower() != ".json" else run_path
        require(target not in (index.resolve(), index.parent / ".writer.lock"), "Output must not overwrite the run index or lock")
        require(not target.is_relative_to((index.parent / "source-snapshots").resolve()),
                "Output must not overwrite immutable source snapshots")
        require(target.parent == index.parent or not target.is_relative_to(index.parent),
                "Keep status output outside internal run subdirectories")
        if args.command in ("run-start", "run-next", "run-status", "run-verify-copy", "run-pages"):
            roles = _helper("run_controller").artifact_paths(index.parent)
            for value in roles.values():
                require(target != Path(value).resolve(), "Output must not overwrite reserved run evidence")
        if index.is_file():
            existing = load(index)
            snapshots = [existing.get("source", {}).get("snapshot")]
            snapshots.extend(record.get("path") for record in existing.get("source_snapshots", [])
                             if isinstance(record, dict))
            for value in snapshots:
                if isinstance(value, str) and value:
                    require(target != Path(value).resolve(), "Output must not overwrite a historical source snapshot")
            for value in (existing.get("source", {}).get("path"), existing.get("previous_baseline")):
                if value:
                    require(target != Path(value).resolve(), "Output must not overwrite the run source or prior baseline")
            for value in existing.get("paths", {}).values():
                if isinstance(value, str):
                    require(target != Path(value).resolve(), "Output must not overwrite run evidence")
                    if existing.get("schema") != _helper("run_controller").SCHEMA and Path(value).suffix.lower() == ".json":
                        queue.append(value)
    seen = set()
    while queue:
        profile_path = queue.pop()
        if not profile_path or str(Path(profile_path).resolve()) in seen:
            continue
        seen.add(str(Path(profile_path).resolve()))
        require(len(seen) <= 64, "Too many linked workflow evidence files")
        try:
            profile = load(profile_path)
        except (OSError, ValueError):
            continue  # Actual validation writes HOLD to safe outputs in the caller.
        def walk(value):
            if isinstance(value, dict):
                for key, item in value.items():
                    if isinstance(item, str) and item and (key in linked or
                            (key == "evidence" and Path(item).is_absolute())):
                        require(target != Path(item).resolve(), "Output must not overwrite bound resource " + key)
                        if Path(item).suffix.lower() == ".json":
                            queue.append(item)
                    elif isinstance(item, (dict, list)):
                        walk(item)
            elif isinstance(value, list):
                for item in value:
                    walk(item)
        walk(profile)


def require(condition, message):
    if not condition:
        raise ValueError(message)


def iso(value):
    require(isinstance(value, str), "Observation timestamp must be an ISO 8601 string")
    dt = datetime.fromisoformat(value.replace("Z", "+00:00"))
    require(dt.tzinfo is not None, "Observation timestamp must include its timezone")
    return dt.astimezone(timezone.utc)


def utcnow():
    return datetime.now(timezone.utc).isoformat()


def canonical_uri(value):
    if not isinstance(value, str):
        return None
    p = urlparse(value)
    if p.scheme not in ("http", "https") or p.netloc.lower() not in ("zotero.org", "www.zotero.org") or p.query or p.fragment:
        return None
    m = URI_RE.fullmatch(p.path)
    return "http://zotero.org/" + m.group(1) + "/items/" + m.group(2) if m else None


def library_of(uri):
    return URI_RE.fullmatch(urlparse(uri).path).group(1)


def doi(value):
    return _cues.normalize_doi(value)


def current_children(node):
    """Current text includes insertions; deleted/moved-from revisions and comments do not."""
    if _helper("revision_history").is_historical(node):
        return
    if node.tag == MC + "AlternateContent":
        selected = node.find(MC + "Choice")
        if selected is None:
            selected = node.find(MC + "Fallback")
        if selected is not None:
            for child in selected:
                yield child
        return
    yield from node


def current_elements(node):
    if _helper("revision_history").is_historical(node):
        return
    yield node
    for child in current_children(node):
        yield from current_elements(child)


def current_text(node):
    parts = []
    def visit(e):
        if e.tag == W + "t":
            parts.append(e.text or "")
        elif e.tag in (W + "tab", W + "br"):
            parts.append(" " if e.tag == W + "tab" else "\n")
        for child in current_children(e):
            visit(child)
    visit(node)
    return "".join(parts)


def _inspect_docx_uncached(path, *, fields_only=False):
    path = Path(path)
    raw = path.read_bytes()
    source_hash = digest(raw)
    fields, simple, issues, blocks, coverage = [], [], [], [], {}
    revisions = Counter()
    reserved_marker_leaks = []
    equations, drawings, table_structures = [], [], []
    with ZipFile(BytesIO(raw) if fields_only else path) as z:
        names = z.namelist()
        require("word/document.xml" in names, "Input must be a readable .docx Word document")
        # Current-text traversal deliberately excludes deleted revisions. A
        # generated locator must nevertheless never survive in revision history
        # or an otherwise unreferenced story in a deliverable.
        if not fields_only:
            for part in names:
                if not (part.startswith("word/") and part.endswith(".xml")):
                    continue
                part_root = ET.fromstring(z.read(part))
                markers = sorted(set(re.findall(r"\bZOTERO_ANCHOR_[A-Za-z0-9_]+\b", "".join(part_root.itertext()))))
                if markers:
                    reserved_marker_leaks.append({"part": part, "markers": markers})
                    issues.append("Generated locator remains in package, including possible deleted revisions: " + part)
        styles = {}
        if not fields_only and "word/styles.xml" in names:
            for s in ET.fromstring(z.read("word/styles.xml")).findall(W + "style"):
                name = s.find(W + "name")
                styles[s.get(W + "styleId")] = name.get(W + "val", "") if name is not None else ""
        prefs = None
        if "docProps/custom.xml" in names:
            props = ET.fromstring(z.read("docProps/custom.xml"))
            chunks = sorted((e for e in props if e.get("name", "").startswith("ZOTERO_PREF")), key=lambda e: int(re.search(r"(\d+)$", e.get("name", "0")).group(1)) if re.search(r"(\d+)$", e.get("name", "")) else -1)
            text = "".join("".join(e.itertext()) for e in chunks)
            if text:
                try:
                    if text.lstrip().startswith("{"):
                        data = json.loads(text)
                        prefs = {"style": {"id": data["style"]["styleID"]}, "data": data}
                    else:
                        p = ET.fromstring(text)
                        style = p.find("style")
                        prefs = {"style": dict(style.attrib) if style is not None else {}, "xml": text}
                except (ET.ParseError, ValueError, KeyError, TypeError):
                    issues.append("Unreadable Zotero document preferences")
        docroot = ET.fromstring(z.read("word/document.xml"))
        stories = {"word/document.xml": docroot}
        rels = {}
        if "word/_rels/document.xml.rels" in names:
            rels = {e.get("Id"): e for e in ET.fromstring(z.read("word/_rels/document.xml.rels"))}
        for e in current_elements(docroot):
            if e.tag not in (W + "headerReference", W + "footerReference"):
                continue
            rel = rels.get(e.get(R + "id"))
            if rel is None or rel.get("TargetMode") == "External":
                issues.append("Cannot resolve an active header/footer reference")
                continue
            target = rel.get("Target", "")
            part = posixpath.normpath(target.lstrip("/") if target.startswith("/") else posixpath.join("word", target))
            if part not in names or not part.startswith("word/"):
                issues.append("Active header/footer part is missing or invalid")
            else:
                stories[part] = ET.fromstring(z.read(part))
        refs = {"footnotes": set(), "endnotes": set()}
        for root in list(stories.values()):
            for e in current_elements(root):
                for kind in refs:
                    if e.tag == W + kind[:-1] + "Reference":
                        refs[kind].add(e.get(W + "id"))
        for kind, selected in refs.items():
            part = "word/" + kind + ".xml"
            if not selected:
                continue
            if part not in names:
                issues.append(f"Referenced {kind} part is missing")
                continue
            root = ET.fromstring(z.read(part)); found = set()
            for child in list(root):
                key = child.get(W + "id")
                if key in selected:
                    if key in found:
                        issues.append(f"Duplicate referenced {kind} ID {key}")
                    found.add(key)
                else:
                    root.remove(child)
            if found != selected:
                issues.append(f"Some current {kind} references have no note body")
            stories[part] = root
        coverage = {"included_parts": list(stories), "referenced_note_ids": {k: sorted(v) for k, v in refs.items()}, "excluded_unreferenced_story_parts": sorted(n for n in names if re.fullmatch(r"word/(?:header\d+|footer\d+|footnotes|endnotes)\.xml", n) and n not in stories), "comments_included": False}
        for story, root in stories.items():
            stack = []
            paragraph_index = 0
            story_rels_path = posixpath.join(posixpath.dirname(story), "_rels", posixpath.basename(story) + ".rels")
            story_rels = {e.get("Id"): e for e in ET.fromstring(z.read(story_rels_path))} if not fields_only and story_rels_path in names else {}

            def public_context(context):
                return {k: context[k] for k in ("story", "paragraph", "location", "paragraph_style", "text")}

            def visible(context, text, kind, run_format=None, *, control=False):
                if context is None:
                    return
                start = context["visible_length"]
                context["visible_parts"].append(text)
                segment = {"start": start, "end": start + len(text), "kind": kind}
                if control:
                    segment["control"] = True
                if run_format and run_format.get("vertAlign") == "superscript":
                    segment["superscript"] = True
                context["visible_segments"].append(segment)
                context["visible_length"] += len(text)
                if kind == "text":
                    context["plain_parts"].append(text)

            def protected_value(node):
                attrs = []
                for key, value in sorted(node.attrib.items()):
                    if key.startswith(R):
                        relationship = story_rels.get(value)
                        if relationship is not None:
                            target = relationship.get("Target", "")
                            part = posixpath.normpath(target.lstrip("/") if target.startswith("/") else posixpath.join(posixpath.dirname(story), target))
                            value = {"type": relationship.get("Type"), "mode": relationship.get("TargetMode"), "payload": digest(z.read(part)) if relationship.get("TargetMode") != "External" and part in names else target}
                    attrs.append([key, value])
                return [node.tag, attrs, node.text or "", [protected_value(c) for c in current_children(node)]]

            def visit(e, context=None, in_table=False, in_simple=False, run_format=None, table_tokens=None):
                nonlocal paragraph_index
                if e.tag in (W + "del", W + "moveFrom"):
                    revisions[e.tag[len(W):]] += 1
                    return
                if e.tag in (W + "ins", W + "moveTo"):
                    revisions[e.tag[len(W):]] += 1
                if not fields_only and e.tag == W + "tbl":
                    topology = []
                    for row in (c for c in current_children(e) if c.tag == W + "tr"):
                        cells = []
                        for cell in (c for c in current_children(row) if c.tag == W + "tc"):
                            properties = cell.find(W + "tcPr")
                            cells.append([(name, dict(value.attrib)) for name in ("gridSpan", "vMerge", "hMerge") if properties is not None and (value := properties.find(W + name)) is not None])
                        topology.append(cells)
                    table_structures.append([story, topology])
                    table_tokens = list(table_tokens or []) + ["table:" + digest(topology)]
                if e.tag == W + "tc":
                    in_table = True
                if not fields_only and e.tag == W + "p":
                    paragraph_index += 1
                    ps = e.find(W + "pPr/" + W + "pStyle")
                    style_name = styles.get(ps.get(W + "val"), "") if ps is not None else ""
                    text = current_text(e)
                    location = "table" if in_table else "figure_caption" if style_name.casefold() in ("caption", "题注", "圖表標題") and re.match(r"\s*(?:图|Figure|Fig\.)", text, re.I) else "table_note" if re.match(r"\s*(?:表注|Table\s+note|Note[s]?\s*:)", text, re.I) else "body" if story == "word/document.xml" else story.rsplit("/", 1)[-1].replace(".xml", "")
                    context = {"story": story, "paragraph": paragraph_index, "location": location, "paragraph_style": style_name, "text": text}
                    blocks.append(dict(context, plain_parts=[], visible_parts=[], visible_segments=[], visible_length=0, links=[], protected_tokens=list(table_tokens or [])))
                    context = blocks[-1]
                if not fields_only and e.tag == "{http://schemas.openxmlformats.org/officeDocument/2006/math}oMath":
                    value = protected_value(e)
                    equations.append([story, value])
                    if context is not None:
                        context["protected_tokens"].append("equation:" + digest(value))
                if not fields_only and e.tag in (W + "drawing", W + "pict", W + "object"):
                    value = protected_value(e)
                    drawings.append([story, value])
                    if context is not None:
                        context["protected_tokens"].append("drawing:" + digest(value))
                if e.tag == W + "hyperlink" and not stack and context is not None:
                    rel = story_rels.get(e.get(R + "id"))
                    if rel is not None and rel.get("TargetMode") == "External":
                        start = context["visible_length"]
                        label = current_text(e)
                        context["links"].append({"start": start, "end": start + len(label), "target": rel.get("Target", ""), "text": label, "source": "hyperlink"})
                if e.tag == W + "fldSimple":
                    code = e.get(W + "instr", "")
                    simple.append({"story": story, "code": code, "display": current_text(e)})
                    if re.search(r"ZOTERO_", code, re.I):
                        issues.append(f"Unexpected simple Zotero field in {story}")
                    if stack:
                        stack[-1]["nested_children"] += 1
                    in_simple = code
                    if context is not None:
                        match = re.match(r'^\s*HYPERLINK\s+"([^"\r\n]+)"', code, re.I)
                        if match:
                            start = context["visible_length"]
                            label = current_text(e)
                            context["links"].append({"start": start, "end": start + len(label), "target": match.group(1), "text": label, "source": "field"})
                if e.tag == W + "fldChar":
                    kind = e.get(W + "fldCharType")
                    if kind == "begin":
                        stack.append({"story": story, "depth": len(stack), "nested_children": 0, "context": dict(public_context(context), start=context["visible_length"]) if context else {}, "code_parts": [], "text_parts": [], "runs": [], "phase": "code", "separated": False})
                    elif kind == "separate":
                        if not stack or stack[-1]["separated"]:
                            issues.append(f"Orphan/duplicate field separator in {story}")
                        else:
                            stack[-1]["phase"] = "result"
                            stack[-1]["separated"] = True
                    elif kind == "end":
                        if not stack:
                            issues.append(f"Orphan field end in {story}")
                        else:
                            f = stack.pop()
                            f["code"] = "".join(f.pop("code_parts"))
                            f["display"] = "".join(f.pop("text_parts"))
                            if context is not None and f["context"].get("paragraph") == context["paragraph"]:
                                f["context"]["end"] = context["visible_length"]
                            f.pop("phase")
                            if stack:
                                stack[-1]["nested_children"] += 1
                            has_zotero = bool(re.search(r"ZOTERO_", f["code"], re.I))
                            if has_zotero and (f["depth"] or f["nested_children"]):
                                issues.append(f"Nested Zotero field in {story}")
                            if has_zotero and not f["separated"]:
                                issues.append(f"Zotero field has no result separator in {story}")
                            link = re.match(r'^\s*HYPERLINK\s+"([^"\r\n]+)"', f["code"], re.I)
                            if link and not stack and context is not None and "end" in f["context"]:
                                context["links"].append({"start": f["context"]["start"], "end": f["context"]["end"], "target": link.group(1), "text": f["display"], "source": "field"})
                            fields.append(f)
                elif e.tag == W + "instrText":
                    if stack and stack[-1]["phase"] == "code":
                        stack[-1]["code_parts"].append(e.text or "")
                    elif (e.text or "").strip():
                        issues.append(f"Orphan field instruction in {story}")
                elif e.tag == W + "r":
                    pr = e.find(W + "rPr")
                    fmt = {}
                    if pr is not None:
                        for name in ("vertAlign", "b", "i", "rStyle"):
                            x = pr.find(W + name)
                            if x is not None:
                                fmt[name] = x.get(W + "val", "true")
                    run_format = fmt
                elif e.tag == W + "t":
                    native = any(re.search(r"\bZOTERO_(?:ITEM|BIBL)\b", "".join(f["code_parts"]), re.I) for f in stack)
                    native = native or bool(in_simple and re.search(r"\bZOTERO_(?:ITEM|BIBL)\b", str(in_simple), re.I))
                    if not stack or stack[-1]["phase"] == "result":
                        visible(context, e.text or "", "zotero" if native else "text", run_format)
                    if stack and stack[-1]["phase"] == "result":
                        stack[-1]["text_parts"].append(e.text or "")
                        stack[-1]["runs"].append({"text": e.text or "", **(run_format or {})})
                elif e.tag in (W + "tab", W + "br"):
                    native = any(re.search(r"\bZOTERO_(?:ITEM|BIBL)\b", "".join(f["code_parts"]), re.I) for f in stack)
                    if not stack or stack[-1]["phase"] == "result":
                        visible(context, " " if e.tag == W + "tab" else "\n", "zotero" if native else "text", control=True)
                for child in current_children(e):
                    visit(child, context, in_table, in_simple, run_format, table_tokens)
            units = list(root) if story in ("word/footnotes.xml", "word/endnotes.xml") else [root]
            for unit in units:
                # Notes are separate Word stories: a field may span paragraphs
                # within one note, but never continue into another note.
                stack = []
                visit(unit)
                if stack:
                    note = f" note {unit.get(W + 'id')}" if unit is not root else ""
                    issues.append(f"{len(stack)} unclosed field(s) in {story}{note}; tracked edits may have split the field")
        media = {} if fields_only else {n: digest(z.read(n)) for n in names if n.startswith("word/media/")}
    citations, bibliographies = [], []
    for f in fields:
        code = f["code"]
        has_zotero = bool(re.search(r"ZOTERO_", code, re.I))
        citation_marker = re.match(r"^\s*(?:ADDIN\s+)?ZOTERO_ITEM\s+CSL_CITATION\s*(?=\{)", code, re.I)
        bibliography_marker = re.match(r"^\s*(?:ADDIN\s+)?ZOTERO_BIBL\s*(?=\{)", code, re.I)
        if re.search(r"ZOTERO_TEMP", code, re.I):
            issues.append("Pending ZOTERO_TEMP field")
        if citation_marker:
            try:
                start = citation_marker.end()
                data, tail = json.JSONDecoder().raw_decode(code[start:])
                require(not code[start + tail:].strip(), "Unexpected trailing citation field code")
                require(isinstance(data, dict) and isinstance(data.get("citationID"), str) and data["citationID"], "Missing citationID")
                require(isinstance(data.get("citationItems"), list) and data["citationItems"], "Empty citationItems")
                citations.append({"citationID": data["citationID"], "citationItems": data["citationItems"], "display": f["display"], "runs": f["runs"], "context": f["context"], "field_data": data})
            except (ValueError, KeyError, TypeError) as exc:
                issues.append("Invalid Zotero citation JSON: " + str(exc))
        elif bibliography_marker:
            try:
                start = bibliography_marker.end()
                data, tail = json.JSONDecoder().raw_decode(code[start:])
                require(isinstance(data, dict), "Bibliography metadata is not an object")
                require(code[start + tail:].strip().upper() in ("", "CSL_BIBLIOGRAPHY"), "Unexpected trailing bibliography field code")
                bibliographies.append({"display": f["display"], "runs": f["runs"], "context": f["context"], "field_data": data})
            except (ValueError, KeyError, TypeError) as exc:
                issues.append("Invalid Zotero bibliography metadata: " + str(exc))
        elif has_zotero:
            issues.append("Unrecognized or malformed Zotero field marker; inspect the original field before computing membership")
    ids = [c["citationID"] for c in citations]
    if len(ids) != len(set(ids)):
        issues.append("Duplicate citationID; Refresh and reassess actual fields before reconciliation")
    if fields_only:
        return {"document_sha256": source_hash, "preferences": prefs, "issues": sorted(set(issues)),
                "citations": citations, "bibliographies": bibliographies}
    cues, placeholders, numeric, identifiers, links, all_cues = [], [], [], [], [], []
    for b in blocks:
        plain = "".join(b.pop("plain_parts"))
        b["plain_text"] = plain
        b["text"] = "".join(b.pop("visible_parts"))
        b.pop("visible_length")
        comparison = list(b["text"])
        for segment in b["visible_segments"]:
            if segment["kind"] == "zotero":
                comparison[segment["start"]:segment["end"]] = " " * (segment["end"] - segment["start"])
        b["comparison_text"] = "".join(comparison)
        context = {k: b[k] for k in ("story", "paragraph", "location")}
        for found in _cues.extract_cues(b["comparison_text"], b["links"]):
            cue = {**context, **found}
            cue["cue_id"] = "cue:" + digest({"document_sha256": source_hash, **cue})
            cue["cue_sha256"] = digest(cue)
            all_cues.append(cue)
            {"author_year": cues, "numeric": numeric, "doi": identifiers, "article_link": links}[cue["kind"]].append(cue)
        if re.search(r"(?:\b(?:TODO|TBD|CITE|REFHERE|BIBHERE|ZOTERO_ANCHOR_[A-Za-z0-9_]+)\b|某某|待(?:补|查|引)|参考文献待|\[\s*\?\s*\])", plain, re.I):
            placeholders.append({**context, "text": plain})
    for group in (placeholders,):
        for cue in group:
            cue["cue_sha256"] = digest(cue)
    other_fields = [{"code": f["code"], "display": f["display"]} for f in fields + simple if not re.search(r"ZOTERO_|^\s*HYPERLINK\b", f["code"], re.I)]
    protected = {"media": digest(sorted(media.values())), "equations": digest(equations), "other_fields": digest(sorted(other_fields, key=lambda f: (f["code"], f["display"]))), "drawings": digest(drawings), "table_structure": digest(table_structures)}
    return {"schema": "zotero-word.inspection.v1", "document": str(path.resolve()), "document_sha256": source_hash, "current_text_policy": "Exclude comments, deleted revisions and moveFrom; include ins and moveTo; inspect only referenced note/header/footer stories; preserve coherent current-text fields", "story_coverage": coverage, "revisions_observed": dict(revisions), "reserved_marker_leaks": reserved_marker_leaks, "issues": sorted(set(issues)), "citation_count": len(citations), "bibliography_count": len(bibliographies), "citations": citations, "bibliographies": bibliographies, "simple_fields": simple, "preferences": prefs, "blocks": blocks, "author_year_cues": cues, "unresolved_placeholders": placeholders, "static_numeric_cues": numeric, "doi_cues": identifiers, "link_cues": links, "citation_cues": all_cues, "media_sha256": media, "protected_content_sha256": protected, "ordering_note": "Stored XML story order is diagnostic; Word/Zotero determines final citation order, including footnote/endnote ordering.", "bibliography_validation_scope": "Field presence/integrity only; Completed native all-field update plus visual review must establish final bibliography content and formatting; a successful style switch can supply that update."}


def inspect_docx(path, *, fields_only=False):
    """A saved render needs field results. The full report is for source analysis."""
    if fields_only:
        return _inspect_docx_uncached(path, fields_only=True)
    return _memo().inspection(path, _inspect_docx_uncached)


def observed_report(path, report=None):
    """Reuse parsing only within a call, after checking the actual saved bytes."""
    if report is None:
        return inspect_docx(path)
    require(report.get("document") == str(Path(path).resolve()) and digest(Path(path).read_bytes()) == report.get("document_sha256"), "Saved document changed; discard the in-call inspection and recheck the actual state")
    return report


def formal_doc(report, cue_review=None, empty_review=None):
    require(not report["issues"], "; ".join(report["issues"]))
    if report["citation_count"] == 0:
        require(empty_review, "No native Zotero citations; do not clear a collection without an exact-state empty review and confirmed existing-baseline update")
        review = load(empty_review)
        require(review.get("schema") == "zotero-word.empty-review.v1" and review.get("document_sha256") == report["document_sha256"] and review.get("unresolved_count") == 0, "Empty-state review is unresolved or belongs to another saved state")
        require(review.get("native_refresh_source") and review.get("review_source"), "Empty-state review must record actual native empty/Refresh observation and final content review")
        require(report["bibliography_count"] <= 1 and all(not b["display"].strip() and not any(b.get("field_data", {}).values()) for b in report["bibliographies"]), "Zero-citation state contains a stale/nonempty bibliography; recheck and Refresh in Word")
    else:
        require(report["bibliography_count"] == 1, "Formal reconciliation requires exactly one native Zotero bibliography")
    require(report["preferences"] and report["preferences"].get("style", {}).get("id"), "Missing Zotero style preferences")
    placeholders = report["unresolved_placeholders"]
    if placeholders:
        require(cue_review, "Placeholder diagnostics require an exact-state review before formal processing")
        review = load(cue_review)
        require(review.get("schema") == "zotero-word.cue-review.v1" and review.get("document_sha256") == report["document_sha256"] and review.get("unresolved_count") == 0,
                "Placeholder review is unresolved or belongs to another saved state")
        entries = review.get("reviewed_placeholders", [])
        require(isinstance(entries, list), "reviewed_placeholders must be a per-position list")
        known = {p["cue_sha256"]: p for p in placeholders}
        reviewed = set()
        for item in entries:
            require(isinstance(item, dict) and item.get("cue_sha256") in known and item["cue_sha256"] not in reviewed,
                    "Placeholder reviews must reference distinct observed positions")
            require(item.get("disposition") == "non_citation" and all(isinstance(item.get(k), str) and item[k].strip() for k in ("reason", "evidence")),
                    "A scientific status such as TBD needs an explicit non-citation reason and evidence; unresolved references remain HOLD")
            require(not re.search(r"ZOTERO_ANCHOR_", known[item["cue_sha256"]]["text"], re.I),
                    "A reserved temporary Zotero anchor can never be reviewed away")
            reviewed.add(item["cue_sha256"])
        require(reviewed == set(known), "Some placeholder positions remain unresolved")
        require(review.get("review_source"), "Record the actual placeholder review source")
    diagnostics = report.get("citation_cues", report["author_year_cues"] + report["static_numeric_cues"])
    if diagnostics:
        require(cue_review, "Plain citation/DOI/link cues remain; review them and supply an exact-state cue-review receipt")
        review = load(cue_review)
        require(review.get("schema") == "zotero-word.cue-review.v1" and review.get("document_sha256") == report["document_sha256"] and review.get("unresolved_count") == 0, "Cue review is unresolved or belongs to another saved state")
        require({c["cue_sha256"] for c in diagnostics} <= set(review.get("reviewed_cue_hashes", [])), "Some plain citation diagnostics have not been reviewed")
        require(review.get("review_source"), "Record the matching manifest or review evidence for plain citation diagnostics")


def snapshot(path, kind, max_age):
    return validate_snapshot(load(path), kind, max_age)


def validate_snapshot(d, kind, max_age):
    require(isinstance(d, dict) and d.get("schema") == SNAPSHOT, "Use a bound snapshot envelope, not an unlabelled CSL JSON export")
    require(d.get("kind") == kind, f"Expected {kind} snapshot")
    require(LIB_RE.fullmatch(d.get("library_id", "")), "Invalid Zotero library identity")
    require(d.get("filters_clear") is True, "Snapshot requires observed clear search/tag filters")
    age = (datetime.now(timezone.utc) - iso(d["observed_at"])).total_seconds()
    require(age >= -300 and age <= max_age * 3600, "Snapshot is stale or future-dated; re-export and re-observe actual state")
    require(isinstance(d.get("items"), list), "Snapshot items must be a list")
    require(d.get("items_sha256") == digest(d["items"]), "Snapshot item payload hash mismatch")
    require(re.fullmatch(r"[a-f0-9]{64}", d.get("source_export_sha256", "")), "Missing source export hash")
    if kind == "collection":
        require(d.get("collection_name"), "Missing observed collection name")
        require(isinstance(d.get("collection_path"), str) and d["collection_path"].strip(), "Missing observed collection parent/path")
        if not d["items"]:
            require(d.get("empty_observed") is True, "Empty collection needs native empty-view observation")
    mapping = d.get("identity_map", {})
    require(isinstance(mapping, dict), "identity_map must map export IDs to proven Zotero URIs")
    by_id = {}
    for item in d["items"]:
        require(isinstance(item, dict) and "id" in item, "CSL item lacks an export ID")
        key = str(item["id"])
        uri = canonical_uri(item["id"]) or canonical_uri(mapping.get(key))
        require(uri, "CSL ID is not a Zotero URI; provide a verified identity_map, never infer identity from DOI alone")
        require(library_of(uri) == d["library_id"], "Snapshot contains cross-library items")
        require(uri not in by_id, "Multiple exported items map to the same Zotero URI")
        # Export citation keys and URI http/https variants are not item identity.
        by_id[uri] = dict(item, id=uri)
    return d, by_id


def identities(report, known, library_id):
    counts, order = Counter(), []
    for citation in report["citations"]:
        for item in citation["citationItems"]:
            require(isinstance(item, dict) and isinstance(item.get("uris"), list) and item["uris"], "Citation item lacks Zotero URIs")
            uris = [canonical_uri(u) for u in item["uris"]]
            require(all(uris), "Citation contains an unrecognized item URI")
            require(all(library_of(u) == library_id for u in uris), "Citation belongs to another Zotero library")
            hits = set(uris) & set(known)
            require(len(hits) == 1, "Citation identity missing or ambiguous in the observed library; investigate orphan/merge mapping")
            uri = hits.pop()
            embedded = item.get("itemData", {})
            require(isinstance(embedded, dict), "Citation itemData must be an object")
            a, b = doi(embedded.get("DOI")), doi(known[uri].get("DOI"))
            require(not (a and b and a != b), "Embedded DOI conflicts with its current Zotero item; verify the record before processing")
            if uri not in counts:
                order.append(uri)
            counts[uri] += 1
    return counts, order


def sources(args):
    lib, known = snapshot(args.library, "library", args.max_age_hours)
    col, members = snapshot(args.collection, "collection", args.max_age_hours)
    require(lib["library_id"] == col["library_id"], "Collection and library snapshots refer to different libraries")
    if getattr(args, "collection_name", None):
        require(args.collection_name == col["collection_name"], "Observed collection name differs from requested target")
    if getattr(args, "collection_path", None):
        require(args.collection_path == col["collection_path"], "Observed collection path differs from requested target")
    for uri, item in members.items():
        require(uri in known and item == known[uri], "Collection export differs from the library snapshot; recheck current metadata/identity")
    return lib, known, col, members


def baseline_change(state, report, col, members):
    require(state.get("schema") == BASELINE, "Unsupported baseline schema")
    require(state.get("library_id") == col["library_id"] and state.get("collection_name") == col["collection_name"], "Baseline is for a different library/collection")
    require(state.get("collection_path") == col["collection_path"], "Baseline collection parent/path differs; confirm the intended target instead of rebinding by name")
    if state.get("collection_key"):
        require(state["collection_key"] == col.get("collection_key"), "Baseline collection key differs or is unproven; do not rebind by name")
    return state.get("document_sha256") != report["document_sha256"] or state.get("collection_fingerprint") != digest(members)


def library_fingerprint(lib, known):
    return digest({"library_id": lib["library_id"], "records": known})


def collection_fingerprint(col, members):
    return digest({"library_id": col["library_id"], "collection_name": col["collection_name"], "collection_path": col["collection_path"], "collection_key": col.get("collection_key"), "records": members})


def bind(args):
    protect_output(args)
    raw = Path(args.export).read_bytes()
    value = json.loads(raw)
    require(isinstance(value, list), "Export must be a native CSL JSON array; use [] only for an actually observed empty view")
    require(LIB_RE.fullmatch(args.library_id), "Use a Zotero library identity such as users/local/AbCd1234 or users/12345")
    require(args.filters_clear, "Observe and clear native quick search/tag filters first")
    if args.kind == "collection":
        require(args.collection_name, "A collection snapshot must name the exact observed native collection")
        require(value or args.empty_observed, "Bind an empty collection only after observing the native empty view")
    iso(args.observed_at)
    d = {"schema": SNAPSHOT, "kind": args.kind, "library_id": args.library_id, "observed_at": args.observed_at, "filters_clear": True, "source_export_sha256": digest(raw), "items_sha256": digest(value), "items": value}
    if args.kind == "collection":
        d.update(collection_name=args.collection_name, collection_path=args.collection_path or args.collection_name, empty_observed=args.empty_observed)
        if args.collection_key:
            d["collection_key"] = args.collection_key
    if args.identity_map:
        d["identity_map"] = load(args.identity_map)
    validate_snapshot(d, args.kind, args.max_age_hours)
    output(args.output, d)
    return {"result": "BOUND", "output": str(Path(args.output).resolve()), "items": len(value)}


def reconcile(args):
    protect_output(args)
    report = inspect_docx(args.document)
    formal_doc(report, args.cue_review, args.empty_review)
    lib, known, col, members = sources(args)
    counts, order = identities(report, known, lib["library_id"])
    reasons = []
    previous = load(args.baseline) if args.baseline else None
    if report["citation_count"] == 0:
        require(previous is not None and args.update_decision == "confirmed" and args.decision_source, "Clearing all collection members requires an existing baseline and actual confirmed update decision")
    changed = baseline_change(previous, report, col, members) if previous else True
    if previous is None and not args.initial_authorized:
        reasons.append("Initial target collection/manuscript scope has not been authorized")
    if args.update_decision == "declined":
        reasons.append("User declined collection-baseline update; pause formal processing")
    elif changed and args.update_decision != "confirmed":
        reasons.append("Ask whether the user wants this collection version updated; wait for their decision")
    if changed and args.update_decision == "confirmed" and not args.decision_source:
        reasons.append("Record the actual user instruction/answer authorizing this baseline update")
    row = lambda u: {"uri": u, "title": known[u].get("title", ""), "citation_occurrences": counts.get(u, 0)}
    d = {"schema": PLAN, "status": "HOLD" if reasons else "READY", "reasons": reasons, "created_at": utcnow(), "document_sha256": report["document_sha256"], "library_id": lib["library_id"], "collection_name": col["collection_name"], "collection_path": col["collection_path"], "collection_key": col.get("collection_key"), "library_snapshot_sha256": digest(Path(args.library).read_bytes()), "collection_snapshot_sha256": digest(Path(args.collection).read_bytes()), "library_record_hashes": {u: digest(i) for u, i in known.items()}, "current_members": sorted(members), "desired_members": sorted(counts), "citation_counts": dict(counts), "first_occurrence_xml_order": order, "add": [row(u) for u in order if u not in members], "remove": [row(u) for u in sorted(set(members) - set(counts))], "baseline_changed": changed, "baseline_sha256": digest(Path(args.baseline).read_bytes()) if previous else None, "update_decision": args.update_decision, "decision_source": args.decision_source, "empty_state_review_sha256": digest(Path(args.empty_review).read_bytes()) if report["citation_count"] == 0 else None, "policy": "Only add/remove membership in this observed collection. Never create, delete, trash or merge library records from this plan."}
    output(args.output, d)
    return {"status": d["status"], "add_count": len(d["add"]), "remove_count": len(d["remove"]), "reasons": reasons}


def verify(args, _report=None):
    protect_output(args)
    p = load(args.plan)
    require(p.get("schema") == PLAN and p.get("status") == "READY", "Only a READY membership plan can be verified/executed")
    report = observed_report(args.document, _report)
    formal_doc(report, args.cue_review, args.empty_review)
    if report["citation_count"] == 0:
        require(p.get("baseline_sha256") and p.get("update_decision") == "confirmed" and p.get("empty_state_review_sha256") == digest(Path(args.empty_review).read_bytes()), "Zero-state verification requires the exact reviewed existing-baseline update plan")
    require(report["document_sha256"] == p["document_sha256"], "Document changed after the plan; recheck state and compute a fresh plan")
    lib, known, col, members = sources(args)
    require(iso(col["observed_at"]) >= iso(p["created_at"]) and iso(lib["observed_at"]) >= iso(p["created_at"]), "Verification requires fresh post-plan collection and library observations")
    require(lib["library_id"] == p["library_id"] and col["collection_name"] == p["collection_name"], "After snapshots target a different library/collection")
    require(col["collection_path"] == p["collection_path"], "After snapshot collection parent/path differs")
    if p.get("collection_key"):
        require(col.get("collection_key") == p["collection_key"], "After snapshot collection key differs")
    counts, order = identities(report, known, lib["library_id"])
    require(sorted(counts) == p["desired_members"] and set(members) == set(counts), "Actual collection membership does not match saved manuscript citations")
    require(all(u in known and digest(known[u]) == h for u, h in p["library_record_hashes"].items()), "Original library items changed or disappeared during membership-only operations")
    d = {"schema": "zotero-word.membership-verification.v1", "result": "PASS", "validation_scope": "Collection membership and source-record preservation only. Not manuscript completion; baseline also requires citation coverage, body protection and quality evidence.", "verified_at": utcnow(), "plan_sha256": digest(Path(args.plan).read_bytes()), "document_sha256": report["document_sha256"], "library_id": lib["library_id"], "collection_name": col["collection_name"], "collection_path": col["collection_path"], "collection_key": col.get("collection_key"), "members": sorted(members), "citation_counts": dict(counts), "all_original_library_records_preserved": True, "library_after_count": len(known), "library_after_fingerprint": library_fingerprint(lib, known), "collection_after_fingerprint": collection_fingerprint(col, members), "empty_state_review_sha256": p.get("empty_state_review_sha256"), "add_count": len(p["add"]), "remove_count": len(p["remove"])}
    output(args.output, d)
    return d


def citation_quality(args):
    protect_output(args)
    report = inspect_docx(args.document)
    require(not report["issues"], "Inspect saved document issues before boundary QA: " + "; ".join(report["issues"]))
    profile = _helper("reference_format").validate_profile(args.format_profile)
    result = _helper("citation_quality").check(report, profile)
    require(digest(Path(args.document).read_bytes()) == report["document_sha256"], "Document changed during citation boundary check")
    output(args.output, result)
    return result


def bibliography_quality(args):
    protect_output(args)
    result = _helper("reference_format").verify_bibliography_typography(args.document, args.format_profile)
    output(args.output, result)
    return result


def style_profile(args):
    protect_output(args)
    result = _helper("csl_quality").build_config(args.style_file, args.parent_file, args.zotero_archive,
        args.node_executable, args.locale, args.evidence, args.abbreviation_override)
    output(args.output, result)
    return {"result": "CSL_CONFIG_BOUND", "style_id": result["style_id"], "scope": "Resource configuration only; not manuscript acceptance"}


def citation_style_quality(args):
    protect_output(args)
    profile = _helper("reference_format").validate_profile(args.format_profile)
    result = _helper("csl_quality").verify(args.document, profile)
    output(args.output, result)
    return result


def completion_check(args, report=None):
    for key in ("source_document", "coverage_manifest", "completion_review"):
        require(getattr(args, key, None), "Completion requires --source-document, --coverage-manifest and --completion-review; membership alone is insufficient")
    require(not any(getattr(args, key, None) is not None for key in ("format_source", "format_role_map", "format_final_role_map")),
            "Stage A accepts citations/References only; use --reference-profile, then offer optional layout after this stage completes")
    require(getattr(args, "format_profile", None), "Citation/References completion requires --reference-profile with citation boundary, bibliography and actual CSL rules")
    target = observed_report(args.document, report)
    formal_doc(target, getattr(args, "cue_review", None), getattr(args, "empty_review", None))
    source = observed_report(args.source_document, target) if Path(args.source_document).resolve() == Path(args.document).resolve() else inspect_docx(args.source_document)
    manifest_raw = Path(args.coverage_manifest).read_bytes()
    quality_raw = Path(args.completion_review).read_bytes()
    manifest, quality = json.loads(manifest_raw), json.loads(quality_raw)
    require(isinstance(quality, dict) and isinstance(quality.get("checks"), dict), "Completion review and checks must be objects")
    reader = _helper("reference_format")
    profile = reader.validate_profile(args.format_profile)
    profile_hash = digest(profile)
    math_equivalence = None
    if source.get("protected_content_sha256", {}).get("equations") != target.get("protected_content_sha256", {}).get("equations"):
        math_equivalence = _helper("math_equivalence").verify(args.source_document, args.document,
                                                             source_inspection=source, target_inspection=target)
        require(math_equivalence.get("result") == "PASS", "Math preservation HOLD: " + "; ".join(math_equivalence.get("issues", [])))
    result = _audit.audit(source, target, manifest, quality, protected_equivalence=math_equivalence)
    require(result.get("result") == "PASS", "Citation/body audit HOLD: " + "; ".join(result.get("issues", [])[:8]))
    body_preservation = _helper("body_format").verify(args.source_document, args.document, source, target,
                                                     result.get("block_alignment", []), manifest)
    require(body_preservation.get("result") == "PASS", "Ordinary formatting preservation HOLD: " + "; ".join(body_preservation.get("issues", [])[:8]))
    boundary_result = _helper("citation_quality").check(target, profile)
    require(boundary_result["result"] == "PASS", "Citation boundary HOLD: " + "; ".join(i["kind"] + " at " + str(i.get("citation_id", "profile")) for i in boundary_result["issues"]))
    bibliography_result = {"result": "NOT_APPLICABLE", "reason": "Reviewed zero-citation state has no nonempty native bibliography", "document_sha256": target["document_sha256"]}
    style_result = {"result": "NOT_APPLICABLE", "reason": "Reviewed zero-citation state has no rendered citation/bibliography", "document_sha256": target["document_sha256"]}
    if target["citation_count"]:
        for name in ("bibliography_format", "citation_style"):
            entry = quality["checks"].get(name, {})
            require(isinstance(entry, dict) and entry.get("passed") is True and isinstance(entry.get("evidence"), str) and entry["evidence"].strip(),
                    "Citation/References completion requires an evidence-backed " + name + " check after native update, save and reopen")
        bibliography_result = reader.verify_bibliography_typography(args.document, profile, inspection=target)
        require(bibliography_result.get("result") == "PASS", "References typography HOLD: " + "; ".join(bibliography_result.get("issues", [])))
        require(bibliography_result["document_sha256"] == target["document_sha256"], "Saved bibliography changed after inspection")
        style_result = _helper("csl_quality").verify(args.document, profile, inspection=target)
        require(style_result.get("result") == "PASS", "Target CSL numbering/rich-text HOLD: " + "; ".join(style_result.get("issues", [])))
    require(digest(Path(args.document).read_bytes()) == target["document_sha256"] and digest(Path(args.source_document).read_bytes()) == source["document_sha256"], "Source or target changed during completion; recheck saved state")
    require(digest(load(args.format_profile)) == profile_hash, "Reference profile changed during completion")
    result.update(citation_quality=boundary_result, bibliography_format=bibliography_result, citation_style=style_result,
                  ordinary_format_preservation=body_preservation, math_equivalence=math_equivalence,
                  completion_scope="citations_and_references", references_format_status=bibliography_result["result"],
                  full_manuscript_format_required=False, full_manuscript_format_performed=False,
                  reference_profile_sha256=profile_hash, reference_profile=str(Path(args.format_profile).resolve()),
                  source_document=str(Path(args.source_document).resolve()), target_document=str(Path(args.document).resolve()),
                  coverage_manifest_sha256=digest(manifest_raw), completion_review_sha256=digest(quality_raw))
    return result

def audit_manuscript(args):
    protect_output(args)
    result = completion_check(args)
    output(args.output, result)
    return result


def save_baseline(args, _report=None):
    protect_output(args)
    require(args.consent == "confirmed" and args.decision_source, "Baseline creation/update requires an actual recorded user decision")
    require(args.native_refresh_verified and args.final_review_source, "Record completed native all-field update and final citation/References review before establishing a baseline")
    report = observed_report(args.document, _report)
    completion = completion_check(args, report)
    require(completion.get("result") == "PASS", "Completion audit HOLD: " + "; ".join(completion.get("issues", [])))
    if report["citation_count"] == 0:
        require(args.previous, "An empty manuscript can only update an existing verified baseline")
    lib, known, col, members = sources(args)
    counts, _ = identities(report, known, lib["library_id"])
    require(all(str(known[u].get("edition", "")).strip().casefold() not in ("0", "0.0", "0th") for u in counts), "A cited Zotero record still contains edition 0; correct and verify metadata before completion")
    require(set(counts) == set(members), "Do not establish a baseline until actual collection and saved manuscript agree")
    p = load(args.verification)
    require(p.get("schema") == "zotero-word.membership-verification.v1" and p.get("result") == "PASS" and p.get("document_sha256") == report["document_sha256"] and p.get("members") == sorted(members) and p.get("library_id") == lib["library_id"] and p.get("collection_name") == col["collection_name"] and p.get("collection_path") == col["collection_path"] and p.get("collection_key") == col.get("collection_key"), "A passing verification for this exact saved state is required")
    require(p.get("library_after_fingerprint") == library_fingerprint(lib, known) and p.get("collection_after_fingerprint") == collection_fingerprint(col, members), "Library or collection content changed after verification; recheck state and produce a fresh plan and verification")
    if report["citation_count"] == 0:
        require(p.get("empty_state_review_sha256") == digest(Path(args.empty_review).read_bytes()), "Empty-state receipt changed after verification")
    previous = load(args.previous) if args.previous else None
    if previous:
        baseline_change(previous, report, col, members)
    require(digest(Path(args.document).read_bytes()) == report["document_sha256"] and digest(Path(args.source_document).read_bytes()) == completion["source_document_sha256"], "Source or target changed before baseline creation")
    require(digest(Path(args.coverage_manifest).read_bytes()) == completion["coverage_manifest_sha256"] and digest(Path(args.completion_review).read_bytes()) == completion["completion_review_sha256"], "Completion receipts changed before baseline creation")
    require(digest(load(args.format_profile)) == completion["reference_profile_sha256"], "Reference profile changed before baseline creation")
    d = {"schema": BASELINE, "revision": int(previous.get("revision", 0)) + 1 if previous else 1, "established_at": utcnow(), "library_id": lib["library_id"], "collection_name": col["collection_name"], "collection_path": col["collection_path"], "collection_key": col.get("collection_key"), "document_sha256": report["document_sha256"], "collection_fingerprint": digest(members), "members": sorted(members), "verification_sha256": digest(Path(args.verification).read_bytes()), "decision_source": args.decision_source, "native_refresh_verified": True, "final_review_source": args.final_review_source, "empty_state_review_sha256": p.get("empty_state_review_sha256"), "completion_audit": completion}
    output(args.output, d)
    return {"result": "BASELINE_SAVED", "revision": d["revision"], "output": str(Path(args.output).resolve()),
            "completion_scope": completion["completion_scope"], "references_format_status": completion["references_format_status"],
            "citation_stage_status": "COMPLETE", "optional_layout_state": "NOT_REQUESTED",
            "next_action": "DELIVER_VERIFIED_CITATION_DOCUMENT"}


def prepare_manuscript(args):
    protect_output(args)
    started = time.perf_counter()
    report = inspect_docx(args.document)
    library_raw = Path(args.library).read_bytes() if args.library else None
    lib, known = validate_snapshot(json.loads(library_raw), "library", args.max_age_hours) if library_raw is not None else (None, None)
    prepared = _helper("execution_plan").prepare(report, known)
    prepared.update(inspection=report, library_snapshot_sha256=digest(library_raw) if library_raw is not None else None,
                    library_id=lib["library_id"] if lib else None,
                    full_manuscript_format_required=False, references_format_required=True, completion_scope="citations_and_references",
                    result="HOLD" if report["issues"] else "PREPARED",
                    elapsed_seconds=round(time.perf_counter() - started, 6))
    require(digest(Path(args.document).read_bytes()) == report["document_sha256"], "Document changed during preparation; re-read the current saved input")
    if lib:
        require(digest(Path(args.library).read_bytes()) == prepared["library_snapshot_sha256"], "Library snapshot changed during preparation")
    output(args.output, prepared)
    return {"result": prepared["result"], "output": str(Path(args.output).resolve()), "cue_count": prepared["cue_count"],
            "unique_query_count": prepared["unique_query_count"], "lookup_count": prepared["lookup_count"],
            "workload": prepared.get("workload"), "placeholder_count": prepared["placeholder_count"],
            "full_manuscript_format_required": False, "references_format_required": True, "completion_scope": "citations_and_references",
            "pending_position_count": len(prepared["unresolved_semantic_positions"]), "issues": report["issues"], "elapsed_seconds": prepared["elapsed_seconds"]}


def finalize_manuscript(args):
    """One model/tool round trip; all membership and completion gates still run."""
    protect_output(args)
    final_path, verification_path = Path(args.output).resolve(), Path(args.verification_output).resolve()
    require(final_path != verification_path, "Final baseline and verification outputs must be different files")
    verification_args = argparse.Namespace(**vars(args))
    verification_args.output = str(verification_path)
    protect_output(verification_args)
    started = time.perf_counter()
    # Invalidate an old receipt before doing work; an unsuccessful verify must
    # never leave an earlier PASS available as this attempt's evidence.
    output(verification_path, {"result": "HOLD", "stage": "membership", "reason": "Finalization has not verified this saved state"})
    verified = False
    try:
        report = inspect_docx(args.document)
        parsed = time.perf_counter()
        verify_args = argparse.Namespace(**vars(args))
        delattr(verify_args, "verification_output")
        verify_args.output = str(verification_path)
        verification = verify(verify_args, _report=report)
        verified = True
        verified_at = time.perf_counter()
        baseline_args = argparse.Namespace(**vars(args))
        delattr(baseline_args, "verification_output")
        baseline_args.verification = str(verification_path)
        result = save_baseline(baseline_args, _report=report)
        result.update(verification_output=str(verification_path), member_count=len(verification["members"]),
                      stage_seconds={"target_inspection": round(parsed-started, 6), "membership": round(verified_at-parsed, 6), "completion_and_baseline": round(time.perf_counter()-verified_at, 6)},
                      target_docx_parse_count=1, validation_scope="Membership, citation coverage, body protection and recorded native/quality evidence; does not perform native application actions")
        return result
    except (ValueError, KeyError, TypeError, OSError, ET.ParseError, BadZipFile) as exc:
        if not verified:
            output(verification_path, {"result": "HOLD", "stage": "membership", "error": str(exc)})
        output(final_path, {"result": "HOLD", "stage": "completion" if verified else "membership", "error": str(exc)})
        # A successful membership receipt may survive a completion failure,
        # but the final baseline cannot. The main handler writes its HOLD.
        raise ValueError("Finalization " + ("completion" if verified else "membership") + " HOLD: " + str(exc)) from exc


def duplicates(args):
    protect_output(args)
    _, known = snapshot(args.library, "library", args.max_age_hours)
    doi_groups, title_groups = defaultdict(list), defaultdict(list)
    for uri, i in known.items():
        if doi(i.get("DOI")):
            doi_groups[doi(i["DOI"])].append(uri)
        title = "".join(c for c in unicodedata.normalize("NFKC", str(i.get("title", ""))).casefold() if c.isalnum())
        year = (i.get("issued", {}).get("date-parts") or [[None]])[0][0]
        authors = [a.get("family", a.get("literal", "")).casefold() for a in i.get("author", []) if isinstance(a, dict)]
        if title:
            title_groups[(title, year, tuple(authors[:1]))].append(uri)
    groups = []
    for key, uris in doi_groups.items():
        if len(uris) > 1:
            groups.append({"reason": "same_normalized_DOI", "DOI": key, "uris": uris, "metadata_conflict": len({digest({k: known[u].get(k) for k in ("title", "author", "issued", "volume", "page")}) for u in uris}) > 1})
    for key, uris in title_groups.items():
        if len(uris) > 1 and len({doi(known[u].get("DOI")) for u in uris}) > 1:
            groups.append({"reason": "same_normalized_title_year_first_author", "uris": uris, "DOIs": sorted({doi(known[u].get("DOI")) for u in uris}), "metadata_conflict": True})
    d = {"schema": "zotero-word.duplicate-candidates.v1", "result": "REVIEW_REQUIRED" if groups else "NO_EXACT_DUPLICATE_CANDIDATES", "groups": groups, "scope": "Deterministic DOI and exact normalized title/year/first-author candidates only; this is not exhaustive fuzzy duplicate detection.", "action": "Do not merge automatically. Consolidate candidates and ask the user which merges to perform in native Zotero."}
    output(args.output, d)
    return {"result": d["result"], "groups": len(groups)}


def candidates(args):
    protect_output(args)
    lib, known = snapshot(args.library, "library", args.max_age_hours)
    query_doi = getattr(args, "doi", None)
    query_url = getattr(args, "url", None)
    identifier = doi(query_doi or query_url) if query_doi or query_url else ""
    if query_doi:
        require(identifier, "The DOI is not a syntactically recognizable DOI; do not guess a replacement")
    author = unicodedata.normalize("NFKC", args.author or "").casefold().strip()
    require(bool(identifier or query_url) != bool(author or args.year), "Use either DOI/URL or both --author and --year")
    require(identifier or query_url or (author and args.year), "Provide DOI, article URL, or author and year")
    if query_url:
        parsed = urlparse(query_url)
        require(parsed.scheme in ("http", "https") and parsed.netloc, "Article URL must be an absolute HTTP(S) URL")
    hits = []
    for uri, item in known.items():
        names = [unicodedata.normalize("NFKC", a.get("family", a.get("literal", ""))).casefold().strip() for a in item.get("author", []) if isinstance(a, dict)]
        years = item.get("issued", {}).get("date-parts") or [[]]
        year = str(years[0][0]) if years and years[0] else ""
        matches = doi(item.get("DOI")) == identifier if identifier else str(item.get("URL", "")).strip().rstrip("/") == query_url.strip().rstrip("/") if query_url else year == str(args.year) and names and names[0] == author
        if matches:
            hits.append({"uri": uri, "DOI": item.get("DOI"), "title": item.get("title"), "author": item.get("author", []), "issued": item.get("issued"), "container-title": item.get("container-title"), "abstract": item.get("abstract", ""), "source_record_sha256": digest(item)})
    d = {"schema": "zotero-word.identifier-candidates.v1" if identifier or query_url else "zotero-word.author-year-candidates.v1", "library_id": lib["library_id"], "author": args.author, "year": args.year, "DOI": identifier or None, "URL": query_url, "context": args.context, "candidates": hits, "status": "NEEDS_SEMANTIC_VERIFICATION" if hits else "NO_LOCAL_CANDIDATES", "scope": "Local exact DOI, article URL, or first-author/year candidates only. DOI syntax and a single result do not establish public resolution, metadata accuracy, claim support or unique Zotero identity. Multiple matching library records require review, not automatic selection."}
    output(args.output, d)
    return {"status": d["status"], "candidate_count": len(hits)}


def batch_library_plan(args):
    preparation = load(args.preparation) if args.preparation else None
    imports = load(args.import_requests) if args.import_requests else None
    membership = load(args.membership_plan) if args.membership_plan else None
    result = _helper("library_batch").build(preparation, load(args.library),
        current_document_sha256=digest(Path(args.document).read_bytes()), imports=imports,
        membership_plan=membership, collection_snapshot=load(args.collection) if args.collection else None,
        batch_size=args.batch_size, max_age_hours=args.max_age_hours)
    output(args.output, result)
    return result


def batch_library_status(args):
    result = _helper("library_batch").status(load(args.batch_plan), load(args.library),
        current_document_sha256=digest(Path(args.document).read_bytes()),
        collection_snapshot=load(args.collection) if args.collection else None,
        observations=load(args.observations) if args.observations else None,
        max_age_hours=args.max_age_hours)
    output(args.output, result)
    return result


def layout_offer(args):
    result = _helper("workflow_state").offer(args.document, args.baseline, args.format_profile)
    output(args.output, result)
    return result


def layout_decision(args):
    result = _helper("workflow_state").decide(args.document, args.offer, args.choice, args.decision_source)
    output(args.output, result)
    return result


def layout_paragraphs(args):
    consent = _helper("workflow_state").authorize(args.decision, args.document, args.baseline)
    result = _helper("layout_format").inspect_paragraphs(args.document)
    observed = inspect_docx(args.document)
    by_position = {(b["story"], b["paragraph"]): b for b in observed["blocks"]}
    for p in result["paragraphs"]:
        block = by_position[p["story"], p["paragraph"]]
        p.update(location=block["location"], paragraph_style=block["paragraph_style"])
    require(result["source_document_sha256"] == observed["document_sha256"] == consent["citation_document_sha256"],
            "Citation checkpoint changed while preparing optional roles")
    result.update(result="LAYOUT_PARAGRAPHS_PREPARED", paragraph_count=len(result["paragraphs"]),
                  citation_stage_status="COMPLETE", layout_stage_status="PREPARING")
    output(args.output, result)
    return result


def _copy_content(source_path, target_path, *, clean=False):
    """Compare the active text, fields and protected objects of two actual files."""
    source, target = inspect_docx(source_path), inspect_docx(target_path)
    require(not source["issues"] and not target["issues"], "Layout source/target has structural field issues")
    left = {c["citationID"]: c for c in source["citations"]}
    right = {c["citationID"]: c for c in target["citations"]}
    require(len(left) == source["citation_count"] and len(right) == target["citation_count"] and set(left) == set(right),
            "Layout changed citation IDs or field coverage")
    for cid in left:
        require(left[cid]["citationItems"] == right[cid]["citationItems"], "Layout changed Zotero identity or citation metadata at " + cid)
    def stable_preferences(preferences):
        require(isinstance(preferences, dict), "Saved copy lacks native Zotero preferences")
        if not preferences.get("xml"):
            data = preferences.get("data")
            if isinstance(data, dict):
                return {**preferences, "data": {k: v for k, v in data.items() if k != "sessionID"}}, None
            return preferences, None
        xml = ET.fromstring(preferences.get("xml", ""))
        sessions = xml.findall("session")
        require(len(sessions) <= 1, "Ambiguous Zotero session data")
        for session in sessions:
            require(set(session.attrib) == {"id"} and session.get("id") and not len(session)
                    and not (session.text or "").strip(), "Unsupported Zotero session metadata")
            session.set("id", "CURRENT_NATIVE_SESSION")
        def tree(e):
            return [e.tag, sorted(e.attrib.items()), e.text or "", [tree(c) for c in e]]
        return {k: v for k, v in preferences.items() if k != "xml"}, tree(xml)
    require(stable_preferences(source["preferences"]) == stable_preferences(target["preferences"]),
            "Layout changed Zotero style or document preferences outside the transient native session ID")
    require(source["bibliography_count"] == target["bibliography_count"], "Layout changed the dynamic bibliography count")
    if clean:
        # Accept All can discard media owned solely by deleted revisions. Active drawing hashes include their actual payloads.
        require(set(target["media_sha256"].values()) <= set(source["media_sha256"].values()),
                "Clean copy introduced a new media payload")
    else:
        require(source["media_sha256"] == target["media_sha256"], "Layout changed embedded images")
    a, b = source["protected_content_sha256"], target["protected_content_sha256"]
    separately_checked = {"equations", "media"}
    require({k: v for k, v in a.items() if k not in separately_checked} == {k: v for k, v in b.items() if k not in separately_checked},
            "Layout changed drawings, other fields or table structure")
    math = None
    if a["equations"] != b["equations"]:
        math = _helper("math_equivalence").verify(source_path, target_path,
            source_inspection=source, target_inspection=target)
        require(math.get("result") == "PASS", "Layout changed mathematical content: " + "; ".join(math.get("issues", [])))
    def sequence(report, source_side):
        anchors = defaultdict(list)
        for c in report["citations"]:
            p = c["context"]
            anchors[p["story"], p["paragraph"]].append(c["citationID"])
        result = []
        for block in report["blocks"]:
            text = _audit._ordinary_text(block)
            tokens = list(block["protected_tokens"])
            if math:
                mapping = math["source_equation_token_map" if source_side else "target_equation_token_map"]
                tokens = ["equation:" + mapping[t.removeprefix("equation:")] if t.startswith("equation:") else t for t in tokens]
            cids = anchors[block["story"], block["paragraph"]]
            native = any(s.get("kind") == "zotero" for s in block["visible_segments"])
            # The native bibliography owns its generated result paragraphs;
            # ordinary blank paragraphs and citation-only paragraphs stay bound.
            if native and not text and not tokens and not cids:
                continue
            result.append((block["story"], text, tokens, cids))
        return result
    require(sequence(source, True) == sequence(target, False), "Layout changed scientific text, controls, protected anchors or citation positions")
    return source, target, math


def _layout_protection(source_path, target_path, reference_profile):
    source, target, math = _copy_content(source_path, target_path)
    profile = _helper("reference_format").validate_profile(reference_profile)
    refs = {"result": "NOT_APPLICABLE", "reason": "Completed zero-citation checkpoint"}
    csl = {"result": "NOT_APPLICABLE", "reason": "Completed zero-citation checkpoint"}
    if target["citation_count"]:
        refs = _helper("reference_format").verify_bibliography_typography(target_path, profile, inspection=target)
        require(refs["result"] == "PASS", "Layout damaged References typography: " + "; ".join(refs.get("issues", [])))
        csl = _helper("csl_quality").verify(target_path, profile, inspection=target)
        require(csl["result"] == "PASS", "Layout damaged target CSL: " + "; ".join(csl.get("issues", [])))
    boundary = _helper("citation_quality").check(target, profile)
    require(boundary["result"] == "PASS", "Layout damaged citation boundaries")
    require(digest(Path(source_path).read_bytes()) == source["document_sha256"]
            and digest(Path(target_path).read_bytes()) == target["document_sha256"], "Layout evidence changed during inspection")
    return {"result": "PASS", "source_document_sha256": source["document_sha256"],
            "target_document_sha256": target["document_sha256"], "citation_count": target["citation_count"],
            "bibliography_count": target["bibliography_count"], "bibliography_format": refs,
            "citation_style": csl, "citation_boundary": boundary, "math_equivalence": math,
            "citation_engine_inputs_unchanged": True,
            "native_session_id_changed": source["preferences"] != target["preferences"],
            "source_preferences_sha256": digest(source["preferences"]),
            "target_preferences_sha256": digest(target["preferences"]),
            "scope": "Saved-file content, identity, CSL and References protection only; not native application or visual evidence"}


def apply_optional_layout(args):
    state = _helper("workflow_state")
    consent = state.authorize(args.decision, args.document, args.baseline)
    require(Path(args.target_document).resolve() != Path(args.output).resolve(), "Layout document and receipt paths must differ")
    result = _helper("layout_format").apply(args.document, args.target_document,
        args.layout_profile, args.role_map, consent)
    protection = _layout_protection(args.document, args.target_document, consent["reference_profile"])
    result.update(result="LAYOUT_APPLIED_PENDING_NATIVE_REVIEW", citation_stage_status="COMPLETE",
        layout_stage_status="PENDING_NATIVE_REVIEW", citation_checkpoint_retained=True,
        citation_checkpoint=consent["citation_document"], citation_baseline=consent["citation_baseline"],
        decision_path=consent["decision_path"], decision_file_sha256=consent["decision_file_sha256"],
        layout_profile=str(Path(args.layout_profile).resolve()), role_map=str(Path(args.role_map).resolve()),
        citation_protection=protection)
    output(args.output, result)
    return result


def audit_optional_layout(args):
    state = _helper("workflow_state")
    consent = state.authorize(args.decision, args.citation_document, args.baseline)
    receipt = load(args.apply_receipt)
    require(receipt.get("result") == "LAYOUT_APPLIED_PENDING_NATIVE_REVIEW"
            and receipt.get("citation_checkpoint") == consent["citation_document"]
            and receipt.get("decision_file_sha256") == consent["decision_file_sha256"],
            "A bound successful layout application receipt is required")
    review_raw = Path(args.completion_review).read_bytes()
    review = json.loads(review_raw)
    require(isinstance(review, dict) and review.get("schema") == "zotero-word.layout-completion-review.v1"
            and review.get("target_document_sha256") == digest(Path(args.document).read_bytes())
            and review.get("unresolved_count") == 0 and review.get("unresolved_issues") == [],
            "Layout completion needs an exact-state review with no unresolved issues")
    checks = review.get("checks", {})
    desktop_run = load(args.baseline).get("schema") == _helper("run_controller").SCHEMA
    for name in (("visual_review",) if desktop_run else ("native_save", "reopen_link", "visual_review")):
        entry = checks.get(name, {}) if isinstance(checks, dict) else {}
        require(isinstance(entry, dict) and entry.get("passed") is True
                and isinstance(entry.get("evidence"), str) and entry["evidence"].strip(),
                "Layout completion lacks actual " + name + " evidence")
    verification = _helper("layout_format").verify(args.citation_document, args.document,
        args.layout_profile, args.role_map, apply_receipt=receipt, allow_native_resave=True)
    require(verification.get("result") == "PASS", "Layout target/range HOLD: " + "; ".join(verification.get("issues", [])))
    protection = _layout_protection(args.citation_document, args.document, consent["reference_profile"])
    update = checks.get("native_update", {})
    native_copy = None
    if desktop_run:
        from types import SimpleNamespace
        guard = _helper("execution_guard")
        directory = Path(args.baseline).resolve().parent
        with guard.exclusive(directory, _helper("file_lock")):
            native_copy = _helper("run_controller").verify_copy(directory, args.document, False,
                SimpleNamespace(**globals()), guard, layout_protection=protection)
        update = {"status": "not_required", "evidence": "Original native render compared on reopen of the saved layout copy"}
    require(isinstance(update, dict) and update.get("status") in ("not_required", "completed")
            and isinstance(update.get("evidence"), str) and update["evidence"].strip(),
            "Record completed necessary native update, or explain why pure layout required none")
    require(update["status"] != "not_required" or protection["citation_engine_inputs_unchanged"],
            "A related citation/bibliography change needs a completed native update")
    state.validate_context(consent)
    require(digest(Path(args.completion_review).read_bytes()) == digest(review_raw), "Layout review changed during acceptance")
    result = {"schema": "zotero-word.layout-completion.v1", "result": "PASS",
        "completion_scope": "optional_full_manuscript_format", "citation_stage_status": "COMPLETE",
        "layout_stage_status": "COMPLETE", "citation_checkpoint_retained": True,
        "citation_checkpoint": consent["citation_document"], "citation_baseline": consent["citation_baseline"],
        "citation_baseline_sha256": consent["citation_baseline_sha256"],
        "target_document": str(Path(args.document).resolve()),
        "target_document_sha256": protection["target_document_sha256"],
        "layout_verification": verification, "citation_protection": protection,
        "completion_review_sha256": digest(review_raw), "native_update": update}
    if native_copy:
        result["native_reopen"] = native_copy
    output(args.output, result)
    return result


def run_control(args):
    from types import SimpleNamespace
    result = _helper("run_controller").run(args, SimpleNamespace(**globals()))
    output(args.output, result)
    return result


def select_windows_runtime():
    """Prefer this interpreter, then look for an already-installed Windows Python with lxml."""
    if os.name != "nt":
        raise ValueError("Use the local Windows desktop for a Word/Zotero operation")
    if sys.version_info >= (3, 10):
        try:
            from lxml import etree
            return
        except ImportError:
            pass
    import subprocess
    candidates = []
    if os.environ.get("VIRTUAL_ENV"):
        candidates.append(Path(os.environ["VIRTUAL_ENV"]) / "Scripts/python.exe")
    candidates.extend(Path(p) / "python.exe" for p in os.get_exec_path() if p)
    if os.environ.get("LOCALAPPDATA"):
        installs = Path(os.environ["LOCALAPPDATA"]) / "Programs/Python"
        candidates.extend(sorted(installs.glob("Python*/python.exe"), reverse=True))
    seen = {Path(sys.executable).resolve()}
    for candidate in candidates:
        if not candidate.is_file() or candidate.parent.name.casefold() == "windowsapps" or candidate.resolve() in seen:
            continue
        seen.add(candidate.resolve())
        try:
            check = subprocess.run([str(candidate), "-c", "from lxml import etree; import sys; sys.exit(sys.version_info < (3, 10))"],
                                   capture_output=True, timeout=4)
        except (OSError, subprocess.TimeoutExpired):
            continue
        if check.returncode == 0:
            raise SystemExit(subprocess.call([str(candidate), "-X", "utf8", *sys.argv]))
    raise ImportError("Use Windows Python 3.10+ with lxml installed in that interpreter")


def library_health(args):
    result = _helper("library_executor").health(load(args.config))
    output(args.output, result)
    return result


def execute_library_batch(args):
    result = _helper("library_executor").execute(load(args.batch_plan), load(args.library), load(args.config),
        current_document_sha256=digest(Path(args.document).read_bytes()),
        collection_snapshot=load(args.collection) if args.collection else None,
        observations=load(args.observations) if args.observations else None,
        batch_index=args.batch_index, max_age_hours=args.max_age_hours)
    output(args.output, result)
    return result


def hybrid_control(args):
    """Deterministic file work only; native acceptance stays a separate gate."""
    writer = _helper("hybrid_citations")
    if args.command == "hybrid-inspect":
        result = writer.inspect(args.document)
    elif args.command == "hybrid-plan":
        entries = load(args.entries)
        if isinstance(entries, dict):
            entries = entries["entries"]
        result = writer.build_plan(args.document, entries, args.library,
            style_config=load(args.style_config) if args.style_config else None,
            preferences_seed=args.preferences_seed)
    elif args.command == "hybrid-apply":
        require(Path(args.target_document).resolve() != Path(args.output).resolve(),
                "Hybrid document and receipt paths must differ")
        result = writer.apply(args.document, load(args.plan), args.library, args.target_document)
        result["write_scope"] = "working_copy_pending_native_validation"
    else:
        result = writer.audit(args.source_document, args.document, load(args.plan),
            allow_native_resave=args.allow_native_resave)
    if args.command == "hybrid-apply":
        result.setdefault("next_action", "RUN_NATIVE_WORD_ZOTERO_VALIDATION")
    output(args.output, result)
    return result


def adapter_select_control(args):
    result = _helper("host_adapter").select(args.host_context)
    output(args.output, result)
    return result


def accept_table_history_control(args):
    result = _helper("revision_history").accept_table_history(
        args.document, args.target_document, authorization_evidence=args.authorization_evidence)
    output(args.output, result)
    return result


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest="command", required=True)
    p = sub.add_parser("adapter-select", help="Select the Windows host from its metadata")
    p.add_argument("--host-context", required=True); p.add_argument("--output", required=True)
    p = sub.add_parser("accept-table-history", help="Accept authorized table property history in a new copy only; not general Accept All")
    p.add_argument("document"); p.add_argument("--target-document", required=True)
    p.add_argument("--authorization-evidence", required=True)
    p.add_argument("--output", required=True)
    p = sub.add_parser("run-verify-copy", help="Verify the saved content and native fields of a separately prepared derivative")
    p.add_argument("document")
    p.add_argument("--run", required=True)
    p.add_argument("--clean", action="store_true", help="Also verify that no tracked revisions or comments remain")
    p.add_argument("--output", required=True)
    p = sub.add_parser("run-pages", help="Render the accepted disk file read-only with Word and report bibliography typography")
    p.add_argument("--run", required=True)
    p.add_argument("--output", required=True)
    for name in ("run-start", "run-next", "run-status"):
        p = sub.add_parser(name, help="Run or resume the Windows Word/Zotero workflow; status only reads saved evidence")
        p.add_argument("--run", required=True)
        p.add_argument("--output", required=True)
        if name == "run-start":
            p.add_argument("document")
            p.add_argument("--journal", required=True)
        if name == "run-next":
            p.add_argument("--saved-copy", help="Resume the same working copy with its recorded native refresh")
        if name != "run-status":
            p.add_argument("--library", help="Optional bound library snapshot; otherwise read Zotero")
            p.add_argument("--reload-library", action="store_true", help="Read Zotero again after resolving missing references")
            p.add_argument("--import-missing", action=argparse.BooleanOptionalAction, default=None,
                           help="Import exact missing references when additions to the local Zotero library are authorized")
            p.add_argument("--entries", help="Reviewed mappings for unresolved source positions")
            p.add_argument("--resolve", action="append", metavar="CUE_ID=ITEM_URI",
                           help="Choose a Zotero item for a reported location; repeat for an ordered group")
            p.add_argument("--style-id", help="Installed Zotero CSL URI")
            p.add_argument("--abbreviations", action=argparse.BooleanOptionalAction, default=None)
    for name in ("hybrid-inspect", "hybrid-plan", "hybrid-apply", "hybrid-audit"):
        p = sub.add_parser(name, help="Batch citation mapping/writing; file results require native Word validation")
        p.add_argument("document"); p.add_argument("--output", required=True)
        if name == "hybrid-plan":
            p.add_argument("--entries", required=True); p.add_argument("--library")
            p.add_argument("--preferences-seed"); p.add_argument("--style-config")
        elif name == "hybrid-apply":
            p.add_argument("--plan", required=True); p.add_argument("--library")
            p.add_argument("--target-document", required=True)
        elif name == "hybrid-audit":
            p.add_argument("--source-document", required=True); p.add_argument("--plan", required=True)
            p.add_argument("--allow-native-resave", action="store_true", help="Allow only validated Word normalization and native bibliography generation")
    p = sub.add_parser("library-health", help="Read-only probe of the configured local Zotero API; never authorize access")
    p.add_argument("--config", required=True); p.add_argument("--output", required=True)
    p = sub.add_parser("batch-execute", help="Execute one verified bounded library batch using existing access")
    p.add_argument("document"); p.add_argument("--config", required=True); p.add_argument("--library", required=True)
    p.add_argument("--collection"); p.add_argument("--batch-plan", required=True); p.add_argument("--observations")
    p.add_argument("--batch-index", type=int); p.add_argument("--max-age-hours", type=float, default=24)
    p.add_argument("--output", required=True)
    p = sub.add_parser("inspect", help="Read actual current-text Word fields and draft cue diagnostics")
    p.add_argument("document"); p.add_argument("--output", required=True)
    p = sub.add_parser("prepare", help="Inspect once and enumerate all local candidate queries in one batch; no automatic choices")
    p.add_argument("document"); p.add_argument("--library"); p.add_argument("--max-age-hours", type=float, default=24); p.add_argument("--output", required=True)
    p = sub.add_parser("citation-quality", help="Read-only journal-sensitive native citation boundary QA; no field edits")
    p.add_argument("document"); p.add_argument("--reference-profile", "--format-profile", dest="format_profile", required=True, help="Citation/References-only profile; --format-profile is a legacy argument-name alias"); p.add_argument("--output", required=True)
    p = sub.add_parser("bibliography-quality", help="Read-only actual native References font, size, spacing and indentation QA; no field edits")
    p.add_argument("document"); p.add_argument("--reference-profile", "--format-profile", dest="format_profile", required=True, help="Citation/References-only profile; --format-profile is a legacy argument-name alias"); p.add_argument("--output", required=True)
    p = sub.add_parser("citation-style-quality", help="Recompute bound target CSL expectations and compare actual numbering and rich text; read-only")
    p.add_argument("document"); p.add_argument("--reference-profile", "--format-profile", dest="format_profile", required=True, help="Citation/References-only profile; --format-profile is a legacy argument-name alias"); p.add_argument("--output", required=True)
    p = sub.add_parser("style-profile", help="Bind actual installed style/parent/processor/locale resources; not manuscript completion")
    p.add_argument("--style-file", required=True); p.add_argument("--parent-file"); p.add_argument("--zotero-archive", required=True)
    p.add_argument("--node-executable", required=True); p.add_argument("--locale", default="en-US"); p.add_argument("--abbreviation-override")
    p.add_argument("--evidence", required=True); p.add_argument("--output", required=True)
    p = sub.add_parser("audit", help="Accept citation coverage, body protection, actual CSL output and dynamic References typography")
    p.add_argument("document"); p.add_argument("--source-document", required=True)
    p.add_argument("--coverage-manifest", required=True); p.add_argument("--completion-review", required=True)
    p.add_argument("--cue-review"); p.add_argument("--empty-review"); p.add_argument("--output", required=True)
    p.add_argument("--reference-profile", "--format-profile", dest="format_profile"); p.add_argument("--format-source", help=argparse.SUPPRESS); p.add_argument("--format-role-map", help=argparse.SUPPRESS); p.add_argument("--format-final-role-map", help=argparse.SUPPRESS)
    p = sub.add_parser("bind-export", help="Bind a native CSL export to the UI state actually observed by the agent")
    p.add_argument("export"); p.add_argument("--kind", choices=("library", "collection"), required=True)
    p.add_argument("--library-id", required=True); p.add_argument("--collection-name"); p.add_argument("--collection-key")
    p.add_argument("--collection-path", help="Observed full collection parent/path; defaults to its name for a root collection")
    p.add_argument("--observed-at", required=True); p.add_argument("--filters-clear", action="store_true")
    p.add_argument("--empty-observed", action="store_true"); p.add_argument("--identity-map")
    p.add_argument("--max-age-hours", type=float, default=24); p.add_argument("--output", required=True)
    for command in ("reconcile", "verify", "baseline", "finalize"):
        p = sub.add_parser(command)
        p.add_argument("document"); p.add_argument("--library", required=True); p.add_argument("--collection", required=True)
        p.add_argument("--collection-name", required=True); p.add_argument("--max-age-hours", type=float, default=24)
        p.add_argument("--collection-path", help="Expected observed full collection parent/path")
        p.add_argument("--cue-review", help="Exact-state reviewed plain author-year/static-number diagnostics, if any")
        p.add_argument("--empty-review", help="Explicit exact-state native empty/Refresh review for a zero-citation existing-baseline update")
        p.add_argument("--output", required=True)
        if command == "reconcile":
            p.add_argument("--baseline"); p.add_argument("--initial-authorized", action="store_true")
            p.add_argument("--update-decision", choices=("confirmed", "declined", "pending"), default="pending")
            p.add_argument("--decision-source", default="")
        elif command == "verify":
            p.add_argument("--plan", required=True)
        else:
            if command == "finalize":
                p.add_argument("--plan", required=True); p.add_argument("--verification-output", required=True)
            else:
                p.add_argument("--verification", required=True)
            p.add_argument("--previous")
            p.add_argument("--consent", choices=("confirmed",), required=True); p.add_argument("--decision-source", required=True)
            p.add_argument("--native-refresh-verified", action="store_true"); p.add_argument("--final-review-source", required=True)
            p.add_argument("--source-document", required=True); p.add_argument("--coverage-manifest", required=True); p.add_argument("--completion-review", required=True)
            p.add_argument("--reference-profile", "--format-profile", dest="format_profile"); p.add_argument("--format-source", help=argparse.SUPPRESS); p.add_argument("--format-role-map", help=argparse.SUPPRESS); p.add_argument("--format-final-role-map", help=argparse.SUPPRESS)
    p = sub.add_parser("duplicates")
    p.add_argument("--library", required=True); p.add_argument("--max-age-hours", type=float, default=24); p.add_argument("--output", required=True)
    p = sub.add_parser("candidates")
    p.add_argument("--library", required=True); p.add_argument("--author"); p.add_argument("--year", type=int)
    group = p.add_mutually_exclusive_group(); group.add_argument("--doi"); group.add_argument("--url")
    p.add_argument("--context", default=""); p.add_argument("--max-age-hours", type=float, default=24); p.add_argument("--output", required=True)
    p = sub.add_parser("batch-plan", help="Build bounded library adapter tasks; does not search/import/mutate applications")
    p.add_argument("document"); p.add_argument("--library", required=True); p.add_argument("--preparation")
    p.add_argument("--import-requests"); p.add_argument("--membership-plan"); p.add_argument("--collection")
    p.add_argument("--batch-size", type=int, default=20); p.add_argument("--max-age-hours", type=float, default=24); p.add_argument("--output", required=True)
    p = sub.add_parser("batch-status", help="Recheck partial/unknown library operations from fresh observations; never blindly retry")
    p.add_argument("document"); p.add_argument("--library", required=True); p.add_argument("--batch-plan", required=True)
    p.add_argument("--collection"); p.add_argument("--observations"); p.add_argument("--max-age-hours", type=float, default=24); p.add_argument("--output", required=True)
    p = sub.add_parser("layout-offer", help="Bind an explicitly requested layout task to a completed citation checkpoint")
    p.add_argument("--reference-profile", dest="format_profile", help="For optional layout from an accepted Windows run.json")
    p.add_argument("document"); p.add_argument("--baseline", required=True); p.add_argument("--output", required=True)
    p = sub.add_parser("layout-decision", help="Record the actual response to a bound post-citation layout question")
    p.add_argument("document"); p.add_argument("--offer", required=True); p.add_argument("--choice", choices=("yes", "no", "pending"), required=True)
    p.add_argument("--decision-source", required=True); p.add_argument("--output", required=True)
    p = sub.add_parser("layout-apply", help="Apply optional manuscript formatting to a new copy only after an explicit yes")
    p.add_argument("document"); p.add_argument("--baseline", required=True); p.add_argument("--decision", required=True)
    p.add_argument("--layout-profile", required=True); p.add_argument("--role-map", required=True)
    p.add_argument("--target-document", required=True); p.add_argument("--output", required=True)
    p = sub.add_parser("layout-paragraphs", help="Prepare complete ordinary role inventory only after an explicit yes")
    p.add_argument("document"); p.add_argument("--baseline", required=True); p.add_argument("--decision", required=True); p.add_argument("--output", required=True)
    p = sub.add_parser("layout-audit", help="Accept the separately saved/reopened optional layout; keep citation checkpoint intact")
    p.add_argument("document"); p.add_argument("--citation-document", required=True); p.add_argument("--baseline", required=True)
    p.add_argument("--decision", required=True); p.add_argument("--layout-profile", required=True); p.add_argument("--role-map", required=True)
    p.add_argument("--apply-receipt", required=True); p.add_argument("--completion-review", required=True); p.add_argument("--output", required=True)
    for command_parser in sub.choices.values():
        command_parser.add_argument("--brief", action="store_true", help="Print a compact receipt; retain complete evidence in --output")
    args = parser.parse_args()
    safe_output = False
    try:
        protect_output(args)
        safe_output = True
        if args.command in ("run-start", "run-next", "run-verify-copy", "layout-apply", "layout-audit", "layout-paragraphs"):
            select_windows_runtime()
        if args.command == "inspect":
            d = inspect_docx(args.document)
            output(args.output, d)
            result = {"result": "HOLD" if d["issues"] else "INSPECTED", "citation_count": d["citation_count"], "bibliography_count": d["bibliography_count"], "issues": d["issues"], "author_year_cue_count": len(d["author_year_cues"]), "doi_cue_count": len(d["doi_cues"]), "link_cue_count": len(d["link_cues"]), "citation_cue_count": len(d["citation_cues"]), "placeholder_count": len(d["unresolved_placeholders"])}
        else:
            result = {"bind-export": bind, "reconcile": reconcile, "verify": verify, "baseline": save_baseline, "duplicates": duplicates, "candidates": candidates, "audit": audit_manuscript, "prepare": prepare_manuscript, "finalize": finalize_manuscript, "citation-quality": citation_quality, "bibliography-quality": bibliography_quality, "style-profile": style_profile, "citation-style-quality": citation_style_quality,
                      "accept-table-history": accept_table_history_control,
                      "adapter-select": adapter_select_control,
                      "batch-plan": batch_library_plan, "batch-status": batch_library_status,
                      "library-health": library_health, "batch-execute": execute_library_batch,
                      **{name: hybrid_control for name in ("hybrid-inspect", "hybrid-plan", "hybrid-apply", "hybrid-audit")},
                      **{name: run_control for name in ("run-start", "run-next", "run-status", "run-verify-copy", "run-pages")},
                      "layout-offer": layout_offer, "layout-decision": layout_decision,
                      "layout-paragraphs": layout_paragraphs,
                      "layout-apply": apply_optional_layout, "layout-audit": audit_optional_layout}[args.command](args)
        if result.get("code") and not result.get("reason_code"):
            result["reason_code"] = result["code"]
            output(args.output, result)
        if result.get("status", result.get("result")) == "HOLD" and not result.get("reason_code"):
            result.update(_helper("execution_errors").explain(ValueError(result.get("reason") or result.get("error", "Required evidence incomplete")), command=args.command))
            output(args.output, result)
        shown = result
        if args.brief:
            shown = {k: v for k, v in result.items() if k in ("result", "status", "output", "revision", "counts", "citation_count", "bibliography_count", "author_year_cue_count", "doi_cue_count", "link_cue_count", "citation_cue_count", "placeholder_count", "items", "cue_count", "unique_query_count", "lookup_count", "pending_position_count", "add_count", "remove_count", "candidate_count", "groups", "target_docx_parse_count", "member_count", "stage_seconds", "elapsed_seconds", "workload", "completion_scope", "references_format_status", "references_format_required", "full_manuscript_format_required", "style_id", "checked_citation_count", "checked_bibliography_entry_count",
                        "citation_stage_status", "optional_layout_state", "layout_stage_status", "next_action", "question", "options", "requires_explicit_yes", "citation_checkpoint_retained", "citation_checkpoint", "plan_fingerprint", "batch_count", "operation_count", "paragraph_count", "summary", "reason_code", "reason", "retry_allowed", "stage", "mode", "run", "run_index", "tool_request", "operation", "operation_id", "observations", "capabilities", "read_verified", "write_authorization_verified") and (not isinstance(v, list) or k == "options")}
            if "members" in result:
                shown["member_count"] = len(result["members"])
            if isinstance(shown.get("workload"), dict):
                shown["workload"] = {k: v for k, v in shown["workload"].items() if k in (
                    "source_cue_position_count", "existing_native_field_count", "unique_query_count", "local_query_candidate_counts", "placeholder_count")}
            shown["evidence_file"] = str(Path(args.output).resolve())
            for key in ("citation_count_inserted", "unique_item_count", "write_scope", "native_validation_verified"):
                if key in result:
                    shown[key] = result[key]
            for key in ("outcome", "path", "initial_sha256", "adapter_id", "agent_host"):
                if key in result:
                    shown[key] = result[key]
            if args.command in ("run-start", "run-next", "run-status", "run-verify-copy", "run-pages"):
                for key in ("document", "document_sha256", "matching", "native_job", "style", "word_window", "native_seconds", "bibliography_entries",
                            "zotero_connection", "library_receipt", "diagnostic", "resume_from", "word_documents_open", "word_control_error",
                            "saved_copy_audited", "native_refresh_observed", "native_update_origin", "sample_verification", "desktop_counts", "image_directory", "page_count", "visual_review"):
                    if key in result:
                        shown[key] = result[key]
                if result.get("native_validation_verified"):
                    observed = result["verified"]
                    shown.update(citation_count=observed["citations"], bibliography_entries=observed["bibliographyEntries"],
                                 sampled_citations=[s["citation"] for s in observed["samples"]])
            for key in ("issues", "reasons"):
                if result.get(key):
                    shown[key] = result[key][:3]
                    shown[key + "_count"] = len(result[key])
        print(json.dumps(shown, ensure_ascii=False))
        return 2 if result.get("status", result.get("result")) == "HOLD" else 0
    except (ValueError, KeyError, TypeError, OSError, ImportError, ET.ParseError, BadZipFile) as exc:
        result = {"result": "HOLD", "error": str(exc), **_helper("execution_errors").explain(exc, command=args.command)}
        if args.command.startswith("layout-"):
            result["layout_stage_status"] = "HOLD"
            try:
                baseline = getattr(args, "baseline", None)
                if not baseline and getattr(args, "offer", None):
                    baseline = load(args.offer).get("citation_baseline")
                retained = _helper("workflow_state").checkpoint(getattr(args, "citation_document", None) or args.document, baseline)
                result.update(citation_stage_status="COMPLETE", citation_checkpoint_retained=True,
                              citation_checkpoint=retained["citation_document"], citation_baseline=retained["citation_baseline"])
            except (ValueError, KeyError, TypeError, OSError):
                pass
        if safe_output:
            try:
                output(args.output, result)
            except OSError:
                pass
        print(json.dumps(result, ensure_ascii=False), file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(main())

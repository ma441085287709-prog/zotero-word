"""Read-only, narrowly bounded equivalence evidence for native Word math saves.

``verify(source, target, source_inspection=..., target_inspection=...)`` binds
the exact DOCX bytes and the inspector's original equation fingerprints.  It
does not replace those fingerprints or approve any non-math difference.

The only normalizations are empty spelling/grammar proof markers inside
oMath and adjacent text-only m:r nodes with exactly equal attributes, m:rPr,
w:rPr and m:t attributes.  Text, structure, order and all other properties
remain literal.  Only a full PASS exposes raw-token -> canonical-token maps.
"""
from __future__ import annotations

import hashlib
import importlib.util
import io
import json
from pathlib import Path
import posixpath
import sys
import xml.etree.ElementTree as ET
from zipfile import BadZipFile, ZipFile

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
MC = "{http://schemas.openxmlformats.org/markup-compatibility/2006}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
SCHEMA = "zotero-word.math-equivalence-verification.v1"
PROOF_TYPES = {"spellStart", "spellEnd", "gramStart", "gramEnd"}


def _require(ok, message):
    if not ok:
        raise ValueError(message)


def _digest(value):
    raw = value if isinstance(value, bytes) else json.dumps(
        value, sort_keys=True, ensure_ascii=False, separators=(",", ":")
    ).encode("utf-8")
    return hashlib.sha256(raw).hexdigest()


def _children(node):
    # Match the existing inspector's current-text and MC selection policies.
    if node.tag in (W + "del", W + "moveFrom") or (
        node.tag.startswith(W) and node.tag.endswith("PrChange")
    ):
        return
    if node.tag == MC + "AlternateContent":
        selected = node.find(MC + "Choice")
        if selected is None:
            selected = node.find(MC + "Fallback")
        if selected is not None:
            yield from selected
        return
    yield from node


def _xml(raw):
    _require(b"<!DOCTYPE" not in raw.upper() and b"<!ENTITY" not in raw.upper(),
             "External XML declarations are unsupported")
    return ET.fromstring(raw)


def _inspect(document):
    # Optional convenience for standalone use; the completion caller should
    # pass its own exact-state inspections to avoid reparsing the manuscript.
    path = Path(__file__).with_name("zotero_docx.py")
    spec = importlib.util.spec_from_file_location("_zotero_math_inspector", path)
    module = importlib.util.module_from_spec(spec)
    previous = sys.dont_write_bytecode
    try:
        sys.dont_write_bytecode = True
        spec.loader.exec_module(module)
    finally:
        sys.dont_write_bytecode = previous
    return module.inspect_docx(document)


def _text_run(value):
    """Return exact properties and sole text node, or refuse run merging."""
    if value[0] != M + "r" or value[2]:
        return None
    children = value[3]
    if not children or children[-1][0] != M + "t":
        return None
    properties, text = children[:-1], children[-1]
    if text[3] or not text[2]:
        return None
    tags = [p[0] for p in properties]
    if any(t not in (M + "rPr", W + "rPr") for t in tags) or len(tags) != len(set(tags)):
        return None
    return properties, text


def _canonical(value, events, path="oMath"):
    """Normalize an inspector-shaped tree without changing the input tree."""
    children = []
    for index, child in enumerate(value[3]):
        child_path = f"{path}/{index + 1}"
        if child[0] == W + "proofErr":
            attrs = dict(child[1])
            _require(not child[2] and not child[3]
                     and set(attrs) == {W + "type"}
                     and attrs[W + "type"] in PROOF_TYPES,
                     "Nonempty or unsupported math proofErr cannot be ignored")
            events.append({"kind": "remove_math_proofErr", "path": child_path,
                           "type": attrs[W + "type"], "raw_sha256": _digest(child)})
            continue
        node = _canonical(child, events, child_path)
        left = _text_run(children[-1]) if children else None
        right = _text_run(node)
        if left and right and children[-1][1:3] == node[1:3] \
                and left[0] == right[0] and left[1][1] == right[1][1]:
            previous = children[-1]
            joined_text = [left[1][0], left[1][1], left[1][2] + right[1][2], []]
            joined = [previous[0], previous[1], previous[2], left[0] + [joined_text]]
            events.append({"kind": "merge_adjacent_equal_property_math_text_runs",
                           "path": child_path, "left_sha256": _digest(previous),
                           "right_sha256": _digest(node), "merged_sha256": _digest(joined),
                           "left_text": left[1][2], "right_text": right[1][2],
                           "merged_text": joined_text[2]})
            children[-1] = joined
        else:
            children.append(node)
    return [value[0], value[1], value[2], children]


def _read(document, inspection):
    path = Path(document).resolve()
    raw = path.read_bytes()
    _require(isinstance(inspection, dict)
             and inspection.get("schema") == "zotero-word.inspection.v1"
             and inspection.get("document") == str(path)
             and inspection.get("document_sha256") == _digest(raw)
             and not inspection.get("issues"),
             "Math comparison requires the current intact saved-document inspection")
    coverage = inspection.get("story_coverage", {})
    parts = coverage.get("included_parts")
    note_ids = coverage.get("referenced_note_ids", {})
    _require(isinstance(parts, list) and parts and parts[0] == "word/document.xml"
             and len(parts) == len(set(parts)), "Missing or invalid inspected story coverage")
    equations = []
    with ZipFile(io.BytesIO(raw)) as archive:
        names = set(archive.namelist())
        for story in parts:
            _require(isinstance(story, str) and story in names,
                     "Inspected equation story is unavailable")
            root = _xml(archive.read(story))
            if story in ("word/footnotes.xml", "word/endnotes.xml"):
                kind = Path(story).stem
                selected = note_ids.get(kind)
                _require(isinstance(selected, list), "Missing inspected note identities")
                for child in list(root):
                    if child.get(W + "id") not in selected:
                        root.remove(child)
            rel_path = posixpath.join(posixpath.dirname(story), "_rels", posixpath.basename(story) + ".rels")
            rels = {e.get("Id"): e for e in _xml(archive.read(rel_path))} if rel_path in names else {}

            def raw_value(node):
                # Keep this representation equal to inspect_docx.protected_value.
                attrs = []
                for key, value in sorted(node.attrib.items()):
                    if key.startswith(R):
                        rel = rels.get(value)
                        if rel is not None:
                            target = rel.get("Target", "")
                            member = posixpath.normpath(target.lstrip("/") if target.startswith("/") else posixpath.join(posixpath.dirname(story), target))
                            value = {"type": rel.get("Type"), "mode": rel.get("TargetMode"),
                                     "payload": _digest(archive.read(member)) if rel.get("TargetMode") != "External" and member in names else target}
                    attrs.append([key, value])
                return [node.tag, attrs, node.text or "", [raw_value(c) for c in _children(node)]]

            paragraph = 0
            story_index = 0

            def visit(node, context=None):
                nonlocal paragraph, story_index
                if node.tag in (W + "del", W + "moveFrom") or (
                    node.tag.startswith(W) and node.tag.endswith("PrChange")
                ):
                    return
                if node.tag == W + "p":
                    paragraph += 1
                    context = paragraph
                if node.tag == M + "oMath":
                    story_index += 1
                    value = raw_value(node)
                    events = []
                    canonical = _canonical(value, events)
                    equations.append({"story": story, "paragraph": context,
                                      "story_equation_index": story_index,
                                      "raw_value": value, "raw_sha256": _digest(value),
                                      "canonical_sha256": _digest(canonical),
                                      "canonical_value": canonical, "normalizations": events})
                for child in _children(node):
                    visit(child, context)

            visit(root)
    fingerprint = _digest([[e["story"], e["raw_value"]] for e in equations])
    _require(fingerprint == inspection.get("protected_content_sha256", {}).get("equations"),
             "Math reader does not reproduce the original inspector equation fingerprint")
    _require(_digest(path.read_bytes()) == _digest(raw), "Saved document changed during math read")
    return raw, fingerprint, equations


def verify(source_document, target_document, *, source_inspection=None, target_inspection=None):
    """Return bound PASS/HOLD evidence; never write files or control applications.

    On PASS, ``source_equation_token_map`` and ``target_equation_token_map`` map
    each original inspector equation token digest to its canonical digest.
    They cover *all* equations.  Callers must retain all raw reports and check
    non-math protected content strictly; this result is not manuscript approval.
    """
    report = {"schema": SCHEMA, "result": "HOLD", "issues": [],
              "source_document": str(Path(source_document).resolve()),
              "target_document": str(Path(target_document).resolve()),
              "source_equation_token_map": {}, "target_equation_token_map": {},
              "equations": [], "changed_equations": [],
              "scope": "Exact saved-state math equivalence only; preserves raw hashes and does not accept non-math differences or establish native execution."}
    try:
        source_inspection = source_inspection if source_inspection is not None else _inspect(source_document)
        target_inspection = target_inspection if target_inspection is not None else _inspect(target_document)
        source_raw, source_fingerprint, source = _read(source_document, source_inspection)
        target_raw, target_fingerprint, target = _read(target_document, target_inspection)
        report.update(source_document_sha256=_digest(source_raw), target_document_sha256=_digest(target_raw),
                      source_raw_equations_sha256=source_fingerprint, target_raw_equations_sha256=target_fingerprint,
                      source_equation_count=len(source), target_equation_count=len(target))
        _require(len(source) == len(target), "Equation count changed")
        source_map, target_map = {}, {}
        for index, (left, right) in enumerate(zip(source, target), 1):
            entry = {"index": index, "story": left["story"], "source_paragraph": left["paragraph"],
                     "target_paragraph": right["paragraph"],
                     "source_raw_sha256": left["raw_sha256"], "target_raw_sha256": right["raw_sha256"],
                     "source_canonical_sha256": left["canonical_sha256"], "target_canonical_sha256": right["canonical_sha256"],
                     "source_normalizations": left["normalizations"], "target_normalizations": right["normalizations"],
                     "raw_equal": left["raw_value"] == right["raw_value"],
                     "canonical_equal": left["canonical_value"] == right["canonical_value"]}
            report["equations"].append(entry)
            if not entry["raw_equal"]:
                report["changed_equations"].append(entry)
            _require((left["story"], left["paragraph"], left["story_equation_index"])
                     == (right["story"], right["paragraph"], right["story_equation_index"]),
                     f"Equation story, paragraph or order changed at equation {index}")
            _require(entry["canonical_equal"],
                     f"Math text, structure or properties changed outside allowed normalization at equation {index}")
            for mapping, value in ((source_map, left), (target_map, right)):
                previous = mapping.setdefault(value["raw_sha256"], value["canonical_sha256"])
                _require(previous == value["canonical_sha256"], "Conflicting equation token mapping")
        _require(_digest(Path(source_document).read_bytes()) == report["source_document_sha256"]
                 and _digest(Path(target_document).read_bytes()) == report["target_document_sha256"],
                 "Source or target changed before math comparison completed")
        report.update(result="PASS", source_equation_token_map=source_map,
                      target_equation_token_map=target_map,
                      changed_equation_count=len(report["changed_equations"]))
    except (ValueError, TypeError, KeyError, AttributeError, OSError, BadZipFile, ET.ParseError) as exc:
        report["issues"].append(str(exc))
    return report

"""Optional, consent-bound DOCX layout after a completed citation checkpoint.

Creates document-local styles only for actual effective differences. Original
styles, native field runs and bibliography paragraphs remain owned by stage A.
This helper neither presents a dialog nor certifies native/visual observations.
"""
from __future__ import annotations

from collections import Counter
import copy
import hashlib
import importlib.util
import io
import json
import math
import os
from pathlib import Path
import re
import sys
import tempfile
import xml.etree.ElementTree as ET
from zipfile import BadZipFile, ZipFile

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
PROFILE = "zotero-word.optional-layout-profile.v1"
ROLE_MAP = "zotero-word.optional-layout-role-map.v1"
ROLES = {"title", "subtitle", "authors", "affiliations", "keywords", "heading1", "heading2", "heading3", "heading4", "abstract", "body", "figure_caption", "table_title", "table_note", "table_body", "equation"}
BOOLS = {"keep_next": "keepNext", "keep_lines": "keepLines", "page_break_before": "pageBreakBefore", "widow_control": "widowControl"}
FONT = {"font_family", "font_size_pt", "bold", "italic"}
CORE_ROLE_RULES = {"font_family", "font_size_pt", "line_spacing", "space_before_pt", "space_after_pt", "left_indent_pt", "right_indent_pt", "first_line_indent_pt", "alignment"}
PARA = {"alignment", "line_spacing", "space_before_pt", "space_after_pt", "left_indent_pt", "right_indent_pt", "first_line_indent_pt", "heading_level"} | set(BOOLS)
PAGE = {"margin_top_pt": ("pgMar", "top"), "margin_bottom_pt": ("pgMar", "bottom"), "margin_left_pt": ("pgMar", "left"), "margin_right_pt": ("pgMar", "right"), "page_width_pt": ("pgSz", "w"), "page_height_pt": ("pgSz", "h")}
_CACHE = {}


def helper(name):
    if name not in _CACHE:
        spec = importlib.util.spec_from_file_location("_optional_layout_" + name, Path(__file__).with_name(name + ".py"))
        module = importlib.util.module_from_spec(spec)
        previous = sys.dont_write_bytecode
        try:
            sys.dont_write_bytecode = True
            spec.loader.exec_module(module)
        finally:
            sys.dont_write_bytecode = previous
        _CACHE[name] = module
    return _CACHE[name]


def require(ok, message):
    if not ok:
        raise ValueError(message)


def digest(value):
    if not isinstance(value, bytes):
        value = json.dumps(value, sort_keys=True, ensure_ascii=False, separators=(",", ":")).encode()
    return hashlib.sha256(value).hexdigest()


def obj(value):
    return value if isinstance(value, dict) else json.loads(Path(value).read_text(encoding="utf-8"))


def validate_profile(value):
    profile = obj(value)
    require(isinstance(profile, dict) and profile.get("schema") == PROFILE, "Unsupported optional layout profile")
    require(set(profile) <= {"schema", "journal", "article_type", "submission_stage", "roles", "page"}, "Layout profile may contain only ordinary role/page rules; stage-A References/CSL stay separate")
    for name in ("journal", "article_type"):
        require(isinstance(profile.get(name), str) and profile[name].strip(), "Layout profile requires " + name)
    require(profile.get("submission_stage") in ("initial", "revision", "final"), "Unknown submission stage")
    require(isinstance(profile.get("roles"), dict) and bool(profile["roles"]) and set(profile["roles"]) <= ROLES, "Explicit ordinary role rules are required")
    require(isinstance(profile.get("page"), dict) and set(profile["page"]) == set(PAGE), "Full layout requires all four margins and both paper dimensions")
    ref = helper("reference_format")
    for role, rules in profile["roles"].items():
        require(isinstance(rules, dict) and set(rules) <= PARA | FONT, "Unknown role typography rule")
        require(CORE_ROLE_RULES <= set(rules), "Each full-layout role requires font, size, line spacing, paragraph spacing, indentation and alignment targets")
        require(not role.startswith("heading") or "heading_level" in rules, "Heading roles require their actual heading level")
        for name, rule in rules.items():
            require(isinstance(rule, dict) and set(rule) == {"value", "basis", "evidence"} and rule.get("basis") in ("required", "workflow_default") and isinstance(rule.get("evidence"), str) and rule["evidence"].strip(), "Every rule needs its actual source or declared working default")
            value = rule["value"]
            if name in ref.BIBLIOGRAPHY_RULES:
                ref._validate_rule(name, rule)
            elif name in BOOLS or name in ("bold", "italic"):
                require(type(value) is bool, "Boolean layout rule required")
            elif name == "alignment":
                require(value in ("left", "center", "right", "both"), "Unsupported alignment")
            elif name == "heading_level":
                require(role in {"heading1", "heading2", "heading3", "heading4"} and type(value) is int and value == int(role[-1]), "Heading role/level mismatch")
    for name, rule in profile.get("page", {}).items():
        require(isinstance(rule, dict) and set(rule) == {"value", "basis", "evidence"} and rule.get("basis") in ("required", "workflow_default") and isinstance(rule.get("evidence"), str) and rule["evidence"].strip(), "Page rule needs source/default evidence")
        value = rule["value"]
        require(type(value) in (int, float) and math.isfinite(value) and (0 <= value <= 360 if name.startswith("margin_") else 144 <= value <= 1440), "Invalid page dimension")
    return profile


def canonical(node):
    if node is None:
        return None
    return [node.tag, sorted(node.attrib.items()), node.text or "", [canonical(c) for c in node]]


def _ensure(parent, name):
    found = parent.find(W + name)
    if found is not None:
        return found
    found = ET.Element(W + name)
    order = {"pPr": "pStyle keepNext keepLines pageBreakBefore framePr widowControl numPr suppressLineNumbers pBdr shd tabs suppressAutoHyphens kinsoku wordWrap overflowPunct topLinePunct autoSpaceDE autoSpaceDN bidi adjustRightInd snapToGrid spacing ind contextualSpacing mirrorIndents suppressOverlap jc textDirection textAlignment textboxTightWrap outlineLvl divId cnfStyle rPr sectPr pPrChange".split(), "rPr": "rStyle rFonts b bCs i iCs caps smallCaps strike dstrike outline shadow emboss imprint noProof snapToGrid vanish webHidden color spacing w kern position sz szCs highlight u effect bdr shd fitText vertAlign rtl cs em lang eastAsianLayout specVanish oMath rPrChange".split(), "sectPr": "headerReference footerReference footnotePr endnotePr type pgSz pgMar paperSrc pgBorders lnNumType pgNumType cols formProt vAlign noEndnote titlePg textDirection bidi rtlGutter docGrid printerSettings sectPrChange".split()}
    names = order.get(parent.tag.removeprefix(W), [])
    position = names.index(name) if name in names else -1
    for index, child in enumerate(parent):
        child_name = child.tag.removeprefix(W)
        if child_name in names and names.index(child_name) > position:
            parent.insert(index, found)
            break
    else:
        if name in ("pPr", "rPr"):
            parent.insert(0, found)
        else:
            parent.append(found)
    return found


def _attrs(parent, name, values):
    node = _ensure(parent, name)
    for key, value in values.items():
        if value is None:
            node.attrib.pop(W + key, None)
        else:
            node.set(W + key, str(value))


def _remove(parent, name, names):
    if parent is None:
        return
    node = parent.find(W + name)
    if node is None:
        return
    for name in names:
        node.attrib.pop(W + name, None)
    if not node.attrib and not len(node) and not node.text:
        parent.remove(node)


def _twip(value):
    return str(round(value * 20))


def _paragraph_rule(name, value):
    if name in BOOLS:
        return BOOLS[name], {"val": "1" if value else "0"}
    if name == "alignment":
        return "jc", {"val": value}
    if name == "heading_level":
        return "outlineLvl", {"val": str(value - 1)}
    if name in ("space_before_pt", "space_after_pt"):
        key = "before" if name == "space_before_pt" else "after"
        return "spacing", {key: _twip(value), key + "Autospacing": "0", key + "Lines": "0"}
    if name == "line_spacing":
        mode = value["mode"]
        return "spacing", {"line": str(round(value["value"] * (240 if mode == "multiple" else 20))), "lineRule": {"multiple": "auto", "exact": "exact", "at_least": "atLeast"}[mode]}
    if name == "first_line_indent_pt":
        return "ind", {"firstLine": _twip(max(value, 0)), "hanging": _twip(max(-value, 0)), "firstLineChars": "0", "hangingChars": "0"}
    key = "left" if name == "left_indent_pt" else "right"
    modern = "start" if key == "left" else "end"
    return "ind", {key: _twip(value), key + "Chars": "0", modern: _twip(value), modern + "Chars": "0"}


def _observed_para(state, name):
    if name in BOOLS:
        return state[BOOLS[name]]
    if name == "alignment":
        return state["jc"]
    if name == "heading_level":
        return state["outlineLvl"] + 1
    if name == "line_spacing":
        a = state["spacing"]
        mode = a["lineRule"]
        return {"mode": {"auto": "multiple", "exact": "exact", "atLeast": "at_least"}.get(mode), "value": a["line"] / (240 if mode == "auto" else 20)}
    if name in ("space_before_pt", "space_after_pt"):
        a = state["spacing"]; key = "before" if name == "space_before_pt" else "after"
        return a[key] / 20 if not a[key + "Autospacing"] and not a[key + "Lines"] else None
    a = state["ind"]
    if name == "first_line_indent_pt":
        return (a["firstLine"] - a["hanging"]) / 20 if not a["firstLineChars"] and not a["hangingChars"] and not(a["firstLine"] and a["hanging"]) else None
    key = "left" if name == "left_indent_pt" else "right"
    return a[key] / 20 if not a[key + "Chars"] else None


def _serialize(root, original):
    declared = {}
    for _, (prefix, uri) in ET.iterparse(io.BytesIO(original), events=("start-ns",)):
        require(prefix not in declared or declared[prefix] == uri, "Scoped namespace rebinding requires native formatting")
        declared[prefix] = uri
    for prefix, uri in declared.items():
        try:
            ET.register_namespace(prefix, uri)
        except ValueError:
            pass
    raw = ET.tostring(root, encoding="utf-8", xml_declaration=True)
    match = re.search(rb"<(?!\?)[^>]+>", raw)
    require(match is not None, "Cannot serialize layout XML")
    head, additions = match.group(0), []
    for prefix, uri in declared.items():
        attr = "xmlns:" + prefix if prefix else "xmlns"
        if not re.search(rb"\s" + re.escape(attr.encode()) + rb"=", head):
            escaped = uri.replace("&", "&amp;").replace('"', "&quot;")
            additions.append(f' {attr}="{escaped}"'.encode())
    return raw[:match.end() - 1] + b"".join(additions) + raw[match.end() - 1:]


def _bib(root):
    ref = helper("reference_format")
    if not re.search(r"\bZOTERO_BIBL\b", "".join(e.text or "" for e in root.iter(W + "instrText"))):
        return set()
    owned = {p for _, p in ref._bibliography_scope(root)[1]}
    paragraphs = ref._paragraphs(root)
    indices = {p: i for i, p in enumerate(paragraphs)}
    stack, spans, visible, anchors = [], [], {}, set()
    graphic = {M + "oMath", M + "oMathPara", W + "drawing", W + "pict", W + "object", W + "tbl"}
    def walk(node, paragraph=None):
        if node.tag in (W + "pPr", W + "rPr"):
            return  # layout tab stops are not visible paragraph content
        if node.tag == W + "p":
            paragraph = node
        if node.tag in graphic and paragraph is not None:
            anchors.add(paragraph)
        if node.tag == W + "fldChar":
            kind = node.get(W + "fldCharType")
            if kind == "begin":
                stack.append({"start": paragraph, "code": [], "phase": "code"})
            elif kind == "separate":
                require(stack, "Unbound bibliography separator")
                stack[-1]["phase"] = "result"
            elif kind == "end":
                require(stack, "Unbound bibliography end")
                field = stack.pop()
                if re.match(r"^\s*(?:ADDIN\s+)?ZOTERO_BIBL\b", "".join(field["code"]), re.I):
                    require(field["start"] in indices and paragraph in indices, "Unbound bibliography paragraph span")
                    spans.append((indices[field["start"]], indices[paragraph], field))
        elif node.tag == W + "instrText":
            require(stack and stack[-1]["phase"] == "code", "Unbound bibliography instruction")
            stack[-1]["code"].append(node.text or "")
        elif node.tag in (W + "t", W + "tab", W + "br", W + "cr"):
            if not stack or stack[-1]["phase"] == "result":
                text = node.text or "" if node.tag == W + "t" else " "
                visible.setdefault(paragraph, []).append((stack[-1] if stack else None, text))
        for child in ref._children(node):
            walk(child, paragraph)
    walk(root)
    require(not stack and len(spans) == 1, "An exact native bibliography span is required")
    first, last, field = spans[0]
    for paragraph in paragraphs[first:last + 1]:
        # Include the field-closing empty paragraph and empty internal result
        # paragraphs. Nonempty prose outside the field is never exempted.
        if paragraph not in anchors and not any(text.strip() and owner is not field for owner, text in visible.get(paragraph, [])):
            owned.add(paragraph)
    last_style = paragraphs[last].find(W + "pPr/" + W + "pStyle")
    last_sid = last_style.get(W + "val") if last_style is not None else None
    tail = paragraphs[last + 1:]
    parents = {c: p for p in root.iter() for c in p}
    def tail_safe(paragraph):
        node = paragraph
        while node in parents:
            node = parents[node]
            if node.tag in (W + "tbl", W + "tc"):
                return False
        return not any(e.tag in (W + "bookmarkStart", W + "bookmarkEnd", W + "footnoteReference", W + "endnoteReference", W + "fldChar", W + "instrText", W + "fldSimple") for e in paragraph.iter())
    # A generated empty list tail must be wholly empty and unanchored, with
    # an explicit Bibliography style connection. Ordinary appendix separators
    # or empty body paragraphs do not qualify merely because they occur last.
    if tail and all(not ref.paragraph_text(p).strip() and p not in anchors and tail_safe(p) for p in tail):
        for paragraph in tail:
            style = paragraph.find(W + "pPr/" + W + "pStyle")
            if last_sid and style is not None and style.get(W + "val") == last_sid:
                owned.add(paragraph)
    return owned


def inspect_paragraphs(document):
    ref = helper("reference_format")
    raw, _, _, root = ref._package(document)
    bib = _bib(root)
    return {"schema": "zotero-word.optional-layout-paragraphs.v1", "source_document_sha256": digest(raw), "paragraphs": [{"story": "word/document.xml", "paragraph": index, "text": ref.paragraph_text(p), "text_sha256": digest(ref.paragraph_text(p).encode()), "native_bibliography": p in bib} for index, p in enumerate(ref._paragraphs(root), 1)], "scope": "Only ordinary main-story paragraphs require a reviewed role; native bibliography paragraphs remain stage-A-owned."}


def _bindings(role_map, raw, root, profile):
    value = obj(role_map)
    require(isinstance(value, dict) and value.get("schema") == ROLE_MAP and set(value) <= {"schema", "source_document_sha256", "paragraphs"}, "Unsupported optional role map")
    require(value.get("source_document_sha256") == digest(raw), "Role map belongs to another citation checkpoint")
    require(isinstance(value.get("paragraphs"), list), "A full role-map list is required")
    ref = helper("reference_format")
    paragraphs = ref._paragraphs(root); bib = _bib(root)
    parents = {c: p for p in root.iter() for c in p}
    expected = {i for i, p in enumerate(paragraphs, 1) if p not in bib}
    seen, bindings = set(), []
    for entry in value["paragraphs"]:
        require(isinstance(entry, dict) and set(entry) == {"story", "paragraph", "text_sha256", "role", "evidence"}, "Each role binding needs exact position/text hash, role and evidence")
        i = entry["paragraph"]
        require(entry["story"] == "word/document.xml" and type(i) is int and i in expected and i not in seen, "Invalid/duplicate ordinary role position or native bibliography overlap")
        require(entry["role"] in profile["roles"] and isinstance(entry["evidence"], str) and entry["evidence"].strip(), "Role rules/evidence missing")
        p = paragraphs[i - 1]
        require(entry["text_sha256"] == digest(ref.paragraph_text(p).encode()), "Stale role text hash")
        ancestor, tags = p, []
        while ancestor in parents:
            ancestor = parents[ancestor]; tags.append(ancestor.tag)
        require(not any(t in (W + "txbxContent", W + "sdt", W + "customXml", ref.MC + "AlternateContent", ref.MC + "Choice", ref.MC + "Fallback") for t in tags), "Unsupported role container")
        require(not any(e.tag in (W + "p", W + "sdt", W + "txbxContent", ref.MC + "AlternateContent") for e in list(p.iter())[1:]), "Nested paragraph/alternate role scope requires native formatting")
        require(entry["role"] != "table_body" or W + "tc" in tags, "table_body requires an actual cell paragraph")
        seen.add(i); bindings.append((entry, p))
    require(seen == expected, "Full ordinary paragraph coverage required; unassigned count=" + str(len(expected - seen)))
    return value, bindings


def _style(styles, kind, base, rules):
    sid = "ZWLayout" + ("P" if kind == "paragraph" else "C") + digest([base, rules])[:18]
    require(styles.find(f"{W}style[@{W}styleId='{sid}']") is None, "A generated style ID already exists; keep the prior citation checkpoint")
    node = ET.SubElement(styles, W + "style", {W + "type": kind, W + "customStyle": "1", W + "styleId": sid})
    ET.SubElement(node, W + "name", {W + "val": "Zotero Word Layout " + sid[-10:]})
    if base:
        ET.SubElement(node, W + "basedOn", {W + "val": base})
    return node, sid


def _font_values(changes, toggles=None):
    values = []
    for name, value in sorted(changes.items()):
        if name == "font_family":
            values.append(("rFonts", {k: value for k in ("ascii", "hAnsi", "eastAsia", "cs")}))
        elif name == "font_size_pt":
            values.extend((n, {"val": str(round(value * 2))}) for n in ("sz", "szCs"))
        elif name in ("bold", "italic"):
            values.extend((n, {"val": "1"}) for n in (("b", "bCs") if name == "bold" else ("i", "iCs")) if toggles is None or toggles.get(n))
    return values


def _font_differences(tokens, rules, preserve_inline=None):
    changed = {}
    for name in set(rules) & FONT:
        expected = rules[name]["value"]
        offset = {"font_family": 1, "font_size_pt": 2, "bold": 3, "italic": 4}[name]
        expected = helper("reference_format")._font_key(expected) if name == "font_family" else round(expected * 2) if name == "font_size_pt" else expected
        values = {c[offset] for c in tokens}
        if name in ("bold", "italic") and preserve_inline and name in preserve_inline:
            continue
        if values != {expected}:
            require(name not in ("bold", "italic") or len(values) == 1, "Mixed-script role emphasis requires native review")
            changed[name] = rules[name]["value"]
    return changed


def _expected(source_document, profile, role_map):
    ref, body = helper("reference_format"), helper("body_format")
    profile = validate_profile(profile)
    raw, infos, parts, root = ref._package(source_document)
    mapping, bindings = _bindings(role_map, raw, root, profile)
    require("word/styles.xml" in parts, "Document styles are required for bounded layout")
    styles = ref._xml(parts["word/styles.xml"])
    resolver = body.Resolver(root, parts)
    protected = ref._field_runs(root)
    cache_p, cache_c, counts, samples = {}, {}, Counter(), []
    already_matching_runs = 0
    original_protected = [canonical(r) for r in root.iter(W + "r") if r in protected]
    for entry, p in bindings:
        rules = profile["roles"][entry["role"]]
        state, mark = resolver.paragraph(p), resolver.mark(p)
        mark_changes = _font_differences([(" ",) + mark], rules)
        old_style = resolver.rich._selected(p.find(W + "pPr"), "pStyle", resolver.rich.default_para)
        ordinary = []
        for run in p.iter(W + "r"):
            text = "".join(e.text or "" for e in run.iter(W + "t"))
            if run in protected or not text:
                continue
            tokens = resolver.chars(p, run, text)
            style_layers, direct = resolver.rich.layers(p, run)
            char_base = resolver.rich._selected(direct, "rStyle", resolver.rich.default_char)
            char_layers = resolver.rich._chain(char_base, "character")
            inline = set()
            for name, prop in (("bold", "b"), ("italic", "i")):
                if any(layer is not None and layer.find(W + prop) is not None
                       and layer.find(W + prop).get(W + "val", "1") in ("1", "true", "on")
                       for layer in [direct] + char_layers):
                    inline.add(name)
            changes = _font_differences(tokens, rules, inline)
            if changes:
                toggles = {prop: resolver.rich.toggle(prop, style_layers, None) != changes[name]
                           for name in ("bold", "italic") if name in changes
                           for prop in (("b", "bCs") if name == "bold" else ("i", "iCs"))}
                ordinary.append((run, char_base, changes, toggles))
            else:
                already_matching_runs += 1
        deltas = {name: rule["value"] for name, rule in rules.items() if name in PARA and _observed_para(state, name) != rule["value"]}
        if deltas:
            cache_key = digest([old_style, deltas])
            if cache_key not in cache_p:
                node, sid = _style(styles, "paragraph", old_style, deltas)
                props = ET.SubElement(node, W + "pPr")
                for name, value in sorted(deltas.items()):
                    local, attrs = _paragraph_rule(name, value); _attrs(props, local, attrs)
                cache_p[cache_key] = sid; counts["paragraph_styles_created"] += 1
            props = _ensure(p, "pPr")
            _attrs(props, "pStyle", {"val": cache_p[cache_key]})
            for name, value in deltas.items():
                local, attrs = _paragraph_rule(name, value); _remove(props, local, attrs)
                counts["paragraph_rule_changes"] += 1
            counts["paragraphs_rebound"] += 1
        for run, char_base, changes, toggles in ordinary:
            cache_key = digest([char_base, changes, toggles])
            if cache_key not in cache_c:
                node, sid = _style(styles, "character", char_base, changes)
                props = ET.SubElement(node, W + "rPr")
                for local, attrs in _font_values(changes, toggles):
                    _attrs(props, local, attrs)
                cache_c[cache_key] = sid; counts["character_styles_created"] += 1
            props = _ensure(run, "rPr")
            _attrs(props, "rStyle", {"val": cache_c[cache_key]})
            for name in changes:
                if name == "font_family":
                    _remove(props, "rFonts", ("ascii", "hAnsi", "eastAsia", "cs", "asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme", "csTheme"))
                elif name == "font_size_pt":
                    for local in ("sz", "szCs"):
                        _remove(props, local, ("val",))
                elif name in ("bold", "italic"):
                    for local in (("b", "bCs") if name == "bold" else ("i", "iCs")):
                        direct = props.find(W + local)
                        if direct is not None:
                            props.remove(direct)
            counts["ordinary_runs_rebound"] += 1
        if mark_changes:
            props = _ensure(_ensure(p, "pPr"), "rPr")
            for name, value in mark_changes.items():
                if name in ("bold", "italic"):
                    for local in (("b", "bCs") if name == "bold" else ("i", "iCs")):
                        _attrs(props, local, {"val": "1" if value else "0"})
                else:
                    for local, attrs in _font_values({name: value}):
                        _attrs(props, local, attrs)
                        if local == "rFonts":
                            _remove(props, local, ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme", "csTheme"))
            counts["paragraph_marks_changed"] += 1
        if (deltas or ordinary or mark_changes) and len(samples) < 12:
            samples.append({"paragraph": entry["paragraph"], "role": entry["role"], "paragraph_rules": sorted(deltas), "ordinary_runs_changed": len(ordinary), "mark_rules": sorted(mark_changes)})
    for section in root.iter(W + "sectPr"):
        for name, rule in profile.get("page", {}).items():
            local, attr = PAGE[name]
            node = section.find(W + local); old = node.get(W + attr) if node is not None else None
            if old != _twip(rule["value"]):
                _attrs(section, local, {attr: _twip(rule["value"])}); counts["page_rule_changes"] += 1
        size, margin = section.find(W + "pgSz"), section.find(W + "pgMar")
        if size is not None and margin is not None:
            for dim, a, b in (("w", "left", "right"), ("h", "top", "bottom")):
                values = [size.get(W + dim), margin.get(W + a), margin.get(W + b)]
                if all(v is not None for v in values):
                    require(int(values[0]) > int(values[1]) + int(values[2]), "Margins leave no writable area")
    require(original_protected == [canonical(r) for r in root.iter(W + "r") if r in protected], "Native field/equation runs changed")
    expected = dict(parts)
    if counts:
        expected["word/document.xml"] = _serialize(root, parts["word/document.xml"])
    if cache_p or cache_c:
        expected["word/styles.xml"] = _serialize(styles, parts["word/styles.xml"])
    summary = {"mapped_paragraph_count": len(bindings), "bibliography_paragraph_count_excluded": len(_bib(root)), "untouched_matching_run_count": already_matching_runs, "counts": dict(counts), "change_samples": samples, "existing_styles_unchanged": True, "field_runs_unchanged": True, "new_styles_only_for_actual_differences": True}
    return raw, infos, parts, expected, profile, mapping, summary


def _effective_parts(parts, inspection):
    ref, body = helper("reference_format"), helper("body_format")
    result = {}
    for name in inspection["story_coverage"]["included_parts"]:
        root = ref._xml(parts[name])
        if name in ("word/footnotes.xml", "word/endnotes.xml"):
            kind = "footnotes" if "footnotes" in name else "endnotes"
            ids = set(inspection["story_coverage"]["referenced_note_ids"][kind])
            for child in list(root):
                if child.get(W + "id") not in ids:
                    root.remove(child)
        records = body.story_data(root, parts)
        if name == "word/document.xml":
            owned = _bib(root)
            for paragraph, record in zip(ref._paragraphs(root), records):
                if paragraph in owned and not record["tokens"]:
                    # These empty paragraphs are structurally inside the native
                    # BIB field, or have the narrow generated-tail ownership
                    # above. Word may restore Bibliography indents on saving.
                    record.pop("properties", None); record.pop("mark", None)
                    record["bibliography"] = True
        result[name] = records
    return result


def _atoms(root, parts):
    """Ordinary characters plus exact source emphasis/field ownership."""
    ref, body = helper("reference_format"), helper("body_format")
    resolver = body.Resolver(root, parts)
    protected = ref._field_runs(root)
    records, stack, inline_cache = [], [], {}
    def walk(node, current=None, paragraph=None, run=None):
        if node.tag in (W + "pPr", W + "rPr", M + "oMath", M + "oMathPara", W + "drawing", W + "pict", W + "object"):
            return
        if node.tag == W + "p":
            paragraph, current = node, []
            records.append(current)
        if node.tag == W + "r":
            run = node
        if node.tag == W + "fldChar":
            kind = node.get(W + "fldCharType")
            if kind == "begin":
                stack.append({"phase": "code", "parts": [], "code": ""})
            elif kind == "separate":
                require(stack, "Unbound field separator")
                stack[-1].update(phase="result", code="".join(stack[-1]["parts"]))
            elif kind == "end":
                require(stack, "Unbound field end"); stack.pop()
        elif node.tag == W + "instrText":
            require(stack and stack[-1]["phase"] == "code", "Unbound field code")
            stack[-1]["parts"].append(node.text or "")
        elif node.tag in (W + "t", W + "tab", W + "br", W + "cr") and (not stack or stack[-1]["phase"] == "result"):
            native = any(re.search(r"\bZOTERO_(?:ITEM|BIBL)\b", f["code"]) for f in stack)
            if not native:
                require(current is not None and paragraph is not None and run is not None, "Unbound ordinary text")
                control = "text" if node.tag == W + "t" else "tab" if node.tag == W + "tab" else "break:" + node.get(W + "type", "textWrapping")
                text = node.text or "" if node.tag == W + "t" else " " if node.tag == W + "tab" else "\n"
                if run not in inline_cache:
                    _, direct = resolver.rich.layers(paragraph, run)
                    sid = resolver.rich._selected(direct, "rStyle", resolver.rich.default_char)
                    layers = [direct] + resolver.rich._chain(sid, "character")
                    inline_cache[run] = {name for name, prop in (("bold", "b"), ("italic", "i"))
                                         if any(layer is not None and layer.find(W + prop) is not None
                                                and layer.find(W + prop).get(W + "val", "1") in ("1", "true", "on") for layer in layers)}
                for token in resolver.chars(paragraph, run, text, control):
                    current.append((token, run in protected, inline_cache[run]))
        for child in ref._children(node):
            walk(child, current, paragraph, run)
    walk(root)
    require(not stack, "Unclosed field during actual-rule check")
    return resolver, records


def _check_actual_rules(before, actual, mapping, profile):
    """Resolve target values independently of the expected-output generator."""
    ref = helper("reference_format")
    source_root, target_root = ref._xml(before["word/document.xml"]), ref._xml(actual["word/document.xml"])
    source_resolver, source = _atoms(source_root, before)
    target_resolver, target = _atoms(target_root, actual)
    source_paragraphs, target_paragraphs = ref._paragraphs(source_root), ref._paragraphs(target_root)
    require(len(source) == len(target) == len(source_paragraphs) == len(target_paragraphs), "Saved paragraph count changed during layout")
    checked, exceptions = Counter(), Counter()
    for entry in mapping["paragraphs"]:
        index = entry["paragraph"] - 1
        rules = profile["roles"][entry["role"]]
        p = target_paragraphs[index]
        state = target_resolver.paragraph(p)
        for name in set(rules) & PARA:
            require(_observed_para(state, name) == rules[name]["value"], "Actual target role rule differs: paragraph " + str(index + 1) + "/" + name)
            checked["paragraph_properties"] += 1
        old, new = source[index], target[index]
        require([(a[0][0], a[0][-1]) for a in old] == [(a[0][0], a[0][-1]) for a in new], "Literal ordinary text/control flow changed during actual-rule check")
        for (a, protected, emphasis), (b, _, _) in zip(old, new):
            if protected or a[-1] != "text":
                continue
            for name in set(rules) & FONT:
                offset = {"font_family": 1, "font_size_pt": 2, "bold": 3, "italic": 4}[name]
                value = rules[name]["value"]
                expected = ref._font_key(value) if name == "font_family" else round(value * 2) if name == "font_size_pt" else value
                if name in emphasis:
                    expected = a[offset]; exceptions[name] += 1
                require(b[offset] == expected, "Actual target text rule differs: paragraph " + str(index + 1) + "/" + name)
                checked["character_properties"] += 1
        mark = target_resolver.mark(p)
        for name in set(rules) & FONT:
            offset = {"font_family": 0, "font_size_pt": 1, "bold": 2, "italic": 3}[name]
            value = rules[name]["value"]
            expected = ref._font_key(value) if name == "font_family" else round(value * 2) if name == "font_size_pt" else value
            require(mark[offset] == expected, "Actual paragraph-mark rule differs: paragraph " + str(index + 1) + "/" + name)
            checked["paragraph_mark_properties"] += 1
    sections = list(target_root.iter(W + "sectPr"))
    require(sections or not profile.get("page"), "Target has no declared page section")
    for section in sections:
        for name, rule in profile.get("page", {}).items():
            local, attr = PAGE[name]; node = section.find(W + local)
            require(node is not None and node.get(W + attr) == _twip(rule["value"]), "Actual target page rule differs: " + name)
            checked["page_properties"] += 1
    return {"result": "PASS", "checked": dict(checked), "preserved_source_inline_emphasis_character_counts": dict(exceptions)}


def verify(source_document, target_document, profile, role_map, apply_receipt=None, allow_native_resave=False):
    report = {"schema": "zotero-word.optional-layout-verification.v1", "result": "HOLD", "issues": [], "scope": "Actual optional ordinary typography and bounded package transformation only; native/visual acceptance belongs to the final stage-B audit."}
    try:
        values = _expected(source_document, profile, role_map)
        raw, _, before, expected, profile, mapping, summary = values
        target_raw, _, actual, _ = helper("reference_format")._package(target_document)
        report.update(source_document_sha256=digest(raw), target_document_sha256=digest(target_raw), profile_sha256=digest(profile), role_map_sha256=digest(mapping), **summary)
        if apply_receipt is not None:
            receipt = obj(apply_receipt)
            require(receipt.get("schema") == "zotero-word.optional-layout-application.v1"
                    and receipt.get("result") in ("LAYOUT_APPLIED_PENDING_REVIEW", "LAYOUT_APPLIED_PENDING_NATIVE_REVIEW")
                    and Path(receipt.get("source_document", "")).resolve() == Path(source_document).resolve()
                    and receipt.get("source_document_sha256") == digest(raw)
                    and Path(receipt.get("target_document", "")).resolve() == Path(target_document).resolve()
                    and receipt.get("profile_sha256") == digest(profile)
                    and receipt.get("role_map_sha256") == digest(mapping),
                    "Layout receipt belongs to another checkpoint, target path or typography plan")
            if not allow_native_resave:
                require(receipt.get("target_document_sha256") == digest(target_raw),
                        "Layout receipt does not bind this generated target")
        require(set(actual) == set(expected), "Package parts added or removed")
        report["actual_rule_check"] = _check_actual_rules(before, actual, mapping, profile)
        if not allow_native_resave:
            for name in actual:
                require(actual[name] == expected[name] or (name in ("word/document.xml", "word/styles.xml") and canonical(helper("reference_format")._xml(actual[name])) == canonical(helper("reference_format")._xml(expected[name]))), "Target differs from the declared bounded layout at " + name)
        else:
            source = helper("reference_format")._inspect_native_document(source_document)
            target = helper("reference_format")._inspect_native_document(target_document)
            require(not source["issues"] and not target["issues"] and source["story_coverage"] == target["story_coverage"], "Native save changed or invalidated story coverage")
            # Compare ordinary text/control streams plus every effective font,
            # emphasis, paragraph/mark/table/page property against the plan.
            # This permits Word's run fragmentation without accepting new prose.
            require(_effective_parts(expected, source) == _effective_parts(actual, target), "Actual saved ordinary typography/text differs from the declared plan")
            require(helper("body_format").page_state(expected) == helper("body_format").page_state(actual), "Saved page/table/global settings differ from the layout plan")
            require(source["media_sha256"] == target["media_sha256"], "Media payload changed")
            protected_a, protected_b = source["protected_content_sha256"], target["protected_content_sha256"]
            require({k: v for k, v in protected_a.items() if k != "equations"} == {k: v for k, v in protected_b.items() if k != "equations"}, "Drawing/non-Zotero field/table topology changed")
            if protected_a["equations"] != protected_b["equations"]:
                math_result = helper("math_equivalence").verify(source_document, target_document, source_inspection=source, target_inspection=target)
                require(math_result["result"] == "PASS", "Math equivalence failed after native save")
                report["math_equivalence"] = {"result": "PASS", "changed_equation_count": math_result["changed_equation_count"]}
            report["resave_comparison"] = "effective_properties_and_literal_ordinary_controls"
        # Independently resolve the expected output so an unresolved or
        # unsupported target cannot become a successful writer receipt.
        observed = helper("reference_format")._inspect_native_document(source_document)
        _effective_parts(expected, observed)
        require(digest(Path(source_document).read_bytes()) == report["source_document_sha256"] and digest(Path(target_document).read_bytes()) == report["target_document_sha256"], "Saved files changed during layout verification")
        report.update(result="PASS", native_validation_required=True, visual_review_required=True)
    except (ValueError, TypeError, KeyError, AttributeError, OSError, BadZipFile, ET.ParseError, json.JSONDecodeError) as exc:
        report["issues"].append(str(exc))
    return report


def apply(source_document, target_document, profile, role_map, consent_context):
    context = helper("workflow_state").validate_context(consent_context)
    source, target = Path(source_document).resolve(), Path(target_document).resolve()
    require(source == Path(context["citation_document"]).resolve() and digest(source.read_bytes()) == context["citation_document_sha256"], "Layout source must be the actual authorized citation checkpoint")
    require(source != target and not target.exists(), "Layout output must be a new separate file; citation checkpoint is immutable")
    profile_value = validate_profile(profile)
    require(profile_value["journal"] == context["journal"], "Layout and citation target journals differ")
    raw, infos, before, expected, p, mapping, summary = _expected(source, profile_value, role_map)
    protected_paths = {source, Path(context["citation_baseline"]).resolve(), Path(context["decision_path"]).resolve(), Path(context["reference_profile"]).resolve()}
    for candidate in (profile, role_map):
        if not isinstance(candidate, dict):
            protected_paths.add(Path(candidate).resolve())
    require(target not in protected_paths, "Layout target overlaps its evidence")
    target.parent.mkdir(parents=True, exist_ok=True)
    temp = None
    try:
        with tempfile.NamedTemporaryFile(dir=target.parent, prefix=".zotero-layout-", suffix=".docx", delete=False) as stream:
            temp = Path(stream.name)
        if expected == before:
            temp.write_bytes(raw)
        else:
            with ZipFile(temp, "w") as archive:
                for info in infos:
                    archive.writestr(info, expected[info.filename])
        receipt = verify(source, temp, p, mapping)
        require(receipt["result"] == "PASS", "Layout candidate failed verification: " + "; ".join(receipt["issues"]))
        helper("workflow_state").validate_context(context)
        require(digest(source.read_bytes()) == digest(raw) and not target.exists(), "Citation checkpoint or target changed before publishing layout copy")
        # Exclusive link publishes a verified local file without overwriting a
        # concurrent target, and the independently retained A remains untouched.
        os.link(temp, target)
        temp.unlink(); temp = None
        return {"schema": "zotero-word.optional-layout-application.v1", "result": "LAYOUT_APPLIED_PENDING_REVIEW", "citation_stage_status": "COMPLETE", "layout_stage_status": "AWAITING_FINAL_REVIEW", "source_document": str(source), "source_document_sha256": digest(raw), "target_document": str(target), "target_document_sha256": digest(target.read_bytes()), "profile_sha256": digest(p), "role_map_sha256": digest(mapping), "decision_file_sha256": context["decision_file_sha256"], "citation_checkpoint_retained": True, "format_verification": receipt, "native_or_visual_completion_claimed": False, **summary}
    finally:
        if temp is not None and temp.exists():
            temp.unlink()


if __name__ == "__main__":
    raise SystemExit("Use the public Zotero Word layout commands; this helper has no unguarded writer CLI.")

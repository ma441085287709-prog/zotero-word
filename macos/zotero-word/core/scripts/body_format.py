"""Read-only preservation of ordinary manuscript typography, paragraphs and pages.

No target typography profile, role map, DOCX writer, native action or completion
claim. Zotero result characters and their bibliography paragraphs are owned by
the separate citation/References checks. Unsupported alignments fail closed.
"""
from __future__ import annotations
from collections import Counter
import copy
from difflib import SequenceMatcher
import hashlib
import importlib.util
import json
from pathlib import Path
import posixpath
import re
import sys
import xml.etree.ElementTree as ET
from zipfile import BadZipFile

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
SCHEMA = "zotero-word.body-format-preservation.v1"
_CACHE = {}


def helper(name):
    if name not in _CACHE:
        spec = importlib.util.spec_from_file_location("_body_preservation_" + name, Path(__file__).with_name(name + ".py"))
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


def key(value):
    require(isinstance(value, dict) and isinstance(value.get("story"), str) and type(value.get("paragraph")) is int, "Invalid paragraph alignment")
    return value["story"], value["paragraph"]


def boolean(value):
    require(value in ("1", "true", "on", "0", "false", "off"), "Invalid Word boolean")
    return value in ("1", "true", "on")


def canonical(node):
    """Ignore XML prefixes/order, and normalize lexical number/boolean values."""
    if node is None:
        return None
    attrs = {}
    for name, value in node.attrib.items():
        if name.startswith(W + "rsid"):
            continue
        attrs[name] = int(value) if re.fullmatch(r"-?\d+", value) else value
    return [node.tag, sorted(attrs.items()), (node.text or "").strip(), [canonical(c) for c in node]]


class Resolver:
    def __init__(self, root, parts, story="word/document.xml"):
        self.ref = helper("reference_format")
        self.rich = helper("rich_text")._EffectiveRichText(root, parts, self.ref)
        self.root, self.parts = root, parts
        self.story = story
        self.paragraph_indices = {p: i for i, p in enumerate(root.iter(W + "p"), 1)}
        styles = self.ref._xml(parts["word/styles.xml"]) if "word/styles.xml" in parts else ET.Element(W + "styles")
        self.dp = styles.find(W + "docDefaults/" + W + "pPrDefault/" + W + "pPr")
        self.paragraph_cache = {}
        self.themes = [self.ref._xml(value) for name, value in parts.items() if re.fullmatch(r"word/theme/theme\d+\.xml", name)]

    def paragraph_layers(self, paragraph):
        sid = self.rich._selected(paragraph.find(W + "pPr"), "pStyle", self.rich.default_para)
        def chain(current, seen=None):
            if current is None:
                return []
            require(current in self.rich.styles, "Missing paragraph style")
            seen = set(seen or ())
            require(current not in seen, "Cyclic paragraph style")
            seen.add(current)
            style = self.rich.styles[current]
            require(style.get(W + "type") == "paragraph", "Wrong paragraph style type")
            base = style.find(W + "basedOn")
            return chain(base.get(W + "val") if base is not None else None, seen) + [style.find(W + "pPr")]
        if sid not in self.paragraph_cache:
            self.paragraph_cache[sid] = [self.dp] + chain(sid)
        # Conditional table paragraph layers cannot be silently omitted.
        for ancestor in self.ancestors(paragraph):
            if ancestor.tag == W + "tbl":
                self.rich._table_safe(ancestor)
                props = ancestor.find(W + "tblPr")
                current = self.rich._selected(props, "tblStyle", self.rich.default_table)
                while current is not None:
                    style = self.rich.styles[current]
                    require(not any(len(p) or p.attrib for p in style.iter(W + "pPr")), "Table paragraph formatting requires a separate preservation resolver")
                    base = style.find(W + "basedOn")
                    current = base.get(W + "val") if base is not None else None
        return self.paragraph_cache[sid] + [paragraph.find(W + "pPr")]

    def ancestors(self, element):
        while element in self.rich.parents:
            element = self.rich.parents[element]
            yield element

    @staticmethod
    def merged(layers, name):
        value = {}
        for layer in layers:
            if layer is not None:
                nodes = layer.findall(W + name)
                require(len(nodes) <= 1, "Duplicate property: " + name)
                if nodes:
                    value.update({k.removeprefix(W): v for k, v in nodes[0].attrib.items()})
        return value

    def paragraph(self, paragraph):
        layers = self.paragraph_layers(paragraph)
        state = {}
        bools = {"keepNext": False, "keepLines": False, "pageBreakBefore": False, "widowControl": True,
                 "contextualSpacing": False, "snapToGrid": True, "bidi": False,
                 "mirrorIndents": False, "suppressLineNumbers": False, "suppressAutoHyphens": False,
                 "adjustRightInd": True, "wordWrap": True, "overflowPunct": True,
                 "autoSpaceDE": True, "autoSpaceDN": True, "kinsoku": True, "topLinePunct": False}
        for name, default in bools.items():
            value = default
            for layer in layers:
                node = layer.find(W + name) if layer is not None else None
                if node is not None:
                    value = boolean(node.get(W + "val", "1"))
            state[name] = value
        state["jc"] = self.merged(layers, "jc").get("val", "left")
        state["outlineLvl"] = int(self.merged(layers, "outlineLvl").get("val", "9"))
        defaults = {"spacing": {"before": "0", "after": "0", "beforeLines": "0", "afterLines": "0",
                                "line": "240", "lineRule": "auto", "beforeAutospacing": "0", "afterAutospacing": "0"},
                    "ind": {"left": "0", "right": "0", "firstLine": "0", "hanging": "0",
                            "leftChars": "0", "rightChars": "0", "firstLineChars": "0", "hangingChars": "0"}}
        for name, default in defaults.items():
            values = {**default, **self.merged(layers, name)}
            if name == "ind":
                for modern, old in (("start", "left"), ("end", "right"), ("startChars", "leftChars"), ("endChars", "rightChars")):
                    if modern in values:
                        values[old] = values.pop(modern)
            for prop, value in list(values.items()):
                values[prop] = boolean(value) if prop.endswith("Autospacing") else int(value) if re.fullmatch(r"-?\d+", value) else value
            state[name] = values
        handled = set(bools) | {"jc", "outlineLvl", "spacing", "ind", "pStyle", "rPr", "sectPr"}
        # Preserve less-common effective properties too; do not rely on style ID.
        extras = {}
        for layer in layers:
            if layer is not None:
                for node in layer:
                    if node.tag.startswith(W) and node.tag.removeprefix(W) not in handled:
                        if node.tag == W + "numPr":
                            num = node.find(W + "numId")
                            require(num is not None and num.get(W + "val") == "0", "Active list numbering requires separate preservation support")
                            continue
                        extras[node.tag] = canonical(node)
        state["other"] = extras
        return state

    def theme(self, name):
        require(len(self.themes) == 1, "Ambiguous theme font")
        family = "majorFont" if name.startswith("major") else "minorFont" if name.startswith("minor") else None
        slot = "ea" if name.endswith("EastAsia") else "cs" if name.endswith("Bidi") else "latin"
        node = self.themes[0].find(".//" + A + "fontScheme/" + A + str(family) + "/" + A + slot)
        require(node is not None and node.get("typeface"), "Unresolved theme font")
        return node.get("typeface")

    def font_size(self, layers, complex_char, paragraph, run):
        """Preserve a declared size or its genuine absence; never guess points.

        OOXML permits sz to be absent throughout the hierarchy. This reader
        compares saved formatting, so equal absence is comparable even though
        an application's chosen display size cannot be inferred from it.
        """
        name = "szCs" if complex_char else "sz"
        runs = paragraph.findall(W + "r")
        location = (self.story + " paragraph " + str(self.paragraph_indices.get(paragraph, "?"))
                    + (" run " + str(runs.index(run) + 1) if run in runs else " paragraph mark/nested run"))
        size = None
        for layer in layers:
            nodes = layer.findall(W + name) if layer is not None else []
            require(len(nodes) <= 1, "Duplicate ordinary font size " + name + " at " + location)
            if nodes:
                node = nodes[0]
                value = node.get(W + "val")
                require(not len(node) and set(node.attrib) == {W + "val"}
                        and isinstance(value, str) and re.fullmatch(r"[0-9]+", value)
                        and int(value) > 0,
                        "Invalid ordinary font size " + name + " at " + location)
                size = int(value)
        return size if size is not None else ("unspecified-size", name)

    def chars(self, paragraph, run, text, control="text"):
        rich = self.rich.characters(paragraph, run, text)
        styles, direct = self.rich.layers(paragraph, run)
        layers = [self.rich.defaults] + styles + [direct]
        forced = self.rich.absolute("cs", styles, direct, False) or self.rich.absolute("rtl", styles, direct, False)
        values = []
        font_cache = {}
        size_cache = {}
        for ch in rich:
            code = ord(ch["char"])
            complex_char = forced or 0x590 <= code <= 0x8ff or 0xfb1d <= code <= 0xfdff or 0xfe70 <= code <= 0xfeff
            slot = "cs" if complex_char else "ascii" if code < 128 else "eastAsia" if (0x2e80 <= code <= 0x9fff or 0xac00 <= code <= 0xd7af or 0xf900 <= code <= 0xfaff or 0x20000 <= code <= 0x3134f) else "hAnsi"
            if slot not in font_cache:
                font = None
                theme_key = {"ascii": "asciiTheme", "hAnsi": "hAnsiTheme", "eastAsia": "eastAsiaTheme", "cs": "cstheme"}[slot]
                for layer in layers:
                    rf = layer.find(W + "rFonts") if layer is not None else None
                    if rf is not None:
                        if rf.get(W + slot) is not None:
                            font = rf.get(W + slot)
                        theme = rf.get(W + theme_key) or (rf.get(W + "csTheme") if slot == "cs" else None)
                        if theme is not None:
                            font = self.theme(theme)
                require(isinstance(font, str) and font.strip(), "Ordinary font is unresolved")
                font_cache[slot] = self.ref._font_key(font)
            if complex_char not in size_cache:
                size_cache[complex_char] = self.font_size(layers, complex_char, paragraph, run)
            values.append((ch["char"], font_cache[slot], size_cache[complex_char], ch["bold"], ch["italic"], ch["vertical_align"], control))
        return values

    def mark(self, paragraph):
        run = ET.Element(W + "r")
        props = paragraph.find(W + "pPr/" + W + "rPr")
        if props is not None:
            run.append(copy.deepcopy(props))
        return self.chars(paragraph, run, " ")[0][1:]


def story_data(root, parts, story="word/document.xml"):
    ref = helper("reference_format")
    resolver = Resolver(root, parts, story)
    records, stack = [], []
    def walk(node, current=None, run=None):
        if node.tag in (W + "pPr", W + "rPr"):
            # Current core inspection represents paragraph tab-stop elements as
            # virtual spaces. Bind that inspection text without treating layout
            # declarations as visible characters; their properties are checked.
            if current is not None:
                current["visible"] += " " * sum(1 for _ in node.iter(W + "tab"))
            return  # layout tab stops/properties are not visible text controls
        if node.tag in (M + "oMath", M + "oMathPara", W + "drawing", W + "pict", W + "object"):
            return  # protected by the separate structural/content audit
        if node.tag == W + "p":
            current = {"paragraph": node, "tokens": [], "visible": "", "native": False, "bibliography": False}
            records.append(current)
        if node.tag == W + "r":
            run = node
        if node.tag == W + "fldChar":
            kind = node.get(W + "fldCharType")
            if kind == "begin":
                stack.append({"phase": "code", "parts": [], "code": ""})
            elif kind == "separate":
                require(stack and stack[-1]["phase"] == "code", "Malformed field separator")
                stack[-1].update(phase="result", code="".join(stack[-1]["parts"]))
            elif kind == "end":
                require(stack, "Unbound field end")
                stack.pop()
        elif node.tag == W + "instrText":
            require(stack and stack[-1]["phase"] == "code", "Unbound field instruction")
            stack[-1]["parts"].append(node.text or "")
        elif node.tag in (W + "t", W + "tab", W + "br", W + "cr"):
            if not stack or stack[-1]["phase"] == "result":
                require(current is not None and run is not None, "Unbound ordinary text")
                native = any(re.search(r"\bZOTERO_(?:ITEM|BIBL)\b", f["code"]) for f in stack)
                bib = any(re.search(r"\bZOTERO_BIBL\b", f["code"]) for f in stack)
                text = (node.text or "") if node.tag == W + "t" else " " if node.tag == W + "tab" else "\n"
                # Upstream inspection does not count w:cr in visible offsets.
                # Retain it in our control stream so insertion/deletion still
                # cannot evade preservation merely because prose hashes match.
                if node.tag != W + "cr":
                    current["visible"] += text
                current["native"] |= native
                current["bibliography"] |= bib
                if not native:
                    control = "tab" if node.tag == W + "tab" else "break:" + node.get(W + "type", "textWrapping") if node.tag in (W + "br", W + "cr") else "text"
                    current["tokens"].extend(resolver.chars(current["paragraph"], run, text, control))
        elif node.tag == W + "fldSimple":
            # Ordinary figure/table cross-references are common in native
            # manuscripts. Read their displayed text/format; the separate
            # inspector still protects the exact instruction and result.
            require(re.match(r"^\s*(?:SEQ|REF|PAGEREF|PAGE|NUMPAGES|HYPERLINK)\b", node.get(W + "instr", ""), re.I),
                    "Unsupported simple field needs separate preservation support")
            require(not any(child.tag in (W + "fldSimple", W + "fldChar", W + "instrText") for child in list(node.iter())[1:]),
                    "Nested simple fields need separate preservation support")
        elif node.tag in (W + "sym", W + "ruby", W + "noBreakHyphen", W + "softHyphen"):
            require(False, "Special glyph needs separate ordinary-format preservation support")
        for child in ref._children(node):
            walk(child, current, run)
    walk(root)
    require(not stack, "Unclosed native field")
    for record in records:
        if not record["bibliography"]:
            record["properties"] = resolver.paragraph(record["paragraph"])
            record["mark"] = resolver.mark(record["paragraph"])
        record.pop("paragraph")
    return records


def package_data(path, report):
    ref = helper("reference_format")
    raw, _, parts, main = ref._package(path)
    require(report.get("schema") == "zotero-word.inspection.v1" and report.get("document_sha256") == digest(raw)
            and report.get("document") == str(Path(path).resolve()) and not report.get("issues"), "Inspection does not bind an issue-free saved document")
    result = {}
    for name in report["story_coverage"]["included_parts"]:
        root = main if name == "word/document.xml" else ref._xml(parts[name])
        if name in ("word/footnotes.xml", "word/endnotes.xml"):
            kind = "footnotes" if "footnotes" in name else "endnotes"
            selected = set(report["story_coverage"]["referenced_note_ids"][kind])
            for child in list(root):
                if child.get(W + "id") not in selected:
                    root.remove(child)
        result[name] = story_data(root, parts, name)
    blocks = {key(b): b for b in report["blocks"]}
    for name, values in result.items():
        for index, record in enumerate(values, 1):
            require((name, index) in blocks and record["visible"] == blocks[name, index]["text"], "Ordinary formatting walk differs from current-text inspection")
    return digest(raw), parts, result


def font_size_diagnostics(stories):
    """Make unspecified sizes visible without claiming an absolute point size."""
    result = {"mode": "declared_or_unspecified_preservation", "unspecified_character_count": 0,
              "unspecified_paragraph_mark_count": 0, "locations": []}
    def unspecified(value):
        return isinstance(value, tuple) and value[:1] == ("unspecified-size",)
    for story, records in stories.items():
        for index, record in enumerate(records, 1):
            chars = sum(unspecified(token[2]) for token in record["tokens"])
            mark = "mark" in record and unspecified(record["mark"][1])
            result["unspecified_character_count"] += chars
            result["unspecified_paragraph_mark_count"] += int(mark)
            if (chars or mark) and len(result["locations"]) < 12:
                result["locations"].append({"story": story, "paragraph": index,
                                            "text_characters": chars, "paragraph_mark": bool(mark)})
    return result


def page_state(parts):
    ref = helper("reference_format")
    root = ref._xml(parts["word/document.xml"])
    sections = []
    rels = {}
    if "word/_rels/document.xml.rels" in parts:
        rels = {r.get("Id"): r for r in ref._xml(parts["word/_rels/document.xml.rels"])}
    R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
    for section in root.iter(W + "sectPr"):
        children = [canonical(c) for c in section if c.tag not in (W + "headerReference", W + "footerReference", W + "printerSettings")]
        for child in section:
            if child.tag in (W + "headerReference", W + "footerReference"):
                rel = rels.get(child.get(R + "id"))
                require(rel is not None and rel.get("TargetMode", "Internal") == "Internal", "Unresolved section header/footer relationship")
                children.append([child.tag, child.get(W + "type"), posixpath.normpath(posixpath.join("word", rel.get("Target", "")))])
        sections.append(sorted(children, key=lambda x: x[0]))
    settings = ref._xml(parts["word/settings.xml"]) if "word/settings.xml" in parts else ET.Element(W + "settings")
    names = {"defaultTabStop", "characterSpacingControl", "compat", "mirrorMargins", "evenAndOddHeaders",
             "autoHyphenation", "consecutiveHyphenLimit", "hyphenationZone", "doNotHyphenateCaps", "gutterAtTop"}
    tables = []
    for name, raw in parts.items():
        if name == "word/document.xml" or re.fullmatch(r"word/(?:header\d+|footer\d+|footnotes|endnotes)\.xml", name):
            story = root if name == "word/document.xml" else ref._xml(raw)
            for table in story.iter(W + "tbl"):
                rows = []
                for row in table.findall(W + "tr"):
                    rows.append([canonical(row.find(W + "trPr")), [canonical(cell.find(W + "tcPr")) for cell in row.findall(W + "tc")]])
                tables.append([name, canonical(table.find(W + "tblPr")), canonical(table.find(W + "tblGrid")), rows])
    return {"sections": sections, "tables": tables, "settings": sorted([canonical(c) for c in settings if c.tag.removeprefix(W) in names], key=lambda x: x[0])}


def verify(source_document, target_document, source_report, target_report, block_alignment, manifest):
    result = {"schema": SCHEMA, "result": "HOLD", "issues": [], "difference_samples": [],
              "scope": "Read-only source-to-target ordinary effective font/size/bold/italic/vertical, paragraph and section-layout preservation; native Zotero result characters/References excluded. Not native/visual/claim acceptance.",
              "unsupported_scope": ["Active numbering and conditional table paragraph styling", "Multi-paragraph approved body rewrites", "Unresolved fonts, legacy symbols or complex revisions"],
              "checked_paragraph_count": 0, "checked_character_count": 0}
    try:
        sh, sp, source = package_data(source_document, source_report)
        th, tp, target = package_data(target_document, target_report)
        result.update(source_document_sha256=sh, target_document_sha256=th,
                      font_size_resolution={"source": font_size_diagnostics(source),
                                            "target": font_size_diagnostics(target)})
        require(page_state(sp) == page_state(tp), "Ordinary section/page/global-layout properties changed")
        require(set(source) == set(target), "Referenced story set changed; explicit story remapping is required")
        pairs, used_target = {}, set()
        for row in block_alignment:
            sk, tk = key(row["source"]), key(row["target"])
            require(sk not in pairs and tk not in used_target, "Duplicate ordinary alignment")
            pairs[sk] = tk; used_target.add(tk)
        rewrites, rewritten_targets = set(), set()
        source_blocks = {key(b): b for b in source_report["blocks"]}
        target_blocks = {key(b): b for b in target_report["blocks"]}
        references = set()
        declared_references = manifest.get("source_reference_blocks", [])
        require(isinstance(declared_references, list), "Source reference-block exclusions must be a list")
        for entry in declared_references:
            sk = key(entry)
            require(sk in source_blocks and sk not in references
                    and isinstance(entry.get("reason"), str) and entry["reason"].strip()
                    and isinstance(entry.get("evidence"), str) and entry["evidence"].strip(),
                    "Each source reference-block exclusion needs a unique observed position, reason and evidence")
            references.add(sk)
        # The coverage audit owns the semantic reference-only decision. Recheck
        # its exact positions here; exclusions cannot hide actual native fields
        # or a cue that the manifest maps into an active manuscript citation.
        require(not references.intersection(key(f["context"]) for f in source_report.get("citations", [])),
                "A source reference-block exclusion overlaps a native manuscript citation")
        result["source_reference_blocks_excluded"] = len(references)
        for change in manifest.get("body_changes", []):
            left, right = change.get("source_blocks"), change.get("target_blocks")
            require(isinstance(left, list) and isinstance(right, list) and len(left) == len(right) == 1, "Multi-block body rewrites need explicit format alignment support")
            sk, tk = key(left[0]), key(right[0])
            require(sk not in references, "A body rewrite overlaps an excluded source reference block")
            require(sk in source_blocks and tk in target_blocks and change.get("meaning_preserved") is True
                    and isinstance(change.get("reason"), str) and change["reason"].strip()
                    and isinstance(change.get("evidence"), str) and change["evidence"].strip(), "Body rewrite lacks exact reviewed evidence")
            require(left[0].get("text_sha256") == hashlib.sha256(source_blocks[sk]["text"].encode()).hexdigest()
                    and right[0].get("text_sha256") == hashlib.sha256(target_blocks[tk]["text"].encode()).hexdigest(), "Body rewrite text binding is stale")
            require(sk not in {a for a, _ in rewrites} and tk not in rewritten_targets, "Repeated body rewrite block")
            require(pairs.get(sk, tk) == tk, "Body rewrite conflicts with audited paragraph alignment")
            pairs[sk] = tk; rewrites.add((sk, tk)); rewritten_targets.add(tk)
        cues = {c.get("cue_id"): c for c in source_report.get("citation_cues", [])}
        converted_blocks = {key(cues[e["source_id"]]) for e in manifest.get("entries", [])
                            if e.get("disposition") in ("converted", "removed") and e.get("source_id") in cues}
        require(not references.intersection(converted_blocks),
                "A source reference-block exclusion overlaps a converted/removed manuscript cue")

        def ordinary_candidates(story, records, blocks, excluded):
            # Only a wholly empty, unanchored tail after the last observed
            # native bibliography result is References-owned generated space.
            # Appendix paragraphs or graphics after a list keep their ordinary
            # empty paragraphs inside the preservation check as well.
            last_bib = max((i for i, p in enumerate(records, 1) if p["bibliography"]), default=0)
            tail = set()
            if last_bib and all(not p["tokens"] and not p["native"]
                                and not blocks[story, i].get("protected_tokens")
                                for i, p in enumerate(records, 1) if i > last_bib):
                tail = {(story, i) for i in range(last_bib + 1, len(records) + 1)}
            return [(story, i) for i, p in enumerate(records, 1)
                    if not p["bibliography"] and (story, i) not in excluded and (story, i) not in tail]

        # Reference migration/regeneration can shift the native bibliography
        # start. Align the remaining ordinary paragraphs by story order rather
        # than requiring source and target bibliography indices to match.
        expected_pairs = {}
        for story in source:
            sa, ta = source[story], target[story]
            left_keys = ordinary_candidates(story, sa, source_blocks, references)
            right_keys = ordinary_candidates(story, ta, target_blocks, set())
            require(len(left_keys) == len(right_keys), "Ordinary paragraph counts differ after reviewed References exclusions at " + story)
            for sk, tk in zip(left_keys, right_keys):
                expected_pairs[sk] = tk
                require(sk not in pairs or pairs[sk] == tk, "Audited paragraph alignment conflicts with ordinary story order")
                if sk in pairs:
                    continue
                left, right = sa[sk[1] - 1], ta[tk[1] - 1]
                empty = (not any(c[0].strip() for c in left["tokens"] + right["tokens"])
                         and not left["native"] and not right["native"]
                         and not source_blocks[sk].get("protected_tokens")
                         and not target_blocks[tk].get("protected_tokens"))
                require(story != "word/document.xml" or empty,
                        "An ordinary paragraph has no audited preservation alignment at " + story + ":" + str(sk[1]))
                pairs[sk] = tk
        require(all(expected_pairs.get(sk) == tk for sk, tk in pairs.items()),
                "Ordinary alignment includes an excluded reference block or unobserved paragraph")
        counters = Counter()
        def different(kind, sk, tk, detail=None):
            counters[kind] += 1
            if len(result["difference_samples"]) < 12:
                result["difference_samples"].append({"kind": kind, "source": list(sk), "target": list(tk), **(detail or {})})
        for sk, tk in pairs.items():
            require(sk[0] in source and tk[0] in target and 1 <= sk[1] <= len(source[sk[0]]) and 1 <= tk[1] <= len(target[tk[0]]), "Alignment points outside an observed story")
            left, right = source[sk[0]][sk[1] - 1], target[tk[0]][tk[1] - 1]
            require(not left["bibliography"] and not right["bibliography"], "Ordinary alignment overlaps a native bibliography")
            result["checked_paragraph_count"] += 1
            if left["properties"] != right["properties"] or left["mark"] != right["mark"]:
                different("paragraph_or_mark_format", sk, tk)
            a, b = left["tokens"], right["tokens"]
            astr, bstr = "".join(c[0] for c in a), "".join(c[0] for c in b)
            if [(c[0], c[-1]) for c in a if c[-1] != "text"] != [(c[0], c[-1]) for c in b if c[-1] != "text"]:
                different("ordinary_control_format", sk, tk)
            if astr != bstr:
                # Citation clue conversion has already been checked by the
                # coverage audit. A reviewed rewrite is not a format exemption.
                require((sk, tk) in rewrites or sk in converted_blocks, "Unreviewed ordinary text differs during format preservation")
            opcodes = [("equal", 0, len(a), 0, len(b))] if astr == bstr else SequenceMatcher(None, astr, bstr, autojunk=False).get_opcodes()
            for tag, i, j, k, l in opcodes:
                if tag == "equal":
                    for offset, (before, after) in enumerate(zip(a[i:j], b[k:l])):
                        result["checked_character_count"] += 1
                        if before[1:] != after[1:]:
                            different("ordinary_character_format", sk, tk, {"source_offset": i + offset, "target_offset": k + offset})
                elif tag in ("replace", "insert") and k < l:
                    # New prose must use the same format as surviving/source
                    # context; a fully changed or mixed-format span is ambiguous.
                    baseline = {c[1:] for c in a[max(0, i - 1):min(len(a), max(j, i + 1))]}
                    require(len(baseline) == 1, "Inserted/rewritten ordinary text has ambiguous source formatting")
                    if any(c[1:] not in baseline for c in b[k:l]):
                        different("rewritten_character_format", sk, tk)
        result["difference_counts"] = dict(counters)
        require(not counters, "Ordinary manuscript formatting changed")
        require(digest(Path(source_document).read_bytes()) == sh and digest(Path(target_document).read_bytes()) == th, "Document changed during ordinary format preservation")
        result["result"] = "PASS"
    except (ValueError, TypeError, KeyError, AttributeError, OSError, BadZipFile, ET.ParseError, json.JSONDecodeError) as exc:
        result["issues"].append(str(exc))
    return result

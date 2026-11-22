#!/usr/bin/env python3
"""Read actual native Word/Zotero result character formatting; never write DOCX."""
from __future__ import annotations

import importlib.util
import json
from pathlib import Path
import re
import sys
import unicodedata
import xml.etree.ElementTree as ET
from zipfile import BadZipFile

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
MC = "{http://schemas.openxmlformats.org/markup-compatibility/2006}"
M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
SCHEMA = "zotero-word.native-rich-text.v1"
_FORMAT = None


def _formatter():
    global _FORMAT
    if _FORMAT is None:
        spec = importlib.util.spec_from_file_location("_zotero_native_rich_text_reader", Path(__file__).with_name("reference_format.py"))
        module = importlib.util.module_from_spec(spec)
        previous = sys.dont_write_bytecode
        try:
            sys.dont_write_bytecode = True
            spec.loader.exec_module(module)
        finally:
            sys.dont_write_bytecode = previous
        _FORMAT = module
    return _FORMAT


def _require(value, message):
    if not value:
        raise ValueError(message)


def _property(rpr, name):
    nodes = rpr.findall(W + name) if rpr is not None else []
    _require(len(nodes) <= 1, "Duplicate rich-text property: " + name)
    if not nodes:
        return None
    node = nodes[0]
    _require(not len(node) and set(node.attrib) <= {W + "val"}, "Unsupported rich-text property shape: " + name)
    if name == "vertAlign":
        value = node.get(W + "val")
        _require(value in ("baseline", "superscript", "subscript"), "Invalid vertical alignment")
        return value
    value = node.get(W + "val", "1")
    _require(value in ("1", "true", "on", "0", "false", "off"), "Invalid rich-text boolean: " + name)
    return value in ("1", "true", "on")


class _EffectiveRichText:
    def __init__(self, root, parts, fmt):
        self.root, self.fmt = root, fmt
        styles_root = fmt._xml(parts["word/styles.xml"]) if "word/styles.xml" in parts else ET.Element(W + "styles")
        values = styles_root.findall(W + "style")
        self.styles = {s.get(W + "styleId"): s for s in values}
        _require(len(self.styles) == len(values) and None not in self.styles, "Missing/duplicate style identifiers")
        self.defaults = styles_root.find(W + "docDefaults/" + W + "rPrDefault/" + W + "rPr")
        self.default_para = self._default("paragraph")
        self.default_char = self._default("character")
        self.default_table = self._default("table")
        self.parents = {child: parent for parent in root.iter() for child in parent}
        self.special = fmt._special_font_runs(root, parts.get("word/styles.xml"))
        self.cache = {}
        self.safe_tables = set()

    def _default(self, kind):
        selected = [s.get(W + "styleId") for s in self.styles.values() if s.get(W + "type") == kind and s.get(W + "default") in ("1", "true", "on")]
        _require(len(selected) <= 1, "Multiple default styles")
        return selected[0] if selected else None

    def _chain(self, style_id, kind, seen=None):
        if style_id is None:
            return []
        key = (style_id, kind)
        if key in self.cache:
            return self.cache[key]
        _require(style_id in self.styles, "Missing applied style: " + str(style_id))
        seen = set(seen or ())
        _require(style_id not in seen, "Cyclic style inheritance")
        seen.add(style_id)
        style = self.styles[style_id]
        _require(style.get(W + "type") == kind, "Wrong applied style type")
        bases = style.findall(W + "basedOn")
        _require(len(bases) <= 1 and (not bases or bases[0].get(W + "val")), "Malformed style inheritance")
        result = self._chain(bases[0].get(W + "val") if bases else None, kind, seen) + [style.find(W + "rPr")]
        self.cache[key] = result
        return result

    def _selected(self, props, name, default):
        nodes = props.findall(W + name) if props is not None else []
        _require(len(nodes) <= 1 and (not nodes or nodes[0].get(W + "val")), "Ambiguous selected rich-text style")
        return nodes[0].get(W + "val") if nodes else default

    def layers(self, paragraph, run):
        tags, tables, ancestor = [], [], paragraph
        while ancestor in self.parents:
            ancestor = self.parents[ancestor]; tags.append(ancestor.tag)
            if ancestor.tag == W + "tbl": tables.append(ancestor)
        _require(not any(tag in (W + "txbxContent", W + "sdt", MC + "Choice", MC + "Fallback", MC + "AlternateContent") for tag in tags), "Native rich-text context has unsupported alternate/container formatting")
        for table in tables: self._table_safe(table)
        ppr, rpr = paragraph.find(W + "pPr"), run.find(W + "rPr")
        para = self._chain(self._selected(ppr, "pStyle", self.default_para), "paragraph")
        char = self._chain(self._selected(rpr, "rStyle", self.default_char), "character")
        return para + char, rpr

    def _table_safe(self, table):
        """Allow editable table citations only with no unresolved inline layers."""
        if table in self.safe_tables: return
        props = table.find(W + "tblPr")
        style_id = self._selected(props, "tblStyle", self.default_table)
        seen = set()
        while style_id is not None:
            _require(style_id not in seen and style_id in self.styles, "Missing/cyclic applied table style")
            seen.add(style_id); style = self.styles[style_id]
            _require(style.get(W + "type") == "table", "Wrong applied table style type")
            # Even an inactive-looking conditional layer may be selected through
            # tblLook/cnfStyle; do not silently ignore its inline character data.
            _require(not any(len(rpr) or rpr.attrib for rpr in style.iter(W + "rPr")), "Conditional/table style inline formatting requires native rich-text review")
            _require(not any(s.tag in (W + "pStyle", W + "rStyle") for s in style.iter()), "Table style has unresolved nested paragraph/character style selection")
            bases = style.findall(W + "basedOn")
            _require(len(bases) <= 1 and (not bases or bases[0].get(W + "val")), "Malformed table style inheritance")
            style_id = bases[0].get(W + "val") if bases else None
        if props is not None:
            _require(not any(len(rpr) or rpr.attrib for rpr in props.iter(W + "rPr")), "Direct table inline formatting is unresolved")
        self.safe_tables.add(table)

    def toggle(self, name, styles, direct):
        initial = _property(self.defaults, name)
        state = bool(initial)
        declarations = [_property(layer, name) for layer in styles]
        override = _property(direct, name)
        # Word documents implementation-specific behavior for default-true
        # multilevel toggles. Unsupported combinations cannot become false PASS.
        _require(not (initial is True and any(v is not None for v in declarations) and override is None), "Default-true multilevel toggle requires native rich-text review: " + name)
        for value in declarations:
            if value is True:
                state = not state
        return override if override is not None else state

    def absolute(self, name, styles, direct, default):
        value = default
        for layer in [self.defaults] + styles + [direct]:
            observed = _property(layer, name)
            if observed is not None:
                value = observed
        return value

    def characters(self, paragraph, run, text):
        _require(run not in self.special, "Legacy symbol/private-use glyph rich text cannot be inferred")
        styles, direct = self.layers(paragraph, run)
        for name in ("vanish", "webHidden", "caps", "smallCaps"):
            _require(not self.toggle(name, styles, direct), "Hidden/case-transformed native result requires native review: " + name)
        for layer in [self.defaults] + styles + [direct]:
            if layer is not None:
                positions = layer.findall(W + "position")
                _require(len(positions) <= 1 and (not positions or positions[0].get(W + "val") == "0"), "Manual text raising/lowering is unsupported")
        forced = self.absolute("cs", styles, direct, False) or self.absolute("rtl", styles, direct, False)
        normal = {"bold": self.toggle("b", styles, direct), "italic": self.toggle("i", styles, direct)}
        complex_flags = {"bold": self.toggle("bCs", styles, direct), "italic": self.toggle("iCs", styles, direct)}
        vertical = self.absolute("vertAlign", styles, direct, "baseline")
        result = []
        complex_present = any(0x590 <= ord(c) <= 0x8ff or 0xfb1d <= ord(c) <= 0xfdff or 0xfe70 <= ord(c) <= 0xfeff for c in text)
        _require(not (complex_present and not forced and any(c.isascii() and c.isalnum() for c in text)), "Mixed complex-script runs need explicit cs/rtl rich-text scope")
        for char in text:
            code = ord(char)
            complex_char = forced or 0x590 <= code <= 0x8ff or 0xfb1d <= code <= 0xfdff or 0xfe70 <= code <= 0xfeff
            name = unicodedata.name(char, "")
            unsupported = any(script in name for script in ("DEVANAGARI", "BENGALI", "GURMUKHI", "GUJARATI", "ORIYA", "TAMIL", "TELUGU", "KANNADA", "MALAYALAM", "SINHALA", "THAI", "LAO", "TIBETAN", "MYANMAR", "KHMER"))
            _require(not (unsupported and not forced), "Complex-script rich-text selection needs explicit cs/rtl")
            result.append({"char": char, **(complex_flags if complex_char else normal), "vertical_align": vertical})
        return result


def _actual_tokens(root, fields, fmt):
    """Walk actual complex-field phases so control whitespace stays owned."""
    by_code = {f["code"]: f for f in fields if re.search(r"\bZOTERO_(?:ITEM|BIBL)\b", f["code"], re.I)}
    _require(len(by_code) == sum(bool(re.search(r"\bZOTERO_(?:ITEM|BIBL)\b", f["code"], re.I)) for f in fields), "Duplicate native field instructions")
    stacks, result, owners = [], {}, {}
    def walk(node, paragraph=None, run=None):
        if node.tag == W + "p": paragraph = node
        if node.tag == W + "r": run = node
        if node.tag == W + "fldChar":
            kind = node.get(W + "fldCharType")
            if kind == "begin": stacks.append({"phase":"code", "parts":[], "code":None})
            elif kind == "separate":
                _require(stacks and stacks[-1]["phase"] == "code", "Malformed field activation")
                stacks[-1].update(phase="result", code="".join(stacks[-1]["parts"]))
            elif kind == "end":
                _require(stacks, "Unbound field end")
                stacks.pop()
            else: _require(False, "Unknown field marker")
        elif node.tag == W + "instrText":
            _require(stacks and stacks[-1]["phase"] == "code", "Unbound field instruction")
            stacks[-1]["parts"].append(node.text or "")
        elif node.tag in (W + "t", W + "tab", W + "br", W + "cr"):
            if not stacks or stacks[-1]["phase"] == "result":
                code = stacks[-1]["code"] if stacks else None
                owners.setdefault(run, set()).add(code)
                if code in by_code:
                    text = (node.text or "") if node.tag == W + "t" else "\t" if node.tag == W + "tab" else "\n"
                    _require(paragraph is not None and run is not None, "Unbound native result text")
                    result.setdefault(code, []).append((paragraph, run, text))
        elif node.tag in (W + "sym", W + "noBreakHyphen", W + "softHyphen", W + "drawing", W + "pict", W + "object", W + "ruby", M + "oMath", M + "oMathPara"):
            _require(not (stacks and stacks[-1]["phase"] == "result" and stacks[-1]["code"] in by_code), "Unsupported visible native glyph would be dropped")
        for child in fmt._children(node): walk(child, paragraph, run)
    walk(root)
    _require(not stacks, "Unclosed native field")
    for code, tokens in result.items():
        _require(not by_code[code]["nested"], "Nested native result fields are unsupported")
        for _, run, _ in tokens:
            _require(owners.get(run) == {code} and not any(e.tag in (W + "fldChar", W + "instrText") for e in run.iter()), "Native result shares a run with body/code/other fields")
    return result


def inspect_native_rich_text(document, inspection=None):
    report = {"schema":SCHEMA, "result":"HOLD", "issues":[], "citations":{}, "bibliography":{},
              "scope":"Read-only effective bold/italic/vertical alignment of actual main-story native Zotero results; no CSL expectation, application control or native Refresh is implied."}
    try:
        fmt = _formatter()
        raw, _, parts, root = fmt._package(document, native_results_readonly=True)
        sha = fmt.digest(raw)
        report["document_sha256"] = sha
        observed = inspection if inspection is not None else fmt._inspect_native_document(document)
        _require(isinstance(observed, dict) and observed.get("schema") == "zotero-word.inspection.v1" and observed.get("document_sha256") == sha and observed.get("document") == str(Path(document).resolve()), "Native inspection is not bound to current saved document")
        _require(not observed.get("issues"), "Native document inspection has unresolved issues")
        _require(observed.get("citation_count", 0) > 0, "No activated native Zotero citation fields")
        _require(all(c.get("context", {}).get("story") == "word/document.xml" for c in observed.get("citations", [])), "Native rich-text citation stories beyond main document are unsupported")
        _require(len(observed.get("bibliographies", [])) == 1 and observed["bibliographies"][0].get("context", {}).get("story") == "word/document.xml", "Exactly one main-story native bibliography is required")
        bib, bib_paragraphs, _, bib_data, fields = fmt._bibliography_scope(root)
        _require(observed["bibliographies"][0].get("field_data") == bib_data and observed["bibliographies"][0].get("display") == "".join(bib["text_parts"]), "Actual bibliography differs from inspection")
        tokens = _actual_tokens(root, fields, fmt)
        effective = _EffectiveRichText(root, parts, fmt)
        paragraphs = fmt._paragraphs(root); indices = {p:i for i,p in enumerate(paragraphs,1)}
        def rich(code):
            values = tokens.get(code, [])
            _require(values and any(text.strip() for _,_,text in values), "Native field has no visible result characters")
            chars, groups = [], {}
            for paragraph, run, text in values:
                current = effective.characters(paragraph, run, text)
                chars.extend(current);groups.setdefault(paragraph, []).extend(current)
            return chars, groups
        observed_by_id = {c["citationID"]:c for c in observed.get("citations", [])}
        for field in fields:
            marker = re.match(r"^\s*(?:ADDIN\s+)?ZOTERO_ITEM\s+CSL_CITATION\s*(?=\{)", field["code"], re.I)
            if marker is None: continue
            data, end = json.JSONDecoder().raw_decode(field["code"][marker.end():])
            _require(not field["code"][marker.end()+end:].strip() and isinstance(data,dict), "Malformed citation native JSON")
            cid = data.get("citationID");old=observed_by_id.get(cid)
            _require(old is not None and old.get("field_data") == data and old.get("display") == "".join(field["text_parts"]), "Native citation ID/data/result differs from inspection")
            _require(not field["nested"], "Nested native citation fields are unsupported")
            chars, groups = rich(field["code"])
            _require(len(groups) == 1, "A citation result crosses paragraphs")
            paragraph = next(iter(groups))
            report["citations"][cid] = {"text":"".join(c["char"] for c in chars),"characters":chars,"context":old["context"],"paragraph":indices[paragraph],"field_data":data}
        _require(set(report["citations"]) == set(observed_by_id), "Not all actual citation fields were inspected")
        chars, groups = rich(bib["code"])
        _require({p for p,g in groups.items() if any(c['char'].strip() for c in g)} == {p for _,p in bib_paragraphs}, "Bibliography paragraph scope differs from actual native result")
        report["bibliography"] = {"text":"".join(c["char"] for c in chars),"characters":chars,"field_data":bib_data,
            "paragraphs":[{"story":"word/document.xml","paragraph":index,"text":"".join(c["char"] for c in groups[p]),"characters":groups[p]} for index,p in bib_paragraphs]}
        report.update(result="INSPECTED", citation_count=len(report["citations"]), bibliography_paragraph_count=len(bib_paragraphs), native_refresh_performed=False,
                      semantics_sources=["https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.bold", "https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.verticaltextalignment", "https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/f7130225-2368-48f3-acae-a9d278d0fb25"])
        _require(fmt.digest(Path(document).read_bytes()) == sha, "Document changed during rich-text inspection")
    except (ValueError, TypeError, KeyError, AttributeError, OSError, BadZipFile, ET.ParseError, json.JSONDecodeError) as exc:
        report["result"] = "HOLD";report["issues"].append(str(exc))
    return report

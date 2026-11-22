"""Recompute expected CSL output from bound installed resources; read-only QA.

No processor code is bundled here. Only the explicitly trusted installed Zotero
distribution is read for offline rendering, never application control or fields.
"""
from __future__ import annotations

import copy
import hashlib
from html.parser import HTMLParser
import importlib.util
import json
import os
from pathlib import Path
import re
import subprocess
import sys
import xml.etree.ElementTree as ET
from zipfile import ZipFile, BadZipFile

CSL = "{http://purl.org/net/xbiblio/csl}"
CONFIG = "zotero-word.csl-configuration.v1"
SCHEMA = "zotero-word.citation-style-verification.v1"
PROCESSOR = "chrome/content/zotero/xpcom/citeproc.js"
CITE = "chrome/content/zotero/xpcom/cite.js"
ABBREVIATIONS = "resource/schema/abbreviations.json"
LOCALE = "chrome/content/zotero/locale/csl/locales-{}.xml"
KEYS = {"schema", "style_id", "style_file", "style_sha256", "parent_file", "parent_sha256",
        "zotero_archive", "archive_sha256", "processor_sha256", "abbreviation_handler_sha256",
        "abbreviations_sha256", "abbreviation_override", "abbreviation_override_sha256",
        "locale", "locale_sha256", "node_executable", "evidence"}


def require(ok, message):
    if not ok:
        raise ValueError(message)


def digest(raw):
    if not isinstance(raw, bytes):
        raw = json.dumps(raw, ensure_ascii=False, sort_keys=True, separators=(",", ":")).encode()
    return hashlib.sha256(raw).hexdigest()


def file_hash(path):
    with Path(path).open("rb") as stream:
        h = hashlib.sha256()
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            h.update(block)
        return h.hexdigest()


def helper(name):
    spec = importlib.util.spec_from_file_location("_zotero_style_" + name, Path(__file__).with_name(name + ".py"))
    m = importlib.util.module_from_spec(spec)
    previous = sys.dont_write_bytecode
    try:
        sys.dont_write_bytecode = True
        spec.loader.exec_module(m)
    finally:
        sys.dont_write_bytecode = previous
    return m


def object_value(value):
    return value if isinstance(value, dict) else json.loads(Path(value).read_text(encoding="utf-8"))


def validate_config(config):
    require(isinstance(config, dict) and set(config) == KEYS and config.get("schema") == CONFIG, "Missing or unsupported bound CSL configuration")
    for name in ("style_id", "style_file", "zotero_archive", "locale", "node_executable", "evidence"):
        require(isinstance(config[name], str) and config[name].strip(), "CSL configuration needs " + name)
    require(re.fullmatch(r"[a-z]{2,3}(?:-[A-Za-z0-9]{2,8})*", config["locale"]), "Unsupported CSL locale identifier")
    for name in ("style_sha256", "archive_sha256", "processor_sha256", "abbreviation_handler_sha256", "abbreviations_sha256", "locale_sha256"):
        require(isinstance(config[name], str) and re.fullmatch(r"[a-f0-9]{64}", config[name]), "Invalid CSL resource hash: " + name)
    for path, sha in (("parent_file", "parent_sha256"), ("abbreviation_override", "abbreviation_override_sha256")):
        require(config[path] is None or (isinstance(config[path], str) and config[path].strip()), "Invalid CSL resource path: " + path)
        require(config[sha] is None or (isinstance(config[sha], str) and re.fullmatch(r"[a-f0-9]{64}", config[sha])), "Invalid optional resource hash: " + sha)
    require(bool(config["parent_file"]) == bool(config["parent_sha256"]), "Parent style path/hash must agree")
    require(config["abbreviation_override"] is not None or config["abbreviation_override_sha256"] is None, "Override hash has no path")
    return config


def xml(raw):
    require(b"<!DOCTYPE" not in raw.upper() and b"<!ENTITY" not in raw.upper(), "External XML declarations are unsupported")
    return ET.fromstring(raw)


def style_bundle(style_file, parent_file):
    selected_raw = Path(style_file).read_bytes()
    selected = xml(selected_raw)
    require(selected.tag == CSL + "style", "Selected file is not CSL")
    sid = selected.findtext(CSL + "info/" + CSL + "id")
    require(sid, "Selected CSL has no identity")
    links = selected.findall(CSL + "info/" + CSL + "link[@rel='independent-parent']")
    require(len(links) <= 1, "Ambiguous dependent CSL parents")
    if links:
        require(parent_file and selected.find(CSL + "citation") is None and selected.find(CSL + "bibliography") is None, "Dependent style needs its actual independent parent")
        parent_raw = Path(parent_file).read_bytes()
        parent = xml(parent_raw)
        require(parent.tag == CSL + "style" and parent.findtext(CSL + "info/" + CSL + "id") == links[0].get("href"), "CSL parent identity differs from selected dependency")
        require(not parent.findall(CSL + "info/" + CSL + "link[@rel='independent-parent']"), "Multi-level CSL dependencies need separate support")
        independent, independent_raw = parent, parent_raw
    else:
        require(parent_file is None, "Independent style must not be replaced by an unrelated parent")
        parent_raw = None
        independent, independent_raw = selected, selected_raw
    require(independent.find(CSL + "citation") is not None and independent.find(CSL + "bibliography") is not None, "A citation and bibliography CSL layout are both required")
    return sid, selected_raw, parent_raw, selected, independent, independent_raw


def resources(archive, locale, override):
    before = file_hash(archive)
    with ZipFile(archive) as z:
        processor = z.read(PROCESSOR)
        source = z.read(CITE).decode("utf-8")
        start = source.find("Zotero.Cite.getAbbreviation = new function () {")
        end = source.find("\n}();", start) if start >= 0 else -1
        require(start >= 0 and end >= 0, "Installed abbreviation handler extraction is unsupported")
        handler = source[start:end + 5].encode()
        default_abbreviations = z.read(ABBREVIATIONS)
        locales = {code: z.read(LOCALE.format(code)) for code in {locale, "en-US"}}
    override_hash = file_hash(override) if override and Path(override).exists() else None
    abbreviations = Path(override).read_bytes() if override_hash else default_abbreviations
    require(file_hash(archive) == before, "Installed Zotero resources changed during read")
    require(not override_hash or file_hash(override) == override_hash, "Abbreviation override changed during read")
    return {"archive_sha256": before, "processor_sha256": digest(processor), "abbreviation_handler_sha256": digest(handler),
            "abbreviations_sha256": digest(abbreviations), "abbreviation_override_sha256": override_hash, "locale_sha256": digest(locales[locale]),
            "processor_js": processor.decode(), "abbreviation_js": handler.decode(), "abbreviations": json.loads(abbreviations),
            "locales": {code: raw.decode() for code, raw in locales.items()}}


def build_config(style_file, parent_file, zotero_archive, node_executable, locale, evidence, abbreviation_override=None):
    sid, selected, parent, _, _, _ = style_bundle(style_file, parent_file)
    r = resources(zotero_archive, locale, abbreviation_override)
    node = Path(node_executable).resolve()
    require(node.is_file() and os.access(node, os.X_OK), "Discover an actual usable Node executable before CSL checks")
    config = {"schema": CONFIG, "style_id": sid, "style_file": str(Path(style_file).resolve()), "style_sha256": digest(selected),
              "parent_file": str(Path(parent_file).resolve()) if parent_file else None, "parent_sha256": digest(parent) if parent else None,
              "zotero_archive": str(Path(zotero_archive).resolve()), "node_executable": str(node), "locale": locale,
              "abbreviation_override": str(Path(abbreviation_override).resolve()) if abbreviation_override else None, "evidence": evidence,
              **{name: r[name] for name in ("archive_sha256", "processor_sha256", "abbreviation_handler_sha256", "abbreviations_sha256", "abbreviation_override_sha256", "locale_sha256")}}
    return validate_config(config)


class RichHTML(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.characters, self.stack = [], [("root", {"bold": False, "italic": False, "vertical_align": "baseline"})]

    def handle_starttag(self, tag, attrs):
        attrs = dict(attrs)
        require(tag in ("div", "span", "a", "b", "strong", "i", "em", "sup", "sub", "br"), "Unsupported expected CSL HTML element: " + tag)
        fmt = dict(self.stack[-1][1])
        if tag in ("b", "strong"): fmt["bold"] = True
        if tag in ("i", "em"): fmt["italic"] = True
        if tag in ("sup", "sub"): fmt["vertical_align"] = "superscript" if tag == "sup" else "subscript"
        for declaration in attrs.get("style", "").split(";"):
            if not declaration.strip(): continue
            name, _, value = declaration.partition(":")
            name, value = name.strip().lower(), value.strip().lower()
            if name == "font-weight":
                require(value in ("bold", "normal", "400", "700"), "Unsupported expected font weight")
                fmt["bold"] = value in ("bold", "700")
            elif name == "font-style":
                require(value in ("italic", "oblique", "normal"), "Unsupported expected font style")
                fmt["italic"] = value != "normal"
            elif name == "vertical-align":
                require(value in ("super", "sub", "baseline"), "Unsupported expected vertical alignment")
                fmt["vertical_align"] = {"super": "superscript", "sub": "subscript", "baseline": "baseline"}[value]
            else:
                require(False, "Unsupported expected CSL CSS: " + name)
        if tag == "br":
            self.characters.append({"char": "\n", **fmt})
        else:
            self.stack.append((tag, fmt))

    def handle_endtag(self, tag):
        require(len(self.stack) > 1 and self.stack[-1][0] == tag, "Unbalanced expected CSL HTML")
        self.stack.pop()
        if tag == "div": self.characters.append({"char": "\n", **self.stack[-1][1]})

    def handle_data(self, data):
        self.characters.extend({"char": c, **self.stack[-1][1]} for c in data)


def expected_characters(html):
    parser = RichHTML(); parser.feed(html); parser.close()
    require(len(parser.stack) == 1, "Incomplete expected CSL HTML")
    return parser.characters


def normalized(characters):
    result = []
    for c in characters:
        if c["char"].isspace():
            if result and result[-1]["char"] != " ": result.append({"char": " "})
        else:
            result.append(c)
    if result and result[-1]["char"] == " ": result.pop()
    return result


def compare(actual, expected, properties, location):
    a, e = normalized(actual), normalized(expected)
    at, et = "".join(c["char"] for c in a), "".join(c["char"] for c in e)
    if at != et:
        return [{**location, "kind": "rendered_text_differs", "actual": at, "expected": et}]
    drifts = []
    for index, (ac, ex) in enumerate(zip(a, e)):
        if ex["char"].isspace(): continue
        different = [p for p in properties if ac[p] != ex[p]]
        if different:
            drifts.append({**location, "kind": "rich_text_differs", "character_index": index,
                           "excerpt": at[max(0, index - 16):index + 17], "properties": different,
                           "actual": {p: ac[p] for p in different}, "expected": {p: ex[p] for p in different}})
    return drifts


def render_payload(observed, rich, independent_raw, settings, r):
    items, citations, parents, uri_metadata = {}, [], {}, {}
    ordered = sorted(observed["citations"], key=lambda c: (c["context"]["paragraph"], c["context"]["start"]))
    def canonical_uris(item):
        uris = sorted(set(u.replace("https://www.zotero.org/", "http://zotero.org/").replace("http://www.zotero.org/", "http://zotero.org/").replace("https://zotero.org/", "http://zotero.org/") for u in item.get("uris", [])))
        require(uris and all(re.fullmatch(r"http://zotero\.org/(?:users/(?:local/[^/]+|[0-9]+)|groups/[0-9]+)/items/[A-Z0-9]{8}", u) for u in uris), "Expected CSL item has no proven Zotero identity")
        return uris
    def find(uri):
        parents.setdefault(uri, uri)
        root = uri
        while parents[root] != root:
            root = parents[root]
        while parents[uri] != uri:
            old = parents[uri]
            parents[uri] = root
            uri = old
        return root
    # Zotero may encode merged-record aliases. Partial alias overlap must never
    # let one URI represent two different papers, nor assign two CSL numbers to
    # the same consistent native identity component.
    for citation in ordered:
        for item in citation["citationItems"]:
            uris = canonical_uris(item)
            meta = copy.deepcopy(item.get("itemData"))
            require(isinstance(meta, dict) and meta.get("type"), "Native citation lacks embedded metadata for expected CSL output")
            meta.pop("id", None)
            signature = digest(meta)
            for uri in uris:
                require(uri not in uri_metadata or uri_metadata[uri] == signature, "A Zotero URI is shared by conflicting article metadata; recheck native identities before style QA")
                uri_metadata[uri] = signature
                parents[find(uri)] = find(uris[0])
    components = {}
    for uri in parents:
        components.setdefault(find(uri), []).append(uri)
    keys = {uri: "zw-" + digest(sorted(group)) for group in components.values() for uri in group}
    for citation in ordered:
        data = copy.deepcopy(citation["field_data"])
        require(data.get("properties", {}).get("noteIndex", 0) == 0, "Note-based CSL ordering needs a separate native adapter")
        for item in data["citationItems"]:
            uris = canonical_uris(item)
            key = keys[uris[0]]
            meta = copy.deepcopy(item.get("itemData"))
            meta["id"] = key
            items[key] = meta
            item["id"] = key
            item.pop("itemData", None); item.pop("uris", None)
        data["properties"] = {k: v for k, v in data.get("properties", {}).items() if k not in ("formattedCitation", "plainCitation")}
        citations.append(data)
    return {"processor_js": r["processor_js"], "abbreviation_js": r["abbreviation_js"], "abbreviations": r["abbreviations"],
            "locales": r["locales"], "style_xml": independent_raw.decode(), "style_id": settings["style_id"], "locale": settings["locale"],
            "automatic_journal_abbreviations": settings["automatic_journal_abbreviations"], "itemsByID": items, "citations": citations}


def verify(document, profile, inspection=None):
    report = {"schema": SCHEMA, "result": "HOLD", "issues": [], "changes": [],
              "scope": "Read-only main-story CSL identity, citation text/vertical alignment and bibliography text/bold/italic/vertical alignment. No native fields or application actions created."}
    try:
        p = object_value(profile); config = validate_config(p.get("csl"))
        sid, selected, parent, selected_xml, independent, independent_raw = style_bundle(config["style_file"], config["parent_file"])
        require(sid == config["style_id"] and digest(selected) == config["style_sha256"] and (digest(parent) if parent else None) == config["parent_sha256"], "Selected CSL identity or source hash changed")
        r = resources(config["zotero_archive"], config["locale"], config["abbreviation_override"])
        for name in ("archive_sha256", "processor_sha256", "abbreviation_handler_sha256", "abbreviations_sha256", "abbreviation_override_sha256", "locale_sha256"):
            require(r[name] == config[name], "Bound installed CSL resource changed: " + name)
        core = helper("zotero_docx")
        observed = inspection if inspection is not None else core.inspect_docx(document)
        require(observed.get("document_sha256") == file_hash(document) and observed.get("document") == str(Path(document).resolve()) and not observed.get("issues"), "CSL inspection does not bind the current intact saved document")
        report["document_sha256"] = observed["document_sha256"]
        require(observed.get("preferences", {}).get("style", {}).get("id") == sid, "Word's actual selected CSL differs from the target profile")
        prefs = xml(observed["preferences"]["xml"].encode())
        options = {x.get("name"): x.get("value") for x in prefs.findall("prefs/pref")}
        abbreviation = options.get("automaticJournalAbbreviations")
        require(abbreviation in ("true", "false"), "Observe an explicit current automaticJournalAbbreviations setting")
        locale = options.get("locale") or observed["preferences"]["style"].get("locale") or selected_xml.get("default-locale") or independent.get("default-locale") or "en-US"
        require(locale == config["locale"], "Expected CSL language differs from current Word/style settings")
        rich = helper("rich_text").inspect_native_rich_text(document, inspection=observed)
        require(rich["result"] == "INSPECTED", "Native rich-text HOLD: " + "; ".join(rich["issues"]))
        require(not any(rich["bibliography"]["field_data"].get(k) for k in ("uncited", "omitted", "custom")), "Custom/uncited/omitted bibliography edits require separate supported expectation handling")
        settings = {**config, "automatic_journal_abbreviations": abbreviation == "true"}
        payload = render_payload(observed, rich, independent_raw, settings, r)
        require(Path(config["node_executable"]).is_file() and os.access(config["node_executable"], os.X_OK), "Configured Node runtime is unavailable")
        run = subprocess.run([config["node_executable"], str(Path(__file__).with_name("csl_render.js"))], input=json.dumps(payload, ensure_ascii=False), capture_output=True, text=True, timeout=45)
        require(run.returncode == 0, "Offline CSL renderer HOLD: " + run.stdout[:500] + run.stderr[:500])
        expected = json.loads(run.stdout)
        report.update(style_id=sid, csl_config_sha256=digest(config), resources={k: config[k] for k in ("style_sha256", "parent_sha256", "processor_sha256", "abbreviations_sha256", "locale_sha256")},
                      processor_version=expected["processor_version"], expected=expected, checked_citation_count=len(observed["citations"]))
        require({x["citationID"] for x in expected["citations"]} == set(rich["citations"]), "Expected CSL citations do not cover actual native fields")
        for citation in expected["citations"]:
            report["changes"].extend(compare(rich["citations"][citation["citationID"]]["characters"], expected_characters(citation["html"]), ["vertical_align"], {"citation_id": citation["citationID"]}))
        entries = expected["bibliography"]["entriesHTML"]
        paragraphs = rich["bibliography"]["paragraphs"]
        require(len(entries) == len(paragraphs), "Expected bibliography entries and actual nonempty result paragraphs differ; multi-paragraph entries require separate support")
        for i, (html, actual) in enumerate(zip(entries, paragraphs), 1):
            report["changes"].extend(compare(actual["characters"], expected_characters(html), ["bold", "italic", "vertical_align"], {"bibliography_entry": i, "paragraph": actual["paragraph"]}))
        report["checked_bibliography_entry_count"] = len(entries)
        require(not report["changes"], "Actual numbering or rich text differs from the bound target CSL")
        require(file_hash(document) == report["document_sha256"] and file_hash(config["style_file"]) == config["style_sha256"] and (file_hash(config["parent_file"]) if config["parent_file"] else None) == config["parent_sha256"], "Document or CSL source changed during style QA")
        require(file_hash(config["zotero_archive"]) == config["archive_sha256"], "Installed processor changed during style QA")
        override = config["abbreviation_override"]
        require((file_hash(override) if override and Path(override).exists() else None) == config["abbreviation_override_sha256"], "Abbreviation override changed during style QA")
        report.update(result="PASS", native_refresh_performed=False)
    except (ValueError, TypeError, KeyError, AttributeError, OSError, BadZipFile, ET.ParseError, json.JSONDecodeError, subprocess.TimeoutExpired) as exc:
        report["issues"].append(str(exc))
    return report

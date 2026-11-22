"""Bound, surgical DOCX citation writing; every output still needs native QA.

This module does not run Word or Zotero and cannot certify live associations.
It writes reviewed citation payloads only to a new copy.  Source coordinates
are exact active-text offsets (one-based main-story paragraphs, including
tables).  No visible locator tokens, revision acceptance, or library writes
are used.  Unsupported boundaries fail before publishing any output.
"""
from __future__ import annotations

import copy
from datetime import datetime, timezone
import hashlib
import importlib.util
import json
import os
from pathlib import Path
import posixpath
import re
import tempfile
from urllib.parse import urlsplit
import uuid
from zipfile import ZipFile

try:
    from lxml import etree as E
except ImportError as exc:
    raise ImportError("Hybrid citation writing needs the discovered runtime's lxml package") from exc

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
M = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"
CP = "{http://schemas.openxmlformats.org/officeDocument/2006/custom-properties}"
VT = "{http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes}"
PKGREL = "{http://schemas.openxmlformats.org/package/2006/relationships}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
CT = "{http://schemas.openxmlformats.org/package/2006/content-types}"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
MAIN = "word/document.xml"
PLAN = "zotero-word.hybrid-plan.v1"
VERSION = "2.4.3"
URI = re.compile(r"^/(users/(?:local/[A-Za-z0-9_-]+|[0-9]+)|groups/[0-9]+)/items/([A-Z0-9]{8})/?$")
MARKER = re.compile(r"\b(?:ZOTERO_ANCHOR_[A-Za-z0-9_]+|BIBHERE|REFHERE)\b")
SKIP = {W + "del", W + "moveFrom"}
OBJECTS = {M + "oMath", M + "oMathPara", W + "drawing", W + "pict", W + "object", W + "fldSimple"}
OPTIONS = {"locator", "label", "prefix", "suffix", "suppress-author", "author-only"}
_revision_spec = importlib.util.spec_from_file_location("_zotero_hybrid_revision_history", Path(__file__).with_name("revision_history.py"))
_revisions = importlib.util.module_from_spec(_revision_spec)
_revision_spec.loader.exec_module(_revisions)
_boundary_spec = importlib.util.spec_from_file_location("_zotero_hybrid_boundaries", Path(__file__).with_name("citation_quality.py"))
_boundaries = importlib.util.module_from_spec(_boundary_spec)
_boundary_spec.loader.exec_module(_boundaries)


def require(value, message):
    if not value:
        raise ValueError(message)


def digest(value):
    if not isinstance(value, bytes):
        value = json.dumps(value, sort_keys=True, ensure_ascii=False, separators=(",", ":")).encode()
    return hashlib.sha256(value).hexdigest()


def _object(value):
    return copy.deepcopy(value) if isinstance(value, dict) else json.loads(Path(value).read_text(encoding="utf-8"))


def _xml(raw):
    require(b"<!DOCTYPE" not in raw.upper() and b"<!ENTITY" not in raw.upper(), "External XML declarations are unsupported")
    return E.fromstring(raw, E.XMLParser(resolve_entities=False, no_network=True, remove_blank_text=False))


def _pack(path):
    with ZipFile(path) as z:
        names = z.namelist()
        require(len(names) == len(set(names)), "Duplicate ZIP part names")
        require(MAIN in names and "[Content_Types].xml" in names, "Not a DOCX package")
        require(all(not n.startswith("/") and ".." not in n.split("/") for n in names), "Unsafe ZIP part name")
        require(not any(n.startswith("_xmlsignatures/") for n in names), "Signed DOCX needs a separately authorized workflow")
        return {n: z.read(n) for n in names}, z.comment


def _skip(e):
    return _revisions.is_historical(e)


def _active(e):
    if _skip(e):
        return
    yield e
    for c in e:
        yield from _active(c)


def canonical_uri(uri):
    require(isinstance(uri, str), "A real Zotero item URI is required")
    u = urlsplit(uri)
    m = URI.fullmatch(u.path)
    require(u.scheme in ("http", "https") and u.netloc.lower() in ("zotero.org", "www.zotero.org") and not u.query and not u.fragment and m, "Invalid Zotero item URI")
    return "http://zotero.org" + m.group(0).rstrip("/")


def _snapshot(value):
    v = _object(value)
    require(v.get("schema") == "zotero-word.snapshot.v1" and v.get("kind") == "library", "Bind an actual current library snapshot first")
    require(v.get("filters_clear") is True and isinstance(v.get("items"), list), "Snapshot must represent the unfiltered library")
    require(v.get("items_sha256") == digest(v["items"]), "Library snapshot content hash differs")
    when = datetime.fromisoformat(v["observed_at"].replace("Z", "+00:00"))
    require(when.tzinfo is not None, "Library observation needs a timezone")
    age = (datetime.now(timezone.utc) - when.astimezone(timezone.utc)).total_seconds()
    require(-300 <= age <= 24 * 3600, "Library snapshot is stale or future-dated; observe the actual library again")
    index = {}
    aliases = v.get("identity_map", {})
    for item in v["items"]:
        raw_id = str(item.get("id", ""))
        try:
            uri = canonical_uri(raw_id)
        except ValueError:
            uri = canonical_uri(aliases.get(raw_id))
        require(urlsplit(uri).path.rsplit("/items/", 1)[0].lstrip("/") == v.get("library_id"), "Cross-library snapshot item")
        require(uri not in index, "Duplicate library item identity")
        index[uri] = item
    return v, index


def _owner(p, leaf):
    owners = []
    e = leaf.getparent()
    while e is not p:
        require(e is not None, "Text node escaped its paragraph")
        if e.tag == W + "ins":
            owners.append(e)
        elif e.tag not in (W + "r", W + "hyperlink"):
            raise ValueError("Target text is inside an unsupported wrapper: " + str(e.tag))
        e = e.getparent()
    require(len(owners) <= 1, "Nested inserted revisions require separate review")
    return owners[0] if owners else p


def _inventory(root):
    paragraphs, fields, stack = [], [], []
    for p in (e for e in _active(root) if e.tag == W + "p"):
        tokens, events, chars = [], [], []
        length = 0
        def visit(e, simple=False):
            nonlocal length
            if _skip(e):
                events.append((length, "original_revision_history", e))
                return
            if e is not p and e.tag == W + "p":
                return
            tag = e.tag
            if tag in {W + name for name in ("bookmarkStart", "bookmarkEnd", "commentRangeStart", "commentRangeEnd", "commentReference", "permStart", "permEnd", "footnoteReference", "endnoteReference", "sdt", "customXml", "smartTag", "sym", "lastRenderedPageBreak")}:
                events.append((length, "unsupported_boundary", e))
            if tag in OBJECTS:
                events.append((length, "protected_object", e))
                if tag != W + "fldSimple":
                    return
                simple = True
            if tag == W + "fldChar":
                kind = e.get(W + "fldCharType")
                events.append((length, "field", e))
                if kind == "begin":
                    if stack:
                        stack[-1]["nested"] = True
                    stack.append({"parts": [], "phase": "code", "display": [], "paragraph": len(paragraphs) + 1, "start": length, "depth": len(stack), "controls": [e], "nested": False})
                elif kind == "separate":
                    require(stack and stack[-1]["phase"] == "code", "Malformed existing field separator")
                    stack[-1]["controls"].append(e)
                    stack[-1]["phase"] = "result"
                elif kind == "end":
                    require(stack, "Malformed existing field end")
                    f = stack.pop()
                    f["controls"].append(e)
                    f["code"] = "".join(f.pop("parts"))
                    f["display"] = "".join(f["display"])
                    f["end"] = length
                    f["end_paragraph"] = len(paragraphs) + 1
                    fields.append(f)
                else:
                    require(False, "Unknown field character")
            elif tag == W + "instrText":
                require(stack, "Orphan field instruction")
                stack[-1]["parts"].append(e.text or "")
                stack[-1]["controls"].append(e)
            elif tag in (W + "t", W + "tab", W + "br"):
                if stack and stack[-1]["phase"] == "code":
                    return
                text = (e.text or "") if tag == W + "t" else (" " if tag == W + "tab" else "\n")
                owner, reason = None, None
                if tag != W + "t" or simple:
                    reason = "field_or_nontext"
                else:
                    try:
                        owner = _owner(p, e)
                        run = e.getparent()
                        require(run.tag == W + "r", "Text is not directly owned by a run")
                        require(not any(_revisions.is_property_history(x) for x in run.iter()), "Target run has property revision history")
                    except ValueError as exc:
                        reason = str(exc)
                    if stack and reason is None:
                        reason = "field_or_nontext"
                tokens.append({"node": e, "start": length, "end": length + len(text), "text": text, "owner": owner, "reason": reason, "field": stack[-1] if stack else None})
                chars.append(text)
                length += len(text)
                if stack:
                    stack[-1]["display"].append(text)
            for child in e:
                visit(child, simple)
        visit(p)
        paragraphs.append({"node": p, "paragraph": len(paragraphs) + 1, "text": "".join(chars), "tokens": tokens, "events": events})
    require(not stack, "Existing document has an unclosed field")
    require(not any(f["depth"] and "ZOTERO_" in f["code"] for f in fields), "Nested existing Zotero field needs native recovery")
    require(not any("ZOTERO_" in f["code"] and not re.match(r"^\s*ADDIN\s+ZOTERO_(?:ITEM\s+CSL_CITATION|BIBL)\s*\{", f["code"]) for f in fields), "Pending or unrecognized existing Zotero field")
    for paragraph in paragraphs:
        paragraph["fields"] = []
    for field in fields:
        paragraphs[field["paragraph"] - 1]["fields"].append(field)
    return paragraphs, fields


def _citation_payloads(fields):
    result = {}
    for f in fields:
        if "ZOTERO_ITEM" not in f["code"]:
            continue
        match = re.match(r"^\s*ADDIN\s+ZOTERO_ITEM\s+CSL_CITATION\s*(\{.*)\s*$", f["code"], re.S)
        require(match, "Malformed Zotero citation instruction")
        data = json.loads(match.group(1))
        cid = data.get("citationID")
        require(isinstance(cid, str) and cid and cid not in result, "Missing or duplicate citation identity")
        require(isinstance(data.get("citationItems"), list) and data["citationItems"], "Empty citation payload")
        for item in data["citationItems"]:
            require(isinstance(item.get("uris"), list) and item["uris"], "Existing citation lacks stable item URI")
            for u in item["uris"]:
                canonical_uri(u)
        result[cid] = (data, f)
    return result


def marker_scan(document, parts=None):
    """Inspect all XML text, including deleted revisions and unreferenced stories."""
    if parts is None:
        parts, _ = _pack(document)
    found = []
    for name, raw in parts.items():
        if not name.endswith(".xml"):
            continue
        root = _xml(raw)
        text = "".join(root.itertext())
        for m in MARKER.finditer(text):
            found.append({"part": name, "marker": m.group()})
    return found


def _edge_whitespace(parts):
    """Find text that Word may trim while loading an otherwise valid package."""
    found = []
    for name, raw in parts.items():
        if not name.startswith("word/") or not name.endswith(".xml"):
            continue
        for node in _xml(raw).iter():
            value = node.text or ""
            if node.tag not in (W + "t", W + "delText", M + "t") or not value or not (value[0] in " \t\r\n" or value[-1] in " \t\r\n"):
                continue
            mode = next((e.get(XML_SPACE) for e in [node] + list(node.iterancestors()) if e.get(XML_SPACE) is not None), None)
            if mode != "preserve":
                found.append({"part": name, "tag": node.tag, "text_sha256": digest(value.encode())})
    return found


def inspect(document):
    parts, _ = _pack(document)
    paragraphs, fields = _inventory(_xml(parts[MAIN]))
    return {"schema": "zotero-word.hybrid-inventory.v1", "source_document": str(Path(document).resolve()), "source_sha256": digest(Path(document).read_bytes()), "scope": "Active main-document paragraphs, including editable tables; deleted/moveFrom text excluded from coordinates but retained and marker-scanned", "paragraphs": [{"story": MAIN, "paragraph": p["paragraph"], "text": p["text"], "paragraph_sha256": digest(p["text"].encode()), "in_table": any(a.tag == W + "tc" for a in p["node"].iterancestors()), "eligible_spans": [{"start": t["start"], "end": t["end"], "revision_id": t["owner"].get(W + "id") if t["owner"] is not None and t["owner"].tag == W + "ins" else None, "reason": t["reason"]} for t in p["tokens"]]} for p in paragraphs], "existing_citations": sum("ZOTERO_ITEM" in f["code"] for f in fields), "existing_bibliographies": sum("ZOTERO_BIBL" in f["code"] for f in fields), "markers_any_revision": marker_scan(document), "edge_whitespace_issues": _edge_whitespace(parts), "result": "INSPECTED"}


def _source_hyperlinks(paragraph, entry):
    """Return only fully covered, simple complex HYPERLINK source fields."""
    selected = []
    for field in paragraph.get("fields", []):
        if not (field["start"] < entry["end"] and entry["start"] < field["end"]):
            continue
        match = re.fullmatch(r'\s*HYPERLINK\s+"(https?://[^"\s]+)"(?:\s+\\h)?\s*', field["code"], re.I)
        require(match and field["depth"] == 0 and not field["nested"], "Range touches a protected non-HYPERLINK or nested field: " + entry["id"])
        require(field["paragraph"] == field["end_paragraph"] and entry["start"] <= field["start"] < field["end"] <= entry["end"], "Range only partly covers a HYPERLINK field: " + entry["id"])
        require(all(not len(n) and _owner(paragraph["node"], n) is paragraph["node"] for n in field["controls"]), "HYPERLINK controls cross an unsupported owner: " + entry["id"])
        field["hyperlink_url"] = match.group(1)
        selected.append(field)
    return selected


def _hyperlink_record(field):
    return {"paragraph": field["paragraph"], "start": field["start"], "end": field["end"], "code_sha256": digest(field["code"].encode()), "url": field["hyperlink_url"]}


def _link_identity(value):
    value = str(value).strip()
    if value.lower().startswith("doi:"):
        value = value[4:].strip()
    if value.startswith("10."):
        return ("doi", value.lower())
    parsed = urlsplit(value)
    if parsed.netloc.lower() in ("doi.org", "dx.doi.org"):
        return ("doi", parsed.path.lstrip("/").lower())
    return ("url", parsed.scheme.lower(), parsed.netloc.lower(), parsed.path, parsed.query, parsed.fragment)


def _range(paragraph, entry):
    start, end = entry["start"], entry["end"]
    require(type(start) is int and type(end) is int and 0 <= start < end <= len(paragraph["text"]), "Invalid exact replacement range: " + entry["id"])
    require(paragraph["text"][start:end] == entry["text"], "Source range text differs: " + entry["id"])
    touched = [t for t in paragraph["tokens"] if t["start"] < end and start < t["end"]]
    require(not any(t["owner"] is not None and t["owner"].tag == W + "ins" for t in touched), "New citation is inside existing tracked revision: " + entry["id"])
    links = _source_hyperlinks(paragraph, entry)
    require(touched and not any(t["reason"] and not (t["reason"] == "field_or_nontext" and t["node"].tag == W + "t" and t.get("field") is not None and any(t["field"] is f for f in links)) for t in touched), "Range touches protected or unsupported content: " + entry["id"])
    owner = touched[0]["owner"]
    require(all(t["owner"] is owner for t in touched), "Range crosses revision owner boundaries: " + entry["id"])
    allowed_controls = {n for field in links for n in field["controls"]}
    require(not any(start < pos < end and node not in allowed_controls for pos, _, node in paragraph["events"]), "Range crosses a field, equation, or drawing: " + entry["id"])
    return owner, touched[0], touched[-1]


def _preferences(parts):
    if "docProps/custom.xml" not in parts:
        return None
    root = _xml(parts["docProps/custom.xml"])
    props = [e for e in root if re.fullmatch(r"ZOTERO_PREF_?\d+", e.get("name", ""))]
    if not props:
        return None
    props.sort(key=lambda e: int(re.search(r"\d+$", e.get("name")).group()))
    value = "".join("".join(e.itertext()) for e in props)
    _preference_style(value)
    return value


def _preference_style(value):
    if value.lstrip().startswith("{"):
        data = json.loads(value)
        style = data.get("style", {}).get("styleID")
        field_type = data.get("prefs", {}).get("fieldType")
    else:
        data = _xml(value.encode())
        require(data.tag == "data" and data.find("style") is not None, "Unrecognized native Zotero preferences")
        style = data.find("style").get("id")
        ft = data.find("prefs/pref[@name='fieldType']")
        field_type = ft.get("value") if ft is not None else None
    require(field_type == "Field" and style, "Native citation work requires Word Fields and a Zotero style")
    return style


def _scientific_separator(paragraph, entry):
    start = entry["start"]
    leading = len(entry["text"]) - len(entry["text"].lstrip(_boundaries.HORIZONTAL))
    end = start + leading
    text = paragraph["text"]
    left = len(text[:end].rstrip(_boundaries.HORIZONTAL))
    if left == end or not left:
        return None
    token = next((t for t in paragraph["tokens"] if t["start"] < left <= t["end"]), None)
    if not token or token["field"] is not None or token["node"].tag != W + "t":
        return None
    for t in paragraph["tokens"]:
        if left < t["end"] and t["start"] < end:
            if t["node"].tag != W + "t" or (t["start"] < start and (t["reason"] or t["owner"] is not paragraph["node"])):
                return None
    vertical = token["node"].getparent().find(W + "rPr/" + W + "vertAlign")
    raised = vertical is not None and vertical.get(W + "val") == "superscript"
    if _boundaries.scientific_end(text, left, raised):
        return {"text": text[left:end], "inside": leading}


def build_plan(document, entries, library_snapshot, *, style_config=None, preferences_seed=None):
    parts, _ = _pack(document)
    root = _xml(parts[MAIN])
    paragraphs, fields = _inventory(root)
    existing_payloads = _citation_payloads(fields)
    require(isinstance(entries, list) and (entries or existing_payloads), "No existing or planned citations")
    snap, items = _snapshot(library_snapshot) if library_snapshot else ({}, {})
    require(not entries or library_snapshot, "New fields require verified library items")
    require(not marker_scan(document, parts), "Input contains unresolved artificial markers, including revision history")
    prepared, occupied, ids = [], {}, set()
    source_hash = digest(Path(document).read_bytes())
    for raw in entries:
        e = copy.deepcopy(raw)
        require(isinstance(e, dict) and isinstance(e.get("id"), str) and e["id"] and e["id"] not in ids, "Entry IDs must be nonempty and unique")
        ids.add(e["id"])
        require(e.get("story", MAIN) == MAIN, "Hybrid writing currently supports the main story only")
        e["story"] = MAIN
        require(type(e.get("paragraph")) is int and 1 <= e["paragraph"] <= len(paragraphs), "Unknown source paragraph")
        p = paragraphs[e["paragraph"] - 1]
        owner, first, last = _range(p, e)
        e.pop("scientific_separator", None)
        edge = _scientific_separator(p, e)
        if edge:
            require(edge["inside"] < len(e["text"]), "Citation span consists only of whitespace")
            e["scientific_separator"] = edge
        spans = occupied.setdefault(e["paragraph"], [])
        require(not any(e["start"] < b and a < e["end"] for a, b in spans), "Overlapping citation ranges")
        spans.append((e["start"], e["end"]))
        require(isinstance(e.get("uris"), list) and e["uris"], "Each citation group needs verified real item URIs")
        e["uris"] = [canonical_uri(u) for u in e["uris"]]
        require(len(e["uris"]) == len(set(e["uris"])), "A citation group repeats the same item")
        require(all(u in items for u in e["uris"]), "Citation identity absent from bound current library")
        options = e.get("citation_options", {})
        require(isinstance(options, dict) and set(options) <= set(e["uris"]), "Unknown citation-option item")
        for opt in options.values():
            require(isinstance(opt, dict) and set(opt) <= OPTIONS, "Unsupported citation options")
            require(all(isinstance(v, (str, bool)) for v in opt.values()), "Invalid citation-option value")
        e["display"] = e.get("display", "1")
        require(isinstance(e["display"], str) and e["display"] and len(e["display"]) <= 1000 and not MARKER.search(e["display"]), "Invalid provisional display")
        e["paragraph_sha256"] = digest(p["text"].encode())
        e["citation_id"] = uuid.uuid5(uuid.NAMESPACE_URL, source_hash + ":" + e["id"] + ":" + ",".join(e["uris"])).hex[:12]
        e["revision_owner"] = dict(owner.attrib) if owner.tag == W + "ins" else None
        e["item_data"] = {u: copy.deepcopy(items[u]) for u in e["uris"]}
        links = _source_hyperlinks(p, e)
        targets = {_link_identity(v) for item in e["item_data"].values() for k, v in item.items() if k.lower() in ("doi", "url") and isinstance(v, str)}
        # A publisher PII and the DOI returned for that exact alternate identifier name the same work.
        aliases = (style_config or {}).get("link_dois", {}) if isinstance(style_config, dict) else {}
        targets.update(_link_identity(url) for url, doi in aliases.items() if ("doi", doi) in targets)
        require(all(_link_identity(f["hyperlink_url"]) in targets for f in links), "Source HYPERLINK target differs from the bound reference metadata: " + e["id"])
        e["source_hyperlink_fields"] = [_hyperlink_record(f) for f in links]
        prepared.append(e)
    existing = {m.group(1) for f in fields if (m := re.search(r'"citationID"\s*:\s*"([^"]+)"', f["code"]))}
    require(not (existing & {e["citation_id"] for e in prepared}), "Generated citation identity collides with existing field")
    prefs = _preferences(parts)
    seed_hash = None
    if preferences_seed:
        seed_parts, _ = _pack(preferences_seed)
        seed_prefs = _preferences(seed_parts)
        require(seed_prefs, "Preferences seed has no genuine native Zotero preferences")
        seed_hash = digest(Path(preferences_seed).read_bytes())
        if prefs:
            require(_preference_style(prefs) == _preference_style(seed_prefs), "Preferences seed and source use different styles")
        elif seed_prefs.lstrip().startswith("{"):
            data = json.loads(seed_prefs)
            data["sessionID"] = uuid.uuid4().hex[:8]
            prefs = json.dumps(data, ensure_ascii=False)
        else:
            data = _xml(seed_prefs.encode())
            session = data.find("session")
            if session is not None:
                session.set("id", uuid.uuid5(uuid.NAMESPACE_URL, source_hash).hex[:8])
            prefs = E.tostring(data, encoding="unicode")
    config = _object(style_config) if style_config else None
    if config and "csl" in config:
        config = config["csl"]
    if config:
        require(isinstance(config.get("style_id"), str) and config["style_id"].startswith("http://www.zotero.org/styles/"), "Bind the actual target CSL style")
    target_style = config.get("style_id") if config else None
    if not target_style and prefs:
        target_style = _preference_style(prefs)
    plan = {"schema": PLAN, "writer_version": VERSION, "source_document": str(Path(document).resolve()), "source_sha256": source_hash, "library_id": snap.get("library_id"), "library_snapshot_sha256": digest(snap), "library_items_sha256": snap.get("items_sha256"), "entries": sorted(prepared, key=lambda e: (e["paragraph"], e["start"])), "preferences_xml": prefs, "preferences_seed_sha256": seed_hash, "style_id": target_style, "existing_citation_count": sum("ZOTERO_ITEM" in f["code"] for f in fields), "existing_bibliography_count": sum("ZOTERO_BIBL" in f["code"] for f in fields), "native_validation_required": True, "result": "PLANNED_PENDING_NATIVE_VALIDATION"}
    if not plan["existing_bibliography_count"]:
        slot, paragraph = _bibliography_slot(root, paragraphs, (config or {}).get("bibliography_paragraph"))
        plan.update(bibliography_slot=slot, bibliography_paragraph=paragraph)
    plan["plan_sha256"] = digest(plan)
    return plan


def _validate_plan(document, plan, library_snapshot=None):
    plan = _object(plan)
    require(plan.get("schema") == PLAN and plan.get("writer_version") in (VERSION, "2.4.2", "2.4.1", "2.4.0-draft", "2.3.1"), "Unsupported hybrid plan")
    expected = plan.get("plan_sha256")
    require(expected == digest({k: v for k, v in plan.items() if k != "plan_sha256"}), "Plan was changed after binding")
    require(plan.get("source_sha256") == digest(Path(document).read_bytes()), "Source changed after plan creation")
    require(str(Path(document).resolve()) == plan.get("source_document"), "Source path differs from bound plan")
    if library_snapshot is not None:
        snap, _ = _snapshot(library_snapshot)
        require(digest(snap) == plan["library_snapshot_sha256"], "Library snapshot changed after planning")
    return plan


def _path(owner, leaf):
    result = []
    while leaf is not owner:
        parent = leaf.getparent()
        require(parent is not None, "Invalid owner path")
        result.append(parent.index(leaf))
        leaf = parent
    return list(reversed(result))


def _shallow(e):
    n = E.Element(e.tag, attrib=dict(e.attrib), nsmap=e.nsmap)
    n.text, n.tail = e.text, e.tail
    return n


def _nonempty(e):
    if e is None:
        return None
    if e.tag == W + "t" and not (e.text or ""):
        return None
    if e.tag == W + "r" and not any(c.tag != W + "rPr" for c in e):
        return None
    if e.tag == W + "hyperlink" and not len(e):
        return None
    return e


def _split(node, path, offset):
    """Split only the boundary branch; no original revision wrapper is published twice."""
    if not path:
        require(node.tag == W + "t", "Boundary must be a text node")
        left, right = copy.deepcopy(node), copy.deepcopy(node)
        left.text, right.text = (node.text or "")[:offset], (node.text or "")[offset:]
        for n in (left, right):
            if n.text and (n.text[0].isspace() or n.text[-1].isspace()):
                n.set(XML_SPACE, "preserve")
        return _nonempty(left), _nonempty(right)
    index = path[0]
    require(index < len(node), "Boundary path invalid after split")
    left, right = _shallow(node), _shallow(node)
    left.tail = None
    a, b = _split(node[index], path[1:], offset)
    for child in list(node)[:index]:
        left.append(copy.deepcopy(child))
    if a is not None:
        left.append(a)
    if b is not None:
        right.append(b)
    for child in list(node)[index + 1:]:
        right.append(copy.deepcopy(child))
    if node.tag == W + "r":
        pr = node.find(W + "rPr")
        if pr is not None and right.find(W + "rPr") is None:
            right.insert(0, copy.deepcopy(pr))
    return _nonempty(left), _nonempty(right)


def _make_run(kind, value=None, properties=None):
    run = E.Element(W + "r")
    if properties is not None:
        run.append(copy.deepcopy(properties))
    if kind in ("begin", "separate", "end"):
        E.SubElement(run, W + "fldChar").set(W + "fldCharType", kind)
    else:
        child = E.SubElement(run, W + kind)
        child.set(XML_SPACE, "preserve")
        child.text = value
    return run


def _field(entry):
    items = []
    for uri in entry["uris"]:
        data = copy.deepcopy(entry["item_data"][uri])
        # Abstracts/annotations are not reference metadata; repeating them in
        # every Word field substantially increases manuscript/update cost.
        data.pop("abstract", None)
        data.pop("annote", None)
        data["id"] = uri
        item = {"id": uri, "uris": [uri], "itemData": data}
        item.update(entry.get("citation_options", {}).get(uri, {}))
        items.append(item)
    payload = {"citationID": entry["citation_id"], "properties": {"unsorted": False, "formattedCitation": entry["display"], "plainCitation": entry["display"], "noteIndex": 0}, "citationItems": items, "schema": "https://github.com/citation-style-language/schema/raw/master/csl-citation.json"}
    instruction = " ADDIN ZOTERO_ITEM CSL_CITATION " + json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
    return [_make_run("begin"), _make_run("instrText", instruction), _make_run("separate"), _make_run("t", entry["display"]), _make_run("end")]


def _replace(paragraph, entry):
    owner, first, last = _range(paragraph, entry)
    links = _source_hyperlinks(paragraph, entry)
    require([_hyperlink_record(f) for f in links] == entry.get("source_hyperlink_fields", []), "Bound source HYPERLINK fields changed")
    if links:
        # Remove exactly the reviewed link's field controls, retaining its
        # displayed runs and all adjacent material until the mapped replace.
        # Removing display text alone would leave orphan field instructions.
        for field in links:
            for node in field["controls"]:
                parent = node.getparent()
                parent.remove(node)
                if parent.tag == W + "r" and _nonempty(parent) is None:
                    parent.getparent().remove(parent)
        local, _ = _inventory(paragraph["node"])
        require(len(local) == 1, "Unexpected paragraphs inside HYPERLINK owner")
        paragraph = local[0]
        owner, first, last = _range(paragraph, entry)
    if edge := entry.get("scientific_separator"):
        inside = edge["inside"]
        pos = entry["start"] + inside
        left = pos - len(edge["text"])
        require(paragraph["text"][left:pos] == edge["text"], "Scientific separator differs from the bound original")
        for token in paragraph["tokens"]:
            if token["start"] < pos and left < token["end"] and token["node"].tag == W + "t":
                token["node"].set(XML_SPACE, "preserve")
        if inside:
            entry = dict(entry, start=pos, text=entry["text"][inside:])
            owner, first, last = _range(paragraph, entry)
    start_path, end_path = _path(owner, first["node"]), _path(owner, last["node"])
    prefix, suffix = _split(owner, end_path, entry["end"] - last["start"])
    require(prefix is not None, "Empty range prefix")
    prefix, discarded = _split(prefix, start_path, entry["start"] - first["start"])
    before = list(prefix) if prefix is not None else []
    after = list(suffix) if suffix is not None else []
    for c in list(owner):
        owner.remove(c)
    field = _field(entry)
    owner.extend(before + field + after)
    return field[0]


def _set_preferences(parts, value, *, replace=False):
    previous = _preferences(parts)
    if not value or previous == value or (previous and not replace):
        return set()
    changed = {"docProps/custom.xml"}
    name = "docProps/custom.xml"
    root = _xml(parts[name]) if name in parts else E.Element(CP + "Properties", nsmap={None: CP[1:-1], "vt": VT[1:-1]})
    if replace:
        for prop in list(root):
            if re.fullmatch(r"ZOTERO_PREF_?\d+", prop.get("name", "")):
                root.remove(prop)
    pid = max([int(e.get("pid", "1")) for e in root] + [1]) + 1
    # Word native custom-property strings are split into bounded chunks.
    for i, start in enumerate(range(0, len(value), 250), 1):
        prop = E.SubElement(root, CP + "property", fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}", pid=str(pid), name="ZOTERO_PREF_" + str(i))
        E.SubElement(prop, VT + "lpwstr").text = value[start:start + 250]
        pid += 1
    parts[name] = E.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    types = _xml(parts["[Content_Types].xml"])
    if not any(e.get("PartName") == "/" + name for e in types):
        E.SubElement(types, CT + "Override", PartName="/" + name, ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml")
        parts["[Content_Types].xml"] = E.tostring(types, xml_declaration=True, encoding="UTF-8", standalone=True)
        changed.add("[Content_Types].xml")
    relname = "_rels/.rels"
    rels = _xml(parts[relname]) if relname in parts else E.Element(PKGREL + "Relationships", nsmap={None: PKGREL[1:-1]})
    reltype = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties"
    if not any(e.get("Type") == reltype for e in rels):
        ids = {e.get("Id") for e in rels}
        i = 1
        while "rId" + str(i) in ids:
            i += 1
        E.SubElement(rels, PKGREL + "Relationship", Id="rId" + str(i), Type=reltype, Target=name)
        parts[relname] = E.tostring(rels, xml_declaration=True, encoding="UTF-8", standalone=True)
        changed.add(relname)
    return changed


def _semantic(e):
    if not isinstance(e.tag, str):
        return None
    attrs = {k: v for k, v in e.attrib.items() if not (k.startswith(W) and k[len(W):].startswith("rsid")) and not k.endswith(("}paraId", "}textId"))}
    return [e.tag, attrs, e.text or "", [v for c in e if (v := _semantic(c)) is not None]]


def _relationship_map(parts, part):
    name = posixpath.join(posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels") if part else "_rels/.rels"
    result = {}
    if name not in parts:
        return result
    for node in _xml(parts[name]):
        rid, target = node.get("Id"), node.get("Target")
        require(rid and rid not in result and target and isinstance(node.get("Type"), str) and node.get("Type"), "Malformed package relationship: " + name)
        mode = node.get("TargetMode", "Internal")
        if mode == "Internal":
            target = posixpath.normpath(posixpath.join(posixpath.dirname(part), target))
            require(not target.startswith(("../", "/")) and target in parts, "Missing or unsafe relationship target: " + name)
        result[rid] = {"Type": node.get("Type"), "Target": target, "TargetMode": mode, "extra": {k: v for k, v in node.attrib.items() if k not in ("Id", "Type", "Target", "TargetMode")}}
    return result


def _resolved_root(parts, part):
    root = _xml(parts[part])
    rels = _relationship_map(parts, part)
    for node in root.iter():
        for key, value in list(node.attrib.items()):
            if key.startswith(R) and value:
                require(value in rels, "Unresolved relationship reference in " + part)
                node.set(key, json.dumps(rels[value], sort_keys=True, separators=(",", ":")))
    return root


def _protected(parts):
    root = _resolved_root(parts, MAIN)
    tags = {M + "oMath", W + "drawing", W + "pict", W + "object", W + "fldSimple", W + "pPr",
            W + "sectPr", W + "tblPr", W + "tblGrid", W + "trPr", W + "tcPr",
            W + "ins", W + "moveTo", W + "del", W + "moveFrom"}
    values = {}
    for tag in tags:
        values[tag] = [_semantic(e) for e in root.iter(tag)]
    values["revision_owners"] = [[e.tag, dict(e.attrib)] for e in root.iter() if e.tag in (W + "ins", W + "moveTo", W + "del", W + "moveFrom")]
    values["property_revisions"] = [_semantic(e) for e in root.iter() if _revisions.is_property_history(e)]
    values["media_and_embeddings"] = {n: digest(b) for n, b in parts.items() if n.startswith(("word/media/", "word/embeddings/"))}
    values["table_topology"] = [[len(tr.findall(W + "tc")) for tr in tbl.findall(W + "tr")] for tbl in root.iter(W + "tbl")]
    return values


def _link_marks(props):
    if props is not None:
        underline, color = props.find(W + "u"), props.find(W + "color")
        if (underline is not None and underline.get(W + "val") == "single"
                and color is not None and color.get(W + "val", "").upper() == "0563C1"):
            return underline, color
    return ()


def _ordinary_format(paragraph, excluded, relationships=None, link_starts=()):
    """Compare actual run properties on every remaining ordinary character.

    Adjacent runs may split/merge without changing this stream.  Citation
    result runs are excluded because the native processor owns their format.
    Paragraph properties, revision history and media are checked separately.
    """
    result = []
    edges = set()
    for offset, length in link_starts:
        require(offset >= length and paragraph["text"][offset-length:offset].isspace(),
                "Changed citation spacing extends outside the bound whitespace")
        edges.update(range(offset-length, offset))
    for token in paragraph["tokens"]:
        run = token["node"].getparent()
        while run is not None and run.tag != W + "r":
            run = run.getparent()
        pr = run.find(W + "rPr") if run is not None else None
        fingerprint = digest(_semantic(pr)) if pr is not None else digest(None)
        edge_fingerprint = fingerprint
        if edges and _link_marks(pr):
            stripped = copy.deepcopy(pr)
            for mark in _link_marks(stripped):
                stripped.remove(mark)
            edge_fingerprint = digest(_semantic(stripped))
        lineage = digest([[e.tag, dict(e.attrib)] for e in token["node"].iterancestors() if e.tag in (W + "ins", W + "del", W + "moveTo", W + "moveFrom")])
        links = []
        for node in token["node"].iterancestors():
            if node.tag == W + "hyperlink":
                attrs = dict(node.attrib)
                if R + "id" in attrs:
                    rid = attrs[R + "id"]
                    require(relationships is not None and rid in relationships, "Ordinary hyperlink has an unresolved relationship")
                    attrs[R + "id"] = relationships[rid]
                links.append(attrs)
        link_fingerprint = digest(links)
        for offset, char in enumerate(token["text"], token["start"]):
            if not any(a <= offset < b for a, b in excluded):
                result.append((char, edge_fingerprint if offset in edges else fingerprint, lineage, link_fingerprint))
    return result


def _package_metadata_protection(a, b, plan, native):
    def custom_values(parts):
        return {e.get("name"): _semantic(e) for e in _xml(parts["docProps/custom.xml"])} if "docProps/custom.xml" in parts else {}
    ac, bc = custom_values(a), custom_values(b)
    original = {k: v for k, v in ac.items() if not k.startswith("ZOTERO_PREF")}
    require(all(bc.get(k) == v for k, v in original.items()), "Unrelated custom document property changed")
    require({k: v for k, v in bc.items() if not k.startswith("ZOTERO_PREF")} == original, "Unplanned custom property added")
    prefs = _preferences(b)
    if plan.get("preferences_xml") and not native:
        require(prefs == plan["preferences_xml"], "Written native preferences differ from reviewed plan")
    if native and plan.get("style_id"):
        require(prefs and _preference_style(prefs) == plan["style_id"], "Saved native target style differs from reviewed plan")
    if not native:
        for name, attribute, permitted in (("[Content_Types].xml", "PartName", "/docProps/custom.xml"), ("_rels/.rels", "Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties")):
            before = [_semantic(e) for e in _xml(a[name]) if e.get(attribute) != permitted] if name in a else []
            after = [_semantic(e) for e in _xml(b[name]) if e.get(attribute) != permitted] if name in b else []
            require(before == after, "Unrelated package relationship/content type changed")


def _word_save_view(root):
    """Canonicalize only Word save representations, on comparison trees."""
    mc = "{http://schemas.openxmlformats.org/markup-compatibility/2006}"
    for holder in root.iter():
        prefixes = holder.get(mc + "Ignorable")
        if not prefixes:
            continue
        namespaces = set()
        for node in holder.iter():
            for name in (node.tag, *node.attrib):
                if isinstance(name, str) and name.startswith("{"):
                    namespaces.add(name[1:name.index("}")])
            # Compatibility rules can reference an otherwise unused namespace by prefix.
            for key in ("Requires", mc + "PreserveElements", mc + "PreserveAttributes", mc + "ProcessContent"):
                for token in node.get(key, "").split():
                    namespaces.add(node.nsmap.get(token.split(":")[0]))
        active = sorted({holder.nsmap[prefix] for prefix in prefixes.split()
                         if prefix in holder.nsmap and holder.nsmap[prefix] in namespaces})
        unknown = [prefix for prefix in prefixes.split() if prefix not in holder.nsmap]
        if active or unknown:
            holder.set(mc + "Ignorable", " ".join(active + unknown))
        else:
            del holder.attrib[mc + "Ignorable"]
    for node in list(root.iter()):
        if isinstance(node.tag, str) and node.tag.startswith(W) and len(node) and node.text and not node.text.strip():
            node.text = None
        if node.tag in (W + "proofErr", W + "lastRenderedPageBreak") and not len(node):
            parent = node.getparent()
            parent.remove(node)
            if (node.tag == W + "lastRenderedPageBreak" and parent.tag == W + "r"
                    and all(key.startswith(W + "rsid") for key in parent.attrib)
                    and all(c.tag == W + "rPr" for c in parent)
                    and parent.find(".//" + W + "rPrChange") is None):
                parent.getparent().remove(parent)
        elif node.tag == W + "rFonts":
            for slot in ("ascii", "hAnsi", "eastAsia", "cs"):
                if node.get(W + slot, "").casefold() in ("等线", "dengxian"):
                    node.set(W + slot, "DengXian")
            if node.get(W + "hint") == "eastAsia":
                run = next(node.iterancestors(W + "r"), None)
                if run is not None:
                    leaves = [child for child in run if child.tag != W + "rPr"]
                    if leaves and all(c.tag in (W + "t", W + "delText") and
                                      re.fullmatch(r"[A-Za-z0-9 ()/.,-]+", c.text or "") for c in leaves):
                        del node.attrib[W + "hint"]
            if not node.attrib and not len(node):
                node.getparent().remove(node)
    for props in list(root.iter(W + "rPr")):
        if not props.attrib and not len(props):
            props.getparent().remove(props)
    # A split run can be folded only within the same structural owner and with identical properties.
    for parent in root.xpath(".//*[w:r]", namespaces={"w": W[1:-1]}):
        previous = key_before = None
        for run in list(parent):
            props = run.find(W + "rPr")
            leaves = [c for c in run if c is not props]
            if (run.tag != W + "r" or not leaves or leaves[0].tag not in (W + "t", W + "delText") or
                    any(c.tag != leaves[0].tag or set(c.attrib) - {XML_SPACE} for c in leaves)):
                previous = key_before = None
                continue
            attrs = {k: v for k, v in run.attrib.items() if not k.startswith(W + "rsid")}
            key = (attrs, _semantic(props) if props is not None else None, leaves[0].tag)
            text = "".join(c.text or "" for c in leaves)
            if previous is not None and key == key_before:
                previous.text = (previous.text or "") + text
                parent.remove(run)
            else:
                previous = leaves[0]
                previous.text = text
                previous.attrib.pop(XML_SPACE, None)
                for extra in leaves[1:]:
                    run.remove(extra)
                key_before = key
    return root


def _native_part_semantic(root, part):
    """Discard only observed save-session metadata, never document content."""
    root = copy.deepcopy(root)
    removable = {W + "rsid"}
    if part == "word/settings.xml":
        removable |= {W + "rsids", "{http://schemas.microsoft.com/office/word/2010/wordml}docId"}
    elif part == "docProps/core.xml":
        removable |= {"{http://purl.org/dc/terms/}modified", "{http://schemas.openxmlformats.org/package/2006/metadata/core-properties}lastModifiedBy", "{http://schemas.openxmlformats.org/package/2006/metadata/core-properties}revision"}
    elif part == "docProps/app.xml":
        ns = "{http://schemas.openxmlformats.org/officeDocument/2006/extended-properties}"
        removable |= {ns + n for n in ("Pages", "Words", "Characters", "CharactersWithSpaces", "Paragraphs", "Lines", "TotalTime", "Application", "AppVersion")}
    for node in list(root.iter()):
        if (part == "word/settings.xml" and node.tag == W + "proofState" and not len(node)
                and not (node.text or "").strip() and set(node.attrib) <= {W + "spelling", W + "grammar"}
                and all(v in ("clean", "dirty") for v in node.attrib.values())):
            node.getparent().remove(node)
            continue
        if node is not root and node.tag in removable:
            node.getparent().remove(node)
        if part == "word/numbering.xml":
            node.attrib.pop("{http://schemas.microsoft.com/office/word/2016/wordml/cid}durableId", None)
    return _semantic(_word_save_view(root))


def _exclusive_bibliography_styles(parts, styles, candidates):
    eligible = set(candidates)
    refs = {W + n for n in ("pStyle", "rStyle", "tblStyle", "basedOn", "link", "next", "styleLink", "numStyleLink")}
    for style in styles:
        own = style.get(W + "styleId")
        for ref in style.iter():
            if ref.tag in refs and ref.get(W + "val") != own:
                eligible.discard(ref.get(W + "val"))
    for name, data in parts.items():
        if not eligible:
            break
        if name.startswith("word/") and name.endswith(".xml") and name not in (MAIN, "word/styles.xml"):
            for ref in _xml(data).iter():
                if ref.tag in refs:
                    eligible.discard(ref.get(W + "val"))
    return eligible


def _native_package_protection(a, b, plan, bibliography_styles, differences=None):
    """Fail closed on non-target stories, styles, assets and package graph."""
    expected = dict(a)
    _set_preferences(expected, _preferences(b))
    changed = sorted(set(expected) ^ set(b))
    if changed:
        detail = "Native save added or removed parts: " + ", ".join(changed)
        if differences is None:
            raise ValueError(detail)
        differences.append(detail)
    reviewed_links = {_link_identity(value) for entry in plan["entries"] for item in entry["item_data"].values() for key, value in item.items() if key.lower() in ("doi", "url") and isinstance(value, str)}
    reviewed_links.update(_link_identity(f["url"]) for e in plan["entries"] for f in e.get("source_hyperlink_fields", []))
    def check_part(name):
        if name in (MAIN, "docProps/custom.xml") or expected[name] == b[name]:
            return
        if name.endswith(".rels"):
            owner = "" if name == "_rels/.rels" else posixpath.join(posixpath.dirname(posixpath.dirname(name)), posixpath.basename(name)[:-5])
            before = {digest(v): v for v in _relationship_map(expected, owner).values()}
            after = {digest(v): v for v in _relationship_map(b, owner).values()}
            require(not (after.keys() - before.keys()), "Native relationship target added or changed: " + name)
            removed = [before[k] for k in before.keys() - after.keys()]
            require(all(v["Type"].endswith("/hyperlink") and v["TargetMode"] == "External" and _link_identity(v["Target"]) in reviewed_links for v in removed), "Unrelated native relationship removed: " + name)
            return
        if name == "[Content_Types].xml":
            def content_types(parts):
                manifest = _xml(parts[name])
                defaults = {e.get("Extension").lower(): e.get("ContentType") for e in manifest.findall(CT + "Default")}
                overrides = {e.get("PartName").lstrip("/"): e.get("ContentType") for e in manifest.findall(CT + "Override")}
                return {n: overrides.get(n, defaults.get(n.rsplit(".", 1)[-1].lower())) for n in parts if n != name}
            require(content_types(expected) == content_types(b), "Native package content types changed")
            return
        if name == "word/styles.xml":
            ar, br = _xml(expected[name]), _xml(b[name])
            old = {e.get(W + "styleId"): e for e in ar if e.tag == W + "style"}
            new = {e.get(W + "styleId"): e for e in br if e.tag == W + "style"}
            require(old.keys() <= new.keys(), "Native save removed an existing style")
            exclusive = _exclusive_bibliography_styles(expected, ar, bibliography_styles)
            for key, style in old.items():
                if _native_part_semantic(style, name) == _native_part_semantic(new[key], name):
                    continue
                title = style.find(W + "name")
                require(key in exclusive and title is not None and title.get(W + "val", "").casefold() == "bibliography"
                        and all(n.get(W + "type") == "paragraph" and n.get(W + "default") in (None, "0", "false", "off")
                                for n in (style, new[key])), "Native save changed an existing style: " + key)
            for key in new.keys() - old.keys():
                style = new[key]
                title = style.find(W + "name")
                require(key in bibliography_styles and title is not None and title.get(W + "val") == "Bibliography" and style.get(W + "type") == "paragraph" and style.get(W + "default") in (None, "0", "false", "off"), "Unplanned native style added")
            for root in (ar, br):
                for node in list(root):
                    if node.tag == W + "style":
                        root.remove(node)
            require(_native_part_semantic(ar, name) == _native_part_semantic(br, name), "Native style defaults or latent styles changed")
            return
        if name.endswith(".xml"):
            require(_native_part_semantic(_resolved_root(expected, name), name) == _native_part_semantic(_resolved_root(b, name), name), "Unrelated native XML part changed: " + name)
        else:
            require(expected[name] == b[name], "Unrelated native binary part changed: " + name)

    for name in sorted(expected.keys() & b.keys()):
        try:
            check_part(name)
        except ValueError as exc:
            if differences is None:
                raise
            differences.append(name + ": " + str(exc))


def _ordinary_blank(paragraph):
    """A bibliography placeholder cannot hide an object or structural break."""
    if paragraph["text"] or paragraph["events"]:
        return False
    if any(node.tag in (W + "ins", W + "del", W + "moveTo", W + "moveFrom", W + "sectPr") for node in paragraph["node"].iter()):
        return False
    for child in paragraph["node"]:
        if child.tag == W + "pPr":
            if any(node.tag == W + "sectPr" for node in child.iter()):
                return False
        elif child.tag == W + "r":
            if any(node.tag not in (W + "rPr", W + "t") or (node.tag == W + "t" and node.text) for node in child):
                return False
        else:
            return False
    return True


def _bibliography_slot(root, paragraphs, requested):
    body = root.find(W + "body")
    rows = {p["node"]: (i, p) for i, p in enumerate(paragraphs, 1) if p["node"].getparent() is body}
    headings = [node for node, (_, p) in rows.items()
                if p["text"].strip().casefold() in ("references", "bibliography", "literature cited", "参考文献")]
    note = re.compile(r"^(?:disclaimer(?:\s*/\s*publisher['’]s note)?|publisher['’]s note)\s*(?:[:.：—–-]|$)", re.I)
    notes = {node for node, (_, p) in rows.items() if note.match(p["text"].strip())}
    section = body.find(W + "sectPr")
    end = body.index(section) if section is not None else len(body)
    require(len(headings) <= 1 or requested is not None, "Select the intended References section")
    if requested is not None:
        require(type(requested) is int and 1 <= requested <= len(paragraphs), "Invalid bibliography paragraph")
        row = paragraphs[requested - 1]
        require(row["node"] in rows and _ordinary_blank(row), "Reviewed bibliography anchor must be an empty body paragraph")
        index = body.index(row["node"])
        add_heading = row["node"].getprevious() not in headings
    elif headings:
        index = body.index(headings[0]) + 1
        after = index
        while after < end and body[after] in rows and _ordinary_blank(rows[body[after]][1]):
            after += 1
        require(after == end or body[after] in notes, "Resolve existing References content before adding a dynamic bibliography")
        add_heading = False
    else:
        index = min((body.index(node) for node in notes), default=end)
        if index and body[index - 1] in rows and _ordinary_blank(rows[body[index - 1]][1]):
            index -= 1
        add_heading = True
    row = rows.get(body[index]) if index < end else None
    require(row is not None or index == end, "Bibliography must be in the main document body")
    first = row[0] if row else len(paragraphs) + 1
    return {"index": index, "use_blank": bool(row and _ordinary_blank(row[1])), "heading": add_heading}, first + int(add_heading)


def _bibliography_nodes(body, slot):
    nodes = []
    if slot["heading"]:
        heading = E.Element(W + "p")
        heading.append(_make_run("t", "References"))
        nodes.append(heading)
    if slot["use_blank"]:
        paragraph = copy.deepcopy(body[slot["index"]])
        row, _ = _inventory(paragraph)
        require(len(row) == 1 and _ordinary_blank(row[0]), "Planned bibliography paragraph changed")
    else:
        paragraph = E.Element(W + "p")
    code = ' ADDIN ZOTERO_BIBL {"uncited":[],"omitted":[],"custom":[]} CSL_BIBLIOGRAPHY '
    paragraph.extend([_make_run("begin"), _make_run("instrText", code), _make_run("separate"),
                      _make_run("t", "Zotero bibliography"), _make_run("end")])
    nodes.append(paragraph)
    return nodes


def _bibliography_span_owned(paragraphs, field):
    """Do not erase material beside a field while comparing its paragraphs."""
    require(not any(token["text"] and token.get("field") is not field for paragraph in paragraphs for token in paragraph["tokens"]), "Unowned active text outside the native bibliography field")
    begin, end = field["controls"][0], field["controls"][-1]
    inside, seen_begin, seen_end = False, False, False
    glyphs = OBJECTS | {W + name for name in ("sym", "noBreakHyphen", "softHyphen", "ruby", "tab", "br", "cr", "ptab", "footnoteReference", "endnoteReference", "commentReference", "separator", "continuationSeparator", "delText", "delInstrText")}
    revisions = {W + name for name in ("ins", "del", "moveTo", "moveFrom", "moveFromRangeStart", "moveFromRangeEnd", "moveToRangeStart", "moveToRangeEnd", "cellIns", "cellDel", "cellMerge", "numberingChange")}
    def visit(node):
        nonlocal inside, seen_begin, seen_end
        if node is begin:
            inside, seen_begin = True, True
        revision = _revisions.is_revision(node)
        require(inside or not (node.tag in glyphs or revision), "Unowned object, glyph, or history outside the native bibliography field")
        for child in node:
            visit(child)
        if node is end:
            inside, seen_end = False, True
    for paragraph in paragraphs:
        visit(paragraph["node"])
    require(seen_begin and seen_end and not inside, "Unbound native bibliography paragraph span")


def audit(source, target, plan, *, allow_native_resave=False, cleaned_spaces=None, saved_copy=False):
    plan = _validate_plan(source, plan)
    if cleaned_spaces is None:
        cleaned_spaces = {}
        if allow_native_resave:
            seed, _, _ = _seed_parts(source, plan)
            _, seed_fields = _inventory(_xml(seed[MAIN]))
            cleaned_spaces = _clean_converted_spaces(_citation_payloads(seed_fields), plan)
    a, _ = _pack(source); b, _ = _pack(target)
    _package_metadata_protection(a, b, plan, allow_native_resave)
    aroot, broot = _xml(a[MAIN]), _xml(b[MAIN])
    bp, bf = _inventory(broot)
    native_bibs = [f for f in bf if "ZOTERO_BIBL" in f["code"]]
    slot = plan.get("bibliography_slot")
    expected_paragraph = plan.get("bibliography_paragraph")
    if saved_copy and not slot and not plan["existing_bibliography_count"] and len(native_bibs) == 1:
        # An older plan left only an empty anchor. A manual save can contain the new backmatter heading.
        source_rows, _ = _inventory(aroot)
        first = native_bibs[0]["paragraph"]
        if first > 1 and bp[first - 2]["text"].strip().casefold() == "references":
            location, paragraph = _bibliography_slot(aroot, source_rows, None)
            if location["heading"] and paragraph == first:
                slot, expected_paragraph = location, paragraph
    if plan.get("append_bibliography_paragraph") and not slot:
        _append_bibliography_paragraph(aroot)
    ap, af = _inventory(aroot)
    converted = []
    for entry in plan["entries"]:
        links = _source_hyperlinks(ap[entry["paragraph"] - 1], entry)
        require([_hyperlink_record(f) for f in links] == entry.get("source_hyperlink_fields", []), "Source link conversion differs from bound plan")
        converted.extend(links)
    def unrelated_fields(fields, omit):
        return [[f["code"], f["depth"], [_semantic(n) for n in f["controls"]]] for f in fields if "ZOTERO_" not in f["code"] and not any(f is old for old in omit)]
    require(unrelated_fields(af, converted) == unrelated_fields(bf, []), "Unplanned non-Zotero field changed or source HYPERLINK controls remain")
    source_payloads = _citation_payloads(af)
    actual_payloads = _citation_payloads(bf)
    require(set(actual_payloads) == set(source_payloads) | {e["citation_id"] for e in plan["entries"]}, "Unplanned citation added, removed, or relabelled")
    for cid, (data, _) in source_payloads.items():
        after_data = actual_payloads[cid][0]
        def identities(payload):
            return [{"uris": sorted(canonical_uri(u) for u in item["uris"]), "options": {k: item[k] for k in OPTIONS if k in item}} for item in payload["citationItems"]]
        require(identities(data) == identities(after_data), "Existing citation composition/options changed")
    # Normalize only the bound heading and field back to their original body position.
    source_bibs = [f for f in af if "ZOTERO_BIBL" in f["code"]]
    if slot:
        require(not source_bibs and len(native_bibs) == 1, "Expected exactly one planned bibliography")
    else:
        require(len(native_bibs) == len(source_bibs) or (allow_native_resave and not source_bibs and len(native_bibs) == 1), "Unplanned bibliography field added or removed")
    generated_bibliography = None
    bibliography_styles = set()
    if (slot or allow_native_resave) and native_bibs:
        require(len(native_bibs) == 1, "Native output has multiple bibliography fields")
        bib = native_bibs[0]
        begin, end = bib["paragraph"], bib["end_paragraph"]
        if slot:
            require(begin == expected_paragraph, "Bibliography moved from its planned paragraph")
            source_body = aroot.find(W + "body")
            expected = _bibliography_nodes(source_body, slot)
            originals = [source_body[slot["index"]]] if slot["use_blank"] else []
        elif source_bibs:
            require(len(source_bibs) == 1 and source_bibs[0]["paragraph"] == begin, "Native bibliography moved away from its original owner")
            source_end = source_bibs[0]["end_paragraph"]
            originals = [ap[i - 1]["node"] for i in range(begin, source_end + 1)]
        else:
            require(1 <= begin <= len(ap), "Native bibliography starts outside bound source paragraphs")
            require(begin == plan.get("bibliography_paragraph", begin), "Native bibliography moved away from the planned anchor")
            require(_ordinary_blank(ap[begin - 1]), "Native bibliography requires an ordinary empty source paragraph")
            originals, source_end = [ap[begin - 1]["node"]], begin
        if not slot:
            require(not any(begin <= e["paragraph"] <= source_end for e in plan["entries"]), "Bibliography overlaps a planned citation paragraph")
        span = [bp[i - 1] for i in range(begin, end + 1)]
        _bibliography_span_owned(span, bib)
        nodes = [paragraph["node"] for paragraph in span]
        bibliography_styles = {e.get(W + "val") for node in nodes for e in node.iter() if e.tag in (W + "pStyle", W + "rStyle")}
        parent = nodes[0].getparent()
        position = parent.index(nodes[0])
        require(list(parent)[position:position + len(nodes)] == nodes, "Bibliography crosses structural owners")
        if slot:
            require(parent is broot.find(W + "body") and position == slot["index"] + int(slot["heading"]),
                    "Bibliography moved from its bound body position")
            for node in nodes:
                _field_content_only(node)
            if not allow_native_resave:
                require(len(nodes) == 1 and _semantic(nodes[0]) == _semantic(expected[-1]), "Planned bibliography field differs")
            else:
                require(bib["display"].strip() not in ("", "Zotero bibliography"), "Bibliography is still a placeholder")
            if slot["heading"]:
                position -= 1
                heading = parent[position]
                if saved_copy:
                    _field_content_only(heading)
                    require(heading.tag == W + "p" and all(n.tag in (W + "pPr", W + "r") for n in heading)
                            and all(n.tag in (W + "rPr", W + "t") for r in heading.findall(W + "r") for n in r)
                            and "".join(n.text or "" for n in heading.iter(W + "t")).strip().casefold() == "references",
                            "Unexpected content at the planned References heading")
                else:
                    require(_semantic(heading) == _semantic(expected[0]), "Planned References heading changed")
        stop = parent.index(nodes[-1]) + 1
        parent[position:stop] = [copy.deepcopy(node) for node in originals]
        bp, normalized_fields = _inventory(broot)
        # Coordinates for the citation comparison must be from this normalized
        # tree, while the bibliography itself was checked from real target data.
        bf = normalized_fields
        bibliography_styles -= {e.get(W + "val") for e in broot.iter() if e.tag in (W + "pStyle", W + "rStyle")}
        if allow_native_resave:
            generated_bibliography = {"paragraph": begin, "end_paragraph": end, "display_sha256": digest(bib["display"].encode())}
    wanted = {e["citation_id"]: e for e in plan["entries"]}
    new = {}
    for field in bf:
        if "ZOTERO_ITEM" not in field["code"]:
            continue
        match = re.search(r"CSL_CITATION\s*(\{.*)", field["code"], re.S)
        require(match, "Malformed target citation field")
        data = json.loads(match.group(1))
        cid = data.get("citationID")
        if cid in wanted:
            require(cid not in new, "Duplicate target citation identity")
            expected = wanted[cid]
            items = data["citationItems"]
            require(len(items) == len(expected["uris"]) and all(
                uri in [canonical_uri(u) for u in item["uris"]]
                and {k: item[k] for k in OPTIONS if k in item} == expected.get("citation_options", {}).get(uri, {})
                for uri, item in zip(expected["uris"], items)), "Target citation item order, identities or options differ: " + cid)
            new[cid] = field
    require(set(new) == set(wanted), "Target is missing a planned citation")
    require(len(ap) == len(bp), "Hybrid insertion changed paragraph count")
    bypara = {}
    for entry in plan["entries"]:
        bypara.setdefault(entry["paragraph"], []).append(entry)
    normalized_payloads = _citation_payloads(bf)
    source_relationships, target_relationships = _relationship_map(a, MAIN), _relationship_map(b, MAIN)
    for i, (srcp, dstp) in enumerate(zip(ap, bp), 1):
        before = srcp["text"]
        after = dstp["text"]
        source_ranges = []
        for entry in bypara.get(i, []):
            edge = entry.get("scientific_separator", {})
            cid = entry["citation_id"]
            start = entry["start"] + edge.get("inside", 0)
            source_ranges.append((start, entry["end"], cid))
            if edge:
                pos = new[cid]["start"]
                require(new[cid]["paragraph"] == i and pos >= len(edge["text"]) and
                        dstp["text"][pos-len(edge["text"]):pos] == edge["text"],
                        "Scientific superscript separator changed before citation " + cid)
        for cid, (_, f) in source_payloads.items():
            if f["paragraph"] == i:
                require(f["end_paragraph"] == i, "Existing citation spans paragraphs; separate native review required")
                source_ranges.append((f["start"], f["end"], cid))
        target_ranges = [(f["start"], f["end"], cid) for cid, (_, f) in normalized_payloads.items() if f["paragraph"] == i]
        for start, end, cid in sorted(source_ranges, reverse=True):
            before = before[:start] + "\x00" + cid + "\x00" + before[end:]
        for start, end, cid in sorted(target_ranges, reverse=True):
            after = after[:start] + "\x00" + cid + "\x00" + after[end:]
        require(before == after, "Unplanned active text change in paragraph " + str(i))
        source_excluded = [(a, b) for a, b, _ in source_ranges]
        target_excluded = [(a, b) for a, b, _ in target_ranges]
        source_edges = [(start, cleaned_spaces[cid]) for start, _, cid in source_ranges if cid in cleaned_spaces]
        target_edges = [(start, cleaned_spaces[cid]) for start, _, cid in target_ranges if cid in cleaned_spaces]
        require(_ordinary_format(srcp, source_excluded, source_relationships, source_edges) == _ordinary_format(dstp, target_excluded, target_relationships, target_edges), "Unplanned ordinary run formatting change or revision/link ownership change in paragraph " + str(i))
    normalized_b = dict(b)
    normalized_b[MAIN] = E.tostring(broot)
    require(_protected(a) == _protected(normalized_b), "Revision, math, drawing, table, paragraph-property, or media protection failed")
    if allow_native_resave:
        _native_package_protection(a, b, plan, bibliography_styles)
    existing = [f["code"] for f in af if "ZOTERO_ITEM" in f["code"] or "ZOTERO_BIBL" in f["code"]]
    if not allow_native_resave:
        require(all(code in [f["code"] for f in bf] for code in existing), "An existing citation/bibliography instruction changed")
        allowed = {MAIN, "docProps/custom.xml", "[Content_Types].xml", "_rels/.rels"}
        require(set(a) <= set(b) and not (set(b) - set(a) - allowed), "Unexpected package part added or removed")
        require(all(a[n] == b[n] for n in a if n not in allowed), "Unrelated ZIP part changed")
    require(not marker_scan(target, b), "Artificial citation markers remain, including deleted revisions")
    return {"schema": "zotero-word.hybrid-protection.v1", "result": "PROTECTION_VERIFIED_PENDING_NATIVE_VALIDATION", "source_sha256": digest(Path(source).read_bytes()), "target_sha256": digest(Path(target).read_bytes()), "plan_sha256": plan["plan_sha256"], "citation_count": len(new), "ordinary_text_preserved": True, "protected_objects_preserved": True, "original_revisions_preserved": True, "markers_any_revision": [], "allow_native_resave": bool(allow_native_resave), "native_bibliography_observed": generated_bibliography, "native_validation_verified": False}


def _append_bibliography_paragraph(root):
    body = root.find(W + "body")
    section = body.find(W + "sectPr")
    body.insert(body.index(section) if section is not None else len(body), E.Element(W + "p"))


def check_clean(path):
    parts, _ = _pack(path)
    markers = {W + name for name in ("comment", "commentRangeStart", "commentRangeEnd", "commentReference", "delText", "delInstrText")}
    for name, data in parts.items():
        if name.startswith("word/") and name.endswith(".xml"):
            for node in _xml(data).iter():
                require(node.tag not in markers and not _revisions.is_revision(node),
                        "Comments or tracked revisions remain in " + name)


def _seed_parts(document, plan, *, clean_link_spaces=False):
    parts, comment = _pack(document)
    root = _xml(parts[MAIN])
    paragraphs, _ = _inventory(root)
    # Validate all original coordinates before the first mutation.  Multiple
    # groups in one paragraph subsequently change its full-text hash while
    # leaving the earlier source offsets valid under reverse-order writing.
    for entry in plan["entries"]:
        p = paragraphs[entry["paragraph"] - 1]
        require(digest(p["text"].encode()) == entry["paragraph_sha256"], "Bound source paragraph changed")
        _range(p, entry)
    # Work backwards so all source offsets remain valid within a paragraph.
    previous_paragraph = None
    for entry in sorted(plan["entries"], key=lambda e: (e["paragraph"], e["start"]), reverse=True):
        p = paragraphs[entry["paragraph"] - 1]
        if previous_paragraph == entry["paragraph"]:
            # A cross-paragraph field needs its original stack context.
            try:
                local, _ = _inventory(p["node"])
                require(len(local) == 1, "Nested paragraphs")
                p = local[0]
                for f in p["fields"]:
                    f["paragraph"] = f["end_paragraph"] = entry["paragraph"]
            except ValueError:
                paragraphs, _ = _inventory(root)
                p = paragraphs[entry["paragraph"] - 1]
        first_run = _replace(p, entry)
        if clean_link_spaces and _converted_link(entry):
            _clean_link_space(first_run)
        previous_paragraph = entry["paragraph"]
    if slot := plan.get("bibliography_slot"):
        body = root.find(W + "body")
        index = slot["index"]
        body[index:index + int(slot["use_blank"])] = _bibliography_nodes(body, slot)
    elif plan.get("append_bibliography_paragraph"):
        _append_bibliography_paragraph(root)
    parts[MAIN] = E.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    prefs_changes = _set_preferences(parts, plan.get("preferences_xml"))
    return parts, comment, prefs_changes


def _publish(document, output, parts, comment, verify, *, resume):
    if resume and output.exists():
        existing, existing_comment = _pack(output)
        require(existing == parts and comment == existing_comment,
                "An existing artifact differs from the planned result: " + str(output))
        return verify(output)
    output.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp = tempfile.mkstemp(prefix=".hybrid-", suffix=".docx", dir=output.parent)
    os.close(fd)
    try:
        with ZipFile(document) as src, ZipFile(tmp, "w") as dst:
            written = set()
            for info in src.infolist():
                dst.writestr(info, parts[info.filename]); written.add(info.filename)
            for name in parts.keys() - written:
                dst.writestr(name, parts[name])
            dst.comment = comment
        protection = verify(tmp)
        # Atomic no-clobber publication, including a concurrent output creator.
        if os.name == "nt":
            os.rename(tmp, output)  # On Windows this fails if the destination exists.
        else:
            os.link(tmp, output)
    finally:
        Path(tmp).unlink(missing_ok=True)
    return protection


def apply(document, plan, library_snapshot, output_document):
    plan = _validate_plan(document, plan, library_snapshot)
    require(not plan["entries"] or library_snapshot, "New fields require the bound library snapshot")
    return write_prepared(document, plan, output_document)


def write_prepared(document, plan, output_document, *, resume=False, for_native_run=False):
    output = Path(output_document).resolve()
    require(output != Path(document).resolve() and (resume or not output.exists()), "Output must be a new independent DOCX path")
    require(output.suffix.lower() == ".docx", "Hybrid output must be DOCX")
    parts, comment, prefs_changes = _seed_parts(document, plan, clean_link_spaces=for_native_run)
    if for_native_run and resume and output.exists() and _pack(output) != (parts, comment):
        # An older run could have published its prepared file just before recording the checkpoint.
        # Its seed used the same fields but deferred link-space cleanup until delivery.
        parts, comment, prefs_changes = _seed_parts(document, plan)
    def verify_prepared(path):
        if not for_native_run:
            return audit(document, path, plan)
        written, saved_comment = _pack(path)
        require(written == parts and saved_comment == comment, "The prepared working copy differs from the planned parts")
        return {"result": "PREPARED_BYTES_VERIFIED", "native_validation_verified": False}
    protection = _publish(document, output, parts, comment, verify_prepared, resume=resume)
    if for_native_run:
        remaining = ["Zotero refresh", "Word save and close"]
    else:
        remaining = ["Word opens without repair", "Zotero Refresh and target style", "Add/Edit Citation recovers true library items", "Bibliography generated by native plugin", "Save close reopen and field/library protection checks"]
    return {"schema": "zotero-word.hybrid-write.v1", "result": "PENDING_NATIVE_VALIDATION", "writer_version": VERSION, "source_document": str(Path(document).resolve()), "target_document": str(output), "source_sha256": plan["source_sha256"], "target_sha256": digest(output.read_bytes()), "plan_sha256": plan["plan_sha256"], "citation_count_inserted": len(plan["entries"]), "unique_item_count": len({u for e in plan["entries"] for u in e["uris"]}), "changed_parts": [MAIN] + sorted(prefs_changes), "protection": protection, "native_validation_required": remaining, "native_validation_verified": False}


def _field_runs(field):
    begin, end = field["controls"][0], field["controls"][-1]
    first, last = begin.getparent(), end.getparent()
    owner = first.getparent()
    require(not field["nested"] and owner is last.getparent() and field["paragraph"] == field["end_paragraph"],
            "Citation field crosses structural owners")
    require(all(n.tag == W + "rPr" for n in begin.itersiblings(preceding=True)) and end.getnext() is None,
            "Citation shares its boundary run with other content")
    runs = list(owner)[owner.index(first):owner.index(last) + 1]
    require(all(r.tag == W + "r" for r in runs), "Non-run content inside citation")
    for run in runs:
        _field_content_only(run)
    return owner, runs


def _field_content_only(node):
    protected = {W + name for name in ("sectPr", "bookmarkStart", "bookmarkEnd", "commentReference",
                 "commentRangeStart", "commentRangeEnd", "permStart", "permEnd", "sdt", "customXml")}
    require(not any(n.tag in OBJECTS or n.tag in protected or _revisions.is_revision(n) for n in node.iter()),
            "Field contains an object, anchor or revision that must remain in place")


def _clean_link_space(begin_run):
    previous = begin_run.getprevious()
    if previous is None or previous.tag != W + "r" or any(
            p.tag in (W + "hyperlink", W + "ins", W + "moveTo") for p in previous.iterancestors()):
        return False
    props = previous.find(W + "rPr")
    marks = _link_marks(props)
    text = [n for n in previous if n is not props]
    if not marks or not text or not all(n.tag == W + "t" and (n.text or "").isspace() for n in text):
        return False
    if any(_revisions.is_revision(n) for n in props.iter()):
        return False
    for mark in marks:
        props.remove(mark)
    return sum(len(n.text) for n in text)


def _converted_link(entry):
    return (entry.get("converted_link") or entry.get("source_hyperlink_fields")
            or re.search(r"https?://|10\.\d{4,9}/", entry["text"], re.I))


def _clean_converted_spaces(citations, plan):
    cleaned = {}
    for entry in plan["entries"]:
        if _converted_link(entry):
            cid = entry["citation_id"]
            spaces = _clean_link_space(citations[cid][1]["controls"][0].getparent())
            if spaces:
                cleaned[cid] = spaces
    return cleaned


def _candidate_differences(projected, native, plan, cleaned):
    """Find edits that projecting the rendered fields back onto the original would discard."""
    differences = []
    wanted = _resolved_root(projected, MAIN)
    actual = _resolved_root(native, MAIN)
    if cleaned:
        _, fields = _inventory(actual)
        _clean_converted_spaces(_citation_payloads(fields), plan)
    roots = [_word_save_view(root) for root in (wanted, actual)]
    paragraphs = [list(root.iter(W + "p")) for root in roots]
    if len(paragraphs[0]) != len(paragraphs[1]):
        differences.append("paragraph count changed outside the projected bibliography")
    for i, (left, right) in enumerate(zip(*paragraphs), 1):
        if _semantic(left) != _semantic(right):
            def words(node):
                return "".join(c.text or "" for c in node.iter() if c.tag in (W + "t", W + "delText"))
            detail = "text" if words(left) != words(right) else "format, revision or structure"
            differences.append("paragraph " + str(i) + " (" + detail + ")")
    for nodes in paragraphs:
        for paragraph in reversed(nodes):
            for child in list(paragraph):
                paragraph.remove(child)
    if _semantic(roots[0]) != _semantic(roots[1]):
        differences.append("document structure outside paragraphs")
    try:
        _package_metadata_protection(projected, native, plan, True)
    except ValueError as exc:
        differences.append(str(exc))
    _native_package_protection(projected, native, plan, set(), differences)
    if differences:
        raise ValueError("Additional changes in the saved working copy; original and candidate retained: "
                         + "; ".join(differences))


def publish_rendered(source, rendered, plan, output):
    """Project native fields onto the protected source and audit the exact published bytes."""
    plan = _validate_plan(source, plan)
    parts, comment, _ = _seed_parts(source, plan)
    native, _ = _pack(rendered)
    root = _xml(parts[MAIN])
    paragraphs, fields = _inventory(root)
    native_paragraphs, native_fields = _inventory(_xml(native[MAIN]))
    citations, fresh = _citation_payloads(fields), _citation_payloads(native_fields)
    require(citations.keys() == fresh.keys(), "Native citation IDs differ from the prepared copy")
    native_rels = _relationship_map(native, MAIN)
    old_rels = {digest(value): key for key, value in _relationship_map(parts, MAIN).items()}

    def copy_field_node(node):
        _field_content_only(node)
        new = copy.deepcopy(node)
        for child in new.iter():
            for key, rid in list(child.attrib.items()):
                if key.startswith(R):
                    original = old_rels.get(digest(native_rels.get(rid)))
                    require(original, "Rendered field references a resource absent from the protected source")
                    child.set(key, original)
        return new

    cleaned = _clean_converted_spaces(citations, plan)
    for cid, (_, field) in citations.items():
        owner, old = _field_runs(field)
        _, new = _field_runs(fresh[cid][1])
        index = owner.index(old[0])
        for r in old:
            owner.remove(r)
        for offset, r in enumerate(new):
            owner.insert(index + offset, copy_field_node(r))

    new_bibs = [f for f in native_fields if "ZOTERO_BIBL" in f["code"]]
    old_bibs = [f for f in fields if "ZOTERO_BIBL" in f["code"]]
    require(len(new_bibs) == 1 and len(old_bibs) <= 1, "Expected one native bibliography")
    bib = new_bibs[0]
    expected_position = old_bibs[0]["paragraph"] if old_bibs else plan["bibliography_paragraph"]
    require(bib["paragraph"] == expected_position, "Rendered bibliography moved from the bound paragraph")
    span = native_paragraphs[bib["paragraph"] - 1:bib["end_paragraph"]]
    _bibliography_span_owned(span, bib)
    if old_bibs:
        old = old_bibs[0]
        anchors = paragraphs[old["paragraph"] - 1:old["end_paragraph"]]
        _bibliography_span_owned(anchors, old)
        for paragraph in anchors:
            _field_content_only(paragraph["node"])
    else:
        anchors = [paragraphs[plan["bibliography_paragraph"] - 1]]
        require(_ordinary_blank(anchors[0]), "Planned bibliography anchor is no longer empty")
    for block in (span, anchors):
        container = block[0]["node"].getparent()
        start = container.index(block[0]["node"])
        require(list(container)[start:start + len(block)] == [p["node"] for p in block],
                "Bibliography crosses a structural boundary")
    parent = anchors[0]["node"].getparent()
    first = parent.index(anchors[0]["node"])
    for p in anchors:
        parent.remove(p["node"])
    for offset, p in enumerate(span):
        parent.insert(first + offset, copy_field_node(p["node"]))
    if "word/styles.xml" in parts and "word/styles.xml" in native:
        styles = _xml(parts["word/styles.xml"])
        old_styles = {e.get(W + "styleId"): e for e in styles if e.tag == W + "style"}
        used = {e.get(W + "val") for p in span for e in p["node"].iter() if e.tag in (W + "pStyle", W + "rStyle")}
        copied_span = set(parent[first:first + len(span)])
        outside = {n.get(W + "val") for n in root.iter() if n.tag in (W + "pStyle", W + "rStyle", W + "tblStyle")
                   and not any(ancestor in copied_span for ancestor in n.iterancestors())}
        exclusive = _exclusive_bibliography_styles(parts, styles, used - outside)
        changed = False
        for style in _xml(native["word/styles.xml"]):
            key = style.get(W + "styleId")
            if style.tag != W + "style" or key not in used:
                continue
            old = old_styles.get(key)
            if old is None:
                styles.append(copy.deepcopy(style))
            elif key in exclusive and _native_part_semantic(old, "word/styles.xml") != _native_part_semantic(style, "word/styles.xml"):
                styles.replace(old, copy.deepcopy(style))
            else:
                continue
            changed = True
        if changed:
            parts["word/styles.xml"] = E.tostring(styles)
    _set_preferences(parts, _preferences(native), replace=True)
    parts[MAIN] = E.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    _candidate_differences(parts, native, plan, cleaned)
    if output is None:  # A read-only inspection of an interrupted, still-open Word copy.
        return

    def verify(path):
        receipt = audit(source, path, plan, allow_native_resave=True, cleaned_spaces=cleaned)
        return {**receipt, "cleaned_link_spaces": cleaned}
    return _publish(source, output, parts, comment, verify, resume=True)

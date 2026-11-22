"""Classify Word revision history; accept table-property history only in a new copy.

This is not Word's Accept All Revisions. Text/row/cell insertion, deletion,
movement and other revision types require a separately reviewed native path.
"""
from __future__ import annotations

from collections import Counter
import hashlib
import io
import os
from pathlib import Path
from zipfile import ZipFile

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
TABLE_HISTORY = {"tblGridChange": "tblGrid", "tblPrExChange": "tblPrEx",
                 "tblPrChange": "tblPr", "trPrChange": "trPr", "tcPrChange": "tcPr"}
REVISION = {"ins", "del", "moveFrom", "moveTo", "moveFromRangeStart", "moveFromRangeEnd",
            "moveToRangeStart", "moveToRangeEnd", "cellIns", "cellDel", "cellMerge",
            "numberingChange", "delText", "delInstrText"}
REVISION |= {"customXml" + kind + "Range" + edge
             for kind in ("Ins", "Del", "MoveFrom", "MoveTo") for edge in ("Start", "End")}


def is_property_history(node):
    tag = node.tag
    return isinstance(tag, str) and tag.startswith(W) and (
        tag.endswith("PrChange") or tag[len(W):] in ("tblGridChange", "tblPrExChange"))


def is_revision(node):
    tag = node.tag
    return is_property_history(node) or (isinstance(tag, str) and tag.startswith(W)
                                         and tag[len(W):] in REVISION)


def is_historical(node):
    return node.tag in (W + "del", W + "moveFrom") or is_property_history(node)


def accept_table_history(source, output, *, authorization_evidence):
    """Remove authorized table property snapshots, preserving current properties.

    Refuse every other revision type, malformed placement and existing output.
    Track Changes settings are retained: this does not disable future tracking.
    """
    if not isinstance(authorization_evidence, str) or not authorization_evidence.strip():
        raise ValueError("Explicit authorization evidence is required")
    from lxml import etree as E

    src, dst = Path(source), Path(output)
    if src.is_symlink() or dst.is_symlink() or not src.is_file():
        raise ValueError("Source/output symlinks or missing input are unsupported")
    if src.resolve() == dst.resolve() or dst.exists():
        raise ValueError("Table history acceptance requires an unused independent output")
    if src.suffix.lower() != ".docx" or dst.suffix.lower() != ".docx":
        raise ValueError("Table history acceptance requires DOCX paths")
    raw = src.read_bytes()
    source_hash = hashlib.sha256(raw).hexdigest()
    removed, modified = Counter(), []
    with ZipFile(io.BytesIO(raw)) as archive:
        infos = archive.infolist()
        if len({i.filename for i in infos}) != len(infos) or sum(i.file_size for i in infos) > 256 * 1024 * 1024:
            raise ValueError("Duplicate parts or oversized DOCX package")
        if any(i.filename.startswith("_xmlsignatures/") for i in infos):
            raise ValueError("Signed DOCX requires a separately reviewed workflow")
        parts = {i.filename: archive.read(i) for i in infos}
        if "word/document.xml" not in parts or "[Content_Types].xml" not in parts:
            raise ValueError("Not a DOCX package")
        for name, data in parts.items():
            if not name.endswith(".xml"):
                continue
            if b"<!DOCTYPE" in data.upper() or b"<!ENTITY" in data.upper():
                raise ValueError("External XML declarations are unsupported: " + name)
            try:
                root = E.fromstring(data, E.XMLParser(resolve_entities=False, no_network=True, remove_blank_text=False))
            except E.XMLSyntaxError as exc:
                raise ValueError("Malformed XML in DOCX part: " + name) from exc
            nodes = [node for node in root.iter() if is_revision(node)]
            for node in nodes:
                local = node.tag[len(W):]
                parent = node.getparent()
                if local not in TABLE_HISTORY or parent is None or parent.tag != W + TABLE_HISTORY[local]:
                    raise ValueError("Unsupported revision needs native review: " + name + ":" + local)
            if nodes:
                for node in reversed(nodes):
                    node.getparent().remove(node)
                    removed[node.tag[len(W):]] += 1
                if any(is_revision(node) for node in root.iter()):
                    raise ValueError("Revision history remains after table acceptance: " + name)
                parts[name] = E.tostring(root, encoding="UTF-8", xml_declaration=True)
                modified.append(name)
        buffer = io.BytesIO()
        with ZipFile(buffer, "w") as result:
            result.comment = archive.comment
            for info in infos:
                result.writestr(info, parts[info.filename])
    if hashlib.sha256(src.read_bytes()).hexdigest() != source_hash:
        raise ValueError("Source changed before output publication")
    # O_EXCL also blocks a raced existing file, symlink or hardlink.
    fd = os.open(dst, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o600)
    try:
        with os.fdopen(fd, "wb") as stream:
            stream.write(buffer.getvalue())
    except BaseException:
        dst.unlink(missing_ok=True)
        raise
    return {"schema": "zotero-word.table-history-acceptance.v1", "result": "TABLE_PROPERTY_HISTORY_ACCEPTED",
            "source_sha256": source_hash, "output_sha256": hashlib.sha256(buffer.getvalue()).hexdigest(),
            "source": str(src.resolve()), "output": str(dst.resolve()), "removed": dict(removed),
            "modified_parts": modified, "authorization_evidence": authorization_evidence,
            "scope": "Table property history only; not general Accept All Revisions"}

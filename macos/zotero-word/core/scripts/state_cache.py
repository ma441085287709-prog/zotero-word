"""Run-local memo for saved DOCX inspection and pure expected layout only.

No cache values are imported from files. This is an optimization of trusted
functions within one process, never an acceptance receipt or a native-state
cache. Every lookup hashes the actual inputs and implementation files again.
"""
from __future__ import annotations

from collections import OrderedDict
import copy
import hashlib
import json
import math
from pathlib import Path
from zipfile import ZipInfo


class CacheInputChanged(ValueError):
    """An input changed while a pure result was being computed."""


def _sha(raw):
    return hashlib.sha256(raw).hexdigest()


def _json(value):
    return json.dumps(value, sort_keys=True, ensure_ascii=False,
                      separators=(",", ":"), allow_nan=False).encode("utf-8")


def _fingerprint(value):
    """Hash supported in-memory values without pickle, eval or custom hooks."""
    result = hashlib.sha256()

    def feed(item):
        kind = type(item)
        if item is None or kind in (str, bool, int, float):
            if kind is float and not math.isfinite(item):
                raise ValueError("Non-finite cache value")
            raw = _json(item)
            result.update(kind.__name__.encode() + b":" + str(len(raw)).encode() + b":" + raw)
        elif kind is bytes:
            result.update(b"bytes:" + str(len(item)).encode() + b":" + item)
        elif kind in (tuple, list):
            result.update(kind.__name__.encode() + b":" + str(len(item)).encode() + b":")
            for member in item:
                feed(member)
        elif kind is dict:
            if not all(type(key) is str for key in item):
                raise ValueError("Cache mappings require string keys")
            result.update(b"dict:" + str(len(item)).encode() + b":")
            for key in sorted(item):
                feed(key)
                feed(item[key])
        elif kind is ZipInfo:
            result.update(b"ZipInfo:")
            for key in ZipInfo.__slots__:
                feed(key)
                feed((True, getattr(item, key)) if hasattr(item, key) else (False,))
        else:
            raise ValueError("Unsupported cache value type: " + kind.__name__)

    feed(value)
    return result.hexdigest()


def _rule_binding(value):
    if type(value) is dict:
        return {"kind": "object", "sha256": _sha(_json(value))}
    path = Path(value).resolve()
    return {"kind": "file", "path": str(path), "sha256": _sha(path.read_bytes())}


class Memo:
    """Cache only trusted pure computations for the current process.

    inspection(path, inspector) calls inspector(path) on a miss.
    layout_expected(path, profile, role_map, compute) calls
    compute(path, profile, role_map) on a miss, preserving those arguments.
    Inputs may be Paths or the same JSON dictionaries supported by layout code.
    Separate CLI invocations deliberately start empty; summaries are not values.
    """

    def __init__(self, code_paths=None, *, max_entries=8):
        if type(max_entries) is not int or max_entries < 1:
            raise ValueError("max_entries must be a positive integer")
        self._code_paths = None if code_paths is None else tuple(sorted({Path(p).resolve() for p in code_paths}))
        if self._code_paths == ():
            raise ValueError("Explicit code_paths must not be empty")
        self._max_entries = max_entries
        self._entries = OrderedDict()
        self._functions = {}
        self._counts = {"hits": 0, "misses": 0, "invalidated": 0, "input_changed": 0}

    def _code(self):
        folder = Path(__file__).resolve().parent
        paths = self._code_paths
        if paths is None:
            # Include added/removed modules and the actual CSL computation code.
            paths = tuple(sorted([*folder.glob("*.py"), *folder.glob("*.js")]))
        return _sha(_json([[str(p), _sha(p.read_bytes())] for p in paths]))

    def _binding(self, slot, document, compute, profile=None, role_map=None):
        if not callable(compute):
            raise ValueError("A trusted computation function is required")
        path = Path(document).resolve()
        identity = id(compute)
        # Retain the callable, preventing reuse of its process-local identity.
        self._functions[identity] = compute
        value = {"slot": slot, "document": str(path),
                 "document_sha256": _sha(path.read_bytes()),
                 "code_sha256": self._code(), "function_identity": identity}
        if slot == "layout_expected":
            value["profile"] = _rule_binding(profile)
            value["role_map"] = _rule_binding(role_map)
        return value

    @staticmethod
    def _validate(slot, value, binding):
        if slot == "inspection":
            required = {"schema", "document", "document_sha256", "story_coverage", "issues",
                        "citation_count", "bibliography_count", "citations", "bibliographies",
                        "blocks", "preferences", "citation_cues", "media_sha256", "protected_content_sha256"}
            if (type(value) is not dict or not required <= value.keys()
                    or value.get("schema") != "zotero-word.inspection.v1"
                    or value.get("document_sha256") != binding["document_sha256"]
                    or value.get("document") != binding["document"]
                    or any(key in value for key in ("result", "status", "final_audit", "native_evidence"))):
                raise ValueError("Only a complete bound saved-DOCX inspection can be memoized")
            for key in ("citations", "bibliographies", "blocks", "issues", "citation_cues"):
                if type(value[key]) is not list:
                    raise ValueError("Invalid saved-DOCX inspection collection")
            if (type(value["citation_count"]) is not int or type(value["bibliography_count"]) is not int
                    or value["citation_count"] != len(value["citations"])
                    or value["bibliography_count"] != len(value["bibliographies"])):
                raise ValueError("Invalid saved-DOCX inspection counts")
        elif slot == "layout_expected":
            if type(value) is not tuple or len(value) != 7:
                raise ValueError("Only a pure expected-layout tuple can be memoized")
            raw, infos, before, expected, profile, mapping, summary = value
            if (type(raw) is not bytes or _sha(raw) != binding["document_sha256"]
                    or type(infos) is not list or not all(type(info) is ZipInfo for info in infos)
                    or any(type(part) is not dict for part in (before, expected, profile, mapping, summary))
                    or not before or before.keys() != expected.keys()
                    or not all(type(k) is str and type(v) is bytes for k, v in before.items())
                    or not all(type(k) is str and type(v) is bytes for k, v in expected.items())
                    or {info.filename for info in infos} != set(before)
                    or profile.get("schema") != "zotero-word.optional-layout-profile.v1"
                    or mapping.get("schema") != "zotero-word.optional-layout-role-map.v1"
                    or any(key in summary for key in ("result", "status", "native_evidence", "final_audit"))):
                raise ValueError("Invalid pure expected-layout cache value")
        else:
            raise ValueError("Unsupported cache slot")

    def _memo(self, slot, document, compute, profile=None, role_map=None):
        before = self._binding(slot, document, compute, profile, role_map)
        key = _sha(_json(before))
        entry = self._entries.get(key)
        if entry is not None:
            try:
                self._validate(slot, entry["value"], before)
                if entry["binding"] != before or entry["value_sha256"] != _fingerprint(entry["value"]):
                    raise ValueError("Changed run-local memo entry")
                result = copy.deepcopy(entry["value"])
            except (KeyError, TypeError, ValueError, AttributeError):
                del self._entries[key]
                self._counts["invalidated"] += 1
            else:
                if before != self._binding(slot, document, compute, profile, role_map):
                    self._counts["input_changed"] += 1
                    raise CacheInputChanged("Cache input changed during lookup; recheck saved state")
                self._entries.move_to_end(key)
                self._counts["hits"] += 1
                return result
        self._counts["misses"] += 1
        value = compute(document) if slot == "inspection" else compute(document, profile, role_map)
        if before != self._binding(slot, document, compute, profile, role_map):
            self._counts["input_changed"] += 1
            raise CacheInputChanged("Cache input changed during computation; recheck saved state")
        self._validate(slot, value, before)
        copied = copy.deepcopy(value)
        self._entries[key] = {"binding": before, "value_sha256": _fingerprint(copied), "value": copied}
        self._entries.move_to_end(key)
        while len(self._entries) > self._max_entries:
            self._entries.popitem(last=False)
        return copy.deepcopy(copied)

    def inspection(self, document, inspector):
        return self._memo("inspection", document, inspector)

    def layout_expected(self, document, profile, role_map, compute):
        return self._memo("layout_expected", document, compute, profile, role_map)

    def clear(self):
        self._entries.clear()
        self._functions.clear()

    def summary(self):
        return {"schema": "zotero-word.run-local-memo-summary.v1", **self._counts,
                "entries": len(self._entries), "value_storage": "current_process_memory_only",
                "acceptance_evidence": False, "persistent_values_loaded": False}

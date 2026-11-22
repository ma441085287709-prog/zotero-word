"""Bound, journalled library batches over Zotero's official local HTTP API.

This module never drives Word, changes API preferences, requests authorization,
executes a shell, or deletes a Zotero object.  HTTP success is not manuscript
completion evidence.  The host-dispatch route emits requests, never tool results.
"""
from __future__ import annotations

from contextlib import contextmanager
from datetime import datetime, timezone
import fcntl
import hashlib
import http.client
import importlib.util
import json
import os
from pathlib import Path
import re
import socket
import sys
import tempfile
from urllib.parse import urlencode, urlsplit


CONFIG = "zotero-word.library-adapter.v1"
HEALTH = "zotero-word.library-adapter-health.v1"
EXECUTION = "zotero-word.library-execution.v1"
JOURNAL = "zotero-word.library-execution-journal.v1"
HOST_REQUEST = "zotero-word.library-host-request.v1"
ACTIONS = frozenset(("lookup", "import_item", "add_membership", "remove_membership"))
MAX_BATCH = 50
MAX_RESPONSE = 32 * 1024 * 1024
KEY = re.compile(r"[A-Z0-9]{8}\Z")
LIBRARY = re.compile(r"(?:users/(?:local/[A-Za-z0-9]+|[0-9]+)|groups/[0-9]+)\Z")
NATIVE_NEXT = ("Library API access and Word/Zotero native integration are separate routes. "
               "Actively discover and connect the authorized local app/plugin or library channel; "
               "diagnose a disabled interface and prepare task-scoped connections under existing authority. "
               "For a genuinely missing approval, state the exact required setting/action, reuse prior consent, "
               "and continue after approval. API 401/403 does not by itself block Word native refresh. "
               "Observe actual library/collection state and run batch-status; never replay unknown writes.")


def _require(condition, message):
    if not condition:
        raise ValueError(message)


def _now():
    return datetime.now(timezone.utc).isoformat()


def _digest(value):
    return hashlib.sha256(json.dumps(value, sort_keys=True, ensure_ascii=False,
                                    separators=(",", ":")).encode()).hexdigest()


def _helper(name):
    spec = importlib.util.spec_from_file_location(
        "_zotero_word_executor_" + name, Path(__file__).with_name(name + ".py"))
    module = importlib.util.module_from_spec(spec)
    before = sys.dont_write_bytecode
    try:
        sys.dont_write_bytecode = True
        spec.loader.exec_module(module)
    finally:
        sys.dont_write_bytecode = before
    return module


def _config(value):
    _require(isinstance(value, dict) and value.get("schema") == CONFIG,
             "Use a zotero-word.library-adapter.v1 configuration")
    allowed = {"schema", "kind", "base_url", "library_id", "probe_item_uri", "server_id",
               "api_key_env", "state_dir", "timeout_seconds", "max_batch_size", "max_items",
               "test_double", "host_tools"}
    _require(not set(value) - allowed, "Unknown adapter configuration fields; shell/command runners are unsupported")
    cfg = dict(value)
    _require(cfg.get("kind") in ("zotero_local_api", "host_dispatch"), "Unsupported library adapter")
    _require(LIBRARY.fullmatch(cfg.get("library_id", "")), "Invalid bound library ID")
    core = _helper("zotero_docx")
    probe = core.canonical_uri(cfg.get("probe_item_uri"))
    _require(probe and core.library_of(probe) == cfg["library_id"],
             "Health needs an actual existing item URI in this bound library, not DOI/key inference")
    cfg["probe_item_uri"] = probe
    cfg["timeout_seconds"] = cfg.get("timeout_seconds", 15)
    _require(type(cfg["timeout_seconds"]) in (int, float) and 0.1 <= cfg["timeout_seconds"] <= 30,
             "Timeout must be 0.1–30 seconds")
    cfg["max_batch_size"] = cfg.get("max_batch_size", 20)
    _require(type(cfg["max_batch_size"]) is int and 1 <= cfg["max_batch_size"] <= MAX_BATCH,
             "Local API batches are limited to at most 50 objects")
    cfg["max_items"] = cfg.get("max_items", 20000)
    _require(type(cfg["max_items"]) is int and 1 <= cfg["max_items"] <= 100000,
             "Invalid local-library read bound")
    _require(type(cfg.get("test_double", False)) is bool, "test_double must be boolean")
    if cfg.get("api_key_env"):
        _require(re.fullmatch(r"[A-Z][A-Z0-9_]{1,99}", cfg["api_key_env"]),
                 "api_key_env names an existing environment variable; never store a key in config")
    if cfg.get("server_id"):
        _require(re.fullmatch(r"[A-Za-z0-9_-]{1,128}", cfg["server_id"]), "Invalid expected server ID")
    if cfg["kind"] == "zotero_local_api":
        p = urlsplit(cfg.get("base_url", "http://127.0.0.1:23119/api/"))
        _require(p.scheme == "http" and p.hostname in ("127.0.0.1", "localhost", "::1")
                 and not p.username and not p.password and not p.query and not p.fragment
                 and p.path == "/api/", "Only the loopback Zotero /api/ endpoint is supported")
        port = p.port or 80
        _require((cfg.get("test_double") and port != 23119) or (not cfg.get("test_double") and port == 23119),
                 "Production requires port 23119; TEST_DOUBLE requires a separate loopback test port")
        cfg["base_url"] = p.geturl()
    else:
        _require(not any(cfg.get(k) for k in ("base_url", "api_key_env", "server_id")),
                 "Host dispatch cannot smuggle an HTTP connection or credentials")
        tools = cfg.get("host_tools", {})
        _require(isinstance(tools, dict) and set(tools) <= ACTIONS
                 and all(isinstance(v, str) and re.fullmatch(r"[A-Za-z0-9_.:-]{1,160}", v) for v in tools.values()),
                 "host_tools must explicitly map whitelisted actions to real host tool names")
    return cfg


class AdapterError(Exception):
    def __init__(self, code, message):
        super().__init__(message)
        self.code = code


class LocalAPI:
    """One persistent connection; no redirects, arbitrary URLs, or implicit retry."""
    def __init__(self, cfg):
        self.cfg = cfg
        p = urlsplit(cfg["base_url"])
        self.connection = http.client.HTTPConnection(p.hostname, p.port, timeout=cfg["timeout_seconds"])
        self.server_id = cfg.get("server_id")
        self.secret = os.environ.get(cfg.get("api_key_env", ""), "")
        self.raw = []
        self.library_path = "users/0" if cfg["library_id"].startswith("users/local/") else cfg["library_id"]

    def close(self):
        self.connection.close()

    def _safe(self, value):
        if isinstance(value, dict):
            return {k: "[REDACTED]" if re.search(r"(?:api.?key|authorization|token|password|secret)", k, re.I)
                    else self._safe(v) for k, v in value.items()}
        if isinstance(value, list):
            return [self._safe(v) for v in value]
        return value.replace(self.secret, "[REDACTED]") if self.secret and isinstance(value, str) else value

    def request(self, method, endpoint="", *, query=None, payload=None, headers=None):
        _require(method in ("GET", "POST"), "Only whitelisted GET/POST routes exist; deletion is unsupported")
        # Endpoints are created by this module, never supplied by an operation.
        _require(endpoint == "" or re.fullmatch(r"(?:users/(?:0|[0-9]+)|groups/[0-9]+)/(?:items(?:/[A-Z0-9]{8})?|collections/[A-Z0-9]{8})", endpoint),
                 "Endpoint is outside the library action whitelist")
        _require(method == "GET" or endpoint == self.library_path + "/items",
                 "Writes may only POST a bounded item-create or partial membership batch")
        path = "/api/" + endpoint
        if query:
            path += "?" + urlencode(query)
        sent_headers = {"Zotero-API-Version": "3", "Accept": "application/json"}
        if self.server_id:
            sent_headers["Zotero-Server-ID"] = self.server_id
        if method != "GET":
            _require(self.server_id and self.secret, "A pre-existing local key and bound server ID are required")
            sent_headers.update({"Content-Type": "application/json", "Zotero-API-Key": self.secret})
        sent_headers.update(headers or {})
        body = json.dumps(payload, ensure_ascii=False, separators=(",", ":")).encode() if payload is not None else None
        record = {"method": method, "endpoint": path, "started_at": _now()}
        self.raw.append(record)
        try:
            self.connection.request(method, path, body=body, headers=sent_headers)
            response = self.connection.getresponse()
            raw = response.read(MAX_RESPONSE + 1)
            _require(len(raw) <= MAX_RESPONSE, "Local API response exceeds the configured safety bound")
            response_headers = {k.lower(): v for k, v in response.getheaders()}
            allowed = ("zotero-api-version", "zotero-server-id", "last-modified-version", "total-results", "retry-after", "content-type")
            record.update(http_status=response.status, headers={k: response_headers[k] for k in allowed if k in response_headers},
                          observed_at=_now())
            try:
                data = json.loads(raw) if raw else None
            except (UnicodeDecodeError, ValueError):
                data = raw.decode("utf-8", errors="replace")[:2000]
            record["body"] = self._safe(data)
            returned_id = response_headers.get("zotero-server-id")
            if self.server_id and returned_id and returned_id != self.server_id:
                raise AdapterError("INSTANCE_CHANGED", "Zotero instance changed; discard cached versions and re-observe")
            return response.status, response_headers, data
        except (OSError, http.client.HTTPException, socket.timeout) as exc:
            self.connection.close()
            record.update(transport_error=type(exc).__name__, observed_at=_now())
            raise AdapterError("TRANSPORT_UNKNOWN", "Connection failed or timed out; no automatic request retry") from exc

    def get(self, endpoint="", *, query=None):
        status, headers, data = self.request("GET", endpoint, query=query)
        if status != 200:
            code = {401: "READ_AUTH_REQUIRED", 403: "READ_ACCESS_DENIED", 404: "IDENTITY_NOT_FOUND",
                    412: "INSTANCE_OR_VERSION_CHANGED", 429: "BUSY_OR_RATE_LIMITED"}.get(status, "READ_FAILED")
            raise AdapterError(code, "Local API read failed (HTTP %s); no preference/permission changes were made" % status)
        return headers, data

    def item(self, key):
        _require(KEY.fullmatch(key), "Invalid item key")
        _, data = self.get(self.library_path + "/items/" + key, query={"include": "data,csljson"})
        return self.record(data)

    def record(self, value):
        _require(isinstance(value, dict) and isinstance(value.get("data"), dict), "Malformed local API item response")
        data = value["data"]
        key = value.get("key", data.get("key"))
        _require(isinstance(key, str) and KEY.fullmatch(key), "Local API returned an invalid item key")
        csl = value.get("csljson")
        if isinstance(csl, str):
            try:
                csl = json.loads(csl)
            except ValueError:
                csl = None
        if isinstance(csl, list) and len(csl) == 1:
            csl = csl[0]
        possible = [value.get("links", {}).get("alternate", {}).get("href")]
        if isinstance(csl, dict):
            possible.append(csl.get("id"))
        core = _helper("zotero_docx")
        uris = {core.canonical_uri(v) for v in possible if core.canonical_uri(v)}
        _require(len(uris) == 1, "API did not prove a unique Zotero item URI; never infer it from DOI/key")
        uri = uris.pop()
        _require(core.library_of(uri) == self.cfg["library_id"] and uri.endswith("/" + key),
                 "API returned another library/item identity")
        version = value.get("version", data.get("version"))
        _require(type(version) is int and version >= 0, "API response lacks an actual object version")
        return {"uri": uri, "key": key, "version": version, "data": data, "csljson": csl}

    def all_items(self):
        records, start, library_version = {}, 0, None
        while True:
            headers, values = self.get(self.library_path + "/items", query={"include": "data,csljson", "limit": 100, "start": start})
            _require(isinstance(values, list), "Local API did not return an item list")
            version = headers.get("last-modified-version")
            _require(version is not None and version.isdigit(), "Missing actual local library version")
            if library_version is not None and version != library_version:
                raise AdapterError("VERSION_CONFLICT", "Library changed during pagination; re-observe before writing")
            library_version = version
            for value in values:
                if isinstance(value, dict) and value.get("data", {}).get("itemType") in ("attachment", "note", "annotation"):
                    continue
                record = self.record(value)
                _require(record["uri"] not in records, "Duplicate item across API pages; re-observe")
                records[record["uri"]] = record
            start += len(values)
            _require(start <= self.cfg["max_items"], "Library exceeds the configured read bound; narrow the native workflow")
            total = headers.get("total-results")
            if len(values) < 100 or (total and total.isdigit() and start >= int(total)):
                break
        return records, int(library_version)


def _health(api):
    headers, _ = api.get()
    _require(headers.get("zotero-api-version") == "3", "Local API v3 was not actually confirmed")
    found_id = headers.get("zotero-server-id")
    if api.cfg.get("server_id") and api.cfg["server_id"] != found_id:
        raise AdapterError("INSTANCE_CHANGED", "Configured Zotero instance does not match the actual response")
    api.server_id = found_id
    record = api.item(api.cfg["probe_item_uri"].rsplit("/", 1)[1])
    _require(record["uri"] == api.cfg["probe_item_uri"], "Health probe returned another item identity")
    supported = bool(found_id)  # Header identifies the v10+ protocol, not authorization.
    return {"read_verified": True, "library_identity_verified": True, "library_id": api.cfg["library_id"],
            "probe_item_uri": record["uri"], "server_id": found_id,
            "local_write_protocol_supported": supported,
            "write_authorization": "CONFIGURED_UNVERIFIED" if supported and api.secret else "MISSING",
            "write_authorization_verified": False,
            "available_capabilities": ["local_library_lookup"]}


def health(config):
    """Actual bound read probe. Never POST /local/authorize or treat ping as health."""
    cfg = _config(config)
    result = {"schema": HEALTH, "observed_at": _now(), "adapter_kind": cfg["kind"],
              "test_evidence": "TEST_DOUBLE" if cfg.get("test_double") else "REAL_ADAPTER",
              "library_id": cfg["library_id"], "read_verified": False,
              "write_authorization_verified": False, "raw_responses": []}
    if cfg["kind"] == "host_dispatch":
        return {**result, "result": "HOST_DISPATCH_REQUIRED", "code": "CAPABILITY_UNVERIFIED",
                "next_action": "Have the host call its actually discovered read tool and verify the exact library/item URI; config is not health evidence."}
    api = LocalAPI(cfg)
    try:
        result.update(_health(api), result="READ_HEALTH_CONFIRMED")
    except (AdapterError, ValueError) as exc:
        result.update(result="CAPABILITY_MISSING", code=getattr(exc, "code", "IDENTITY_UNPROVEN"),
                      reason=str(exc), next_action=NATIVE_NEXT)
    finally:
        result["raw_responses"] = api.raw
        api.close()
    return result


def _observation(op, outcome, evidence, **values):
    return {"operation_id": op["operation_id"], "outcome": outcome,
            "evidence": evidence, "observed_at": _now(), **values}


def _observations(plan, entries):
    return {"schema": "zotero-word.batch-observations.v1",
            "plan_fingerprint": plan["plan_fingerprint"], "results": list(entries)}


@contextmanager
def _state(cfg, plan):
    """One executor batch lock; does not claim to lock Word, Zotero UI or sync."""
    value = cfg.get("state_dir")
    _require(isinstance(value, str) and Path(value).is_absolute(),
             "A durable absolute state_dir is required before dispatching writes")
    directory = Path(value)
    _require(not directory.is_symlink(), "Execution state directory cannot be a symlink")
    directory.mkdir(parents=True, exist_ok=True, mode=0o700)
    token = _digest([cfg.get("base_url"), cfg["library_id"]])
    descriptor = os.open(directory / (token + ".lock"), os.O_CREAT | os.O_RDWR | getattr(os, "O_NOFOLLOW", 0), 0o600)
    try:
        try:
            fcntl.flock(descriptor, fcntl.LOCK_EX | fcntl.LOCK_NB)
        except BlockingIOError as exc:
            raise AdapterError("EXECUTOR_BUSY", "Another executor batch holds this library lock; do not start a competing writer") from exc
        path = directory / (token + "-" + plan["plan_fingerprint"] + ".json")
        _require(not path.is_symlink(), "Execution journal cannot be a symlink")
        if path.exists():
            journal = json.loads(path.read_text())
            _require(journal.get("schema") == JOURNAL and journal.get("plan_fingerprint") == plan["plan_fingerprint"]
                     and journal.get("library_id") == cfg["library_id"], "Journal binding mismatch")
        else:
            journal = {"schema": JOURNAL, "plan_fingerprint": plan["plan_fingerprint"],
                       "library_id": cfg["library_id"], "test_evidence": "TEST_DOUBLE" if cfg.get("test_double") else "REAL_ADAPTER",
                       "server_id": None, "operations": {}}
        def save():
            handle, temporary = tempfile.mkstemp(prefix=".library-execution-", suffix=".json", dir=directory)
            try:
                with os.fdopen(handle, "w") as stream:
                    json.dump(journal, stream, ensure_ascii=False, sort_keys=True, indent=2)
                    stream.flush(); os.fsync(stream.fileno())
                os.replace(temporary, path)
                dir_fd = os.open(directory, os.O_RDONLY)
                try:
                    os.fsync(dir_fd)
                finally:
                    os.close(dir_fd)
            finally:
                if os.path.exists(temporary):
                    os.unlink(temporary)
        yield journal, save, path
    finally:
        os.close(descriptor)


def _merge_observations(plan, previous, journal):
    source = previous or _observations(plan, [])
    _require(source.get("schema") == "zotero-word.batch-observations.v1"
             and source.get("plan_fingerprint") == plan["plan_fingerprint"], "Observations belong to another batch plan")
    rows = {row["operation_id"]: row for row in source.get("results", [])}
    _require(len(rows) == len(source.get("results", [])), "Duplicate observation IDs")
    for oid, entry in journal.get("operations", {}).items():
        row = entry.get("observation")
        if entry.get("state") == "STARTED":
            row = {"operation_id": oid, "outcome": "unknown", "observed_at": entry["started_at"],
                   "evidence": "Durable journal records a started write without a definitive response; inspect actual state, never replay"}
        if row:
            existing = rows.get(oid)
            if existing and existing != row:
                _require(existing.get("observed_at", "") >= row.get("observed_at", "")
                         and (existing.get("outcome") != "not_applied" or existing.get("no_write_verified") is True),
                         "Caller observations cannot erase an unknown journaled write")
            else:
                rows[oid] = row
    return _observations(plan, rows.values())


def _hits(identifier, records):
    batch = _helper("library_batch")
    core = _helper("zotero_docx")
    return sorted(uri for uri, record in records.items()
                  if (core.doi(record["data"].get("DOI")) == identifier.get("DOI") if identifier.get("kind") == "doi"
                      else batch._url(record["data"].get("url")) == identifier.get("URL")))


def _metadata(op):
    value = op.get("metadata")
    _require(isinstance(value, dict) and isinstance(value.get("itemType"), str) and value.get("title"),
             "CAPABILITY_MISSING: local API import requires reviewed Zotero editable JSON metadata; DOI translation belongs to an actually available host tool")
    forbidden = {"key", "version", "deleted", "collections", "parentItem", "relations", "attachments", "notes", "dateAdded", "dateModified"}
    _require(not set(value) & forbidden and value["itemType"] not in ("attachment", "note", "annotation"),
             "Import payload cannot update an existing item, assign memberships, or add child objects")
    core = _helper("zotero_docx")
    batch = _helper("library_batch")
    identifier = op["identifier"]
    if identifier["kind"] == "doi":
        _require(core.doi(value.get("DOI")) == identifier["DOI"], "Reviewed import metadata DOI conflicts with the verified source clue")
    else:
        _require(batch._url(value.get("url")) == identifier["URL"], "Reviewed import metadata URL conflicts with the verified source clue")
    # Copy: do not mutate the bound plan.
    return {**value, "collections": []}


def _membership(api, op, known, binding):
    uri = op["uri"]
    core = _helper("zotero_docx")
    _require(uri in known and op.get("expected_record_sha256") == core.digest(known[uri])
             and op.get("expected_record_version") == known[uri].get("version"),
             "PLAN_PRECONDITION_CONFLICT: operation and its bound source record disagree")
    target = op["target_collection"]
    _require(isinstance(binding, dict) and target == {k: binding.get(k) for k in
             ("library_id", "collection_name", "collection_path", "collection_key")},
             "COLLECTION_CONFLICT: operation targets a collection other than its authorized plan")
    allowed = (set(binding["desired_members"]) - set(binding["current_members"]) if op["action"] == "add_membership"
               else set(binding["current_members"]) - set(binding["desired_members"]))
    _require(uri in allowed, "MEMBERSHIP_CONFLICT: operation is outside the authorized membership difference")
    record = api.item(uri.rsplit("/", 1)[1])
    _require(record["uri"] == uri, "Membership preflight changed item identity")
    expected_version = op.get("expected_record_version")
    _require(type(expected_version) is int and expected_version == record["version"],
             "VERSION_CONFLICT: re-read item/collection and rebuild; never substitute a new version automatically")
    expected = known[uri]
    csl = record["csljson"]
    _require(isinstance(csl, dict), "Live CSL metadata is required to prove the planned item has not changed")
    live_content = {**csl, "id": uri}
    before_content = {k: v for k, v in expected.items() if k not in ("version", "collections")}
    _require(live_content == before_content, "METADATA_CONFLICT: library metadata differs from the bound snapshot; preserve current state and rebuild")
    actual = record["data"].get("collections")
    before = expected.get("collections")
    _require(isinstance(actual, list) and isinstance(before, list) and set(actual) == set(before)
             and len(actual) == len(set(actual)) and all(KEY.fullmatch(k) for k in actual),
             "MEMBERSHIP_CONFLICT: fresh API-derived snapshot must expose every current membership")
    key = target.get("collection_key")
    _require(target.get("library_id") == api.cfg["library_id"] and isinstance(key, str) and KEY.fullmatch(key),
             "Membership target lacks a bound collection key")
    _, collection = api.get(api.library_path + "/collections/" + key)
    _require(isinstance(collection, dict) and collection.get("key", collection.get("data", {}).get("key")) == key
             and collection.get("data", {}).get("name") == target.get("collection_name"),
             "COLLECTION_CONFLICT: collection identity/name changed; re-observe instead of rebinding")
    changed = set(actual)
    if op["action"] == "add_membership":
        changed.add(key)
    else:
        changed.discard(key)
    return {"key": record["key"], "version": record["version"], "collections": sorted(changed)}


def _write_results(api, ops, payloads, data, http_status):
    rows = []
    for index, op in enumerate(ops):
        if http_status == 200 and isinstance(data, dict):
            maps = [data.get(k, {}) for k in ("successful", "success", "unchanged", "failed")]
            _require(all(isinstance(v, dict) for v in maps), "Malformed per-object write response")
            successful, success, unchanged, failed = maps
            index_text = str(index)
            positive = sum(index_text in m for m in maps[:3])
            if positive == 1 and index_text not in failed:
                result = next(m[index_text] for m in maps[:3] if index_text in m)
                key = result.get("key", result.get("data", {}).get("key")) if isinstance(result, dict) else result
                if not isinstance(key, str) or not KEY.fullmatch(key):
                    rows.append(_observation(op, "unknown", "Success response lacks a valid actual item key; inspect current library")); continue
                if op["action"] != "import_item" and key != payloads[index]["key"]:
                    rows.append(_observation(op, "unknown", "Tool response item identity disagrees with the membership target")); continue
                # Read after response: success text alone is never enough.
                try:
                    actual = api.item(key)
                    if op["action"] == "import_item":
                        _require(actual["uri"] in _hits(op["identifier"], {actual["uri"]: actual}), "Imported identity does not match DOI/URL")
                    else:
                        _require(set(actual["data"].get("collections", [])) == set(payloads[index]["collections"]),
                                 "Read-back membership differs, including other collections")
                    rows.append(_observation(op, "succeeded", "Per-object API response and a real read-back confirmed the planned library effect; final fresh snapshot/audit still required", uri=actual["uri"]))
                except (AdapterError, ValueError) as exc:
                    rows.append(_observation(op, "unknown", "API reported success but read-back was not conclusive: " + str(exc)))
                continue
            if index_text in failed and not positive:
                rows.append(_observation(op, "failed", "API reported a per-object failure; re-read actual state before any retry")); continue
        if http_status in (401, 403, 409, 412, 428, 429):
            rows.append(_observation(op, "failed", "API rejected the batch (HTTP %s); re-read state/version, do not automatically replay" % http_status))
        else:
            rows.append(_observation(op, "unknown", "Write result was missing, malformed or uncertain (HTTP %s); inspect actual state, never replay" % http_status))
    return rows


def _dispatch(api, operations, known, journal, save, snapshot_server_id, collection_binding):
    rows, sending, payloads = [], [], []
    action = operations[0]["action"]
    if action == "lookup":
        # A local library query is not the requested public metadata service.
        records, _ = api.all_items()
        for op in operations:
            query = op.get("query", {})
            candidates = _hits(query, records) if query.get("kind") in ("doi", "url") else []
            rows.append(_observation(op, "not_applied", "CAPABILITY_MISSING: this adapter only read local candidates; use a real public metadata lookup host tool", no_write_verified=True, candidate_uris=candidates))
        return rows
    if not api.server_id or not api.secret:
        return [_observation(op, "not_applied", "CAPABILITY_MISSING: local write protocol or pre-existing local write authorization unavailable; no write request was sent", no_write_verified=True) for op in operations]
    if api.cfg.get("server_id") != api.server_id or snapshot_server_id != api.server_id:
        return [_observation(op, "not_applied", "INSTANCE_VERSION_PARTITION_UNPROVEN: both adapter and source snapshot must bind this actual Zotero-Server-ID; discard versions from another local/Web instance; no write request was sent", no_write_verified=True) for op in operations]
    if action == "import_item":
        records, library_version = api.all_items()
    else:
        records, library_version = {}, None
    for op in operations:
        try:
            if action == "import_item":
                hits = _hits(op["identifier"], records)
                if len(hits) == 1:
                    rows.append(_observation(op, "succeeded", "Immediate exact DOI/URL read proved one existing identity; reused it without sending an import", uri=hits[0])); continue
                if hits:
                    rows.append(_observation(op, "not_applied", "IDENTITY_REVIEW_REQUIRED: multiple exact DOI/URL identities; no import was sent", no_write_verified=True, candidate_uris=hits)); continue
                payload = _metadata(op)
            else:
                payload = _membership(api, op, known, collection_binding)
            sending.append(op); payloads.append(payload)
        except ValueError as exc:
            rows.append(_observation(op, "not_applied", str(exc) + "; no write request was sent for this operation", no_write_verified=True))
    if not sending:
        return rows
    started = _now()
    for op in sending:
        journal["operations"][op["operation_id"]] = {"state": "STARTED", "started_at": started}
    save()  # durable before a mutating byte can be sent
    try:
        headers = {"If-Unmodified-Since-Version": str(library_version)} if action == "import_item" else {}
        status, _, data = api.request("POST", api.library_path + "/items", payload=payloads, headers=headers)
        rows.extend(_write_results(api, sending, payloads, data, status))
        if status == 412:
            # Re-read only. Never replace the version and repeat the request.
            for op in sending:
                if op.get("uri"):
                    try:
                        api.item(op["uri"].rsplit("/", 1)[1])
                    except (AdapterError, ValueError):
                        pass
    except (AdapterError, ValueError) as exc:
        rows.extend(_observation(op, "unknown", "Write was started but no conclusive result was obtained: " + str(exc)) for op in sending)
    for row in rows:
        journal["operations"][row["operation_id"]] = {"state": "OBSERVED", "observation": row}
    save()
    return rows


def _host_request(cfg, plan, status, operations):
    requests = []
    for op in operations:
        tool = cfg.get("host_tools", {}).get(op["action"])
        if tool:
            requests.append({"schema": HOST_REQUEST, "request_id": _digest([plan["plan_fingerprint"], op["operation_id"], tool]),
                             "plan_fingerprint": plan["plan_fingerprint"], "source_document_sha256": plan["source_document_sha256"],
                             "library_id": plan["library_id"], "tool_name": tool, "action": op["action"], "operation": op,
                             "requirements": "Host must actually read-bind library identity, enforce DOI/URL reuse and version preconditions, preserve other memberships, and record the real raw response. This request is not execution evidence."})
    return {"result": "HOST_DISPATCH_REQUIRED" if requests else "CAPABILITY_MISSING", "host_requests": requests,
            "next_action": "Dispatch the named existing host tool only after its exact schema/capabilities are verified; record actual observations and recheck state." if requests else NATIVE_NEXT}


def execute(plan, library_snapshot, config, *, current_document_sha256,
            collection_snapshot=None, observations=None, batch_index=None, max_age_hours=24):
    """Dispatch at most one bounded *original-plan* batch and return observations.

    Persist the returned observations and obtain fresh bound snapshots before
    batch-status/resume.  A journal prevents omitted observations from replaying
    a started write.  This function cannot produce a Word/native/final PASS.
    """
    cfg = _config(config)
    batch = _helper("library_batch")
    core = _helper("zotero_docx")
    _require(isinstance(plan, dict) and plan.get("library_id") == cfg["library_id"], "Plan and adapter target different libraries")
    # Validate all bindings before touching an API or journal.
    checked = batch.status(plan, library_snapshot, current_document_sha256=current_document_sha256,
                           collection_snapshot=collection_snapshot, observations=observations, max_age_hours=max_age_hours)
    _, known = core.validate_snapshot(library_snapshot, "library", max_age_hours)
    _require(all(op.get("action") in ACTIONS for op in plan["operations"]), "Unsupported action; delete/merge is never executable")
    result = {"schema": EXECUTION, "plan_fingerprint": plan["plan_fingerprint"],
              "source_document_sha256": current_document_sha256, "library_id": cfg["library_id"],
              "test_evidence": "TEST_DOUBLE" if cfg.get("test_double") else "REAL_ADAPTER",
              "raw_responses": [], "observations": observations or _observations(plan, []),
              "policy": "Library stage only. Not a citation, native Word, collection-finalization or manuscript completion PASS."}
    if cfg["kind"] == "host_dispatch":
        ids = {o["operation_id"] for o in checked["ready_operations"]}
        chosen = next((b for b in plan["batches"] if (batch_index is None or b["batch_index"] == batch_index)
                       and ids.intersection(b["operation_ids"])), None)
        operations = [o for o in checked["ready_operations"] if chosen and o["operation_id"] in chosen["operation_ids"]][:cfg["max_batch_size"]]
        return {**result, **_host_request(cfg, plan, checked, operations)}
    try:
        with _state(cfg, plan) as (journal, save, path):
            result["journal_path"] = str(path)
            merged = _merge_observations(plan, observations, journal)
            result["observations"] = merged
            if merged["results"] != (observations or _observations(plan, []))["results"]:
                try:
                    checked = batch.status(plan, library_snapshot, current_document_sha256=current_document_sha256,
                                           collection_snapshot=collection_snapshot, observations=merged, max_age_hours=max_age_hours)
                except ValueError as exc:
                    return {**result, "result": "RECHECK_REQUIRED", "reason": str(exc), "next_action": "Obtain real post-write library/collection snapshots, then run batch-status; never relabel old snapshots."}
            ids = {o["operation_id"] for o in checked["ready_operations"]}
            chosen = next((b for b in plan["batches"] if (batch_index is None or b["batch_index"] == batch_index)
                           and ids.intersection(b["operation_ids"])), None)
            if not chosen:
                return {**result, "result": checked["status"], "batch_status": checked,
                        "next_action": "Follow batch-status review/recheck requirements; no pending operation was dispatched."}
            operations = [o for o in checked["ready_operations"] if o["operation_id"] in chosen["operation_ids"]][:cfg["max_batch_size"]]
            result.update(batch_index=chosen["batch_index"], operation_ids=[o["operation_id"] for o in operations])
            api = LocalAPI(cfg)
            try:
                result["health"] = _health(api)
                _require(not journal.get("server_id") or journal["server_id"] == api.server_id,
                         "Journal belongs to another Zotero instance; discard snapshots/versions and rebuild")
                journal["server_id"] = api.server_id
                rows = _dispatch(api, operations, known, journal, save, library_snapshot.get("server_id"), plan.get("collection"))
                # Persist non-write outcomes as well, including DOI reuse.
                for row in rows:
                    journal["operations"][row["operation_id"]] = {"state": "OBSERVED", "observation": row}
                save()
                by_id = {row["operation_id"]: row for row in merged["results"]}
                by_id.update({row["operation_id"]: row for row in rows})
                result["observations"] = _observations(plan, by_id.values())
                if any(row["outcome"] in ("unknown", "failed") or (row["outcome"] == "not_applied" and "CONFLICT" in row["evidence"])
                       or "INSTANCE_VERSION_PARTITION_UNPROVEN" in row["evidence"] for row in rows):
                    result["result"] = "RECHECK_REQUIRED"
                elif any("CAPABILITY_MISSING" in row["evidence"] for row in rows):
                    result["result"] = "CAPABILITY_MISSING"
                elif any(row["outcome"] == "not_applied" for row in rows):
                    result["result"] = "REVIEW_REQUIRED"
                else:
                    result["result"] = "BATCH_DISPATCHED"
                result["next_action"] = NATIVE_NEXT if result["result"] == "CAPABILITY_MISSING" else "Obtain fresh actual library/collection snapshots and run batch-status; then continue native Word workflow/final audit. Do not replay unknown writes."
            except (AdapterError, ValueError) as exc:
                result.update(result="CAPABILITY_MISSING" if not any(r.get("method") == "POST" for r in api.raw) else "RECHECK_REQUIRED",
                              code=getattr(exc, "code", "PRECONDITION_UNPROVEN"), reason=str(exc), next_action=NATIVE_NEXT)
            finally:
                result["raw_responses"] = api.raw
                api.close()
    except AdapterError as exc:
        result.update(result="RECHECK_REQUIRED", code=exc.code, reason=str(exc), next_action="Wait for the current batch, then re-observe state; do not run a competing writer.")
    return result

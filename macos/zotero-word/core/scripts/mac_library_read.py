"""Read the personal library over loopback, with the existing Mac reader as fallback."""
from datetime import datetime, timezone
from http.client import IncompleteRead
import subprocess
import time


PAGE_SIZE = 100
SOCKET_TIMEOUT = 3
STARTUP_SECONDS = 12


def _await_first_page(client, query, transport, report, failure):
    started = time.monotonic()
    deadline = started + STARTUP_SECONDS
    startup = report["startup"] = {"wait_limit_seconds": STARTUP_SECONDS, "reads": 0}
    try:
        while time.monotonic() < deadline:
            if isinstance(failure.__cause__, ConnectionRefusedError) and "launch" not in startup:
                # Launch Services reuses a running Zotero and leaves the foreground app in place.
                try:
                    opened = subprocess.run(["/usr/bin/open", "-g", "-b", "org.zotero.zotero"],
                                            capture_output=True, text=True, timeout=min(3, max(0, deadline - time.monotonic())))
                    startup["launch"] = "requested" if opened.returncode == 0 else opened.stderr.strip() or "failed"
                except (OSError, subprocess.TimeoutExpired) as error:
                    startup["launch"] = type(error).__name__
            time.sleep(min(0.5, max(0, deadline - time.monotonic())))
            remaining = deadline - time.monotonic()
            if remaining <= 0:
                break
            client.close()
            client.connection.timeout = remaining
            client.raw.clear()
            startup["reads"] += 1
            try:
                return client.get("users/0/items", query=query)
            except transport.AdapterError as error:
                status = client.raw[-1].get("http_status") if client.raw else None
                if status != 503 and not isinstance(error.__cause__, (TimeoutError, ConnectionError, IncompleteRead)):
                    raise
                failure = error
        raise transport.AdapterError("ZOTERO_STARTUP_TIMEOUT", "Zotero did not serve the first library page in time") from failure.__cause__
    finally:
        client.close()
        client.connection.timeout = SOCKET_TIMEOUT
        startup["seconds"] = round(time.monotonic() - started, 3)


def read(job, request, expected_library, api):
    transport = api._helper("library_executor")
    client = transport.LocalAPI({"base_url": "http://127.0.0.1:23119/api/", "library_id": "users/0",
                                 "timeout_seconds": SOCKET_TIMEOUT})
    query = [("format", "csljson"), ("limit", PAGE_SIZE), ("sort", "dateAdded"),
             ("itemType", "-attachment"), ("itemType", "-note"), ("itemType", "-annotation")]
    started = time.monotonic()
    report = {"request": request, "page_size": PAGE_SIZE, "socket_timeout_seconds": SOCKET_TIMEOUT}
    records = {}
    library = expected_library
    signature = None
    export = None
    try:
        while True:
            offset = len(records)
            report["offset"] = offset
            client.raw.clear()  # The batch client journals bodies. Keep only this page in memory.
            page_query = [*query, ("start", offset)]
            try:
                headers, page = client.get("users/0/items", query=page_query)
            except transport.AdapterError as error:
                cause = error.__cause__
                status = client.raw[-1].get("http_status") if client.raw else None
                if offset == 0 and (isinstance(cause, ConnectionRefusedError) or status == 503):
                    headers, page = _await_first_page(client, page_query, transport, report, error)
                else:
                    if "retry" in report or not isinstance(cause, (TimeoutError, ConnectionError, IncompleteRead)):
                        raise
                    report["retry"] = {"offset": offset, "reason": type(cause).__name__}
                    time.sleep(0.25)
                    continue  # Repeat only the failed GET, before adding any of its rows.

            version = headers.get("last-modified-version", "")
            total = headers.get("total-results", "")
            if headers.get("zotero-api-version") != "3" or not version.isdecimal() or not total.isdecimal():
                raise ValueError("Response lacks Zotero API v3, the library version or the item count")
            total = int(total)
            observed = (headers.get("zotero-server-id"), version, total)
            if signature is None:
                signature = observed
                client.server_id = observed[0]
            elif observed != signature:
                raise ValueError("Zotero instance, library version or count changed between pages")
            if not isinstance(page, list) or len(page) != min(PAGE_SIZE, total - offset):
                raise ValueError("Incomplete Zotero library page")
            if total == 0:
                raise ValueError("Use Zotero's native reader to identify the empty library")
            for row in page:
                uri = api.canonical_uri(row.get("id")) if isinstance(row, dict) else None
                if not uri or not row.get("type"):
                    raise ValueError("CSL export did not supply a real Zotero item URI and type")
                owner = api.library_of(uri)
                if not owner.startswith("users/") or owner == "users/0" or (library and owner != library):
                    raise ValueError("Zotero item belongs to a different or unidentified personal library")
                if uri in records:
                    raise ValueError("Duplicate Zotero item across library pages")
                library = owner
                records[uri] = dict(row, id=uri)
            if len(records) == total:
                items = list(records.values())
                export = {"transport": "local_api", "library_id": library, "items": items, "import_receipt": [],
                          "observed_at": datetime.now(timezone.utc).isoformat()}
                report.update(result="complete", library_id=library, library_version=version, total=total,
                              items_sha256=api.digest(items))
                break
    except (transport.AdapterError, ValueError) as error:
        report.update(result="native_fallback", code=getattr(error, "code", "INVALID_LIBRARY_PAGE"),
                      reason=str(error.__cause__ or error))
        if error.__cause__:
            report["cause"] = type(error.__cause__).__name__
    finally:
        client.close()
    if client.raw:
        report["http_status"] = client.raw[-1].get("http_status")
    report["seconds"] = round(time.monotonic() - started, 3)
    api.output(job / "read-connection.json", report)
    if export is not None:
        api.output(job / "library.json", export)
    return export

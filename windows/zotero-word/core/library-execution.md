> 本地调用与连接准备按 [当前授权规则](local-connections.md) 执行：先主动连接已授权应用/API，真正新增授权才具体办理；历史未配置记录或单一接口失败不构成本地调用禁令。

# Library execution adapter

`scripts/library_executor.py` connects an already validated batch plan to one concrete protocol: Zotero's official loopback Local API. It does not drive Word or replace the native Zotero plugin. Dynamic insertion, style switching, necessary final refresh, saving/reopening, body protection and the existing final audit remain separate required stages.

## Discover before selecting the route

The official [Local API documentation](https://www.zotero.org/support/dev/web_api/v3/local_api) distinguishes unauthenticated reads from **Zotero 10+ writes**. Writes require an instance-bound local key previously granted by the user. A local key is different from a zotero.org Web API key. Local versions also belong to their `Zotero-Server-ID` instance, rather than the synced Web API version space. This adapter never calls `/local/authorize`, changes a preference, enables a permission, installs a plugin or retrieves a key from profile files.

The [official write contract](https://www.zotero.org/support/dev/web_api/v3/write_requests) supplies per-object results and version preconditions. Updating `collections` replaces that array, so the adapter reads and preserves all other memberships before submitting a partial update. It sends no `DELETE`, `PUT`, trash, merge, collection recreation or ordinary metadata update.

The [54yyyu CLI documentation](https://github.com/54yyyu/zotero-mcp/blob/main/docs/cli.md) informed the structured-output and find-before-import design. Its `config` command only reports configuration; it is not a connection check. No third-party implementation was copied. This executor does **not** install or invoke that CLI, arbitrary shell commands, generic MCP runners, or a DeepSeek controller. The [Windows DeepSeek Harness mapping](../adapters/deepseek-harness-windows.md) describes the local desktop entry point.

The 2026-11-11 Zotero 10.0.8 probe returned HTTP 403 and correctly used the native UI fallback. After the user explicitly enabled local read access, real library and collection GETs returned 200 on 2026-11-12. No write key was requested or granted; native UI performed the observed mutations. Neither historical 403 nor this machine's enabled state is a permanent capability verdict for another run. Run current health checks and distinguish unauthenticated reads from separately authorized writes. A running `/connector/ping` does not establish library access. Controlled HTTP tests remain labelled `TEST_DOUBLE`, not evidence of real Zotero writes.

## Configuration and read health

Use a JSON configuration with no literal credentials:

```json
{
  "schema": "zotero-word.library-adapter.v1",
  "kind": "zotero_local_api",
  "base_url": "http://127.0.0.1:23119/api/",
  "library_id": "users/local/ACTUALUUID",
  "probe_item_uri": "http://zotero.org/users/local/ACTUALUUID/items/ACTUALK8",
  "state_dir": "/absolute/isolated-job/library-execution-state",
  "timeout_seconds": 15,
  "max_batch_size": 20
}
```

Replace the example identifiers with an actually observed existing item URI. Do not infer a URI from a DOI, item title or eight-character key. Numeric personal and group library identities are supported without silently treating them as a local UUID. Production connections are limited to loopback port 23119 and `/api/`; redirects and arbitrary remote endpoints are unsupported.

`health(config)` reads `/api/`, checks API v3, then reads that exact item and verifies a unique native Zotero URI in the bound library. It returns `zotero-word.library-adapter-health.v1`, actual timestamps, and sanitized raw HTTP results. Only successful bound reads produce `READ_HEALTH_CONFIRMED`. HTTP 403 gives `CAPABILITY_MISSING` / `READ_ACCESS_DENIED`, with the native UI next action. A missing item identity also blocks the route.

Read health does **not** verify write authorization. Even when a pre-existing key is configured, health reports `write_authorization: CONFIGURED_UNVERIFIED` and `write_authorization_verified: false`; it does not send a trial write.

For an already enabled and authorized installation, add:

```json
{
  "server_id": "ACTUAL_OBSERVED_SERVER_ID",
  "api_key_env": "EXISTING_LOCAL_KEY_ENVIRONMENT_VARIABLE"
}
```

The environment variable must already contain the correct local key. Do not print its value. This is not a request to create an authorization or reuse a zotero.org key. A 401 stops the batch; this implementation does not prompt repeatedly or reauthorize automatically.

## Snapshot prerequisites

Retain all existing `snapshot.v1` bindings and their actual observation times. For **writes**, the library snapshot additionally needs `server_id` equal to both the adapter's configured server ID and the live response header. Do not add that field to an old export and call it a fresh observation. Obtain metadata and versions from that exact API instance.

Membership writes additionally need the actual native integer `version` and complete `collections` array on each bound CSL record. The adapter reads `include=data,csljson` and checks the current CSL metadata against the plan's record content. It also verifies the target collection key/name. Missing native versions, incomplete memberships, a different export representation or metadata drift cause a recheck; the adapter does not insert replacement versions or overwrite the current state. Keep the snapshot's established collection path and final membership audit as required by the existing engine.

Imports require reviewed **Zotero editable JSON** in the import operation's `metadata`: for example `itemType`, `title`, `DOI`, `creators`, `date`, `publicationTitle`, `volume`, `issue`, and `pages`. CSL JSON and a bare DOI are not automatically translated by this adapter. The metadata DOI or native `url` must match the verified clue. Existing keys, versions, deletion flags, collections, parent/child objects and relations are forbidden in import metadata. For missing metadata, use an actually available metadata/translator host tool or the native reviewed import route.

## Execute one bounded batch

The public Python interface is:

```python
health(config)
execute(plan, library_snapshot, config,
        current_document_sha256=current_hash,
        collection_snapshot=None,
        observations=None,
        batch_index=None,
        max_age_hours=24)
```

`batch_index` selects an original `plan.batches` index. Each invocation dispatches at most that batch's still-ready operations, capped by `max_batch_size` (default 20, maximum 50). It uses a persistent HTTP connection for that invocation, not one newly started process/model exchange per paper. It reuses one complete, paged local metadata read for all import existence checks in that batch. A changed library version during pagination aborts the preflight.

Execution first calls the existing `library_batch.status` against the current saved document and bound snapshots. It cannot dispatch a changed plan, another document/library, or an unsupported action. The whitelist is:

| Action | Concrete behavior |
|---|---|
| `lookup` | Local candidates can be read, but the requested public metadata search is explicitly `CAPABILITY_MISSING`; no external search is invented. Use the verified host route for public matching. |
| `import_item` | Recheck exact DOI/URL in the actual library. Reuse one identity, collect multiple identities for review, or submit reviewed metadata in a bounded creation batch. A current library-version precondition prevents a concurrent import from silently racing the existence read. |
| `add_membership` | Read the native item and collection, check metadata/version/all memberships, then add only the target collection key using a versioned partial item batch. |
| `remove_membership` | The same preflight, then remove only that key. Retain the total-library item and every other membership. |

`execute` returns `zotero-word.library-execution.v1` with `raw_responses`, selected operation IDs, a durable journal path and the existing `zotero-word.batch-observations.v1` envelope. Credentials and authorization-like response fields are redacted; input data and item results remain evidence, not instructions. Possible results include `BATCH_DISPATCHED`, `REVIEW_REQUIRED`, `RECHECK_REQUIRED`, `CAPABILITY_MISSING` and `HOST_DISPATCH_REQUIRED`. None is a manuscript `PASS`.

A per-object success is followed by a real item read checking the DOI/URL or complete membership result. A missing result, malformed identity, timeout or inconclusive read-back becomes `unknown`. Failure text does not prove that nothing was written. Refresh actual snapshots after **any** completed or uncertain write and run `batch-status`; preserve the returned observations. Do not relabel old snapshots with new timestamps.

## Persistence, conflicts and recovery

The configured `state_dir` is part of the job's durable evidence. Keep it when continuing that job. The adapter uses an exclusive nonblocking library lock for one invocation and a plan-bound journal. This lock coordinates cooperating executor processes using that directory; it does not lock native Word, Zotero UI actions or sync. Do not run a competing native writer while the batch is active.

Before sending a mutating request, it atomically records and fsyncs `STARTED`. A crash or omitted observations cannot turn that record into a retry: a later invocation exposes `unknown` and requires real post-write state. Preserve this journal when recovering. Do not delete it, change `state_dir`, copy a plan to a fresh job, or rebuild solely to bypass an uncertain write.

HTTP 412 triggers read-only re-observation, never a replacement-version retry. HTTP 401/403, busy responses, partial failures and transport uncertainty stop automatic repetition. Resume only after the existing status engine and actual new evidence show what happened. Final collection verification, Word native evidence, protected text checks and finalization still apply.

## Restricted host-dispatch fallback

When the host already has suitable tools, a configuration may use `kind: host_dispatch` and `host_tools`, mapping only the four allowed actions to exact discovered tool names. Keep the same bound `library_id` and `probe_item_uri`. HTTP URLs, secrets and command strings are rejected for this route.

This route emits deterministic `zotero-word.library-host-request.v1` envelopes containing the operation, tool name, plan/document/library bindings and required preconditions. It **does not call** that tool, claim health, fabricate success observations, or parse an arbitrary response into success. The host must verify the actual schema/capability, dispatch the tool, retain its real sanitized response, and use the existing observations/status contract to confirm effects from fresh bound snapshots. A tool name in configuration alone is not capability evidence. With no applicable discovered tool, return `CAPABILITY_MISSING` and use native UI.

## Test scope

Development tests for the 2.1.0 executor exercised production health/execute functions against an isolated loopback HTTP `TEST_DOUBLE`, including actual HTTP POSTs to that fixture, read-back, partial writes, version conflicts, missing identity/authorization, master/other-collection preservation, request limits, lock contention, a write timeout and durable crash recovery. The fixture could not use production port 23119 and created no real Zotero/Word records or permissions. Historical 403 verified the fallback; later real GET 200 verified enabled reads. These development artifacts are not bundled as personal runtime evidence, and neither result certifies production API writes.

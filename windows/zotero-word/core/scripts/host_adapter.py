"""Select an authored Windows mapping from observed host metadata, never a model.

Host context is a recorded observation, not independent proof of native tools
or authorization. Unknown hosts have no default mapping.
"""
from datetime import datetime, timezone
import hashlib
import json
import os
from pathlib import Path

ROUTES = {
    "codex": "codex-windows",
    "deepseek-harness": "deepseek-harness-windows",
    "deepseek-local-agent": "deepseek-harness-windows",
}
SCHEMA = "zotero-word.adapter-selection.v1"


def require(value, message):
    if not value:
        raise ValueError(message)


def select(host_context):
    require(os.name == "nt", "Windows host selection requires Windows Python")
    require(isinstance(host_context, (str, Path)) and str(host_context).strip(),
            "Observed host context is required; no default Windows adapter")
    path = Path(host_context).resolve()
    raw = path.read_bytes()
    context = json.loads(raw)
    require(isinstance(context, dict) and context.get("schema") == "zotero-word.host-context.v1",
            "Use a real zotero-word.host-context.v1 observation")
    host = context.get("agent_host")
    require(host in ROUTES, "Unrecognized agent host; do not infer it from the model provider")
    require(context.get("application_os") == "windows" and context.get("execution_os") == "windows",
            "Both application desktop and command execution must be Windows")
    require(isinstance(context.get("evidence"), str) and context["evidence"].strip(),
            "Host selection needs the actual runtime/tool metadata source")
    tools = context.get("tools")
    require(isinstance(tools, list) and tools and all(isinstance(t, str) and t.strip() for t in tools)
            and len(tools) == len(set(tools)), "Record the actual discovered tool names")
    stamp = context.get("observed_at")
    require(isinstance(stamp, str), "Host observation needs a timestamp")
    observed = datetime.fromisoformat(stamp.replace("Z", "+00:00"))
    require(observed.tzinfo is not None and observed <= datetime.now(timezone.utc),
            "Host observation timestamp must have a timezone and not be in the future")
    return {"schema": SCHEMA, "result": "HOST_SELECTED", "adapter_id": ROUTES[host],
            "agent_host": host, "application_os": "windows", "execution_os": "windows",
            "host_context": str(path), "host_context_sha256": hashlib.sha256(raw).hexdigest(),
            "instructions": "adapters/" + ROUTES[host] + ".md",
            "native_capabilities": "NOT_PROVEN_BY_HOST_SELECTION"}

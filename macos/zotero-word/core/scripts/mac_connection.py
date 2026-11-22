"""Word Apple events and one run-scoped Zotero executor, shared by the native jobs."""
from contextlib import contextmanager
from datetime import datetime, timezone
import errno
import fcntl
import json
import os
from pathlib import Path
import re
import subprocess
import sys
import tempfile
import time

SCRIPTS = Path(__file__).resolve().parents[2] / "scripts"
JOB_NAME = re.compile(r"native-[0-9a-f]{32}")


class DesktopError(ValueError):
    def __init__(self, code, message, action, *, retry=False, observation=None):
        super().__init__(message)
        self.reason_code = code
        self.next_action = action
        self.retry_allowed = retry
        self.observation = observation or {}


def _script_error(script, action, message, info=None):
    info = info or ({"bundle": "com.microsoft.Word", "phase": action} if script == "word.applescript" else {})
    number = info.get("appleEvent")
    if script == "word.applescript":
        match = re.search(r"\((-?\d+)\)\s*$", message)
        if match:
            number = info["appleEvent"] = int(match[1])
        if number in (-10004, -1743):
            info["host_hint"] = os.environ.get("__CFBundleIdentifier") or os.environ.get("TERM_PROGRAM")
            if action == "connect":
                info["command"] = "get documents"
            return DesktopError("WORD_CONTROL_DENIED", message,
                                "Check the calling host's Automation access to Microsoft Word, then resume this run",
                                observation=info)
    if number in (-1743, -25211) or any(s in message for s in ("-1743", "-25211")) or (
            "-1719" in message and "assistive" in message.lower()):
        return DesktopError("DESKTOP_ACCESS_DENIED", message,
                            "Check this host's macOS Automation and Accessibility access", observation=info)
    if script == "word.applescript":
        if action == "connect":
            return DesktopError("WORD_CONNECT", message, "Launch Word from the host's authorized macOS desktop session",
                                observation=info)
        return DesktopError("WORD_" + action.upper().replace("-", "_"), message, "Check the bound Word working copy",
                            observation=info)
    if info.get("type") in ("TypeError", "ReferenceError", "SyntaxError"):
        return DesktopError("ZOTERO_BRIDGE_ERROR", message, "Correct the bundled connector before resuming", observation=info)
    phase = info.get("phase")
    if info.get("runSent"):
        return DesktopError("ZOTERO_DISPATCH_UNCONFIRMED", message,
                            "Inspect this job's native receipts before resuming", observation=info)
    if phase in ("process", "menu", "window", "focus", "clipboard"):
        return DesktopError("ZOTERO_" + phase.upper(), message, "Resolve the recorded desktop state and resume this run",
                            retry=True, observation=info)
    return DesktopError("ZOTERO_BRIDGE_ERROR", message, "Inspect the bundled connector and its macOS error", observation=info)


def _osa(script, *args, timeout=45):
    language = "JavaScript" if script.endswith(".js") else "AppleScript"
    try:
        result = subprocess.run(["/usr/bin/osascript", "-l", language, str(SCRIPTS / script), *map(str, args)],
                                capture_output=True, text=True, timeout=timeout)
    except subprocess.TimeoutExpired as exc:
        code = "ZOTERO_DISPATCH_UNCONFIRMED" if language == "JavaScript" else "WORD_TIMEOUT"
        raise DesktopError(code, f"macOS timed out during {script}.", "Inspect the last confirmed native operation") from exc
    if result.returncode:
        raise _script_error(script, args[0], result.stderr.strip() or f"macOS could not complete {script}")
    return result.stdout.strip("\r\n")


def _start(job, operation):
    # Loaded code always comes from this distribution. The handoff contains data only.
    loader = """(() => {
        const directory = %s;
        const scope = {Zotero};
        const win = Services.wm.getMostRecentWindow("zotero:run-js");
        const editor = win?.codeEditor;
        const text = editor?.getValue();
        try {
            const options = {target: scope, ignoreCache: true};
            Services.scriptloader.loadSubScriptWithOptions(%s, options);
            Services.scriptloader.loadSubScriptWithOptions(%s, options);
            void scope.serve(directory, scope.runJob);
        } catch (error) {
            void Zotero.File.putContentsAsync(directory + "/error.json",
                JSON.stringify({stage: "bootstrap", type: error.name, error: String(error)}));
        }
        if (text && text.includes(%s)) {
            setTimeout(() => { if (!win.closed && editor.getValue() === text) win.close(); }, 0);
        }
    })()""" % (json.dumps(str(job)), json.dumps((SCRIPTS / "zotero-native.js").as_uri()),
               json.dumps((SCRIPTS / "zotero-session.js").as_uri()), json.dumps(job.name))
    raw = _osa("zotero-runner.js", loader, job, operation, timeout=50)
    try:
        reply = json.loads(raw)
    except ValueError:
        reply = None
    if isinstance(reply, dict):
        error = reply.get("error")
        if isinstance(error, dict):
            raise _script_error("zotero-runner.js", None, error["detail"], error)
        pid = reply.get("pid")
        if type(pid) is int and pid > 0:
            return pid
    raise DesktopError("ZOTERO_BRIDGE_ERROR", "Invalid response from the Zotero launcher.",
                       "Inspect the job receipt before resuming")


class Connection:
    def __init__(self, root, api, report):
        self.root = root
        self.api = api
        self.report = report
        self.first = self.last = self.pid = None

    @staticmethod
    def word(action, *args, timeout=45):
        if sys.platform != "darwin":
            raise DesktopError("WORD_CONNECT", "MacWord requires an interactive macOS desktop.", "Use a macOS desktop session")
        return _osa("word.applescript", action, *args, timeout=timeout)

    @contextmanager
    def desktop_lock(self, job):
        # Different run directories still share Word's active selection on this desktop.
        path = Path(tempfile.gettempdir()) / f"zotero-word-desktop-{os.getuid()}.lock"
        fd = os.open(path, os.O_CREAT | os.O_RDWR | os.O_NOFOLLOW, 0o600)
        with os.fdopen(fd, "r+b") as lock:
            try:
                try:
                    fcntl.flock(lock, fcntl.LOCK_EX | fcntl.LOCK_NB)
                except BlockingIOError as exc:
                    (job / "abort").touch(exist_ok=True)
                    raise DesktopError("DESKTOP_BUSY", "Another run is using this Word/Zotero desktop.",
                                       "Resume after that native operation finishes", retry=True) from exc
                yield
            except DesktopError as error:
                self._record_failure(job, error)
                raise

    def close(self):
        if self.first is not None:
            try:
                (self.first / "session.stop").touch(exist_ok=True)
            except FileNotFoundError:
                pass  # The run directory itself has gone.
        self.first = self.last = self.pid = None

    def _handoff(self, job):
        try:
            os.kill(self.pid, 0)  # Check existence before submitting any work to this process.
        except ProcessLookupError:
            return False
        slot = self.last / "next.json"
        try:
            os.link(job / "request.json", slot)
        except FileExistsError:
            if slot.stat().st_size:
                raise DesktopError("ZOTERO_HANDOFF_CONFLICT", "The preceding job already has a successor.",
                                   "Inspect the preceding job before resuming")
            return False  # The executor retired before this request was submitted.
        except OSError as error:
            if error.errno in (errno.ENOTSUP, errno.EXDEV, errno.EPERM):
                return False  # This filesystem cannot publish the request by hard link.
            raise
        return True

    def dispatch(self, job, cfg):
        if job.parent != self.root or job != job.resolve() or not JOB_NAME.fullmatch(job.name):
            raise ValueError("Native job must be a private child of the current run")
        if cfg["libraryOnly"]:
            operation = "library_import" if cfg.get("imports") else "library_read"
        else:
            operation = "reopen_verify" if cfg["reopenOnly"] else "refresh"
        request = {"job": job.name, "operation": operation, "document": cfg.get("document")}
        while True:
            request["session"] = self.first.name if self.first else job.name
            self.api.output(job / "request.json", request)
            status = {"job": str(job), "session": request["session"], "operation": operation,
                      "via": "session" if self.first else "editor"}
            self.report({**status, "state": "awaiting_start"})
            if self.first is None:
                self.first = job
                self.pid = _start(job, operation)
                break
            if self._handoff(job):
                break
            self.close()
        self.last = job
        receipt = self.wait(job, "started.json", 12)
        if receipt != request:
            raise DesktopError("ZOTERO_RECEIPT_MISMATCH", "Zotero acknowledged a different job, session or document.",
                               "Inspect request.json and started.json for this job")
        self.report({**status, "state": "accepted", "pid": self.pid, "receipt": str(job / "started.json")})

    def wait(self, job, name, seconds):
        end = time.monotonic() + seconds
        while time.monotonic() < end:
            receipt = job / name
            if name == "started.json" and receipt.is_file():
                return self.api.load(receipt)
            for error_path in (job / "error.json", self.first / "session-error.json"):
                if error_path.is_file():
                    error = self.api.load(error_path)
                    code = ("ZOTERO_BRIDGE_ERROR" if error.get("type") in ("TypeError", "ReferenceError", "SyntaxError")
                            else "ZOTERO_NATIVE")
                    raise DesktopError(code, f"Zotero {error.get('stage', 'session')}: {error['error']}",
                                       "Inspect the native error and last confirmed operation", observation=error)
            if (job / "abort").exists():
                raise DesktopError("ZOTERO_STOPPED", "This native job was stopped.", "Resume from the last confirmed operation")
            if receipt.is_file():
                return self.api.load(receipt)
            time.sleep(0.1)
        code = "ZOTERO_DISPATCH_UNCONFIRMED" if name == "started.json" else "ZOTERO_RECEIPT_TIMEOUT"
        raise DesktopError(code, f"Zotero did not confirm {name} within {seconds} seconds.",
                           "Check this job's receipts before repeating any native operation")

    def _record_failure(self, job, error):
        diagnostic = job / "connection-error.json"
        if diagnostic.exists():
            diagnostic = job / f"connection-error-{time.time_ns()}.json"
        self.api.output(diagnostic, {"time": datetime.now(timezone.utc).isoformat(), "job": job.name,
            "code": error.reason_code, "detail": str(error), "observation": error.observation,
            "native_started": (job / "started.json").is_file()})
        error.diagnostic = str(diagnostic)

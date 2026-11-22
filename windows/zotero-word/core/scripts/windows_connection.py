"""Run-scoped Windows desktop connection. One STA owns Word and the Zotero handoffs."""
from contextlib import contextmanager
from datetime import datetime, timezone
import json
import os
from pathlib import Path
import queue
import re
import subprocess
import sys
import tempfile
import threading
import time
import uuid

SCRIPTS = Path(__file__).resolve().parents[2] / "scripts"
JOB_NAME = re.compile(r"native-[0-9a-f]{32}")
EXECUTOR = "windows-native-selection-2.4.2"  # Existing Word selection and reopen receipts use this ID.


class DesktopError(ValueError):
    def __init__(self, code, message, action, *, retry=False, observation=None):
        super().__init__(message)
        self.reason_code = code
        self.next_action = action
        self.retry_allowed = retry
        self.observation = observation or {}


def _failure(action, reply):
    phase = reply.get("phase", action)
    message = reply.get("error") or "The Windows desktop host could not complete the request."
    diagnostic = {k: reply[k] for k in ("phase", "hresult", "runSent", "focusMethod", "focusRecovery") if k in reply}
    if reply.get("runSent"):
        code, hint = "ZOTERO_DISPATCH_UNCONFIRMED", "Inspect this job's native receipts before resuming"
    elif action in ("bootstrap", "launch-zotero", "zotero-alive"):
        code, hint = "ZOTERO_" + phase.upper(), "Inspect the recorded Zotero window or process and resume this run"
    elif reply.get("hresult") == "80070005":
        diagnostic["operation"] = action
        code, hint = "WORD_CONTROL_DENIED", "Grant the calling task access to Word in the same signed-in Windows desktop, then resume the same run directory"
    else:
        code, hint = "WORD_" + action.replace("-", "_").upper(), "Check the bound Word working copy and resume this run"
    return DesktopError(code, message, hint, observation=diagnostic)


def _desktop_shell():
    if os.name != "nt":
        raise ValueError("The WinWord integration requires Windows and the interactive user's desktop")
    shell = Path(os.environ["SystemRoot"]) / "System32/WindowsPowerShell/v1.0/powershell.exe"
    args = [str(shell), "-NoProfile", "-NonInteractive", "-STA"]
    # RemoteSigned only for this child on machines with all five scopes unset.
    # A configured process, user, machine or enterprise policy stays in force.
    try:
        probe = subprocess.run([*args, "-Command",
            "if (@(Get-ExecutionPolicy -List | Where-Object ExecutionPolicy -ne Undefined).Count -eq 0) { 'default' }"],
            capture_output=True, timeout=5)
    except subprocess.TimeoutExpired as error:
        raise ValueError("Windows PowerShell did not respond during connection") from error
    if probe.returncode == 0 and probe.stdout.strip() == b"default":
        args += ["-ExecutionPolicy", "RemoteSigned"]
    return args


def _loader(job):
    # Loaded code always comes from this distribution. The handoff contains data only.
    return """(() => {
        const directory = %s;
        const scope = {Zotero};
        const win = Services.wm.getMostRecentWindow("zotero:run-js");
        const editor = win?.codeEditor;
        const text = editor?.getValue();
        try {
            const options = {target: scope, ignoreCache: true};
            Services.scriptloader.loadSubScriptWithOptions(%s, options);
            if (scope.winWordExecutorVersion !== %s) throw new Error("The loaded WinWord executor version differs from this package.");
            Services.scriptloader.loadSubScriptWithOptions(%s, options);
            void scope.serve(directory, scope.runJob);
        } catch (error) {
            void Zotero.File.putContentsAsync(%s,
                JSON.stringify({stage: "bootstrap", type: error.name, error: String(error), scriptVersion: scope.winWordScriptVersion ?? null}));
        }
        if (text && text.includes(%s)) {
            setTimeout(() => { if (!win.closed && editor.getValue() === text) win.close(); }, 0);
        }
    })()""" % (json.dumps(str(job)), json.dumps((SCRIPTS / "windows-native.js").as_uri()), json.dumps(EXECUTOR),
               json.dumps((SCRIPTS / "zotero-session.js").as_uri()), json.dumps(str(job / "error.json")), json.dumps(job.name))


def _pump(stream, replies):
    try:
        for line in stream:
            replies.put(line.rstrip("\r\n"))
    finally:
        replies.put(None)


class Connection:
    executor = EXECUTOR

    def __init__(self, root, api, report):
        self.root, self.api, self.report = root, api, report
        self.first = self.last = self.pid = None
        self.host = None
        self.broken = False
        self.replies = queue.Queue()
        self.protected = {}
        self.prompted_job = None

    def _rpc(self, action, *args, timeout=45, observation=None):
        if self.broken:
            raise DesktopError("DESKTOP_DISCONNECTED", "The desktop host is no longer connected.",
                               "Resume this run from its last durable receipt")
        if self.host is None:
            self.host = subprocess.Popen([*_desktop_shell(), "-File", str(SCRIPTS / "Invoke-WordZotero.ps1"), "-Bridge"],
                stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                text=True, encoding="utf-8", errors="replace", bufsize=1)
            self.report({"desktop_event": "session_starts"})
            threading.Thread(target=_pump, args=(self.host.stdout, self.replies), daemon=True).start()
        request = {"id": uuid.uuid4().hex, "action": action, "args": list(map(str, args))}
        try:
            self.host.stdin.write(json.dumps(request, ensure_ascii=True) + "\n")
            self.host.stdin.flush()
        except (BrokenPipeError, OSError) as exc:
            self.broken = True
            detail = self.host.stdout is not None and self.host.poll()
            raise DesktopError("DESKTOP_DISCONNECTED", f"Windows PowerShell disconnected (exit {detail}).",
                               "Inspect the last recorded operation and local PowerShell policy") from exc
        deadline = time.monotonic() + timeout
        startup = []
        while True:
            try:
                remaining = max(0, deadline - time.monotonic())
                line = self.replies.get(timeout=min(0.25, remaining) if action == "document-preferences" else remaining)
            except queue.Empty:
                if action == "document-preferences" and time.monotonic() < deadline:
                    self._preference_notice(self.last)
                    continue
                self.broken = True
                self.host.terminate()  # Only the owned bridge process.
                code = ("ZOTERO_DISPATCH_UNCONFIRMED" if action == "bootstrap" else
                        "ZOTERO_TIMEOUT" if action in ("launch-zotero", "zotero-alive") else "WORD_TIMEOUT")
                raise DesktopError(code, f"The Windows desktop did not confirm {action} in {timeout}s.",
                                   "Inspect the last confirmed native operation before resuming")
            if line is None:
                self.broken = True
                detail = " ".join(startup)[-1500:]
                raise DesktopError("DESKTOP_DISCONNECTED", detail or "Windows PowerShell closed the session.",
                                   "Check this desktop's PowerShell policy and the last native receipt")
            try:
                reply = json.loads(line)
            except ValueError:
                startup = (startup + [line])[-8:]
                continue
            if not isinstance(reply, dict) or reply.get("id") != request["id"]:
                raise DesktopError("DESKTOP_REPLY_MISMATCH", "The desktop reply belongs to another request.",
                                   "Resume from the last matching receipt")
            if not reply.get("ok"):
                raise _failure(action, reply)
            if observation is not None and reply.get("focusRecovery"):
                observation["focusRecovery"] = reply["focusRecovery"]
            return reply.get("value")

    def word(self, action, *args, timeout=45):
        if action in ("connect", "state"):
            return self._rpc(action, *args, timeout=timeout)
        source, snapshot = self.protected[str(args[0])]
        return self._rpc(action, args[0], source, snapshot, *args[1:], timeout=timeout)

    def bind_copy(self, target, cfg):
        self.protected[str(target)] = cfg["source"], cfg["snapshot"]

    def process_alive(self, pid):
        if type(pid) is not int or pid <= 0:
            raise ValueError("Invalid Zotero process identity in the native job")
        return self._rpc("zotero-alive", pid, timeout=8)

    def launch_zotero(self):
        return self._rpc("launch-zotero", timeout=10)

    def render_pages(self, document, folder):
        observed = self._rpc("pages-open", document, timeout=60)
        try:
            count = observed["page_count"]
            if type(count) is not int or count <= 0:
                raise ValueError("Word did not report a positive page count")
            for number in range(1, count + 1):
                destination = folder / f"page-{number:04}.png"
                if self._rpc("page-image", number, destination, timeout=30) != number or not destination.stat().st_size:
                    raise ValueError(f"Word did not return a complete image for page {number}")
            return observed
        finally:
            if not self.broken:
                self._rpc("pages-close", timeout=15)

    @contextmanager
    def desktop_lock(self, job):
        path = Path(tempfile.gettempdir()) / "zotero-word-desktop.lock"
        locks = self.api._helper("file_lock")
        with path.open("a+b") as lock:
            try:
                try:
                    locks.lock_descriptor(lock.fileno(), blocking=False)
                except BlockingIOError as exc:
                    raise DesktopError("DESKTOP_BUSY", "Another run is using the Word/Zotero desktop.",
                                       "Resume when the active native operation finishes", retry=True) from exc
                try:
                    yield
                finally:
                    locks.unlock_descriptor(lock.fileno())
            except DesktopError as error:
                self._record_failure(job, error)
                raise

    def _retire_session(self):
        if self.first is not None and self.first.is_dir():
            (self.first / "session.stop").touch(exist_ok=True)
        self.first = self.last = self.pid = None

    def close(self):
        try:
            self._retire_session()
        finally:
            if self.host:
                try:
                    self.host.stdin.close()
                except OSError:
                    pass
                try:
                    self.host.wait(timeout=3)
                except subprocess.TimeoutExpired:
                    self.host.terminate()
                    self.host.wait()
                self.host = None

    def _handoff(self, job):
        if not self.process_alive(self.pid):
            return False
        slot = self.last / "next.json"
        temp = self.last / (".next-" + uuid.uuid4().hex)
        try:
            temp.write_bytes((job / "request.json").read_bytes())
            temp.rename(slot)  # On Windows this never replaces a claimed or retired slot.
        except FileExistsError:
            if slot.stat().st_size:
                raise DesktopError("ZOTERO_HANDOFF_CONFLICT", "The preceding job already has a successor.",
                                   "Inspect the preceding job before resuming")
            return False
        finally:
            temp.unlink(missing_ok=True)
        return True

    def dispatch(self, job, cfg):
        if job.parent != self.root or job != job.resolve() or not JOB_NAME.fullmatch(job.name):
            raise ValueError("Native job must be a private child of the current run")
        if cfg["libraryOnly"]:
            operation = "library_import" if cfg.get("imports") else "library_read"
        else:
            operation = "reopen_verify" if cfg["reopenOnly"] else "refresh"
        request = {"job": job.name, "operation": operation, "document": cfg.get("document"), "executor": EXECUTOR}
        while True:
            request["session"] = self.first.name if self.first else job.name
            self.api.output(job / "request.json", request)
            status = {"job": str(job), "session": request["session"], "operation": operation,
                      "via": "session" if self.first else "editor"}
            self.report({**status, "state": "awaiting_start"})
            if self.first is None:
                self.first = job
                process = {}
                self.pid = self._rpc("bootstrap", _loader(job), job, operation, timeout=55, observation=process)
                if type(self.pid) is not int or self.pid <= 0:
                    raise DesktopError("ZOTERO_PROCESS", "Invalid bound Zotero process.", "Inspect the job receipt")
                self.api.output(job / "process.json", {"pid": self.pid, **process})
                self.report({"desktop_event": "zotero_session_starts"})
                break
            if self._handoff(job):
                break
            self._retire_session()
        self.last = job
        receipt = self.wait(job, "started.json", 12)
        if receipt != request:
            raise DesktopError("ZOTERO_RECEIPT_MISMATCH", "Zotero acknowledged a different job, session or document.",
                               "Inspect request.json and started.json for this job")
        self.report({**status, "state": "accepted", "pid": self.pid, "executor": EXECUTOR, "receipt": str(job / "started.json")})

    def _preference_notice(self, job):
        if job and job != self.prompted_job and (job / "preferences-user.json").is_file():
            notice = self.api.load(job / "preferences-user.json")
            self.prompted_job = job
            self.report({"desktop_event": "preference_prompts"})
            self.report({"job": str(job), "state": "awaiting_document_preferences", **notice})
            print(notice["message"], notice["style"], file=sys.stderr, flush=True)

    def wait(self, job, name, seconds):
        end = time.monotonic() + seconds
        while time.monotonic() < end:
            self._preference_notice(job)
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

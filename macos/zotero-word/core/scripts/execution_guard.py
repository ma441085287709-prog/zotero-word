"""Original and working-file protections for a macOS citation run."""
from contextlib import contextmanager
import fcntl
import hashlib
import os
from pathlib import Path
import stat
import uuid


def sha(path):
    with Path(path).open("rb") as stream:
        value = hashlib.sha256()
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            value.update(chunk)
    return value.hexdigest()


def independent(path, *protected):
    path = Path(path).absolute()
    if path != path.resolve() or path.drive.startswith("\\\\") or path.suffix.lower() != ".docx":
        raise ValueError("Use an independent local DOCX path: " + str(path))
    if path.exists():
        st = path.stat()
        if not stat.S_ISREG(st.st_mode) or st.st_nlink != 1:
            raise ValueError("Working document must be an independent regular file")
    for other in map(Path, protected):
        if path == other.resolve() or (path.exists() and other.exists() and path.samefile(other)):
            raise ValueError("Working document aliases a protected file")
    return path


def snapshot(source, destination):
    """Publish a read-only snapshot exclusively; never chmod the original."""
    return publish_copy(source, destination, read_only=True)


def publish_copy(source, destination, *, expected=None, read_only=False):
    """Link a complete, verified copy into place without replacing an existing file."""
    source = Path(source)
    destination = independent(destination, source)
    if destination.exists():
        saved = sha(destination)
        if ((not read_only or not destination.stat().st_mode & stat.S_IWUSR)
                and sha(source) == saved and (expected is None or saved == expected)):
            return saved
        raise ValueError("The existing destination differs from the bound copy: " + str(destination))
    destination.parent.mkdir(parents=True, exist_ok=True)
    temp = destination.parent / (".snapshot-" + uuid.uuid4().hex + ".tmp")
    out = temp.open("xb")
    try:
        with out:
            with source.open("rb") as stream:
                digest = hashlib.sha256()
                for chunk in iter(lambda: stream.read(1024 * 1024), b""):
                    digest.update(chunk)
                    out.write(chunk)
            out.flush()
            os.fsync(out.fileno())
        value = digest.hexdigest()
        if sha(source) != value or (expected is not None and value != expected):
            raise ValueError("Source changed during capture: " + str(source))
        if read_only:
            temp.chmod(stat.S_IRUSR)
        os.link(temp, destination)
        return value
    finally:
        temp.unlink(missing_ok=True)


@contextmanager
def exclusive(directory):
    """An advisory lock is released by macOS when the controller exits."""
    with (directory / ".writer.lock").open("a+b") as lock:
        fcntl.flock(lock, fcntl.LOCK_EX | fcntl.LOCK_NB)
        try:
            yield
        finally:
            fcntl.flock(lock, fcntl.LOCK_UN)

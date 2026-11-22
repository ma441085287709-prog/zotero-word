"""Cooperating-process locks for local saved-file indexes and batch journals.

These short file locks do not replace durable native-operation leases. Windows
uses a fixed one-byte region at offset zero; unrelated I/O errors propagate.
"""
from __future__ import annotations

import errno
import os
import time


def lock_descriptor(descriptor, *, blocking, timeout_seconds=10):
    if os.name != "nt":
        import fcntl
        flags = fcntl.LOCK_EX | (0 if blocking else fcntl.LOCK_NB)
        fcntl.flock(descriptor, flags)
        return

    import msvcrt
    contention = {errno.EACCES, errno.EAGAIN, errno.EDEADLK}
    deadline = time.monotonic() + timeout_seconds
    while True:
        os.lseek(descriptor, 0, os.SEEK_SET)
        try:
            msvcrt.locking(descriptor, msvcrt.LK_NBLCK, 1)
            return
        except OSError as exc:
            if exc.errno not in contention:
                raise
            if not blocking or time.monotonic() >= deadline:
                raise BlockingIOError(errno.EAGAIN, "Another process holds the file lock") from exc
            time.sleep(0.05)


def unlock_descriptor(descriptor):
    if os.name != "nt":
        import fcntl
        fcntl.flock(descriptor, fcntl.LOCK_UN)
        return
    import msvcrt
    os.lseek(descriptor, 0, os.SEEK_SET)
    msvcrt.locking(descriptor, msvcrt.LK_UNLCK, 1)

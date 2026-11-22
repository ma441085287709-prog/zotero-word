#!/usr/bin/env python3
"""Compatibility entrypoint; the sole implementation lives in core/scripts."""
from importlib.util import module_from_spec, spec_from_file_location
from pathlib import Path
import sys

_core_path = Path(__file__).resolve().parents[1] / "core" / "scripts" / "zotero_docx.py"
_spec = spec_from_file_location("_zotero_word_core", _core_path)
if _spec is None or _spec.loader is None:
    raise RuntimeError("Cannot load the bundled Zotero Word core")
_core = module_from_spec(_spec)
_spec.loader.exec_module(_core)

if __name__ == "__main__":
    sys.exit(_core.main())

# Preserve existing read-only Python callers as well as command-line callers.
globals().update({name: value for name, value in vars(_core).items()
                  if not name.startswith("_")})

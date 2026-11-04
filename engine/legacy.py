"""Bridge to the proven ``apply.py`` pipeline.

The engine deliberately *reuses* rather than re-implements the hard-won parts
of `apply.py`: the KrutiDev→Unicode converter, the English/KrutiDev
classifier, the reportlab styles, ``build_table`` and the paged canvas.
Importing that module here (once) keeps a single source of truth.

``apply.py`` only runs its CLI under ``if __name__ == '__main__'``, so
importing it is side-effect free.
"""
from __future__ import annotations

import importlib.util
import sys
from pathlib import Path
from types import ModuleType

_MODULE: ModuleType | None = None


def _find_apply() -> Path | None:
    """Locate ``apply.py`` — project root, or one level above this package."""
    here = Path(__file__).resolve()
    for cand in (here.parent.parent / 'apply.py',        # <root>/apply.py
                 Path.cwd() / 'apply.py'):
        if cand.is_file():
            return cand
    return None


def apply_module() -> ModuleType:
    """Return the imported ``apply`` module (cached)."""
    global _MODULE
    if _MODULE is not None:
        return _MODULE

    try:                                   # normal case: root on sys.path
        import apply as _a                 # type: ignore
        _MODULE = _a
        return _a
    except ImportError:
        pass

    path = _find_apply()
    if path is None:
        raise RuntimeError('apply.py not found next to the engine/ package')
    spec = importlib.util.spec_from_file_location('apply', str(path))
    assert spec and spec.loader
    mod = importlib.util.module_from_spec(spec)
    sys.modules['apply'] = mod
    spec.loader.exec_module(mod)
    _MODULE = mod
    return mod


# Frequently used names, resolved lazily so importing this module is cheap.
def __getattr__(name: str):
    return getattr(apply_module(), name)

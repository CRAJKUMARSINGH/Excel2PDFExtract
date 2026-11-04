"""Per-sheet configuration schema + ``layouts.json`` persistence.

A ``WorkbookConfig`` is the single source of truth the web UI edits and the
renderer consumes.  Everything is plain dataclasses + JSON so it round-trips
cleanly through the browser.

    WorkbookConfig
      defaults : SheetConfig            # applied to any sheet without an override
      sheets   : {name: SheetConfig}    # per-sheet overrides
"""
from __future__ import annotations

import json
from dataclasses import asdict, dataclass, field, fields
from pathlib import Path
from typing import Any

SCHEMA_VERSION = 1


# ---------------------------------------------------------------------------
# column helpers
# ---------------------------------------------------------------------------
def col_letter(idx: int) -> str:
    """0 -> 'A', 25 -> 'Z', 26 -> 'AA' (0-based index)."""
    s = ''
    idx += 1
    while idx:
        idx, rem = divmod(idx - 1, 26)
        s = chr(65 + rem) + s
    return s


def col_index(letter: str) -> int:
    """'A' -> 0, 'AA' -> 26 (inverse of :func:`col_letter`)."""
    n = 0
    for ch in letter.strip().upper():
        n = n * 26 + (ord(ch) - 64)
    return n - 1


# ---------------------------------------------------------------------------
# nested config blocks
# ---------------------------------------------------------------------------
@dataclass
class ColumnConfig:
    """Which columns print, in what order, at what width."""
    mode: str = 'visible'                 # 'visible' | 'native' | 'explicit'
    include: list[str] = field(default_factory=list)
    exclude: list[str] = field(default_factory=list)   # e.g. ["M","N","O","P","Q"]
    order: list[str] = field(default_factory=list)     # drag-reordered letters
    widths: dict[str, float] = field(default_factory=dict)  # letter -> pt


@dataclass
class RowConfig:
    """1-based inclusive row window."""
    start: int = 1
    end: int | None = None                # None = to the last row
    skip_blank: bool = True


@dataclass
class PageBreakConfig:
    """How the sheet is cut into pages."""
    strategy: str = 'signature'           # 'signature'|'title'|'n_rows'|'continuous'
    every: int = 25                       # used when strategy == 'n_rows'


@dataclass
class TextConfig:
    """KrutiDev / font handling."""
    kruti: str = 'auto'                   # 'auto' | 'on' | 'off'
    kruti_columns: list[str] = field(default_factory=list)   # force-convert
    force_english: list[str] = field(default_factory=list)   # never convert
    font: str = 'Deva'


@dataclass
class PageConfig:
    orientation: str = 'landscape'        # 'landscape' | 'portrait'
    margin_mm: float = 12.0
    repeat_header: bool = True


# ---------------------------------------------------------------------------
# per-sheet config
# ---------------------------------------------------------------------------
@dataclass
class SheetConfig:
    layout_id: str = 'auto'
    columns: ColumnConfig = field(default_factory=ColumnConfig)
    rows: RowConfig = field(default_factory=RowConfig)
    page_breaks: PageBreakConfig = field(default_factory=PageBreakConfig)
    text: TextConfig = field(default_factory=TextConfig)
    page: PageConfig = field(default_factory=PageConfig)

    # -- (de)serialisation --------------------------------------------------
    def to_dict(self) -> dict:
        return asdict(self)

    @classmethod
    def from_dict(cls, d: dict | None) -> 'SheetConfig':
        """Build from a (possibly partial) dict, ignoring unknown keys.

        ``from __future__ import annotations`` makes ``field.type`` a *string*,
        so the nested-block classes are looked up explicitly instead.
        """
        d = d or {}
        nested = {
            'columns': ColumnConfig,
            'rows': RowConfig,
            'page_breaks': PageBreakConfig,
            'text': TextConfig,
            'page': PageConfig,
        }
        kwargs: dict[str, Any] = {}
        for f in fields(cls):
            if f.name not in d:
                continue
            val = d[f.name]
            sub = nested.get(f.name)
            if isinstance(val, dict) and sub is not None:
                kwargs[f.name] = sub(**{k: v for k, v in val.items()
                                        if k in sub.__dataclass_fields__})
            else:
                kwargs[f.name] = val
        return cls(**kwargs)

    # -- merge --------------------------------------------------------------
    def merged(self, override: 'SheetConfig') -> 'SheetConfig':
        """Return a copy where *override*'s non-default fields win.

        Used to layer: LAYOUT preset  <-  WorkbookConfig.defaults  <-  sheet.
        """
        base = SheetConfig.from_dict(self.to_dict())
        for f in fields(SheetConfig):
            ov = getattr(override, f.name)
            bv = getattr(base, f.name)
            if hasattr(ov, '__dataclass_fields__'):
                # nested block: field-by-field, keeping base where override is default
                for sf in fields(type(ov)):
                    if getattr(ov, sf.name) != getattr(type(ov)(), sf.name):
                        setattr(bv, sf.name, getattr(ov, sf.name))
            elif ov != getattr(type(base)(), f.name):
                setattr(base, f.name, ov)
        return base


# ---------------------------------------------------------------------------
# workbook config (the layouts.json payload)
# ---------------------------------------------------------------------------
@dataclass
class WorkbookConfig:
    version: int = SCHEMA_VERSION
    source: str = ''
    defaults: SheetConfig = field(default_factory=SheetConfig)
    sheets: dict[str, SheetConfig] = field(default_factory=dict)

    # -- helpers ------------------------------------------------------------
    def for_sheet(self, name: str) -> SheetConfig:
        """Effective config = LAYOUT preset <- defaults <- sheet override."""
        from . import layouts as _layouts            # local import: avoid cycle
        cfg = self.defaults.merged(self.sheets.get(name, SheetConfig()))
        if cfg.layout_id != 'auto':
            cfg = _layouts.preset(cfg.layout_id).merged(cfg)
        return cfg

    # -- (de)serialisation --------------------------------------------------
    def to_dict(self) -> dict:
        return {
            'version': self.version,
            'source': self.source,
            'defaults': self.defaults.to_dict(),
            'sheets': {k: v.to_dict() for k, v in self.sheets.items()},
        }

    @classmethod
    def from_dict(cls, d: dict) -> 'WorkbookConfig':
        return cls(
            version=int(d.get('version', SCHEMA_VERSION)),
            source=d.get('source', ''),
            defaults=SheetConfig.from_dict(d.get('defaults')),
            sheets={k: SheetConfig.from_dict(v)
                    for k, v in (d.get('sheets') or {}).items()},
        )

    def save(self, path: str | Path) -> Path:
        p = Path(path)
        p.write_text(json.dumps(self.to_dict(), indent=2, ensure_ascii=False),
                     encoding='utf-8')
        return p

    @classmethod
    def load(cls, path: str | Path) -> 'WorkbookConfig':
        return cls.from_dict(json.loads(Path(path).read_text(encoding='utf-8')))

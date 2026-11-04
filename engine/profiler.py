"""Sheet structure profiler.

Turns a raw :class:`~engine.reader.Grid` into a :class:`SheetProfile`:
where the titles, the column-header row, the signature lines and the numeric
blocks are, how much KrutiDev text the sheet holds, and *which layout preset
fits best*.  Nothing is hard-coded to the `Details` sheet, so arbitrary
workbooks profile correctly.

The heuristics reuse the regexes and classifiers from ``apply.py`` (via
:mod:`engine.legacy`) so profiling and rendering never disagree.
"""
from __future__ import annotations

import re
from dataclasses import asdict, dataclass, field

from .reader import Grid
from . import legacy

# ---------------------------------------------------------------------------
# small predicates
# ---------------------------------------------------------------------------
NUMERIC_RE = re.compile(r'^-?\s*[\d,]+(\.\d+)?\s*%?$')
SHORT_LEN = 28            # a header cell is "short"
HEAD_MIN_CELLS = 3        # at least this many non-empty cells to look like a header


def _is_numeric(s: str) -> bool:
    return bool(NUMERIC_RE.match(s.strip()))


def _join(row) -> str:
    return ' '.join(str(c) for c in row if str(c).strip())


# ---------------------------------------------------------------------------
# profile
# ---------------------------------------------------------------------------
@dataclass
class SheetProfile:
    name: str
    n_rows: int
    n_cols: int
    col_names: list[str] = field(default_factory=list)      # real Excel letters
    hidden_letters: list[str] = field(default_factory=list)
    title_rows: list[int] = field(default_factory=list)
    header_row: int | None = None
    signature_rows: list[int] = field(default_factory=list)
    numeric_density: float = 0.0
    kruti_ratio: float = 0.0
    merged_count: int = 0
    guessed_layout: str = 'auto'
    confidence: float = 0.0
    reasons: list[str] = field(default_factory=list)

    def to_dict(self) -> dict:
        return asdict(self)


def profile(grid: Grid) -> SheetProfile:
    """Analyse *grid* and guess the best layout preset."""
    a = legacy.apply_module()
    TITLE_RE, HEAD_RE = a.TITLE_RE, a.HEAD_RE
    SIG = (a.SIG1_RE, a.SIG2_RE, a.SIG3_RE)
    is_kruti, is_english = a.is_kruti, a.is_english

    rows = grid.values
    names = grid.col_names()
    prof = SheetProfile(name=grid.name, n_rows=grid.n_rows, n_cols=grid.n_cols,
                        col_names=names,
                        hidden_letters=[names[i] for i in sorted(grid.hidden_cols)
                                        if i < len(names)],
                        merged_count=len(grid.merged))

    if not rows:
        prof.guessed_layout = 'plain-grid'
        prof.reasons.append('empty sheet')
        return prof

    # -- scan rows ----------------------------------------------------------
    text_cells = kruti_cells = 0
    for i, row in enumerate(rows):
        joined = _join(row)

        if TITLE_RE.search(joined) and not HEAD_RE.search(joined):
            prof.title_rows.append(i)

        if any(p.search(joined) for p in SIG):
            prof.signature_rows.append(i)

        if HEAD_RE.search(joined) and prof.header_row is None:
            prof.header_row = i

        for c in row:
            s = str(c).strip()
            if not s:
                continue
            if any(ch.isalpha() for ch in s):
                text_cells += 1
                # a cell is KrutiDev when it is ASCII-Latin AND not English
                if is_kruti(s) and not is_english(s):
                    kruti_cells += 1

    if text_cells:
        prof.kruti_ratio = round(kruti_cells / text_cells, 3)

    # -- numeric density (over non-empty cells) -----------------------------
    total = numeric = 0
    for row in rows:
        for c in row:
            s = str(c).strip()
            if s:
                total += 1
                numeric += _is_numeric(s)
    prof.numeric_density = round(numeric / total, 3) if total else 0.0

    # -- header row fallback -------------------------------------------------
    if prof.header_row is None:
        prof.header_row = _guess_header_row(rows)

    # -- layout guess --------------------------------------------------------
    prof.guessed_layout, prof.confidence = _guess_layout(prof, rows)
    return prof


def _guess_header_row(rows) -> int | None:
    """First row that looks like a column-header band.

    Heuristic: >=3 non-empty cells, most of them short and non-numeric, and the
    next few rows carry numbers (i.e. a table follows).
    """
    for i, row in enumerate(rows[:-1]):
        cells = [str(c).strip() for c in row if str(c).strip()]
        if len(cells) < HEAD_MIN_CELLS:
            continue
        short = sum(1 for c in cells if len(c) <= SHORT_LEN and not _is_numeric(c))
        if short / len(cells) < 0.6:
            continue
        below = rows[i + 1:i + 4]
        if any(_is_numeric(str(c)) for r in below for c in r):
            return i
    return None


def _guess_layout(prof: SheetProfile, rows) -> tuple[str, float]:
    """Map the profile onto a preset id + a 0..1 confidence."""
    r = prof.reasons
    if prof.n_cols <= 2 and prof.numeric_density < 0.25:
        r.append('wide prose in <=2 columns')
        return 'cover-letter', 0.75
    if prof.title_rows and prof.signature_rows and prof.n_cols >= 6:
        r.append(f'{len(prof.title_rows)} titled blocks x {len(prof.signature_rows)} signature lines')
        return 'budget-details', 0.9
    if prof.kruti_ratio >= 0.5 and prof.header_row is None:
        r.append(f'{int(prof.kruti_ratio * 100)}% KrutiDev, no header row')
        return 'cover-letter', 0.8
    if prof.numeric_density < 0.05 and not prof.title_rows and not prof.signature_rows:
        r.append('almost no numbers, no titles/signatures')
        return 'plain-grid', 0.55
    if prof.n_cols >= 10:
        r.append(f'{prof.n_cols} columns - continuous table')
        return 'flow-table', 0.6
    if prof.header_row is not None and prof.numeric_density >= 0.20:
        r.append(f'header row at {prof.header_row + 1}, numeric density {prof.numeric_density}')
        return 'summary-table', 0.7
    r.append('no strong signal - safe default')
    return 'plain-grid', 0.4

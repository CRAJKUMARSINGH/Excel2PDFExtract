"""Workbook I/O — produce a :class:`Grid` for any worksheet, any format.

``openpyxl`` is used for ``.xlsx`` (it exposes hidden columns and merged
ranges), ``xlrd`` for legacy ``.xls``.  The returned :class:`Grid` is the only
thing the profiler/renderer ever see, so sheet-format differences stop here.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

try:
    from openpyxl import load_workbook
except ImportError:  # pragma: no cover
    load_workbook = None

try:
    import xlrd as _xlrd
    _HAVE_XLRD = True
except ImportError:  # pragma: no cover
    _HAVE_XLRD = False

from .schema import col_letter


@dataclass
class Grid:
    """A worksheet reduced to what printing needs."""
    name: str
    values: list[list[str]]                 # trimmed values, row-major, 0-based
    hidden_cols: set[int] = field(default_factory=set)   # 0-based col indices
    merged: list[tuple[int, int, int, int]] = field(default_factory=list)  # r1,c1,r2,c2
    kind: str = 'xlsx'
    col_offset: int = 0                     # leading empty cols trimmed off
    orig_cols: int = 0                      # column count before trimming

    @property
    def n_rows(self) -> int:
        return len(self.values)

    @property
    def n_cols(self) -> int:
        return len(self.values[0]) if self.values else 0

    def col_names(self) -> list[str]:
        """**Original** Excel column letters, e.g. ['A','N','O',...]."""
        return [col_letter(self.col_offset + i) for i in range(self.n_cols)]


# ---------------------------------------------------------------------------
# public API
# ---------------------------------------------------------------------------
def list_sheets(path: str | Path) -> list[str]:
    """Worksheet names of *path* (``.xls`` or ``.xlsx``)."""
    p = Path(path)
    if p.suffix.lower() == '.xls':
        if not _HAVE_XLRD:
            raise RuntimeError('Reading .xls requires xlrd (pip install xlrd==1.2.0)')
        return list(_xlrd.open_workbook(str(p)).sheet_names())
    _require_openpyxl()
    wb = load_workbook(str(p), read_only=True)
    try:
        return list(wb.sheetnames)
    finally:
        wb.close()


def read_grid(path: str | Path, sheet: str) -> Grid:
    """Read one worksheet into a :class:`Grid` (empty edge rows/cols trimmed)."""
    p = Path(path)
    if p.suffix.lower() == '.xls':
        return _read_xls(p, sheet)
    return _read_xlsx(p, sheet)


def read_all(path: str | Path) -> dict[str, Grid]:
    return {name: read_grid(path, name) for name in list_sheets(path)}


# ---------------------------------------------------------------------------
# xlsx
# ---------------------------------------------------------------------------
def _read_xlsx(path: Path, sheet: str) -> Grid:
    _require_openpyxl()
    # read_only=False: hidden column_dimensions and merged_cells are only
    # fully available on a normally-loaded workbook.
    wb = load_workbook(str(path), data_only=True, read_only=False)
    try:
        ws = wb[sheet]
        raw = [['' if c is None else str(c) for c in row]
               for row in ws.iter_rows(values_only=True)]

        hidden = {i for i in range(ws.max_column)
                  if ws.column_dimensions[col_letter(i)].hidden}
        merged = [(mr.min_row - 1, mr.min_col - 1, mr.max_row - 1, mr.max_col - 1)
                  for mr in ws.merged_cells.ranges]
        orig_cols = ws.max_column
    finally:
        wb.close()

    values, off = _trim(raw)
    # Re-base hidden columns and merged ranges onto the trimmed grid.
    hidden = {i - off for i in hidden if i - off >= 0}
    merged = [(r1, c1 - off, r2, c2 - off)
              for (r1, c1, r2, c2) in merged if c2 - off >= 0]
    return Grid(sheet, values, hidden, merged, kind='xlsx',
                col_offset=off, orig_cols=orig_cols)


# ---------------------------------------------------------------------------
# xls
# ---------------------------------------------------------------------------
def _read_xls(path: Path, sheet: str) -> Grid:
    if not _HAVE_XLRD:
        raise RuntimeError('Reading .xls requires xlrd (pip install xlrd==1.2.0)')
    book = _xlrd.open_workbook(str(path), formatting_info=True)
    sh = book.sheet_by_name(sheet)
    values = [[sh.cell_value(r, c) for c in range(sh.ncols)]
              for r in range(sh.nrows)]
    values = [[str(v) for v in row] for row in values]

    hidden: set[int] = set()
    # TODO: xlrd exposes colinfo via sh.colinfo_map (hidden flag).  Map it to
    # 0-based offsets here; left as a stub until an .xls sample with hidden
    # columns is available to validate against.
    merged = [(rlo, clo, rhi - 1, chi - 1)
              for (rlo, rhi, clo, chi) in sh.merged_cells]
    orig_cols = sh.ncols

    values, off = _trim(values)
    hidden = {i - off for i in hidden if i - off >= 0}
    merged = [(r1, c1 - off, r2, c2 - off)
              for (r1, c1, r2, c2) in merged if c2 - off >= 0]
    return Grid(sheet, values, hidden, merged, kind='xls',
                col_offset=off, orig_cols=orig_cols)


# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------
def _trim(values: list[list[str]]) -> tuple[list[list[str]], int]:
    """Drop all-empty leading/trailing rows & cols; return (rows, col_offset).

    ``col_offset`` is how many leading columns were removed, so callers can map
    trimmed indices back to the workbook's real column letters.
    """
    rows = [[str(c) for c in r] for r in values]
    while rows and not any(c.strip() for c in rows[0]):
        rows.pop(0)
    while rows and not any(c.strip() for c in rows[-1]):
        rows.pop()
    if not rows:
        return [], 0
    ncols = max(len(r) for r in rows)
    rows = [(r + [''] * (ncols - len(r))) for r in rows]
    used = [i for i in range(ncols) if any(r[i].strip() for r in rows)]
    if not used:
        return [], 0
    lo, hi = used[0], used[-1]
    return [r[lo:hi + 1] for r in rows], lo


def _require_openpyxl() -> None:
    if load_workbook is None:
        raise RuntimeError('Reading .xlsx requires openpyxl (pip install openpyxl)')

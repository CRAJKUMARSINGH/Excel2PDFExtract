"""Generalised renderer — :class:`Grid` + :class:`RenderPlan` -> PDF.

The heavy lifting (Devanagari fonts, KrutiDev mapping, LongTable pagination,
table styling) is delegated to ``apply.py`` through :mod:`engine.legacy`; this
module only adds the *variation*: orientation, segmentation strategy, column
subsetting and per-column text handling.

Themes currently wired: ``ruled`` (the proven budget-details look) and
``light`` (plain grid).  ``prose`` and ``totals`` share the ruled path today —
see the TODOs.
"""
from __future__ import annotations

import datetime as dt
from dataclasses import dataclass
from pathlib import Path

from reportlab.lib.pagesizes import A4, landscape, portrait
from reportlab.platypus import (
    BaseDocTemplate, Frame, PageBreak, PageTemplate,
)

from . import legacy
from .layouts import RenderPlan
from .reader import Grid
from .schema import col_index


@dataclass
class RenderResult:
    path: Path
    pages: int
    sections: int


# ---------------------------------------------------------------------------
# public
# ---------------------------------------------------------------------------
def render(grid: Grid, plan: RenderPlan, out: Path, font_dir: Path) -> RenderResult:
    """Render one sheet to *out* and return page/section counts."""
    a = legacy.apply_module()

    rows = _project(grid, plan, a)
    if not rows:
        raise ValueError(f'[{plan.sheet}] no printable rows after column/row selection')

    segs = _segment(rows, plan, a) or [rows]
    mapping = a.register_fonts(Path(font_dir))
    pages = _build(segs, out, grid, mapping, plan, a)
    return RenderResult(path=Path(out), pages=pages, sections=len(segs))


def estimate_sections(grid: Grid, plan: RenderPlan) -> int:
    """Cheap page-block estimate for the UI (no PDF written)."""
    a = legacy.apply_module()
    rows = _project(grid, plan, a)
    return len(_segment(rows, plan, a) or ([rows] if rows else []))


# ---------------------------------------------------------------------------
# row/column projection + text handling
# ---------------------------------------------------------------------------
def _project(grid: Grid, plan: RenderPlan, a) -> list[list[str]]:
    """Apply the row window, the column selection and the text conversion."""
    rs = max(1, plan.row_start) - 1
    re_ = plan.row_end if plan.row_end else grid.n_rows
    window = grid.values[rs:re_]

    if plan.theme == 'prose' or (plan.col_letters and len(plan.col_letters) == 1):
        # Prose sheets: collapse every row to a single wide cell.
        rows = [['  '.join(c for c in r if str(c).strip())] for r in window]
        return [[_convert(v, 'A', plan, a) for v in r] for r in rows]

    # plan.col_letters are the workbook's real letters; map them onto the
    # trimmed grid (which may have dropped leading empty columns).
    idxs = [i for i in (col_index(l) - grid.col_offset for l in plan.col_letters)
            if 0 <= i < grid.n_cols]
    out: list[list[str]] = []
    for r in window:
        row = [(r[i] if i < len(r) else '') for i in idxs]
        row = [_convert(v, plan.col_letters[j], plan, a)
               for j, v in enumerate(row)]
        out.append(row)
    return out


def _convert(text: str, letter: str, plan: RenderPlan, a) -> str:
    """Per-cell KrutiDev policy for the resolved plan."""
    s = str(text)
    if not s.strip():
        return s
    if letter.upper() in {c.upper() for c in plan.force_english}:
        return s
    if plan.kruti_mode == 'off':
        return s
    if plan.kruti_mode == 'on':
        return a.krutidev_to_unicode(s)
    if plan.kruti_columns and letter.upper() not in {c.upper() for c in plan.kruti_columns}:
        return s
    # auto: convert only what really looks like KrutiDev
    if a.is_kruti(s) and not a.is_english(s):
        return a.krutidev_to_unicode(s)
    return s


# ---------------------------------------------------------------------------
# segmentation strategies
# ---------------------------------------------------------------------------
def _segment(rows, plan: RenderPlan, a) -> list[list]:
    st = plan.break_strategy
    if st == 'signature':
        return a.segment(rows)
    if st == 'title':
        segs, cur = [], []
        for r in rows:
            joined = ' '.join(str(c) for c in r if str(c).strip())
            if a.TITLE_RE.search(joined) and not a.HEAD_RE.search(joined) and cur:
                segs.append(cur)
                cur = []
            cur.append(r)
        if cur:
            segs.append(cur)
        return segs
    if st == 'n_rows':
        n = max(1, plan.break_every)
        return [rows[i:i + n] for i in range(0, len(rows), n)]
    return [rows]          # 'continuous' / 'whole'


# ---------------------------------------------------------------------------
# document build
# ---------------------------------------------------------------------------
def _build(segs, out: Path, grid: Grid, mapping, plan: RenderPlan, a) -> int:
    page = landscape(A4) if plan.orientation == 'landscape' else portrait(A4)
    margin = plan.margin_mm
    doc = BaseDocTemplate(
        str(out), pagesize=page,
        leftMargin=margin * a.mm, rightMargin=margin * a.mm,
        topMargin=10 * a.mm, bottomMargin=12 * a.mm,
        title=f'{grid.name} - {plan.layout_id}',
        author='Excel-PDF Print Studio',
        subject=f'layout={plan.layout_id} theme={plan.theme}',
    )
    frame = Frame(doc.leftMargin, doc.bottomMargin, doc.width, doc.height,
                  leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0,
                  id='main')
    doc.addPageTemplates([PageTemplate(id='p', frames=[frame])])

    styles = a.make_styles(mapping)
    ncols = len(segs[0][0]) if segs and segs[0] else 1

    # Per-column widths: explicit override -> proportion -> uniform fallback
    usable = doc.width
    if plan.col_widths:
        vals = [plan.col_widths.get(l, 0.0) for l in plan.col_letters] or [0.0]
        if sum(vals) <= 0:
            cw = a.col_widths(ncols, usable)
        else:
            total = sum(v for v in vals if v) or 1.0
            cw = [(v or total / ncols) / total * usable for v in vals]
        cw = (cw + [usable / ncols] * ncols)[:ncols]
    else:
        cw = a.col_widths(ncols, usable)

    story = []
    for i, seg in enumerate(segs):
        story.append(a.build_table(seg, cw, styles, mapping))
        # 'continuous' -> no forced break between blocks
        if i < len(segs) - 1 and plan.break_strategy != 'continuous':
            story.append(PageBreak())

    label = (f'{grid.name}  |  {plan.layout_id}  |  '
             f'{dt.datetime.now():%d-%b-%Y %H:%M}  |  {len(segs)} sections')
    doc.build(story, canvasmaker=lambda *ar, **kw: a._PagedCanvas(
        *ar, label=label, font=a.F_LATIN, **kw))
    return doc.page

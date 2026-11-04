"""Layout presets + plan resolution.

A **layout** is a named, data-only printing recipe (segmentation rule, column
strategy, theme).  Adding a new sheet shape means adding a row to ``LAYOUTS``,
never a new script.

``resolve()`` merges three layers into a single :class:`RenderPlan` that the
renderer can execute without guessing:

    LAYOUT preset  <  WorkbookConfig.defaults  <  sheet override  <  profile

"""
from __future__ import annotations

from dataclasses import asdict, dataclass, field

from .profiler import SheetProfile
from .schema import SheetConfig, col_letter, col_index


# ---------------------------------------------------------------------------
# layout preset
# ---------------------------------------------------------------------------
@dataclass(frozen=True)
class Layout:
    id: str
    label_en: str
    label_hi: str
    description: str
    segmentation: str          # 'signature'|'title'|'n_rows'|'continuous'|'whole'
    col_strategy: str          # 'visible'|'native'|'single'
    header_rows: str = 'auto'  # 'auto' | 'none'
    repeat_headers: bool = True
    theme: str = 'ruled'
    orientation: str = 'landscape'


LAYOUTS: dict[str, Layout] = {
    'budget-details': Layout(
        'budget-details', 'Budget details', 'बजट विवरण',
        'One page per signature block, 13 visible columns, ruled title band.',
        segmentation='signature', col_strategy='visible', theme='ruled'),
    'flow-table': Layout(
        'flow-table', 'Continuous table', 'सतत सारणी',
        'Continuous rows with the header repeated on every page.',
        segmentation='continuous', col_strategy='native', theme='zebra'),
    'cover-letter': Layout(
        'cover-letter', 'Cover letter', 'आवरण पत्र',
        'Prose sheet: one wide column, wrapped, no grid.',
        segmentation='whole', col_strategy='single', header_rows='none',
        repeat_headers=False, theme='prose', orientation='portrait'),
    'summary-table': Layout(
        'summary-table', 'Summary', 'सारांश',
        'Per-title tables with totals highlighted.',
        segmentation='title', col_strategy='native', theme='totals'),
    'plain-grid': Layout(
        'plain-grid', 'Plain grid', 'सादा ग्रिड',
        'Light grid, no segmentation, no text conversion.',
        segmentation='continuous', col_strategy='native', theme='light'),
    'auto': Layout(
        'auto', 'Auto-detect', 'स्वतः पहचान',
        'Use the layout suggested by the sheet profiler.',
        segmentation='signature', col_strategy='visible', theme='ruled'),
}


def preset(layout_id: str) -> SheetConfig:
    """A :class:`SheetConfig` seeded with *layout_id*'s defaults."""
    lay = LAYOUTS.get(layout_id, LAYOUTS['auto'])
    cfg = SheetConfig(layout_id=lay.id)
    cfg.page.orientation = lay.orientation
    cfg.page.repeat_header = lay.repeat_headers
    cfg.page_breaks.strategy = lay.segmentation
    cfg.columns.mode = {'visible': 'visible',
                        'native': 'native',
                        'single': 'native'}.get(lay.col_strategy, 'visible')
    if lay.theme == 'light':           # 'plain grid' never converts text
        cfg.text.kruti = 'off'
    return cfg


def catalogue() -> list[dict]:
    """UI-facing list of layouts (used to populate the dropdown)."""
    return [asdict(l) for l in LAYOUTS.values()]


# ---------------------------------------------------------------------------
# resolved plan
# ---------------------------------------------------------------------------
@dataclass
class RenderPlan:
    sheet: str
    layout_id: str
    theme: str
    orientation: str
    margin_mm: float
    col_letters: list[str] = field(default_factory=list)
    col_widths: dict[str, float] = field(default_factory=dict)
    header_rows: int = 0
    repeat_header: bool = True
    break_strategy: str = 'continuous'
    break_every: int = 25
    row_start: int = 1
    row_end: int | None = None
    kruti_mode: str = 'auto'
    kruti_columns: list[str] = field(default_factory=list)
    force_english: list[str] = field(default_factory=list)
    font: str = 'Deva'

    def to_dict(self) -> dict:
        return asdict(self)


def resolve(prof: SheetProfile, cfg: SheetConfig) -> RenderPlan:
    """Merge *cfg* with *prof* into an executable :class:`RenderPlan`."""
    layout_id = cfg.layout_id if cfg.layout_id != 'auto' else prof.guessed_layout
    lay = LAYOUTS.get(layout_id, LAYOUTS['plain-grid'])
    # Layer the preset under the user's config so unset fields inherit the
    # layout's behaviour (orientation, break strategy, text handling, ...).
    eff = preset(layout_id).merged(cfg)
    eff.layout_id = layout_id

    plan = RenderPlan(
        sheet=prof.name,
        layout_id=lay.id,
        theme=lay.theme,
        orientation=eff.page.orientation,
        margin_mm=eff.page.margin_mm,
        header_rows=0 if lay.header_rows == 'none' else (
            (prof.header_row + 1) if prof.header_row is not None else 0),
        repeat_header=eff.page.repeat_header,
        break_strategy=eff.page_breaks.strategy,
        break_every=eff.page_breaks.every,
        row_start=eff.rows.start,
        row_end=eff.rows.end,
        kruti_mode=eff.text.kruti,
        kruti_columns=list(eff.text.kruti_columns),
        force_english=list(eff.text.force_english),
        font=eff.text.font,
    )

    plan.col_letters = _select_columns(prof, eff, lay)
    plan.col_widths = {k.upper(): v for k, v in eff.columns.widths.items()}
    return plan


def _select_columns(prof: SheetProfile, cfg: SheetConfig, lay: Layout) -> list[str]:
    all_letters = list(prof.col_names) or [col_letter(i) for i in range(prof.n_cols)]
    hidden = set(prof.hidden_letters)

    if cfg.columns.include:                       # explicit wins
        letters = [c.upper() for c in cfg.columns.include]
    elif cfg.columns.mode == 'native' or lay.col_strategy in ('native', 'single'):
        letters = list(all_letters)
    else:                                          # 'visible' = drop hidden cols
        letters = [c for c in all_letters if c not in hidden]

    drop = {c.upper() for c in cfg.columns.exclude}
    letters = [c for c in letters if c not in drop]

    if cfg.columns.order:                          # drag-reorder (partial ok)
        want = [c.upper() for c in cfg.columns.order]
        rank = {c: i for i, c in enumerate(want)}
        letters.sort(key=lambda c: (rank.get(c, len(want)), col_index(c)))
    return letters

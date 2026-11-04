# MASTER-PLAN — Excel → A4 PDF Print Studio

> Draft master architecture for **varied** Excel sheet structures, with a buttoned
> web UI to pick a layout per sheet and tune sheet-specific options.
> Runtime: **local Python app + browser UI**. Engine: the existing `apply.py`
> pipeline (openpyxl → KrutiDev→Unicode → reportlab LongTable).

---

## 1. Goal

Today `apply.py` prints **one** shape of sheet (the 13-column `Details` budget
sheet) with one implicit layout. Real workbooks mix *many* shapes:

| Sheet in the sample workbook | Shape | Needs |
|---|---|---|
| `Letters1` | KrutiDev cover letter, prose | no table grid, full-width text blocks |
| `Cover3` | memo/notice, merged cells | full-width paragraphs, cheap grid |
| `Saransh` / `Saransh (2)` | summary table, own columns | native columns, totals |
| `Details` | 1001×50, 13 visible cols, titled blocks | one page per signature block |
| `Details (2)` | 1311×24, repeated header each page | continuous flow, repeating header |
| `FORMAT` | template/format sheet | plain grid, no conversion |

The master code delivers: **profile → suggest layout → user tunes → render**.

---

## 2. Layered architecture

```
                       ┌──────────────────────────────────────────┐
   browser (web/)      │  Sheet Board  ·  Config Drawer  ·  Preview│
                       └───────────────┬──────────────────────────┘
                                       │ JSON over HTTP
                       ┌───────────────▼──────────────────────────┐
   app.py (FastAPI)    │  /api/workbook  /api/profile  /api/plan   │
                       │  /api/preview   /api/render   /download   │
                       └───────────────┬──────────────────────────┘
                                       │
     ┌─────────────────────────────────┼───────────────────────────────┐
     │                                 │                               │
┌────▼─────┐   ┌──────────┐   ┌────────▼────────┐   ┌──────────────┐   │
│ reader   │──▶│ profiler │──▶│ layouts (registry│──▶│ renderer     │───┘
│ raw grid │   │ structure│   │ + plan resolver) │   │ reportlab    │
└──────────┘   └──────────┘   └──────────────────┘   └──────────────┘
     │                │                 │                   │
     └──── schema.py (SheetConfig / WorkbookConfig, JSON) ──┘
                            │
                    apply.py  (reused helpers:
              register_fonts · make_styles · krutidev_to_unicode ·
              col_widths · build_table · segment · _PagedCanvas)
```

**Rule:** `apply.py` stays as the legacy CLI and is *imported*, never modified by
this layer. New behaviour lives in `engine/`. (Migration note in §8.)

### Module responsibilities

| Module | Responsibility | Real logic now | TODO |
|---|---|---|---|
| `engine/reader.py` | Open xls/xlsx, return a `Grid` (values + hidden cols + merged ranges) | full | `.xls` merged ranges |
| `engine/profiler.py` | Detect titles/headers/signatures/numeric density/KrutiDev ratio → `SheetProfile` + layout guess | full | ML-free heuristics can be tuned |
| `engine/layouts.py` | `LAYOUTS` registry (6 presets) + `resolve(profile, config) → RenderPlan` | full | more presets |
| `engine/schema.py` | Dataclasses for config; load/save `layouts.json` | full | schema versioning |
| `engine/renderer.py` | `RenderPlan → PDF`; segmentation strategies; column subsetting; orientation | partial (delegates to `apply`) | `cover-letter`, `summary` themes |
| `app.py` | REST + static serving | partial | auth, job queue |
| `web/` | Buttoned UI (vanilla JS, no build step) | screens 1–4 | drag-reorder columns |

---

## 3. Data model (`engine/schema.py`)

```
WorkbookConfig
├── source: str                     # workbook filename
├── defaults: SheetConfig           # applied to any sheet without overrides
└── sheets: dict[str, SheetConfig]  # per-sheet overrides

SheetConfig
├── layout_id: str        = "auto"          # see LAYOUTS
├── columns:
│   ├── mode:  "visible" | "native" | "explicit"
│   ├── include: list[str] = []             # e.g. ["A".."L","R"] when explicit
│   ├── exclude: list[str] = []             # e.g. ["M","N","O","P","Q"]
│   ├── order:   list[str] = []             # drag-reordered letters
│   └── widths:  dict[str,float] = {}       # per-column pt override
├── rows:
│   ├── start: int = 1                      # 1-based, inclusive
│   ├── end:   int | None = None            # None = last
│   └── skip_blank: bool = True
├── page_breaks:
│   ├── strategy: "signature" | "title" | "n_rows" | "continuous"
│   └── every: int = 25                     # used by "n_rows"
├── text:
│   ├── kruti: "auto" | "on" | "off"        # conversion mode
│   ├── kruti_columns: list[str] = []       # force-convert only these
│   ├── force_english: list[str] = []       # never convert these
│   └── font: str = "Deva"
└── page:
    ├── orientation: "landscape" | "portrait"
    ├── margin_mm: float = 12.0
    └── repeat_header: bool = True
```

`layouts.json` is exactly a serialised `WorkbookConfig`, e.g.

```json
{
  "version": 1,
  "source": "FINAL Budget Anuman 31 October 2026.xlsx",
  "defaults": { "layout_id": "auto", "text": { "kruti": "auto" } },
  "sheets": {
    "Details":     { "layout_id": "budget-details",
                     "columns": { "mode": "explicit", "exclude": ["M","N","O","P","Q"] } },
    "Details (2)": { "layout_id": "flow-table",
                     "page_breaks": { "strategy": "continuous" } },
    "Letters1":    { "layout_id": "cover-letter" },
    "Saransh":     { "layout_id": "summary-table" },
    "FORMAT":      { "layout_id": "plain-grid", "text": { "kruti": "off" } }
  }
}
```

---

## 4. Layout registry (`engine/layouts.py`)

| id | Label (en / hi) | Segmentation | Columns | Theme |
|---|---|---|---|---|
| `budget-details` | Budget details / बजट विवरण | one page per signature block | visible (drop hidden M–Q) | ruled title band, repeated header |
| `flow-table` | Continuous table / सतत सारणी | continuous, repeat header | native (trim empty edges) | zebra rows |
| `cover-letter` | Cover letter / आवरण पत्र | whole sheet, prose wrap | single wide column | no grid, 12–14 pt line height |
| `summary-table` | Summary / सारांश | per title | native | totals highlighted |
| `plain-grid` | Plain grid / सादा ग्रिड | continuous | native | light grid, no conversion |
| `auto` | Auto-detect | from profile | from profile | inherits the guessed layout |

Each preset is a frozen `Layout` dataclass (`page`, `segmentation`, `col_strategy`,
`header_rows`, `repeat_headers`, `theme`, `text_defaults`). `resolve()` merges
`LAYOUTS[layout_id]` ← `SheetConfig` ← `profile` into a **`RenderPlan`** the
renderer consumes; `auto` picks `profile.guessed_layout`.

---

## 5. UX — buttoned interface (screens)

Style follows the deployed PWD Tools Suite: 🪔 warm cream palette (`#E8E2D5`
title band, `#F2EDE4` headers, `#8A8377` rules — the same palette as the PDF),
`🏗️ PWD Tools Suite` header, "Initiative: Mrs. Premlata Jain, AAO, PWD Udaipur",
bilingual labels, big emoji tile per card, prominent **Launch / Render** buttons.

### Screen 1 — Upload (home)
```
┌──────────────────────────────────────────────────────────────┐
│ 🏗️  PWD Tools Suite            विवरण: श्रीमती प्रेमलता जैन, AAO │
│ ──────────────────────────────────────────────────────────── │
│            🪔  Excel → A4 PDF  Print Studio  🪔               │
│   ┌────────────────────────────────────────────────────┐     │
│   │   ⬆  Drop .xlsx / .xls here   [ Choose Excel ]     │     │
│   └────────────────────────────────────────────────────┘     │
│   [ Open last workbook ]   [ Load layouts.json ]             │
└──────────────────────────────────────────────────────────────┘
```

### Screen 2 — Sheet Board (the buttoned core)
One **card per worksheet**; every card is self-contained and actionable.

```
┌───────────────────────────────────────────────────────────────────────┐
│ 🔍 Scanned “FINAL Budget Anuman 31 October 2026.xlsx” — 7 sheets       │
│ [ Render All ]  [ Print All ]  [ Save layouts.json ]  [ Reload ]       │
├───────────────────────────────────────────────────────────────────────┤
│ 📄 Details                      1311×24 · titles ✓ sig ✓ kruti 62%     │
│    layout [ Budget details ▾ ]   [ Configure ] [ Preview ] [ Render ]  │
│    ● detected: budget-details    ● 41 pages                           │
├───────────────────────────────────────────────────────────────────────┤
│ 🧾 Letters1                       45×15 · prose · kruti 91%           │
│    layout [ Cover letter  ▾ ]    [ Configure ] [ Preview ] [ Render ]  │
│    ● detected: cover-letter      ● 1 page                             │
├───────────────────────────────────────────────────────────────────────┤
│ 📊 FORMAT                          3×13 · template · kruti 0%         │
│    layout [ Plain grid    ▾ ]    [ Configure ] [ Preview ] [ Render ]  │
└───────────────────────────────────────────────────────────────────────┘
```
Card controls: **layout dropdown** (instant re-plan), **Configure** (drawer),
**Preview** (first-page thumbnail + page count), **Render** (this sheet only).

### Screen 3 — Config Drawer (per-sheet, tabbed)
```
┌── Details ───────────────────────────────────── [ Apply ] [ Reset ] ✕ ─┐
│  ┌─Layout─┬─Columns─┬─Rows & Breaks─┬─Language & Font─┐               │
│  ├──────────────────────────────────────────────────────────────────────┤
│  │  Layout      ▾ Budget details                                        │
│  │  Orientation ▾ Landscape   Margin  [ 12 ] mm                         │
│  │  Page break  ▾ Per signature block                                   │
│  │  ☑ Repeat header row on every page                                   │
│  └──────────────────────────────────────────────────────────────────────┘
│  (Columns tab) A B C D E F G H I J K L [M N O P Q excluded] R           │
│     ☑/☐ per column · ⇅ drag to reorder · width [__]pt · [Auto widths]    │
│  (Rows tab)    From [ 1 ]  To [ 1001 ]  ☑ skip blank  Every [25] rows   │
│  (Language)    kruti ● auto ○ on ○ off   force-English cols [R]          │
│                font ▾ Deva (NotoSansDevanagari)                          │
│  [ Preview ]                          [ Render this sheet ]             │
└─────────────────────────────────────────────────────────────────────────┘
```

### Screen 4 — Preview & Render bar
```
[ ◀ prev ]  page 3 / 41  [ next ▶ ]      [ Render All ]  [ 🖨 Print ]  [ ⬇ Download ZIP ]
  (first-page A4 thumbnail rendered server-side)
```

---

## 6. HTTP API (`app.py`)

| Method + path | Body | Returns |
|---|---|---|
| `POST /api/workbook` | multipart `file` | `workbook_id`, `sheets[]` (profile + guessed layout) |
| `GET  /api/workbook/{id}` | — | refreshed sheet profiles |
| `PUT  /api/workbook/{id}/sheets/{name}` | `SheetConfig` | resolved `RenderPlan` + estimated pages |
| `POST /api/workbook/{id}/{name}/preview` | — | first-page PNG (base64) + `pages` |
| `POST /api/workbook/{id}/render` | `{sheets: [...]}` | `{sheet: download_url}` |
| `GET  /download/{id}/{sheet}.pdf` | — | PDF bytes |
| `GET  /api/workbook/{id}/layouts.json` | — | saved `WorkbookConfig` |

Server keeps an in-memory (and `_work/` on-disk) registry: `id → {path, profiles,
configs, outputs}`. Uploads are copied to `_work/{id}/` so the original INPUT
folder is never touched.

---

## 7. Why this handles *varied* structures

1. **Profile first** — nothing is assumed about column count or position; the
   title/header/signature/anchor rows are found per sheet, so `Details (2)`'s
   repeated headers and `Letters1`'s prose both work.
2. **Layout is data, not code** — a new sheet shape = a new `Layout` entry, not a
   new script.
3. **Per-sheet overrides** — the user resolves ambiguity (e.g. "column R is
   English, never convert") in the UI; the engine stays deterministic.
4. **Graceful fallbacks** — `plain-grid` renders anything; every sheet has a
   working default even with zero configuration.

---

## 8. Build order / migration

1. ✅ `apply.py` engine proven (KrutiDev, LongTable pagination, fonts).
2. **This pass** — architecture + skeletons (`engine/*`, `app.py`, `web/*`).
3. Next — implement `renderer` themes + `preview` (fitz PNG), then column
   drag-reorder.
4. Later — fold `apply.py` in as `engine/renderer.py::cli()` and delete the
   duplicate CLI (one engine, two entry points: web + CLI).

---

*Draft v1 — architecture + UI screens + module skeleton.*

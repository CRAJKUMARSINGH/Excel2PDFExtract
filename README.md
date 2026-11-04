<div align="center">

# 🗂️ Excel → A4 PDF Print Studio

### Turn messy government Excel sheets into beautiful, print-ready A4 PDFs — correct Devanagari, hidden columns dropped, one click. No Excel gymnastics, no re-typing.

[![Python](https://img.shields.io/badge/python-3.9%2B-3776AB?logo=python&logoColor=white)](https://www.python.org/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)
[![API: FastAPI](https://img.shields.io/badge/API-FastAPI-009688?logo=fastapi&logoColor=white)](https://fastapi.tiangolo.com/)
[![PDF: ReportLab](https://img.shields.io/badge/PDF-ReportLab-1f6feb)](https://www.reportlab.com/)
[![PRs Welcome](https://img.shields.io/badge/PRs-welcome-brightgreen.svg)](CONTRIBUTING.md)
[![GitHub stars](https://img.shields.io/github/stars/CRAJKUMARSINGH/PDFExcelExtract?style=social)](https://github.com/CRAJKUMARSINGH/PDFExcelExtract/stargazers)
[![GitHub forks](https://img.shields.io/github/forks/CRAJKUMARSINGH/PDFExcelExtract?style=social)](https://github.com/CRAJKUMARSINGH/PDFExcelExtract/network/members)
[![Last commit](https://img.shields.io/github/last-commit/CRAJKUMARSINGH/PDFExcelExtract)](https://github.com/CRAJKUMARSINGH/PDFExcelExtract/commits/main)

</div>

---

## ✨ What makes this brilliant?

Every district office has the same spreadsheet problem: a 1000-row Excel workbook, thirteen visible columns, five hidden legacy columns, a **legacy KrutiDev-encoded Hindi** font, and a signature block that must start on a fresh page. Printing it means half a day of column juggling and still-mangled Devanagari.

**Excel → A4 PDF Print Studio** reads the workbook, *profiles* each sheet (titles, header band, signature lines, numeric density, KrutiDev ratio), *suggests the right layout*, and renders a clean A4 PDF — **one page per signature block**, with every KrutiDev byte converted to real Unicode Devanagari. You can tune each sheet in a buttoned browser UI and save the recipe to `layouts.json`. It also still runs as a plain one-command CLI. 🇮🇳📄

> Built for Public Works Department (PWD) budget estimates — sane for *any* awkward workbook.

---

## 🚀 Key Features

- 🧠 **Automatic sheet profiling** — detects title rows, the column-header band, signature lines and numeric density with zero configuration.
- 🕵️ **Live hidden-column detection** — hidden columns (`M N O P Q`, or anything else) are found per sheet at runtime, never hard-coded.
- 🔤 **KrutiDev → Unicode converter** — legacy Remington/Walkman-Chanakya Hindi text is decoded to real Devanagari, so headings render instead of `lM+d` gibberish.
- 🅰️ **Smart English vs. KrutiDev classifier** — genuine English words, work-IDs and numbers are left alone, never mangled.
- 🧾 **Six layout presets** — cover letter, budget details, continuous flow table, summary, plain grid, and auto-detect.
- 🖨️ **Pixel-precise pagination** — ReportLab `LongTable` with a correctly-bounded `repeatRows`, so 70-page tables split cleanly instead of throwing `LayoutError`.
- 🎛️ **Buttoned web UI** — one card per worksheet, per-sheet config drawer, live page-1 preview thumbnails.
- 💾 **Reproducible configs** — every choice round-trips through `layouts.json` (save, reload, re-render).
- 🧩 **Non-destructive** — uploads are copied to a scratch folder; your `INPUT/` folder is never touched.
- 🖱️ **Windows-friendly launchers** — `run_pdf.bat` auto-installs dependencies and opens the output folder.

---

## 📸 Demo / Screenshots

> **Slots ready — drop your images in `docs/screenshots/` and they will appear here.**
> Screenshot shots render in the code below (currently commented out so nothing shows broken).

<!-- Uncomment once the files exist in docs/screenshots/
| Upload | Sheet Board | Config Drawer |
|---|---|---|
| ![Upload](docs/screenshots/1-upload.png) | ![Sheet Board](docs/screenshots/2-sheet-board.png) | ![Config Drawer](docs/screenshots/3-config-drawer.png) |

![Preview](docs/screenshots/4-preview.png)
-->

**Suggested captures:** the drop-zone (`1-upload.png`), the per-sheet board with badges (`2-sheet-board.png`), the config drawer (`3-config-drawer.png`), and a rendered Hindi budget page (`4-preview.png`).

💡 *Want maximum reach? Record a 20–40 s screen capture of “drop a workbook → Render All → open the PDF” and embed it as a GIF at the top of this README.*

---

## ⚡ Quick Start (under 90 seconds)

**Prerequisites:** Python **3.9+** and `pip`.

```bash
# 1 — clone
git clone https://github.com/CRAJKUMARSINGH/PDFExcelExtract.git
cd PDFExcelExtract

# 2 — install dependencies
python -m pip install -r requirements.txt

# 3 — run the studio (web UI)
python app.py
# → open http://127.0.0.1:8000
```

Prefer the command line? Render a workbook straight to PDF:

```bash
python apply.py -i "INPUT/your-workbook.xlsx" -o "OUTPUT/budget_paged.pdf"
```

**Windows, zero terminal:** double-click **`run_pdf.bat`** — it checks Python, installs the packages, renders, and opens `OUTPUT\`.

---

## 📖 How to Use

### Option A — the buttoned web UI (recommended)

| Step | What happens |
|---|---|
| **1. Upload** | Drop an `.xlsx` / `.xls` into the drop-zone (or click **Choose Excel**). |
| **2. Sheet Board** | One card per worksheet with badges: `rows×cols`, `titles`, `signatures`, `hidden M N O P Q`, `kruti 62%` — plus the auto-detected layout. |
| **3. Configure** | Open the drawer and tune **Layout · Columns · Rows & Breaks · Language & Font** for that one sheet. |
| **4. Preview / Render** | **Preview page 1** for a thumbnail, **Render this sheet** for one PDF, or **Render All** for the whole workbook. |
| **5. Save the recipe** | **💾 Save layouts.json** persists every choice for next time. |

### Option B — the CLI

```bash
python apply.py --report                              # dry-run: list every page block
python apply.py -s Details                            # render one named sheet
python apply.py --all-sheets                          # render EVERY sheet -> OUTPUT/<sheet>_paged.pdf
python apply.py --keep-hidden                         # include hidden columns M N O P Q
python apply.py -i INPUT/wb.xls -o OUTPUT/out.pdf     # explicit input / output
python apply.py --margin 10 --font-dir fonts          # 10 mm margins, custom font folder
```

`--report` is the fastest way to sanity-check a new workbook: it prints the detected page blocks without writing a PDF.

---

## 🔧 Configuration / Options

### CLI flags (`apply.py`)

| Flag | Default | Description |
|---|---|---|
| `-i, --input` | `INPUT/FINAL Budget Anuman 31 October 2026.xlsx` | Input workbook (`.xls` **or** `.xlsx`) |
| `-o, --output` | `OUTPUT/budget_paged.pdf` | Output PDF path |
| `-s, --sheet` | `Details` | Worksheet to read |
| `--all-sheets` | off | Render every worksheet to `OUTPUT/<sheet>_paged.pdf` |
| `--keep-hidden` | off | Include the hidden columns (M–Q) |
| `--margin MM` | `12` | Page margin in millimetres |
| `--font-dir DIR` | `fonts` | Folder holding `NotoSansDevanagari-Regular.ttf` |
| `--report` | off | Dry-run: print the page-block list, write nothing |

### Layout presets (`engine/layouts.py`)

| Preset id | English / हिंदी | Used for |
|---|---|---|
| `budget-details` | Budget details / बजट विवरण | One page per signature block, ruled title band |
| `flow-table` | Continuous table / सतत सारणी | Long sheet, header repeated every page |
| `cover-letter` | Cover letter / आवरण पत्र | Prose / memo sheets, single wide column |
| `summary-table` | Summary / सारांश | Per-title tables with totals |
| `plain-grid` | Plain grid / सादा ग्रिड | Light grid, no conversion |
| `auto` | Auto-detect / स्वतः पहचान | Uses the profiler's guess |

### `layouts.json`

Exactly a serialised `WorkbookConfig` — edit it, commit it, and share a printing recipe:

```json
{
  "version": 1,
  "source": "FINAL Budget Anuman 31 October 2026.xlsx",
  "defaults": { "layout_id": "auto", "text": { "kruti": "auto" } },
  "sheets": {
    "Details":  { "layout_id": "budget-details",
                  "columns": { "mode": "explicit", "exclude": ["M", "N", "O", "P", "Q"] } },
    "Letters1": { "layout_id": "cover-letter" },
    "Saransh":  { "layout_id": "summary-table" },
    "FORMAT":   { "layout_id": "plain-grid", "text": { "kruti": "off" } }
  }
}
```

Key `SheetConfig` knobs: `columns.mode` (`visible` \| `native` \| `explicit`), `columns.exclude` / `include` / `order` / `widths`, `rows.start` / `end`, `page_breaks.strategy` (`signature` \| `title` \| `n_rows` \| `continuous`), `text.kruti` (`auto` \| `on` \| `off`), `text.force_english`, `page.orientation` / `margin_mm` / `repeat_header`.

> There are **no environment variables** — configuration lives in CLI flags and `layouts.json`.

### REST API (`app.py`)

| Method + path | Purpose |
|---|---|
| `POST /api/workbook` | Upload a workbook → `workbook_id` + per-sheet profiles |
| `GET /api/workbook/{id}` | Re-read profiles |
| `PUT /api/workbook/{id}/sheets/{name}` | Save one sheet's config → resolved plan + page estimate |
| `POST /api/workbook/{id}/preview/{name}` | Page-1 PNG (base64) for the drawer |
| `POST /api/workbook/{id}/render` | Render selected sheets |
| `GET /download/{id}/{file}.pdf` | Download a rendered PDF |
| `GET /api/workbook/{id}/layouts.json` | Saved `WorkbookConfig` |
| `GET /api/layouts` | Preset catalogue (dropdown data) |

---

## 🛠 Tech Stack

| Layer | Choice |
|---|---|
| Language | **Python 3.9+** |
| PDF engine | **ReportLab** (`LongTable`, `BaseDocTemplate`, custom paged canvas) |
| Excel I/O | **openpyxl** (`.xlsx`) · **xlrd 1.2.0** (legacy `.xls`) |
| Web / API | **FastAPI** + **Uvicorn** |
| Thumbnails | **PyMuPDF** (`fitz`) → PNG |
| Front end | **Vanilla JS + CSS** — no build step, no npm |
| Fonts | **Noto Sans Devanagari** (Unicode) + built-in KrutiDev→Unicode transliteration |

### Architecture

```
reader → profiler → layouts (registry + resolver) → renderer
   │         │              │                          │
   └──────── schema.py (SheetConfig / WorkbookConfig / layouts.json) ───────┘
                              │
                      apply.py  ← single source of truth for
             register_fonts · make_styles · krutidev_to_unicode ·
             col_widths · build_table · segment · _PagedCanvas
```

`apply.py` is the proven legacy CLI and is **imported, never modified**, by the new `engine/` layer — one engine, two entry points (web + CLI). See [MASTER-PLAN.md](MASTER-PLAN.md) for the full design.

---

## 🗺 Roadmap

- [ ] Dedicated `prose` and `totals` render themes (today they share the `ruled` path).
- [ ] Hidden-column detection for legacy `.xls` (`xlrd` `colinfo_map`) — currently a stub.
- [ ] Drag-to-reorder column persistence in the config drawer.
- [ ] Fold `apply.py` in as `engine/renderer.py::cli()` and retire the duplicate CLI.
- [ ] Job queue + auth for multi-user deployments.
- [ ] Word (`.docx`) export alongside PDF.
- [ ] Optional OCR pass for scanned annexures.

---

## 🤝 How to Contribute

Contributions are warmly welcome — bug reports, layout presets, docs, and translations alike.

1. Fork → create a branch: `git checkout -b feat/my-idea`
2. Keep changes focused; add a note to the README/roadmap if user-visible.
3. Open a Pull Request describing **what** and **why**.

Please read [CONTRIBUTING.md](CONTRIBUTING.md) and our [Code of Conduct](CODE_OF_CONDUCT.md) first. Found a security issue? See [SECURITY.md](SECURITY.md).

---

## 📜 License

Released under the **MIT License** — see [LICENSE](LICENSE).

---

## ❤️ Shout-outs

- **PWD Tools Suite** — initiative of **Mrs. Premlata Jain, AAO, PWD Udaipur**; the warm cream UI palette is shared with our printed budgets.
- **Walkman Chanakya 905** — the public MIT KrutiDev→Unicode conversion table this project's transliteration is ported from.
- **ReportLab**, **openpyxl**, **FastAPI**, and **Noto Sans Devanagari** — the shoulders this stands on.

---

<div align="center">

### ⭐ If this saves you a day of copy-paste misery — star the repo & share it!

Every star helps another district office find it. 🚀

</div>

<!-- Session log: see creat.md (mirrored from team working notes). -->

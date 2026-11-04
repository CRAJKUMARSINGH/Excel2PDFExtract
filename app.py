"""Excel → A4 PDF Print Studio — local HTTP server.

Serves the buttoned web UI in ``web/`` and a small JSON API that drives the
engine:

    POST /api/workbook                 upload a workbook -> profiles
    GET  /api/workbook/{id}            re-read profiles
    PUT  /api/workbook/{id}/sheets/{n} save one sheet's config -> resolved plan
    POST /api/workbook/{id}/preview/{n} first-page PNG for the drawer preview
    POST /api/workbook/{id}/render     render selected sheets
    GET  /download/{id}/{file}.pdf     download one PDF
    GET  /api/workbook/{id}/layouts.json   saved WorkbookConfig
    GET  /api/layouts                  preset catalogue (dropdown data)

Run::

    pip install fastapi uvicorn python-multipart   # + openpyxl reportlab pymupdf
    python app.py                                  # -> http://127.0.0.1:8000
"""
from __future__ import annotations

import io
import shutil
import uuid
from dataclasses import dataclass, field
from pathlib import Path

try:
    from fastapi import FastAPI, HTTPException, UploadFile, File
    from fastapi.responses import FileResponse, HTMLResponse
    from fastapi.staticfiles import StaticFiles
    from pydantic import BaseModel
except ImportError:  # pragma: no cover
    raise SystemExit('pip install fastapi uvicorn python-multipart')

from engine import layouts as layouts_mod
from engine import profiler, reader, renderer
from engine.schema import SheetConfig, WorkbookConfig

ROOT = Path(__file__).resolve().parent
WEB = ROOT / 'web'
FONTS = ROOT / 'fonts'
WORK = ROOT / '_work'
WORK.mkdir(exist_ok=True)

app = FastAPI(title='Excel → A4 PDF Print Studio', version='0.1')


# ---------------------------------------------------------------------------
# in-memory registry (swap for a DB in production)
# ---------------------------------------------------------------------------
@dataclass
class Session:
    id: str
    path: Path
    config: WorkbookConfig
    grids: dict = field(default_factory=dict)       # name -> Grid
    profiles: dict = field(default_factory=dict)    # name -> SheetProfile
    outputs: dict = field(default_factory=dict)     # name -> Path


STORE: dict[str, Session] = {}


def _session(wid: str) -> Session:
    s = STORE.get(wid)
    if s is None:
        raise HTTPException(404, f'unknown workbook id {wid}')
    return s


def _profile_all(sess: Session) -> list[dict]:
    out = []
    for name, grid in sess.grids.items():
        prof = sess.profiles[name]
        cfg = sess.config.for_sheet(name)
        plan = layouts_mod.resolve(prof, cfg)
        out.append({
            'profile': prof.to_dict(),
            'config': cfg.to_dict(),
            'plan': plan.to_dict(),
            'est_sections': renderer.estimate_sections(grid, plan),
        })
    return out


# ---------------------------------------------------------------------------
# models
# ---------------------------------------------------------------------------
class RenderRequest(BaseModel):
    sheets: list[str] | None = None      # None = every sheet


# ---------------------------------------------------------------------------
# API
# ---------------------------------------------------------------------------
@app.get('/api/layouts')
def api_layouts():
    return {'layouts': layouts_mod.catalogue()}


@app.post('/api/workbook')
async def api_upload(file: UploadFile = File(...)):
    wid = uuid.uuid4().hex[:12]
    dest = WORK / wid
    dest.mkdir(parents=True, exist_ok=True)
    src = dest / (file.filename or 'workbook.xlsx')
    with src.open('wb') as fh:
        shutil.copyfileobj(file.file, fh)

    try:
        names = reader.list_sheets(src)
    except Exception as exc:                     # noqa: BLE001 - surfaced to UI
        raise HTTPException(400, f'cannot read workbook: {exc}')

    sess = Session(id=wid, path=src, config=WorkbookConfig(source=src.name))
    for name in names:
        grid = reader.read_grid(src, name)
        sess.grids[name] = grid
        sess.profiles[name] = profiler.profile(grid)
        sess.config.sheets[name] = SheetConfig(
            layout_id=sess.profiles[name].guessed_layout)
    STORE[wid] = sess

    return {'workbook_id': wid, 'source': src.name,
            'sheets': _profile_all(sess)}


@app.get('/api/workbook/{wid}')
def api_get(wid: str):
    sess = _session(wid)
    return {'workbook_id': wid, 'source': sess.path.name,
            'sheets': _profile_all(sess)}


@app.put('/api/workbook/{wid}/sheets/{name}')
def api_put_config(wid: str, name: str, cfg: dict):
    sess = _session(wid)
    if name not in sess.grids:
        raise HTTPException(404, f'no sheet {name!r}')
    sess.config.sheets[name] = SheetConfig.from_dict(cfg)
    eff = sess.config.for_sheet(name)
    plan = layouts_mod.resolve(sess.profiles[name], eff)
    return {'config': eff.to_dict(), 'plan': plan.to_dict(),
            'est_sections': renderer.estimate_sections(sess.grids[name], plan)}


@app.post('/api/workbook/{wid}/preview/{name}')
def api_preview(wid: str, name: str):
    """Render one sheet and return page 1 as a base64 PNG."""
    sess = _session(wid)
    if name not in sess.grids:
        raise HTTPException(404, f'no sheet {name!r}')
    out = _render_sheet(sess, name)

    try:
        import fitz                              # PyMuPDF
    except ImportError:
        return {'png': None, 'pages': None,
                'note': 'pip install pymupdf for thumbnails',
                'pdf': f'/download/{wid}/{out.name}'}

    with fitz.open(out) as doc:
        page = doc[0]
        pix = page.get_pixmap(dpi=110)
        png = io.BytesIO(pix.tobytes('png')).getvalue()
    import base64
    return {'png': 'data:image/png;base64,' + base64.b64encode(png).decode(),
            'pages': None, 'pdf': f'/download/{wid}/{out.name}'}


@app.post('/api/workbook/{wid}/render')
def api_render(wid: str, req: RenderRequest):
    sess = _session(wid)
    wanted = req.sheets or list(sess.grids)
    made = {}
    for name in wanted:
        if name not in sess.grids:
            continue
        out = _render_sheet(sess, name)
        made[name] = f'/download/{wid}/{out.name}'
    if not made:
        raise HTTPException(400, 'nothing rendered')
    return {'outputs': made}


@app.get('/api/workbook/{wid}/layouts.json')
def api_save_config(wid: str):
    sess = _session(wid)
    p = WORK / wid / 'layouts.json'
    sess.config.save(p)
    return FileResponse(p, filename='layouts.json', media_type='application/json')


@app.get('/download/{wid}/{fname}')
def api_download(wid: str, fname: str):
    _session(wid)
    f = (WORK / wid / fname).resolve()
    if not f.is_file() or WORK.resolve() not in f.parents:
        raise HTTPException(404, 'no such file')
    return FileResponse(f, filename=fname, media_type='application/pdf')


# ---------------------------------------------------------------------------
# static UI
# ---------------------------------------------------------------------------
@app.get('/', response_class=HTMLResponse)
def index():
    return (WEB / 'index.html').read_text(encoding='utf-8')


if WEB.is_dir():
    app.mount('/static', StaticFiles(directory=str(WEB)), name='static')


# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------
def _render_sheet(sess: Session, name: str) -> Path:
    grid = sess.grids[name]
    cfg = sess.config.for_sheet(name)
    plan = layouts_mod.resolve(sess.profiles[name], cfg)
    safe = ''.join(ch if ch.isalnum() or ch in '._-' else '_' for ch in name)
    out = WORK / sess.id / f'{safe}.pdf'
    renderer.render(grid, plan, out, FONTS)
    sess.outputs[name] = out
    return out


if __name__ == '__main__':
    import uvicorn
    uvicorn.run(app, host='127.0.0.1', port=8000)

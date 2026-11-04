/* Excel → A4 PDF Print Studio — UI controller (vanilla JS, no build step) */
'use strict';

const S = {
  wid: null, source: '', sheets: [], layouts: [],
  current: null,        // sheet name in the drawer
  draft: null,          // SheetConfig copy being edited
  cols: [],             // ordered letters currently ON
  colsTouched: false,
  dragIdx: null,
};

const $  = (sel) => document.querySelector(sel);
const $$ = (sel) => Array.from(document.querySelectorAll(sel));

/* ---------------------------------------------------------------- helpers */
async function api(path, opts = {}) {
  const r = await fetch(path, opts);
  if (!r.ok) throw new Error((await r.text()) || r.statusText);
  return r.headers.get('content-type')?.includes('json') ? r.json()
                                                         : r.text();
}
function toast(msg, ms = 2600) {
  const t = $('#toast');
  t.textContent = msg; t.classList.remove('hidden');
  clearTimeout(t._h); t._h = setTimeout(() => t.classList.add('hidden'), ms);
}
const esc = (s) => String(s).replace(/[&<>"]/g, (c) =>
  ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const pct = (v) => Math.round((v || 0) * 100) + '%';

/* ------------------------------------------------------------ Screen 1 */
$('#file').addEventListener('change', (e) => {
  if (e.target.files[0]) upload(e.target.files[0]);
});
const drop = $('#drop');
['dragenter', 'dragover'].forEach((ev) => drop.addEventListener(ev, (e) => {
  e.preventDefault(); drop.classList.add('over');
}));
['dragleave', 'drop'].forEach((ev) => drop.addEventListener(ev, (e) => {
  e.preventDefault(); drop.classList.remove('over');
}));
drop.addEventListener('drop', (e) => {
  const f = e.dataTransfer.files[0];
  if (f) upload(f);
});

async function upload(file) {
  $('#upload-hint').textContent = `Uploading ${file.name} …`;
  const fd = new FormData();
  fd.append('file', file);
  try {
    const d = await api('/api/workbook', { method: 'POST', body: fd });
    S.wid = d.workbook_id; S.source = d.source; S.sheets = d.sheets;
    $('#screen-upload').classList.add('hidden');
    $('#screen-board').classList.remove('hidden');
    $('#wb-name').textContent = d.source;
    renderBoard();
  } catch (err) {
    toast('Upload failed: ' + err.message);
    $('#upload-hint').textContent = 'Upload failed — see the toast.';
  }
}

/* ------------------------------------------------------------ Screen 2 */
function renderBoard() {
  const total = S.sheets.reduce((n, s) => n + (s.est_sections || 0), 0);
  $('#board-summary').textContent =
    `${S.sheets.length} sheets · ${total} page-blocks detected`;
  $('#cards').innerHTML = S.sheets.map(cardHTML).join('');
}

function cardHTML(s) {
  const p = s.profile, pl = s.plan;
  const badges = [
    `<span class="badge">${p.n_rows}×${p.n_cols}</span>`,
    p.title_rows.length ? `<span class="badge ok">${p.title_rows.length} titles</span>` : '',
    p.signature_rows.length ? `<span class="badge ok">${p.signature_rows.length} signatures</span>` : '',
    p.hidden_letters.length ? `<span class="badge warn">hidden ${p.hidden_letters.join('')}</span>` : '',
    `<span class="badge">kruti ${pct(p.kruti_ratio)}</span>`,
    `<span class="badge">numeric ${pct(p.numeric_density)}</span>`,
  ].join('');

  const opts = S.layouts.map((l) =>
    `<option value="${l.id}" ${l.id === pl.layout_id ? 'selected' : ''}>${esc(l.label_en)} · ${esc(l.label_hi)}</option>`
  ).join('');

  return `
  <article class="card" data-sheet="${esc(p.name)}">
    <div class="card-top">
      <div>
        <h3>📄 ${esc(p.name)}</h3>
        <div class="badges">${badges}</div>
      </div>
      <div class="meta">↳ auto-detected
        <b>${esc(pl.layout_id)}</b><br>${pl.col_letters.length} cols · ${pl.orientation}
        · ~${s.est_sections} blocks</div>
    </div>
    <div class="card-actions">
      <label style="margin:0">layout</label>
      <select class="card-layout" data-sheet="${esc(p.name)}">${opts}</select>
      <button class="btn" data-act="configure" data-sheet="${esc(p.name)}">⚙ Configure</button>
      <button class="btn" data-act="preview"   data-sheet="${esc(p.name)}">👁 Preview</button>
      <button class="btn btn-primary" data-act="render" data-sheet="${esc(p.name)}">🖨 Render</button>
    </div>
  </article>`;
}

$('#cards').addEventListener('click', async (e) => {
  const btn = e.target.closest('button[data-act]');
  if (!btn) return;
  const name = btn.dataset.sheet;
  if (btn.dataset.act === 'configure') return openDrawer(name);
  if (btn.dataset.act === 'preview')   return preview(name);
  if (btn.dataset.act === 'render')    return renderSheets([name]);
});
$('#cards').addEventListener('change', async (e) => {
  const sel = e.target.closest('select.card-layout');
  if (!sel) return;
  await saveConfig(sel.dataset.sheet, { layout_id: sel.value });
  toast(`${sel.dataset.sheet} → ${sel.value}`);
});

$('#btn-render-all').addEventListener('click', () => renderSheets(null));
$('#btn-reload').addEventListener('click', async () => {
  if (!S.wid) return;
  S.sheets = (await api(`/api/workbook/${S.wid}`)).sheets;
  renderBoard();
});
$('#btn-save').addEventListener('click', () => {
  if (S.wid) window.open(`/api/workbook/${S.wid}/layouts.json`, '_blank');
});

/* ------------------------------------------------------------ Drawer */
async function saveConfig(name, patch) {
  const cur = S.sheets.find((s) => s.profile.name === name);
  const merged = deepMerge(structuredClone(cur.config), patch);
  const d = await api(`/api/workbook/${S.wid}/sheets/${encodeURIComponent(name)}`,
    { method: 'PUT', headers: { 'content-type': 'application/json' },
      body: JSON.stringify(merged) });
  Object.assign(cur, { config: d.config, plan: d.plan, est_sections: d.est_sections });
  renderBoard();
  return d;
}

function openDrawer(name) {
  const s = S.sheets.find((x) => x.profile.name === name);
  S.current = name;
  S.draft = structuredClone(s.config);
  S.cols = s.plan.col_letters.slice();
  S.colsTouched = false;

  $('#drawer-title').textContent = name;
  $('#drawer').classList.remove('hidden');
  switchTab('layout');
  fillDrawer();
  $('#drawer-status').textContent =
    `detected: ${s.profile.guessed_layout} · ${s.profile.n_rows}×${s.profile.n_cols}`;
}

function fillDrawer() {
  const c = S.draft;
  $('#f-layout').innerHTML = S.layouts.map((l) =>
    `<option value="${l.id}" ${l.id === c.layout_id ? 'selected' : ''}>${esc(l.label_en)} · ${esc(l.label_hi)}</option>`).join('');
  $('#f-orientation').value = c.page.orientation;
  $('#f-margin').value = c.page.margin_mm;
  $('#f-repeat').value = String(c.page.repeat_header);

  $('#f-colmode').value = c.columns.include.length ? 'explicit' : c.columns.mode;
  $('#f-exclude').value = c.columns.exclude.join(',');
  $('#f-include').value = c.columns.include.join(',');
  $('#f-rowstart').value = c.rows.start;
  $('#f-rowend').value = c.rows.end ?? '';
  $('#f-break').value = c.page_breaks.strategy;
  $('#f-every').value = c.page_breaks.every;
  $('#f-kruti').value = c.text.kruti;
  $('#f-force-en').value = c.text.force_english.join(',');
  renderChips();
}

function renderChips() {
  const s = S.sheets.find((x) => x.profile.name === S.current);
  const hidden = new Set(s.profile.hidden_letters);
  $('#col-chips').innerHTML = s.profile.col_names.map((L) => {
    const on = S.cols.includes(L);
    return `<span class="chip ${on ? 'on' : ''} ${hidden.has(L) ? 'hidden-col' : ''}"
      draggable="true" data-col="${L}" title="${hidden.has(L) ? 'hidden in Excel' : 'visible'}">
      ${L}</span>`;
  }).join('');
}

$('#col-chips').addEventListener('click', (e) => {
  const chip = e.target.closest('.chip'); if (!chip) return;
  const L = chip.dataset.col;
  S.cols = S.cols.includes(L) ? S.cols.filter((x) => x !== L) : [...S.cols, L];
  S.colsTouched = true;
  renderChips();
});
$('#col-chips').addEventListener('dragstart', (e) => {
  const chip = e.target.closest('.chip'); if (!chip) return;
  S.dragIdx = S.cols.indexOf(chip.dataset.col);
  chip.classList.add('drag');
});
$('#col-chips').addEventListener('dragend', (e) => e.target.classList?.remove('drag'));
$('#col-chips').addEventListener('dragover', (e) => {
  e.preventDefault();
  const chip = e.target.closest('.chip'); if (!chip || S.dragIdx < 0) return;
  const to = S.cols.indexOf(chip.dataset.col);
  if (to < 0 || to === S.dragIdx) return;
  const [m] = S.cols.splice(S.dragIdx, 1);
  S.cols.splice(to, 0, m);
  S.dragIdx = to; S.colsTouched = true;
  renderChips();
});

/* tabs */
function switchTab(name) {
  $$('.tab').forEach((t) => t.classList.toggle('active', t.dataset.tab === name));
  $$('.pane').forEach((p) => p.classList.toggle('hidden', p.dataset.pane !== name));
}
$$('.tab').forEach((t) => t.addEventListener('click', () => switchTab(t.dataset.tab)));

/* draft <-> form */
function collectDraft() {
  const d = S.draft;
  d.layout_id = $('#f-layout').value;
  d.page.orientation = $('#f-orientation').value;
  d.page.margin_mm = parseFloat($('#f-margin').value) || 12;
  d.page.repeat_header = $('#f-repeat').value === 'true';

  const list = (v) => v.split(',').map((x) => x.trim().toUpperCase()).filter(Boolean);
  d.columns.exclude = list($('#f-exclude').value);
  d.columns.include = list($('#f-include').value);
  d.columns.mode = $('#f-colmode').value === 'explicit' ? 'native' : $('#f-colmode').value;
  if (S.colsTouched) d.columns.include = S.cols.slice();

  d.rows.start = parseInt($('#f-rowstart').value, 10) || 1;
  d.rows.end = $('#f-rowend').value ? parseInt($('#f-rowend').value, 10) : null;
  d.page_breaks.strategy = $('#f-break').value;
  d.page_breaks.every = parseInt($('#f-every').value, 10) || 25;
  d.text.kruti = $('#f-kruti').value;
  d.text.force_english = list($('#f-force-en').value);
  return d;
}

$('#btn-apply').addEventListener('click', async () => {
  const d = await saveConfig(S.current, collectDraft());
  S.draft = structuredClone(d.config);
  S.cols = d.plan.col_letters.slice(); S.colsTouched = false;
  fillDrawer();
  $('#drawer-status').textContent = `✔ applied · ${d.plan.col_letters.length} cols · ~${d.est_sections} blocks`;
  toast('Applied');
});
$('#btn-reset').addEventListener('click', async () => {
  const s = S.sheets.find((x) => x.profile.name === S.current);
  S.draft = { ...structuredClone(s.config),
    layout_id: s.profile.guessed_layout, columns: { mode:'visible', include:[], exclude:[], order:[], widths:{} } };
  S.colsTouched = false;
  fillDrawer();
  toast('Reset to detected defaults');
});
$('#btn-close').addEventListener('click', () => $('#drawer').classList.add('hidden'));
$('#drawer').addEventListener('click', (e) => {
  if (e.target.id === 'drawer') $('#drawer').classList.add('hidden');
});

/* ------------------------------------------------------------ render + preview */
async function renderSheets(sheets) {
  if (!S.wid) return;
  toast('Rendering …');
  const d = await api(`/api/workbook/${S.wid}/render`,
    { method: 'POST', headers: { 'content-type': 'application/json' },
      body: JSON.stringify({ sheets }) });
  const names = Object.keys(d.outputs);
  toast(`✔ rendered: ${names.join(', ')}`);
  if (names.length === 1) window.open(d.outputs[names[0]], '_blank');
}
$('#btn-render-one').addEventListener('click', async () => {
  await saveConfig(S.current, collectDraft());
  await renderSheets([S.current]);
});

async function preview(name) {
  toast('Building preview …');
  const d = await api(`/api/workbook/${S.wid}/preview/${encodeURIComponent(name)}`,
    { method: 'POST' });
  $('#preview-title').textContent = 'Preview · ' + name;
  $('#btn-open-pdf').href = d.pdf;
  $('#preview-img').src = d.png || '';
  $('#preview').classList.remove('hidden');
  if (!d.png) toast(d.note || 'thumbnail unavailable — open the PDF');
}
$('#btn-preview').addEventListener('click', async () => {
  await saveConfig(S.current, collectDraft());
  preview(S.current);
});
$('#btn-preview-close').addEventListener('click', () => $('#preview').classList.add('hidden'));
$('#preview').addEventListener('click', (e) => {
  if (e.target.id === 'preview') $('#preview').classList.add('hidden');
});

/* ------------------------------------------------------------ boot */
function deepMerge(a, b) {
  for (const k of Object.keys(b || {})) {
    if (b[k] && typeof b[k] === 'object' && !Array.isArray(b[k])) {
      a[k] = deepMerge(a[k] || {}, b[k]);
    } else { a[k] = b[k]; }
  }
  return a;
}
(async function boot() {
  try { S.layouts = (await api('/api/layouts')).layouts; }
  catch { S.layouts = []; }
})();

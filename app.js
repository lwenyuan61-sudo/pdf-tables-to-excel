// PDF to Excel — everything runs in the browser.
// Text PDFs: read the text layer with pdf.js; scanned pages (no text layer): OCR with tesseract.js.
// Words → rows (same baseline) → cells (split at wide gaps) → columns (x ranges shared by many rows) → sheet.
import * as pdfjsLib from 'https://cdn.jsdelivr.net/npm/pdfjs-dist@4.4.168/build/pdf.min.mjs';
pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdn.jsdelivr.net/npm/pdfjs-dist@4.4.168/build/pdf.worker.min.mjs';

const DONATE_URL = 'https://ko-fi.com/wenyuanli';                     // Ko-fi / GitHub Sponsors link; button stays hidden while empty
const CELL_GAP = 0.6;                        // gap wider than 0.6 × text height starts a new cell
const ROW_TOL = 0.45;                        // baselines closer than 0.45 × text height are one row
const OCR_SCALE = 3;                         // render scanned pages at 3 × 72 = 216 dpi
const OCR_GAP = 0.9;                         // OCR heights are whole-line heights, and a missed full-width comma leaves ~1 em → wider cell gap
const CJK = /[⺀-鿿豈-﫿＀-￯　-〿]/;

const $ = (id) => document.getElementById(id);
const state = { files: [], wb: null, base: 'tables' };

// ---------- reading a page ----------
function splitWide(str) {                    // "Apple    12.50   3" → pieces with their character offsets
  const out = []; const re = /\S+(?: \S+)*/g; let m;
  while ((m = re.exec(str))) out.push({ s: m[0], at: m.index });
  return out;
}

async function textItems(page) {
  const vp = page.getViewport({ scale: 1 });
  const tc = await page.getTextContent();
  const items = [];
  for (const it of tc.items) {
    if (!it.str || !it.str.trim()) continue;
    const t = pdfjsLib.Util.transform(vp.transform, it.transform);
    const h = Math.hypot(t[2], t[3]) || Math.abs(it.height) || 8;
    const cw = it.width / Math.max(it.str.length, 1);
    for (const p of splitWide(it.str)) {
      const x0 = t[4] + p.at * cw;
      items.push({ s: p.s, x0, x1: x0 + p.s.length * cw, y: t[5], h });
    }
  }
  return items;
}

function eraseRules(ctx, W, H) {             // whiten table rules (dark runs longer than 1/5 of the page) so OCR does not read them as quotes
  const im = ctx.getImageData(0, 0, W, H); const d = im.data;
  const dark = new Uint8Array(W * H);
  for (let i = 0, p = 0; i < W * H; i++, p += 4) dark[i] = d[p] * 0.3 + d[p + 1] * 0.59 + d[p + 2] * 0.11 < 160 ? 1 : 0;
  const kill = new Uint8Array(W * H); const minH = W / 5, minV = H / 5;
  for (let y = 0; y < H; y++) { let s = -1; for (let x = 0; x <= W; x++) { const on = x < W && dark[y * W + x]; if (on && s < 0) s = x; if (!on && s >= 0) { if (x - s >= minH) for (let k = s; k < x; k++) kill[y * W + k] = 1; s = -1; } } }
  for (let x = 0; x < W; x++) { let s = -1; for (let y = 0; y <= H; y++) { const on = y < H && dark[y * W + x]; if (on && s < 0) s = y; if (!on && s >= 0) { if (y - s >= minV) for (let k = s; k < y; k++) kill[k * W + x] = 1; s = -1; } } }
  let n = 0;
  for (let i = 0, p = 0; i < W * H; i++, p += 4) if (kill[i]) { d[p] = d[p + 1] = d[p + 2] = 255; n++; }
  if (n) ctx.putImageData(im, 0, 0);
}

const workers = {};
async function ocrItems(page, lang, onStep) {
  const vp = page.getViewport({ scale: OCR_SCALE });
  const cv = document.createElement('canvas'); cv.width = Math.ceil(vp.width); cv.height = Math.ceil(vp.height);
  const ctx = cv.getContext('2d'); ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, cv.width, cv.height);
  await page.render({ canvasContext: ctx, viewport: vp }).promise;
  eraseRules(ctx, cv.width, cv.height);
  if (!workers[lang]) {
    onStep('Loading text recognition (first time only)…');
    workers[lang] = await Tesseract.createWorker(lang, 1);
  }
  const { data } = await workers[lang].recognize(cv, {}, { tsv: true });
  const items = []; let line = null;
  for (const row of (data.tsv || '').split('\n')) {
    const f = row.split('\t'); if (f.length < 12) continue;
    const [level, , , , , , left, top, width, height, conf] = f.map(Number); const text = f.slice(11).join('\t').trim();
    if (level === 4) line = { top, height };
    if (level !== 5 || !text || conf < 0) continue;
    if (/^[\p{P}\p{S}]+$/u.test(text) && conf < 70) continue;   // table rules / specks read as stray quotes or dots
    let s = text, x0 = left, x1 = left + width; const cw = width / [...text].length;
    while (/^[”“"'’‘]/.test(s) && CJK.test(s[1] || '')) { s = s.slice(1); x0 += cw; }   // a quote glued in front of Chinese is a rule line
    while (/[”“"'’‘]$/.test(s) && CJK.test(s[s.length - 2] || '')) { s = s.slice(0, -1); x1 -= cw; }
    const L = line || { top, height };
    items.push({ s, x0: x0 / OCR_SCALE, x1: x1 / OCR_SCALE, y: (L.top + L.height) / OCR_SCALE, h: L.height / OCR_SCALE, g: OCR_GAP, o: items.length });
  }
  cv.width = cv.height = 0;
  return items;
}

// ---------- words → grid ----------
function median(a) { if (!a.length) return 0; const s = [...a].sort((x, y) => x - y); return s[s.length >> 1]; }

function toRows(items) {
  items.sort((a, b) => a.y - b.y || a.x0 - b.x0);
  const rows = [];
  for (const it of items) {
    const r = rows[rows.length - 1];
    if (r && Math.abs(it.y - r.y) <= ROW_TOL * Math.min(it.h, r.h)) { r.items.push(it); r.h = Math.max(r.h, it.h); }
    else rows.push({ y: it.y, h: it.h, items: [it] });
  }
  for (const r of rows) {
    r.items.sort((a, b) => (a.o ?? a.x0) - (b.o ?? b.x0));   // OCR: keep the recogniser's reading order (its boxes can be a little off)
    const cells = [];
    for (const it of r.items) {
      const c = cells[cells.length - 1]; const gap = c ? it.x0 - c.x1 : Infinity;
      if (c && gap < (it.g || CELL_GAP) * Math.max(it.h, c.h)) glue(c, it, gap);
      else cells.push({ s: it.s, x0: it.x0, x1: it.x1, h: it.h });
    }
    r.cells = cells;
  }
  return rows;
}

function glue(c, it, gap) {                  // append a word to a cell; no space between Chinese characters
  const sp = gap > 0.12 * it.h && !(CJK.test(c.s.slice(-1)) && CJK.test(it.s[0]));
  c.s += (sp ? ' ' : '') + it.s; c.x1 = Math.max(c.x1, it.x1); c.h = Math.max(c.h, it.h);
}

function bandsOf(rows) {                     // column ranges = x ranges covered by cells in enough multi-cell rows
  const multi = rows.filter((r) => r.cells.length >= 2);
  if (!multi.length) return [];
  let lo = Infinity, hi = -Infinity;
  for (const r of multi) for (const c of r.cells) { lo = Math.min(lo, c.x0); hi = Math.max(hi, c.x1); }
  const n = Math.ceil(hi - lo) + 2; const cov = new Float32Array(n);
  for (const r of multi) for (const c of r.cells) for (let x = Math.floor(c.x0 - lo); x < Math.ceil(c.x1 - lo); x++) cov[x] += 1;
  const thr = multi.length >= 4 ? Math.max(2, Math.ceil(0.15 * multi.length)) : 1;   // a column = x range used by ≥ 15 % of rows
  const bands = []; let s = -1;
  for (let x = 0; x <= n; x++) {
    const on = x < n && cov[x] >= thr;
    if (on && s < 0) s = x;
    if (!on && s >= 0) { bands.push([s + lo, x + lo]); s = -1; }
  }
  return bands;
}

function place(cell, bands) {
  let best = -1, bo = 0;
  bands.forEach(([a, b], i) => { const o = Math.min(b, cell.x1) - Math.max(a, cell.x0); if (o > bo) { bo = o; best = i; } });
  if (best >= 0) return best;
  const mid = (cell.x0 + cell.x1) / 2; let d = Infinity;
  bands.forEach(([a, b], i) => { const dd = Math.abs((a + b) / 2 - mid); if (dd < d) { d = dd; best = i; } });
  return Math.max(best, 0);
}

function toGrid(items) {
  if (!items.length) return [];
  const rows = toRows(items);
  let lo = Infinity, hi = -Infinity;
  for (const r of rows) for (const c of r.cells) { lo = Math.min(lo, c.x0); hi = Math.max(hi, c.x1); }
  const width = hi - lo; const gapMed = median(rows.slice(1).map((r, i) => r.y - rows[i].y));
  const isPara = (r) => r.cells.length === 1 && r.cells[0].x1 - r.cells[0].x0 > 0.4 * width;
  const blocks = []; let cur = [];           // a block ends at a paragraph line or a big vertical gap
  rows.forEach((r, i) => {
    const gap = i ? r.y - rows[i - 1].y : 0;
    if (isPara(r) || (cur.length && gap > 3 * Math.max(gapMed, r.h))) { if (cur.length) blocks.push(cur); cur = []; }
    if (isPara(r)) blocks.push([r]); else cur.push(r);
  });
  if (cur.length) blocks.push(cur);
  const grid = [];
  for (const b of blocks) {
    const bands = bandsOf(b);
    for (const r of b) {
      if (!bands.length) { grid.push(r.cells.map((c) => c.s)); continue; }
      const out = bands.map(() => null);     // table rows: each word goes to its own column (two cells glued by a narrow gap come apart)
      for (const it of (r.cells.length >= 2 ? r.items : r.cells)) {
        const k = place(it, bands); const c = out[k];
        if (c) glue(c, it, it.x0 - c.x1); else out[k] = { s: it.s, x0: it.x0, x1: it.x1, h: it.h };
      }
      grid.push(out.map((c) => (c ? c.s : '')));
    }
  }
  return grid;
}

// ---------- grid → workbook ----------
const NUM = /^\(?\s*([-+−])?\s*([$€£¥₹])?\s*((?:\d{1,3}(?:,\d{3})+)|\d+)(\.\d+)?\s*(%)?\s*\)?$/;
function value(s, conv) {                    // → [cell value, number format or null]; the format keeps the PDF's look
  s = s.replace(/\s+/g, ' ').trim().replace(new RegExp(`(${CJK.source}) (?=${CJK.source})`, 'gu'), '$1');
  if (!conv || !s) return [s, null];
  const m = s.match(NUM); if (!m) return [s, null];
  const int = m[3].replace(/,/g, '');
  if (/^0\d/.test(int) || (int.length > 15)) return [s, null];   // IDs / codes with leading zeros stay text
  let v = parseFloat(int + (m[4] || ''));
  const neg = m[1] === '-' || m[1] === '−' || (s[0] === '(' && s[s.length - 1] === ')');
  if (neg) v = -v;
  const dec = m[4] ? '.' + '0'.repeat(m[4].length - 1) : '';
  if (m[5]) return [v / 100, '0' + dec + '%'];
  let z = (m[3].includes(',') ? '#,##0' : '0') + dec;
  if (m[2]) z = '"' + m[2] + '"' + z;
  if (s[0] === '(') z = z + ';(' + z + ')';
  return [v, z === '0' ? null : z];
}

function sheetOf(grid, conv) {
  const w = Math.max(1, ...grid.map((r) => r.length));
  const aoa = []; const fmt = [];
  grid.forEach((r, i) => { const row = []; for (let j = 0; j < w; j++) { const [v, z] = value(r[j] || '', conv); row.push(v); if (z) fmt.push([i, j, z]); } aoa.push(row); });
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  for (const [i, j, z] of fmt) { const a = XLSX.utils.encode_cell({ r: i, c: j }); if (ws[a]) ws[a].z = z; }
  ws['!cols'] = Array.from({ length: w }, (_, j) => ({ wch: Math.min(60, Math.max(6, ...grid.map((r) => (r[j] || '').length + 1))) }));
  return ws;
}

function buildWorkbook() {
  const mode = document.querySelector('input[name=mode]:checked').value; const conv = $('num').checked;
  const wb = XLSX.utils.book_new(); const used = new Set();
  const name = (s) => { s = s.replace(/[\\/?*[\]:]/g, ' ').slice(0, 31) || 'Sheet'; let n = s, k = 2; while (used.has(n)) n = (s.slice(0, 27) + ' (' + k++ + ')'); used.add(n); return n; };
  state.files.forEach((f, fi) => {
    const tag = state.files.length > 1 ? (fi + 1) + ' ' : '';
    if (mode === 'all') {
      const g = []; f.pages.forEach((p) => { if (p.length) { if (g.length) g.push([]); g.push(...p); } });
      if (g.length) XLSX.utils.book_append_sheet(wb, sheetOf(g, conv), name(tag + f.name.replace(/\.pdf$/i, '')));
    } else f.pages.forEach((p, pi) => { if (p.length) XLSX.utils.book_append_sheet(wb, sheetOf(p, conv), name(tag + 'Page ' + (pi + 1))); });
  });
  state.wb = wb.SheetNames.length ? wb : null;
  return state.wb;
}

// ---------- UI ----------
function status(msg, err) { const s = $('status'); s.textContent = msg; s.className = err ? 'err' : ''; }
function progress(x) { const b = $('bar'); b.style.display = x == null ? 'none' : 'block'; b.firstElementChild.style.width = ((x || 0) * 100).toFixed(1) + '%'; }

function preview() {
  const wb = state.wb; const sel = $('sheet'); const t = $('prev'); t.textContent = '';
  if (!wb) return;
  const ws = wb.Sheets[sel.value]; const R = XLSX.utils.decode_range(ws['!ref']); const nr = R.e.r + 1;
  const frag = document.createDocumentFragment();
  for (let r = 0; r < Math.min(nr, 300); r++) {
    const tr = document.createElement('tr');
    for (let c = 0; c <= R.e.c; c++) {
      const cell = ws[XLSX.utils.encode_cell({ r, c })]; const td = document.createElement('td');
      if (cell) { td.textContent = XLSX.utils.format_cell(cell); if (cell.t === 'n') td.className = 'n'; }
      tr.appendChild(td);
    }
    frag.appendChild(tr);
  }
  t.appendChild(frag);
  $('note').textContent = nr > 300 ? `Showing the first 300 of ${nr} rows. The download has all of them.` : '';
}

function refresh() {
  if (!state.files.length) return;
  const wb = buildWorkbook(); const sel = $('sheet'); const keep = sel.value; sel.textContent = '';
  if (!wb) { $('out').style.display = 'none'; status('No text or tables found in this file.', true); return; }
  for (const n of wb.SheetNames) { const o = document.createElement('option'); o.value = o.textContent = n; sel.appendChild(o); }
  if (wb.SheetNames.includes(keep)) sel.value = keep;
  $('out').style.display = 'block'; preview();
}

async function convert(files) {
  files = [...files].filter((f) => /pdf$/i.test(f.type) || /\.pdf$/i.test(f.name));
  if (!files.length) { status('Please choose a PDF file.', true); return; }
  state.files = []; state.base = files.length === 1 ? files[0].name.replace(/\.pdf$/i, '') : 'tables';
  $('out').style.display = 'none'; const lang = $('lang').value; let ocrPages = 0;
  try {
    for (const [fi, f] of files.entries()) {
      const doc = await pdfjsLib.getDocument({ data: new Uint8Array(await f.arrayBuffer()) }).promise;
      const pages = [];
      for (let p = 1; p <= doc.numPages; p++) {
        const lead = (files.length > 1 ? `${f.name}: ` : '') + `page ${p} of ${doc.numPages}`;
        status('Reading ' + lead + '…'); progress((fi + (p - 1) / doc.numPages) / files.length);
        const page = await doc.getPage(p);
        let items = await textItems(page);
        if (items.length < 5) {                // no text layer → scanned page
          status('Recognising text on ' + lead + '…'); ocrPages++;
          items = await ocrItems(page, lang, status);
        }
        pages.push(toGrid(items)); page.cleanup();
      }
      state.files.push({ name: f.name, pages }); await doc.destroy();
    }
    progress(null);
    const nPages = state.files.reduce((a, f) => a + f.pages.length, 0);
    status(`Done — ${nPages} page${nPages > 1 ? 's' : ''}` + (ocrPages ? `, ${ocrPages} scanned (check those numbers)` : '') + '.');
    refresh();
  } catch (e) {
    progress(null); console.error(e);
    status(e && e.name === 'PasswordException' ? 'This PDF is password-protected.' : 'Could not read this file: ' + (e && e.message || e), true);
  }
}

function download(blob, fname) { const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = fname; a.click(); setTimeout(() => URL.revokeObjectURL(a.href), 4000); }

$('file').addEventListener('change', (e) => { convert(e.target.files); e.target.value = ''; });
const drop = $('drop');
drop.addEventListener('keydown', (e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); $('file').click(); } });
['dragenter', 'dragover'].forEach((t) => drop.addEventListener(t, (e) => { e.preventDefault(); drop.classList.add('over'); }));
['dragleave', 'drop'].forEach((t) => drop.addEventListener(t, (e) => { e.preventDefault(); drop.classList.remove('over'); }));
drop.addEventListener('drop', (e) => convert(e.dataTransfer.files));
document.querySelectorAll('input[name=mode], #num').forEach((el) => el.addEventListener('change', refresh));
$('sheet').addEventListener('change', preview);
$('dl').addEventListener('click', () => state.wb && XLSX.writeFile(state.wb, state.base + '.xlsx'));
$('csv').addEventListener('click', () => {
  if (!state.wb) return; const n = $('sheet').value;
  download(new Blob(['﻿' + XLSX.utils.sheet_to_csv(state.wb.Sheets[n])], { type: 'text/csv;charset=utf-8' }), `${state.base} - ${n}.csv`);
});
if (DONATE_URL) { const d = $('donate'); d.href = DONATE_URL; d.style.display = 'inline-flex'; }
window.__convert = convert; window.__state = state;   // (used by the test page)

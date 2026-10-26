// PDF & image tables to Excel — everything runs in the browser.
// Text PDFs: words come from the text layer (pdf.js); scanned pages, photos and screenshots: OCR (tesseract.js).
// Ruled tables: the drawn lines (found in the rendered page) give the cells, so merged cells and multi-line cells come out right.
// Other text: words → rows (same baseline) → cells (split at wide gaps) → columns (x ranges shared by many rows).
import * as pdfjsLib from 'https://cdn.jsdelivr.net/npm/pdfjs-dist@4.4.168/build/pdf.min.mjs';
pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdn.jsdelivr.net/npm/pdfjs-dist@4.4.168/build/pdf.worker.min.mjs';

const DONATE_URL = 'https://ko-fi.com/wenyuanli';                     // Ko-fi / GitHub Sponsors link; button stays hidden while empty
const CELL_GAP = 0.6;                        // gap wider than 0.6 × text height starts a new cell
const ROW_TOL = 0.45;                        // baselines closer than 0.45 × text height are one row
const OCR_SCALE = 3;                         // render scanned pages at 3 × 72 = 216 dpi
const RULE_SCALE = 2;                        // render text pages at 144 dpi just to find the table lines
const TEXT_PX = 40;                          // images are scaled so a line of text is ~40 px tall (28 / 40 / 56 tested on the test images: 40 best)
const OCR_GAP = 0.9;                         // OCR heights are whole-line heights, and a missed full-width comma leaves ~1 em → wider cell gap
const NOSPACE = /[⺀-⿿　-ヿ㐀-鿿豈-﫿＀-￯]/;   // Chinese / Japanese: no spaces between characters
const CJK = /[⺀-⿿　-ヿ㐀-鿿가-힯豈-﫿＀-￯]/;   // … and Korean (which does use spaces)

const $ = (id) => document.getElementById(id);
const state = { files: [], wb: null, base: 'tables' };
const isPdf = (f) => /pdf$/i.test(f.type) || /\.pdf$/i.test(f.name);
const isImg = (f) => /^image\//i.test(f.type) || /\.(png|jpe?g|webp|gif|bmp)$/i.test(f.name);

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

async function pageCanvas(page, scale) {
  const vp = page.getViewport({ scale });
  const cv = document.createElement('canvas'); cv.width = Math.ceil(vp.width); cv.height = Math.ceil(vp.height);
  const ctx = cv.getContext('2d', { willReadFrequently: true }); ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, cv.width, cv.height);
  await page.render({ canvasContext: ctx, viewport: vp, intent: 'print' }).promise;   // 'print' does not wait for screen frames, so it keeps going in a background tab
  return { cv, ctx, k: scale };
}

function lineHeight(bmp) {                   // median height of the ink bands between blank rows ≈ height of one line of text, in image pixels
  const W = bmp.width, H = bmp.height; const c = document.createElement('canvas'); c.width = W; c.height = H;
  const x = c.getContext('2d', { willReadFrequently: true }); x.fillStyle = '#fff'; x.fillRect(0, 0, W, H); x.drawImage(bmp, 0, 0);
  const d = x.getImageData(0, 0, W, H).data; const bands = []; let s = -1;
  for (let y = 0; y <= H; y++) {
    let ink = 0;
    if (y < H) for (let i = y * W * 4, e = i + W * 4; i < e; i += 4) if (d[i] * 0.3 + d[i + 1] * 0.59 + d[i + 2] * 0.11 < 160) ink++;
    const on = ink > 0 && ink < 0.6 * W;     // (a row that is nearly all ink is a rule or a filled bar, not text)
    if (on && s < 0) s = y;
    if (!on && s >= 0) { if (y - s >= 4) bands.push(y - s); s = -1; }
  }
  c.width = c.height = 0;
  return bands.length ? median(bands) : 0;
}

async function imageCanvas(file) {           // photos / screenshots: enlarge or shrink so a line of text is ~TEXT_PX tall (what OCR reads best)
  const bmp = await createImageBitmap(file);
  const lh = lineHeight(bmp);
  const k = lh ? Math.max(0.3, Math.min(4, TEXT_PX / lh)) : Math.max(0.5, Math.min(4, 3000 / bmp.width));
  const cv = document.createElement('canvas'); cv.width = Math.round(bmp.width * k); cv.height = Math.round(bmp.height * k);
  const ctx = cv.getContext('2d', { willReadFrequently: true }); ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, cv.width, cv.height);
  ctx.imageSmoothingQuality = 'high'; ctx.drawImage(bmp, 0, 0, cv.width, cv.height); bmp.close();
  return { cv, ctx, k, lh };
}

// ---------- table lines ----------
function findRules(ctx, W, H, k, minLen = 18) {   // thin dark runs: horizontal ≥ minLen units (images: 2.5 lines of text, so a long dash is not a rule), vertical ≥ 8; ≤ ~2.5 units thick. Units = page points (images: pixels)
  const im = ctx.getImageData(0, 0, W, H); const d = im.data; const dark = new Uint8Array(W * H);
  for (let i = 0, p = 0; i < W * H; i++, p += 4) dark[i] = d[p] * 0.3 + d[p + 1] * 0.59 + d[p + 2] * 0.11 < 215 ? 1 : 0;
  const thick = Math.max(3, Math.round(2.5 * k + 2));
  function segs(along, across, at, minLen) {
    const out = []; let act = [];
    for (let a = 0; a < across; a++) {
      const next = []; let s = -1;
      for (let b = 0; b <= along; b++) {
        const on = b < along && dark[at(a, b)];
        if (on && s < 0) s = b;
        if (!on && s >= 0) {
          if (b - s >= minLen) {
            const e = b - 1; const m = act.find((g) => Math.min(g.e, e) - Math.max(g.s, s) > 0.5 * Math.min(g.e - g.s, e - s));
            if (m) { m.s = Math.min(m.s, s); m.e = Math.max(m.e, e); m.a1 = a; next.push(m); }
            else { const g = { s, e, a0: a, a1: a }; out.push(g); next.push(g); }
          }
          s = -1;
        }
      }
      act = next;
    }
    return out.filter((g) => g.a1 - g.a0 + 1 <= thick);
  }
  const hs = segs(W, H, (y, x) => y * W + x, Math.max(24, Math.round(minLen * k))).map((g) => ({ y: (g.a0 + g.a1) / 2 / k, x0: g.s / k, x1: g.e / k, g }));
  let vs = segs(H, W, (x, y) => y * W + x, Math.max(12, Math.round(8 * k))).map((g) => ({ x: (g.a0 + g.a1) / 2 / k, y0: g.s / k, y1: g.e / k, g }));
  const tol = (thick + 2) / k; const px2 = 2 / k;
  const touch = (x, y) => hs.some((h) => Math.abs(h.y - y) <= (h.g.a1 - h.g.a0 + 1) / k / 2 + px2 && h.x0 - px2 <= x && x <= h.x1 + px2);
  vs = vs.filter((v) => touch(v.x, v.y0) && touch(v.x, v.y1));   // a table border joins a horizontal rule at both ends (within 2 px); letter strokes stop short
  return { hs, vs, tol, W: W / k, erase() {                       // whiten the lines before OCR (otherwise they are read as quotes / dashes)
    for (const h of hs) for (let y = h.g.a0; y <= h.g.a1; y++) for (let x = h.g.s; x <= h.g.e; x++) { const p = (y * W + x) * 4; d[p] = d[p + 1] = d[p + 2] = 255; }
    for (const v of vs) for (let x = v.g.a0; x <= v.g.a1; x++) for (let y = v.g.s; y <= v.g.e; y++) { const p = (y * W + x) * 4; d[p] = d[p + 1] = d[p + 2] = 255; }
    ctx.putImageData(im, 0, 0);
  } };
}

function cluster(vals, tol) {
  const s = [...vals].sort((a, b) => a - b); const out = [];
  for (const v of s) { const c = out[out.length - 1]; if (c && v - c.last <= tol) { c.sum += v; c.n++; c.last = v; } else out.push({ sum: v, n: 1, last: v }); }
  return out.map((c) => c.sum / c.n);
}

function tablesOf(R) {                       // → full grids (lines on both axes) and row-ruled stacks (horizontal rules only)
  const { hs, vs, tol } = R; const n = hs.length + vs.length; const par = Array.from({ length: n }, (_, i) => i);
  const find = (i) => (par[i] === i ? i : (par[i] = find(par[i])));
  vs.forEach((v, j) => hs.forEach((h, i) => { if (h.x0 - tol <= v.x && v.x <= h.x1 + tol && v.y0 - tol <= h.y && h.y <= v.y1 + tol) par[find(hs.length + j)] = find(i); }));
  const comp = new Map();
  hs.forEach((h, i) => { const r = find(i); if (!comp.has(r)) comp.set(r, { H: [], V: [] }); comp.get(r).H.push(h); });
  vs.forEach((v, j) => { const r = find(hs.length + j); if (!comp.has(r)) comp.set(r, { H: [], V: [] }); comp.get(r).V.push(v); });
  const grids = []; const loose = [];
  for (const { H, V } of comp.values()) {
    const g = H.length >= 2 && V.length >= 2 ? gridOf(H, V, tol) : null;
    if (g) grids.push(g); else loose.push(...H);
  }
  return { grids, stacks: stacksOf(loose, tol, R.W) };
}

function gridOf(H, V, tol) {
  const xs = cluster(V.map((v) => v.x), tol), ys = cluster(H.map((h) => h.y), tol);
  const hx0 = Math.min(...H.map((h) => h.x0)), hx1 = Math.max(...H.map((h) => h.x1));
  const vy0 = Math.min(...V.map((v) => v.y0)), vy1 = Math.max(...V.map((v) => v.y1));
  if (hx0 < xs[0] - 3 * tol) xs.unshift(hx0);                       // tables drawn without outer side borders
  if (hx1 > xs[xs.length - 1] + 3 * tol) xs.push(hx1);
  if (vy0 < ys[0] - 3 * tol) ys.unshift(vy0);
  if (vy1 > ys[ys.length - 1] + 3 * tol) ys.push(vy1);
  const nr = ys.length - 1, nc = xs.length - 1;
  if (nr < 1 || nc < 1 || nr * nc < 2) return null;
  const vAt = (j, i) => V.some((v) => Math.abs(v.x - xs[j]) <= tol && v.y0 <= ys[i] + tol && v.y1 >= ys[i + 1] - tol);
  const hAt = (i, j) => H.some((h) => Math.abs(h.y - ys[i]) <= tol && h.x0 <= xs[j] + tol && h.x1 >= xs[j + 1] - tol);
  const par = Array.from({ length: nr * nc }, (_, i) => i); const find = (i) => (par[i] === i ? i : (par[i] = find(par[i])));
  for (let i = 0; i < nr; i++) for (let j = 1; j < nc; j++) if (!vAt(j, i)) par[find(i * nc + j)] = find(i * nc + j - 1);   // no line between → merged cell
  for (let i = 1; i < nr; i++) for (let j = 0; j < nc; j++) if (!hAt(i, j)) par[find(i * nc + j)] = find((i - 1) * nc + j);
  const reg = new Map();
  for (let i = 0; i < nr; i++) for (let j = 0; j < nc; j++) {
    const r = find(i * nc + j); const g = reg.get(r) || { r0: i, c0: j, r1: i, c1: j, n: 0 };
    g.r0 = Math.min(g.r0, i); g.c0 = Math.min(g.c0, j); g.r1 = Math.max(g.r1, i); g.c1 = Math.max(g.c1, j); g.n++; reg.set(r, g);
  }
  for (const g of reg.values()) g.rect = g.n === (g.r1 - g.r0 + 1) * (g.c1 - g.c0 + 1);
  const cellReg = (i, j) => { const g = reg.get(find(i * nc + j)); return g.rect ? g : { r0: i, c0: j, r1: i, c1: j, rect: true, n: 1, solo: true }; };
  return { xs, ys, nr, nc, cellReg, regions: [...reg.values()] };
}

function stacksOf(H, tol, pageW) {           // ≥ 3 rules with the same left / right ends = a table ruled between rows
  const lim = 0.03 * pageW + tol; const stacks = [];
  for (const h of [...H].sort((a, b) => a.y - b.y)) {
    const s = stacks.find((st) => { const l = st.lines[st.lines.length - 1]; return Math.abs(l.x0 - h.x0) <= lim && Math.abs(l.x1 - h.x1) <= lim && h.y - l.y > 2 * tol && h.y - l.y < 120; });
    if (s) s.lines.push(h); else stacks.push({ lines: [h] });
  }
  return stacks.filter((s) => s.lines.length >= 3).map((s) => ({ x0: median(s.lines.map((l) => l.x0)), x1: median(s.lines.map((l) => l.x1)), ys: s.lines.map((l) => l.y), tol }));
}

// ---------- OCR ----------
const workers = {};
async function ocr(cv, k, lang, onStep, extra = []) {   // extra = other sizes to read again (images: small commas in numbers come and go with size)
  if (!workers[lang]) {
    onStep('Loading text recognition (first time only)…');
    workers[lang] = await Tesseract.createWorker(lang, 1);
  }
  const a = await ocrPass(cv, k, lang);
  const NUMLIKE = /^[(\-−+$€£¥₹₩]?\d[\d.,]*\d%?\)?$/; const THOU = /^\D*\d{1,3}(,\d{3})+(\.\d+)?\D*$/; const digits = (s) => s.replace(/\D/g, '');
  const want = a.filter((it) => NUMLIKE.test(it.s) && digits(it.s).length >= 4 && !THOU.test(it.s));
  for (const k2 of extra) {
    if (!want.length) break;
    const c2 = document.createElement('canvas'); c2.width = Math.round(cv.width * k2 / k); c2.height = Math.round(cv.height * k2 / k);
    const x = c2.getContext('2d'); x.imageSmoothingQuality = 'high'; x.drawImage(cv, 0, 0, c2.width, c2.height);
    const b = await ocrPass(c2, k2, lang); c2.width = c2.height = 0;
    for (let i = want.length - 1; i >= 0; i--) {                       // the same number read with proper thousands commas wins
      const it = want[i];
      const m = b.find((o) => Math.min(o.x1, it.x1) - Math.max(o.x0, it.x0) > 0.5 * Math.min(o.x1 - o.x0, it.x1 - it.x0) && Math.abs(o.y - it.y) < 0.5 * it.h);
      if (m && THOU.test(m.s) && digits(m.s) === digits(it.s)) { it.s = m.s; want.splice(i, 1); }
    }
  }
  return a;
}

async function ocrPass(cv, k, lang) {
  const { data } = await workers[lang].recognize(cv, {}, { tsv: true });
  const items = []; let line = null;
  for (const row of (data.tsv || '').split('\n')) {
    const f = row.split('\t'); if (f.length < 12) continue;
    const [level, , , , , , left, top, width, height, conf] = f.map(Number); const text = f.slice(11).join('\t').trim();
    if (level === 4) line = { top, height };
    if (level !== 5 || !text || conf < 0) continue;
    if (/^[\p{P}\p{S}]+$/u.test(text) && conf < 70) continue;   // specks read as stray quotes or dots
    let s = text, x0 = left, x1 = left + width; const cw = width / [...text].length;
    while (/^[”“"'’‘]/.test(s) && CJK.test(s[1] || '')) { s = s.slice(1); x0 += cw; }   // a quote glued in front of Chinese is a rule line
    while (/[”“"'’‘]$/.test(s) && CJK.test(s[s.length - 2] || '')) { s = s.slice(0, -1); x1 -= cw; }
    const L = line || { top, height };
    items.push({ s, x0: x0 / k, x1: x1 / k, y: (L.top + L.height) / k, h: L.height / k, g: OCR_GAP, o: items.length });
  }
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
  rows.forEach((r, ln) => {
    r.items.sort((a, b) => (a.o ?? a.x0) - (b.o ?? b.x0));   // OCR: keep the recogniser's reading order (its boxes can be a little off)
    const cells = [];
    for (const it of r.items) {
      it.ln = ln;
      const c = cells[cells.length - 1]; const gap = c ? it.x0 - c.x1 : Infinity;
      if (c && gap < (it.g || CELL_GAP) * Math.max(it.h, c.h)) glue(c, it, gap);
      else cells.push({ s: it.s, x0: it.x0, x1: it.x1, h: it.h, ln });
    }
    r.cells = cells;
  });
  return rows;
}

function glue(c, it, gap) {                  // append a word to a cell (a word from the next line also gets a space); none between Chinese / Japanese characters
  const sp = (it.o != null || gap > 0.12 * it.h || it.ln !== c.ln) && !(NOSPACE.test(c.s.slice(-1)) && NOSPACE.test(it.s[0]));   // OCR words are already whole words
  c.s += (sp ? ' ' : '') + it.s; c.x1 = Math.max(c.x1, it.x1); c.h = Math.max(c.h, it.h); c.ln = it.ln;
}

function joinWords(ws) {                     // all the words of one table cell, line by line
  let c = null;
  for (const r of toRows(ws)) for (const it of r.items) { if (c) glue(c, it, it.x0 - c.x1); else c = { s: it.s, x0: it.x0, x1: it.x1, h: it.h, ln: it.ln }; }
  return c ? c.s : '';
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

function mergeRuled(rows, stacks) {          // tables with a rule under every row: text lines between two rules are one row (wrapped cells)
  for (const st of stacks) {
    const inside = (r) => r.cells.every((c) => c.x0 >= st.x0 - 3 * st.tol && c.x1 <= st.x1 + 3 * st.tol);
    const slot = (r) => { for (let k = 0; k + 1 < st.ys.length; k++) if (r.y > st.ys[k] && r.y - r.h < st.ys[k + 1]) return k; return -1; };
    const by = new Map();
    for (const r of rows) { if (!inside(r)) continue; const k = slot(r); if (k >= 0) by.set(k, [...(by.get(k) || []), r]); }
    const groups = [...by.values()];
    const perRow = groups.filter((g) => g.length === 1).length >= 0.6 * groups.length;   // most gaps hold one line → a rule under every row
    for (const g of groups) if (g.length > 1) {
      const firstEnd = Math.min(...g[0].cells.map((c) => c.x1));
      const wrapped = g.slice(1).every((r) => Math.min(...r.cells.map((c) => c.x0)) > firstEnd);   // later lines leave the first column empty
      if (!perRow && !wrapped) continue;
      const [a, ...rest] = g;
      for (const r of rest) { a.items.push(...r.items); a.cells.push(...r.cells); r.gone = true; }
      a.h = Math.max(a.h, ...rest.map((r) => r.h));
    }
  }
  return rows.filter((r) => !r.gone);
}

function freeGrid(items, stacks) {           // text that is not inside a fully ruled table → [{ y, cells }]
  if (!items.length) return [];
  const rows = mergeRuled(toRows(items), stacks);
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
      if (!bands.length) { grid.push({ y: r.y, cells: r.cells.map((c) => c.s) }); continue; }
      const out = bands.map(() => null);     // table rows: each word goes to its own column (two cells glued by a narrow gap come apart)
      let prev = null, pk = -1;
      for (const it of (r.cells.length >= 2 ? r.items : r.cells)) {
        const stick = prev && it.ln === prev.ln && it.x0 - prev.x1 < 0.5 * it.h && NOSPACE.test(prev.s.slice(-1)) && NOSPACE.test(it.s[0]);
        const k = stick ? pk : place(it, bands); const c = out[k];   // (Chinese / Japanese characters touching the previous one stay in its column; OCR boxes can be off)
        if (c) glue(c, it, it.x0 - c.x1); else out[k] = { s: it.s, x0: it.x0, x1: it.x1, h: it.h, ln: it.ln };
        prev = it; pk = k;
      }
      grid.push({ y: r.y, cells: out.map((c) => (c ? c.s : '')) });
    }
  }
  return grid;
}

function pageGrid(items, rules) {            // → { rows: string[][], merges: Excel merge ranges }
  const { grids, stacks } = rules ? tablesOf(rules) : { grids: [], stacks: [] };
  const used = new Set(); const tables = [];
  const at = (edges, v) => { for (let i = 0; i + 1 < edges.length; i++) if (v >= edges[i] && v < edges[i + 1]) return i; return -1; };
  for (const g of grids) {
    const words = new Map();
    items.forEach((it, idx) => {
      if (used.has(idx)) return;
      const j = at(g.xs, (it.x0 + it.x1) / 2), i = at(g.ys, it.y - 0.4 * it.h); if (i < 0 || j < 0) return;
      used.add(idx); const r = g.cellReg(i, j); const key = r.solo ? 's' + i + ',' + j : r;
      if (!words.has(key)) words.set(key, { r, ws: [] }); words.get(key).ws.push(it);
    });
    if (!words.size) continue;
    const rows = Array.from({ length: g.nr }, (_, i) => ({ y: g.ys[i], cells: new Array(g.nc).fill('') }));
    for (const { r, ws } of words.values()) rows[r.r0].cells[r.c0] = joinWords(ws);
    tables.push({ g, rows });
  }
  const all = [...freeGrid(items.filter((_, i) => !used.has(i)), stacks), ...tables.flatMap((t) => t.rows)].sort((a, b) => a.y - b.y);
  const index = new Map(all.map((r, i) => [r, i])); const merges = [];
  for (const { g, rows } of tables) for (const r of g.regions) {
    if (!r.rect || (r.r1 === r.r0 && r.c1 === r.c0)) continue;
    const a = index.get(rows[r.r0]), b = index.get(rows[r.r1]);
    if (b - a === r.r1 - r.r0) merges.push({ s: { r: a, c: r.c0 }, e: { r: b, c: r.c1 } });
  }
  return { rows: all.map((r) => r.cells), merges };
}

// ---------- grid → workbook ----------
const NUM = /^\(?\s*([-+−])?\s*([$€£¥₹])?\s*((?:\d{1,3}(?:,\d{3})+)|\d+)(\.\d+)?\s*(%)?\s*\)?$/;
const NOSP_RE = new RegExp(`(${NOSPACE.source}) (?=${NOSPACE.source})`, 'gu');
function value(s, conv) {                    // → [cell value, number format or null]; the format keeps the PDF's look
  s = s.replace(/\s+/g, ' ').trim().replace(NOSP_RE, '$1');
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

function sheetOf(grid, merges, conv) {
  const w = Math.max(1, ...grid.map((r) => r.length));
  const aoa = []; const fmt = [];
  grid.forEach((r, i) => { const row = []; for (let j = 0; j < w; j++) { const [v, z] = value(r[j] || '', conv); row.push(v); if (z) fmt.push([i, j, z]); } aoa.push(row); });
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  for (const [i, j, z] of fmt) { const a = XLSX.utils.encode_cell({ r: i, c: j }); if (ws[a]) ws[a].z = z; }
  if (merges.length) ws['!merges'] = merges;
  ws['!cols'] = Array.from({ length: w }, (_, j) => ({ wch: Math.min(60, Math.max(6, ...grid.map((r) => (r[j] || '').length + 1))) }));
  return ws;
}

function fixDots(rows) {                     // OCR pages: "32.500" in a column of whole numbers is a misread thousands comma
  const w = Math.max(0, ...rows.map((r) => r.length)); const NUMC = /^\D?\d[\d.,]*\d?\D?$/; const DOT3 = /^(\D*\d{1,3})\.(\d{3}\D*)$/;
  for (let j = 0; j < w; j++) {
    const nums = rows.map((r) => (r[j] || '').trim()).filter((s) => NUMC.test(s) && /\d/.test(s));
    const others = nums.filter((s) => !DOT3.test(s));
    if (others.length < 2 || others.some((s) => s.includes('.'))) continue;
    rows.forEach((r) => { const s = (r[j] || '').trim(); if (DOT3.test(s)) r[j] = s.replace(DOT3, '$1,$2'); });
  }
  return rows;
}

function buildWorkbook() {
  const mode = document.querySelector('input[name=mode]:checked').value; const conv = $('num').checked;
  const wb = XLSX.utils.book_new(); const used = new Set();
  const name = (s) => { s = s.replace(/[\\/?*[\]:]/g, ' ').slice(0, 31) || 'Sheet'; let n = s, k = 2; while (used.has(n)) n = (s.slice(0, 27) + ' (' + k++ + ')'); used.add(n); return n; };
  const stem = (f) => f.name.replace(/\.[a-z0-9]+$/i, '');
  state.files.forEach((f, fi) => {
    for (const p of f.pages) if (p.ocr && !p.fixed) { fixDots(p.rows); p.fixed = true; }
    const tag = state.files.length > 1 ? (fi + 1) + ' ' : '';
    if (mode === 'all' || f.image) {
      const g = []; const mg = [];
      f.pages.forEach((p) => { if (!p.rows.length) return; if (g.length) g.push([]); const off = g.length; g.push(...p.rows); for (const m of p.merges) mg.push({ s: { r: m.s.r + off, c: m.s.c }, e: { r: m.e.r + off, c: m.e.c } }); });
      if (g.length) XLSX.utils.book_append_sheet(wb, sheetOf(g, mg, conv), name(tag + stem(f)));
    } else f.pages.forEach((p, pi) => { if (p.rows.length) XLSX.utils.book_append_sheet(wb, sheetOf(p.rows, p.merges, conv), name(tag + 'Page ' + (pi + 1))); });
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
  const span = new Map(); const hidden = new Set();
  for (const m of ws['!merges'] || []) {
    span.set(m.s.r + ',' + m.s.c, [m.e.r - m.s.r + 1, m.e.c - m.s.c + 1]);
    for (let r = m.s.r; r <= m.e.r; r++) for (let c = m.s.c; c <= m.e.c; c++) if (r !== m.s.r || c !== m.s.c) hidden.add(r + ',' + c);
  }
  const frag = document.createDocumentFragment();
  for (let r = 0; r < Math.min(nr, 300); r++) {
    const tr = document.createElement('tr');
    for (let c = 0; c <= R.e.c; c++) {
      if (hidden.has(r + ',' + c)) continue;
      const cell = ws[XLSX.utils.encode_cell({ r, c })]; const td = document.createElement('td');
      const sp = span.get(r + ',' + c); if (sp) { td.rowSpan = sp[0]; td.colSpan = sp[1]; }
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
  if (!wb) { $('out').style.display = 'none'; status('No text or tables found.', true); return; }
  for (const n of wb.SheetNames) { const o = document.createElement('option'); o.value = o.textContent = n; sel.appendChild(o); }
  if (wb.SheetNames.includes(keep)) sel.value = keep;
  $('out').style.display = 'block'; preview();
}

async function convert(files) {
  files = [...files].filter((f) => isPdf(f) || isImg(f));
  if (!files.length) { status('Please choose a PDF or an image (PNG, JPG…).', true); return; }
  state.files = []; state.base = files.length === 1 ? files[0].name.replace(/\.[a-z0-9]+$/i, '') : 'tables';
  $('out').style.display = 'none'; const lang = $('lang').value; let ocrPages = 0, tiny = false;
  try {
    for (const [fi, f] of files.entries()) {
      const lead0 = files.length > 1 ? `${f.name}: ` : '';
      if (isImg(f)) {                                                  // photo / screenshot = one scanned page
        status('Recognising text in ' + (lead0 || 'the image') + '…'); progress(fi / files.length); ocrPages++;
        const { cv, ctx, k, lh } = await imageCanvas(f); if (lh && lh < 18 && /chi|jpn|kor/.test(lang)) tiny = true;   // (tests: 1× screenshots of small Chinese / Japanese text lose words; 2× read fully; English 1× was fine)
        const rules = findRules(ctx, cv.width, cv.height, k, Math.max(18, 2.5 * lh)); rules.erase();
        const items = await ocr(cv, k, lang, status, k > 1.6 ? [1, 1.5] : []); cv.width = cv.height = 0;
        state.files.push({ name: f.name, image: true, pages: [{ ...pageGrid(items, rules), ocr: true }] });
        continue;
      }
      const doc = await pdfjsLib.getDocument({ data: new Uint8Array(await f.arrayBuffer()) }).promise;
      const pages = [];
      for (let p = 1; p <= doc.numPages; p++) {
        const lead = lead0 + `page ${p} of ${doc.numPages}`;
        status('Reading ' + lead + '…'); progress((fi + (p - 1) / doc.numPages) / files.length);
        const page = await doc.getPage(p);
        let items = await textItems(page); let rules; const scanned = items.length < 5;
        if (scanned) {                         // no text layer → scanned page
          status('Recognising text on ' + lead + '…'); ocrPages++;
          const { cv, ctx, k } = await pageCanvas(page, OCR_SCALE);
          rules = findRules(ctx, cv.width, cv.height, k); rules.erase();
          items = await ocr(cv, k, lang, status); cv.width = cv.height = 0;
        } else {
          const { cv, ctx, k } = await pageCanvas(page, RULE_SCALE);
          rules = findRules(ctx, cv.width, cv.height, k); cv.width = cv.height = 0;
        }
        pages.push({ ...pageGrid(items, rules), ocr: scanned }); page.cleanup();
      }
      state.files.push({ name: f.name, pages }); await doc.destroy();
    }
    progress(null);
    const nPages = state.files.reduce((a, f) => a + f.pages.length, 0);
    status(`Done — ${nPages} page${nPages > 1 ? 's' : ''}` + (ocrPages ? `, ${ocrPages} read by text recognition (check those numbers)` : '') + '.'
      + (tiny ? ' The text in this image is small — zoom in before taking the screenshot for better accuracy.' : ''));
    refresh();
  } catch (e) {
    progress(null); console.error(e);
    status(e && e.name === 'PasswordException' ? 'This PDF is password-protected.' : 'Could not read this file: ' + (e && e.message || e), true);
  }
}

function download(blob, fname) { const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = fname; a.click(); setTimeout(() => URL.revokeObjectURL(a.href), 4000); }

const LANGS = { 'zh-tw': 'eng+chi_tra', 'zh-hk': 'eng+chi_tra', 'zh-hant': 'eng+chi_tra', zh: 'eng+chi_sim', ja: 'eng+jpn', ko: 'eng+kor', fr: 'eng+fra+deu+spa', de: 'eng+fra+deu+spa', es: 'eng+fra+deu+spa' };
(function pickLang() {                       // recognition language: last choice, else the browser's language
  let v = null; try { v = localStorage.getItem('ocrLang'); } catch (e) { /* storage blocked */ }
  if (!v) { const l = (navigator.language || 'en').toLowerCase(); v = LANGS[l] || LANGS[l.split('-').slice(0, 2).join('-')] || LANGS[l.split('-')[0]] || 'eng'; }
  if ([...$('lang').options].some((o) => o.value === v)) $('lang').value = v;
  $('lang').addEventListener('change', () => { try { localStorage.setItem('ocrLang', $('lang').value); } catch (e) { /* storage blocked */ } });
})();

$('file').addEventListener('change', (e) => { convert(e.target.files); e.target.value = ''; });
const drop = $('drop');
drop.addEventListener('keydown', (e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); $('file').click(); } });
['dragenter', 'dragover'].forEach((t) => drop.addEventListener(t, (e) => { e.preventDefault(); drop.classList.add('over'); }));
['dragleave', 'drop'].forEach((t) => drop.addEventListener(t, (e) => { e.preventDefault(); drop.classList.remove('over'); }));
drop.addEventListener('drop', (e) => convert(e.dataTransfer.files));
document.addEventListener('paste', (e) => {                            // Ctrl+V a screenshot
  const fs = [...(e.clipboardData && e.clipboardData.items || [])].filter((i) => i.kind === 'file').map((i) => i.getAsFile()).filter(Boolean);
  if (!fs.length) return; e.preventDefault();
  convert(fs.map((f, i) => (f.name && f.name !== 'image.png' ? f : new File([f], `pasted image${fs.length > 1 ? ' ' + (i + 1) : ''}.png`, { type: f.type || 'image/png' }))));
});
document.querySelectorAll('input[name=mode], #num').forEach((el) => el.addEventListener('change', refresh));
$('sheet').addEventListener('change', preview);
$('dl').addEventListener('click', () => state.wb && XLSX.writeFile(state.wb, state.base + '.xlsx'));
$('csv').addEventListener('click', () => {
  if (!state.wb) return; const n = $('sheet').value;
  download(new Blob(['﻿' + XLSX.utils.sheet_to_csv(state.wb.Sheets[n])], { type: 'text/csv;charset=utf-8' }), `${state.base} - ${n}.csv`);
});
if (DONATE_URL) { const d = $('donate'); d.href = DONATE_URL; d.style.display = 'inline-flex'; }
window.__convert = convert; window.__state = state;   // (used by the test page)

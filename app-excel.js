/* Excel viewer (xlsx / xls / ods)
 *
 * Uses SheetJS to read workbooks in the browser. Unlike CSV, every Excel
 * cell knows its type (number, text, date, boolean, error) and whether it
 * holds a formula, so the inconsistency checks can be far more precise.
 *
 * Edits are written straight into the SheetJS workbook object, so the
 * export keeps all sheets, formulas and number formats of the original.
 *
 * Security notes
 * - Everything taken from the file (sheet names, headers, values, formulas)
 *   is HTML-escaped or set via textContent, so a crafted workbook cannot
 *   inject markup or scripts.
 * - Macros (VBA) are never loaded or run, hyperlinks in cells are ignored.
 * - No inline event handlers, so the page runs under a strict
 *   Content-Security-Policy (see excel-viewer.html).
 * - The file is read locally; this script makes no network requests.
 */
'use strict';

/* ── State ─────────────────────────────────────────────────────── */
let workbook    = null;   // SheetJS workbook, edits are written into it
let sheets      = [];     // one model per worksheet (see buildSheet)
let current     = null;   // active sheet model
let viewRows    = [];     // rows shown after search, rule filter and sort
let sortCol     = -1;
let sortDir     = 1;
let highlightOn = true;
let activeRule  = null;   // rule id, 'any' or null
let currentFile = '';
let dirty       = false;

/* ── Inconsistency rules ───────────────────────────────────────── */
const RULES = {
  error:      { sev: 'error', label: 'Excel error' },
  empty:      { sev: 'error', label: 'Gap in filled column' },
  numtext:    { sev: 'warn',  label: 'Number stored as text' },
  notnum:     { sev: 'warn',  label: 'Non-number in number column' },
  datetext:   { sev: 'warn',  label: 'Date stored as text' },
  notdate:    { sev: 'warn',  label: 'Non-date in date column' },
  hardcoded:  { sev: 'warn',  label: 'Typed value in formula column' },
  duplicate:  { sev: 'warn',  label: 'Duplicate ID' },
  negative:   { sev: 'warn',  label: 'Negative amount' },
  whitespace: { sev: 'warn',  label: 'Leading / trailing spaces' },
  outlier:    { sev: 'info',  label: 'Statistical outlier' },
};
const SEV_RANK  = { error: 3, warn: 2, info: 1 };
const SEV_CLASS = { error: 'empty', warn: 'suspect', info: 'info' };   // CSS classes from style.css

const AMOUNT_WORDS = ['amount', 'total', 'price', 'sum', 'revenue', 'cost',
                      'betrag', 'summe', 'preis', 'netto', 'brutto', 'umsatz', 'kosten'];

const EXCEL_ERRORS = { 0x00: '#NULL!', 0x07: '#DIV/0!', 0x0F: '#VALUE!', 0x17: '#REF!',
                       0x1D: '#NAME?', 0x24: '#NUM!', 0x2A: '#N/A', 0x2B: '#GETTING_DATA' };

/* ── Theme (shares its setting with the CSV viewer) ────────────── */
const htmlEl = document.documentElement;
setTheme(readStoredTheme(), false);

function readStoredTheme() {
  try { return localStorage.getItem('csvviewer-theme') || 'dark'; } catch { return 'dark'; }
}

function setTheme(t, save = true) {
  htmlEl.setAttribute('data-theme', t);
  $('themeIcon').textContent  = t === 'dark' ? '☀️' : '🌙';
  $('themeLabel').textContent = t === 'dark' ? 'Light' : 'Dark';
  if (save) { try { localStorage.setItem('csvviewer-theme', t); } catch {} }
}

function toggleTheme() {
  setTheme(htmlEl.getAttribute('data-theme') === 'dark' ? 'light' : 'dark');
}

/* ── Buttons ───────────────────────────────────────────────────── */
$('themeBtn').addEventListener('click', toggleTheme);
$('reloadBtn').addEventListener('click', resetAll);
$('hlBtn').addEventListener('click', toggleHighlight);
$('resetSortBtn').addEventListener('click', resetSort);
$('csvBtn').addEventListener('click', exportSheetCSV);
$('exportBtn').addEventListener('click', exportXLSX);

/* ── Loading ───────────────────────────────────────────────────── */
const dz = $('dropzone');
dz.addEventListener('click', () => $('fileInput').click());
dz.addEventListener('dragover',  e => { e.preventDefault(); dz.classList.add('over'); });
dz.addEventListener('dragleave', ()  => dz.classList.remove('over'));
dz.addEventListener('drop', e => {
  e.preventDefault();
  dz.classList.remove('over');
  handleFile(e.dataTransfer.files[0]);
});
$('fileInput').addEventListener('change', e => handleFile(e.target.files[0]));

if (typeof XLSX === 'undefined') {
  dz.querySelector('p').textContent = 'SheetJS is missing. Put xlsx.full.min.js into the vendor folder (see vendor/README-vendor.md).';
} else if (versionBelow(XLSX.version, '0.20.2')) {
  // 0.18.5 from npm/cdnjs still circulates and has known security issues
  dz.querySelector('p').textContent =
    `SheetJS ${XLSX.version} has known security issues with crafted files. Replace vendor/xlsx.full.min.js with 0.20.3 (see vendor/README-vendor.md).`;
  dz.classList.add('warning');
}

function versionBelow(v, min) {
  const a = String(v).split('.').map(Number), b = min.split('.').map(Number);
  for (let i = 0; i < 3; i++) if ((a[i] || 0) !== b[i]) return (a[i] || 0) < b[i];
  return false;
}

function handleFile(file) {
  if (!file) return;
  setStatus('Loading ' + file.name + ' …');

  file.arrayBuffer().then(buf => {
    // cellHTML: false, SheetJS would otherwise build HTML for rich text we never use
    workbook = XLSX.read(buf, { type: 'array', cellDates: true, cellNF: true, cellHTML: false });
    currentFile = file.name;
    dirty = false;
    sheets = workbook.SheetNames.map((name, i) => buildSheet(name, workbook.Sheets[name], i));
    sheets.forEach(analyze);

    hide('dropzone');
    show('sheetTabs');
    show('toolbar');
    show('tableWrapper');
    show('reloadBtn');

    updateFileBadge();
    renderTabs();
    selectSheet(0);
  }).catch(err => {
    setStatus('Could not read ' + file.name + ': ' + err.message);
  });
}

/* Turn a SheetJS worksheet into a row/column model.
 * The first non-empty row is taken as the header row. Cell objects are
 * the worksheet's own objects, so edits land directly in the workbook. */
function buildSheet(name, ws, index) {
  let minR = Infinity, maxR = -1, minC = Infinity, maxC = -1;

  // Scan real cells instead of trusting !ref, which is often far too large
  for (const addr of Object.keys(ws)) {
    if (addr[0] === '!') continue;
    if (isBlank(ws[addr])) continue;
    const { r, c } = XLSX.utils.decode_cell(addr);
    if (r < minR) minR = r;
    if (r > maxR) maxR = r;
    if (c < minC) minC = c;
    if (c > maxC) maxC = c;
  }

  const model = { name, index, ws, headers: [], cols: [], rows: [], profiles: [], ruleCounts: {} };
  if (maxR < 0) return model;

  for (let c = minC; c <= maxC; c++) model.cols.push(c);

  model.headers = model.cols.map(c => {
    const h = ws[XLSX.utils.encode_cell({ r: minR, c })];
    return isBlank(h) ? 'Column ' + XLSX.utils.encode_col(c) : displayValue(h).trim();
  });

  for (let r = minR + 1; r <= maxR; r++) {
    const cells = model.cols.map(c => ws[XLSX.utils.encode_cell({ r, c })] || null);
    if (cells.every(isBlank)) continue;
    model.rows.push({ i: model.rows.length, r, cells, issues: [], total: false });
  }

  // A SUM/SUBTOTAL row at the very end is a totals row, not data
  model.rows.slice(-3).forEach(row => {
    row.total = row.cells.some(cell => cell && cell.f && /\b(SUM|SUBTOTAL|AGGREGATE)\(/i.test(cell.f));
  });

  return model;
}

/* ── Analysis ──────────────────────────────────────────────────── */
function analyze(sheet) {
  const dataRows = sheet.rows.filter(r => !r.total);
  sheet.profiles = sheet.headers.map((h, c) => profileColumn(h, dataRows.map(r => r.cells[c])));
  sheet.ruleCounts = {};

  sheet.rows.forEach(row => {
    row.issues = row.cells.map((cell, c) => {
      const found = row.total
        ? (cell && cell.t === 'e' ? [['error', 'Excel error ' + errorText(cell)]] : [])
        : checkCell(cell, sheet.profiles[c]);
      found.forEach(([id]) => { sheet.ruleCounts[id] = (sheet.ruleCounts[id] || 0) + 1; });
      return found;
    });
  });
}

/* What does a "normal" cell in this column look like? */
function profileColumn(header, cells) {
  const filled = cells.filter(cell => !isBlank(cell));
  const n      = filled.length;
  const nums   = filled.filter(cell => cell.t === 'n');
  const dates  = filled.filter(cell => cell.t === 'd');

  const p = {
    n,
    fill: cells.length ? n / cells.length : 0,
    kind: 'text',
    formulaShare: n ? filled.filter(cell => cell.f).length / n : 0,
    isAmount: AMOUNT_WORDS.some(w => header.toLowerCase().includes(w)),
    dupCounts: null,
    magMedian: null,   // median of log10(|value|), for the outlier check
    magMad: null,
    typical: null,     // median of |value|, shown in the tooltip
  };

  if (n) {
    if (nums.length / n > 0.7)       p.kind = 'number';
    else if (dates.length / n > 0.7) p.kind = 'date';
  }

  // Duplicates only matter in ID-like columns that are (almost) unique anyway
  if (isIdHeader(header) && n >= 5) {
    const counts = new Map();
    filled.forEach(cell => { const k = idKey(cell); counts.set(k, (counts.get(k) || 0) + 1); });
    if (counts.size / n >= 0.9) p.dupCounts = counts;
  }

  // Outliers by order of magnitude: robust z-score (Iglewicz & Hoaglin) on
  // log10(|value|), so 450 next to 140 is normal but 99,999 is not.
  // Only for amounts or columns with decimals; account numbers and other
  // integer codes would otherwise flag every code that differs.
  const hasDecimals = nums.some(cell => !Number.isInteger(cell.v));
  if (p.kind === 'number' && (p.isAmount || hasDecimals)) {
    const nonZero = nums.map(cell => Math.abs(cell.v)).filter(v => v > 0);
    if (nonZero.length >= 10) {
      const mags = nonZero.map(Math.log10).sort((a, b) => a - b);
      const med  = median(mags);
      const mad  = median(mags.map(m => Math.abs(m - med)).sort((a, b) => a - b));
      if (mad > 0) {
        p.magMedian = med;
        p.magMad    = mad;
        p.typical   = median(nonZero.sort((a, b) => a - b));
      }
    }
  }

  return p;
}

/* Returns a list of [ruleId, reason] for one cell. */
function checkCell(cell, p) {
  const out = [];

  if (isBlank(cell)) {
    if (p.n >= 5 && p.fill >= 0.8) out.push(['empty', `Empty, while ${pct(p.fill)} of this column is filled`]);
    return out;
  }

  if (cell.t === 'e') {
    out.push(['error', 'Excel error ' + errorText(cell)]);
    return out;
  }

  if (cell.t === 's') {
    const s = String(cell.v);
    if (s !== s.trim()) out.push(['whitespace', 'Leading or trailing spaces']);
    const parsed = parseLoose(s);
    if (p.kind === 'number') {
      out.push(parsed && parsed.t === 'n'
        ? ['numtext', 'Number stored as text, Excel will not calculate with it']
        : ['notnum',  'Text in a column of numbers']);
    } else if (p.kind === 'date') {
      out.push(parsed && parsed.t === 'd'
        ? ['datetext', 'Date stored as text, sorting and filtering will treat it as text']
        : ['notdate',  'Text in a column of dates']);
    }
  } else if (p.kind === 'number' && cell.t !== 'n') {
    out.push(['notnum', 'Not a number in a column of numbers']);
  } else if (p.kind === 'date' && cell.t !== 'd') {
    out.push(['notdate', 'Not a date in a column of dates']);
  }

  if (p.n >= 3 && p.formulaShare >= 0.7 && !cell.f) {
    out.push(['hardcoded', `Typed-in value, while ${pct(p.formulaShare)} of this column are formulas`]);
  }

  if (p.dupCounts) {
    const count = p.dupCounts.get(idKey(cell));
    if (count > 1) out.push(['duplicate', `Appears ${count}× in an ID column`]);
  }

  if (cell.t === 'n') {
    if (p.isAmount && cell.v < 0) out.push(['negative', 'Negative value in an amount column']);
    if (p.magMad !== null && cell.v !== 0) {
      const z = 0.6745 * (Math.log10(Math.abs(cell.v)) - p.magMedian) / p.magMad;
      if (Math.abs(z) > 3.5) {
        out.push(['outlier', `Unusually ${z > 0 ? 'large' : 'small'} for this column (typical about ${fmtNum(p.typical)})`]);
      }
    }
  }

  return out;
}

function isIdHeader(header) {
  const h = header.toLowerCase().replace(/[^a-zäöüß0-9]+$/, '');
  const tokens = h.split(/[^a-zäöüß0-9]+/);
  return tokens.some(t => ['id', 'nr', 'no', 'ref'].includes(t)) ||
         /(nummer|number|nr)$/.test(h);
}

function idKey(cell) {
  return cell.t === 'd' ? 'd' + cell.v.getTime() : String(cell.v).trim();
}

/* ── Sheets ────────────────────────────────────────────────────── */
function renderTabs() {
  $('sheetTabs').innerHTML = sheets.map(s => {
    const n = (s.ruleCounts.error || 0) + Object.keys(s.ruleCounts)
      .filter(id => RULES[id].sev === 'warn')
      .reduce((sum, id) => sum + s.ruleCounts[id], 0);
    const badge = n ? `<span class="count" title="${n} cells flagged">${n}</span>` : '';
    return `<button class="sheet-tab${s === current ? ' active' : ''}" data-sheet="${s.index}">${esc(s.name)}${badge}</button>`;
  }).join('');
}

$('sheetTabs').addEventListener('click', e => {
  const tab = e.target.closest('[data-sheet]');
  if (tab) selectSheet(+tab.dataset.sheet);
});

function selectSheet(index) {
  current    = sheets[index];
  sortCol    = -1;
  sortDir    = 1;
  activeRule = null;
  $('searchInput').value = '';
  renderTabs();
  applyView();
  setStatus(current.rows.length
    ? `${current.name}: ${current.rows.length} rows, ${current.headers.length} columns`
    : `${current.name} is empty`);
}

/* ── Search, rule filter, sort ─────────────────────────────────── */
$('searchInput').addEventListener('input', applyView);

function applyView() {
  const q = $('searchInput').value.toLowerCase();

  viewRows = current.rows.filter(row => {
    if (q && !row.cells.some(cell => displayValue(cell).toLowerCase().includes(q))) return false;
    if (!activeRule) return true;
    return row.issues.some(list => list.some(([id]) => activeRule === 'any' || id === activeRule));
  });

  if (sortCol >= 0) {   // totals rows stay at the bottom
    const body   = viewRows.filter(r => !r.total).sort((a, b) => compareCells(a.cells[sortCol], b.cells[sortCol]));
    const totals = viewRows.filter(r => r.total);
    viewRows = body.concat(totals);
  }

  renderTable();
  renderIssues();
  fitTable();          // after the chips, whose height depends on the screen width
}

function sortBy(c) {
  sortDir = (sortCol === c) ? sortDir * -1 : 1;
  sortCol = c;
  applyView();
}

function resetSort() {
  sortCol = -1;
  sortDir = 1;
  applyView();
}

/* Blank cells always go last; numbers and dates before text. */
function compareCells(a, b) {
  const ka = sortKey(a), kb = sortKey(b);
  if (ka === null && kb === null) return 0;
  if (ka === null) return 1;
  if (kb === null) return -1;
  if (typeof ka === 'number' && typeof kb === 'number') return (ka - kb) * sortDir;
  if (typeof ka === 'number') return -1 * sortDir;
  if (typeof kb === 'number') return  1 * sortDir;
  return ka.localeCompare(kb, undefined, { numeric: true, sensitivity: 'base' }) * sortDir;
}

function sortKey(cell) {
  if (isBlank(cell)) return null;
  if (cell.t === 'n') return cell.v;
  if (cell.t === 'd') return cell.v.getTime();
  if (cell.t === 'b') return cell.v ? 1 : 0;
  return displayValue(cell);
}

/* ── Highlight toggle and issue chips ──────────────────────────── */
function toggleHighlight() {
  highlightOn = !highlightOn;
  $('hlBtn').textContent = highlightOn ? '🔍 Inconsistencies' : '○ Inconsistencies';
  if (!highlightOn) activeRule = null;
  applyView();
}

function renderIssues() {
  const bar = $('issuesBar');
  if (!highlightOn || !current.rows.length) { hide('issuesBar'); return; }
  show('issuesBar');

  const ids = Object.keys(current.ruleCounts)
    .sort((a, b) => SEV_RANK[RULES[b].sev] - SEV_RANK[RULES[a].sev]);

  if (!ids.length) {
    bar.innerHTML = '<span class="issues-ok">No inconsistencies found in this sheet</span>';
    return;
  }

  const flaggedRows = current.rows.filter(row => row.issues.some(list => list.length)).length;
  const chip = (id, label, n, sev) =>
    `<button class="chip sev-${sev}${activeRule === id ? ' active' : ''}" data-rule="${id}">` +
    `${label}<span class="n">${n}</span></button>`;

  bar.innerHTML =
    '<span class="issues-label">Show rows with</span>' +
    chip('any', 'Any issue', flaggedRows, 'any') +
    ids.map(id => chip(id, RULES[id].label, current.ruleCounts[id], RULES[id].sev)).join('');
}

$('issuesBar').addEventListener('click', e => {
  const chip = e.target.closest('[data-rule]');
  if (!chip) return;
  activeRule = activeRule === chip.dataset.rule ? null : chip.dataset.rule;
  applyView();
});

/* ── Table ─────────────────────────────────────────────────────── */
function renderTable() {
  const wrapper = $('tableWrapper');

  if (!current.rows.length) {
    wrapper.innerHTML = '<p class="empty-sheet">This sheet has no data.</p>';
    updateStats();
    return;
  }

  let out = '<table><thead><tr><th class="rn" title="Row number in Excel">Row</th>';
  current.headers.forEach((h, c) => {
    const sorted = sortCol === c;
    const icon   = sorted ? (sortDir === 1 ? '▲' : '▼') : '⇅';
    const letter = XLSX.utils.encode_col(current.cols[c]);
    out += `<th class="${sorted ? 'sorted' : ''}" data-c="${c}">` +
           `<span class="col-letter">${letter}</span>${esc(h)}<span class="sort-icon">${icon}</span></th>`;
  });
  out += '</tr></thead><tbody>';

  viewRows.forEach(row => {
    const rnTitle = row.total ? ' title="Totals row, skipped by the checks"' : '';
    out += `<tr><td class="rn${row.total ? ' total' : ''}"${rnTitle}>${row.r + 1}</td>`;
    row.cells.forEach((cell, c) => {
      const { cls, title } = cellLook(row, c);
      out += `<td class="${cls}"${title ? ` title="${esc(title)}"` : ''} contenteditable="true" ` +
             `data-i="${row.i}" data-c="${c}">${esc(displayValue(cell))}</td>`;
    });
    out += '</tr>';
  });

  out += '</tbody></table>';
  wrapper.innerHTML = out;
  updateStats();
}

function cellLook(row, c) {
  const cell   = row.cells[c];
  const issues = highlightOn ? row.issues[c] : [];
  const classes = [];
  if (cell && (cell.t === 'n' || cell.t === 'd')) classes.push('num');
  if (issues.length) {
    const top = issues.reduce((best, [id]) =>
      SEV_RANK[RULES[id].sev] > SEV_RANK[best] ? RULES[id].sev : best, 'info');
    classes.push(SEV_CLASS[top]);
  }
  return { cls: classes.join(' '), title: issues.map(([, reason]) => reason).join('\n') };
}

/* Update colours and tooltips in place, without rebuilding the table
 * (a rebuild would steal focus from the cell the user is moving to). */
function refreshHighlights() {
  $('tableWrapper').querySelectorAll('td[data-c]').forEach(td => {
    const { cls, title } = cellLook(current.rows[+td.dataset.i], +td.dataset.c);
    td.className = cls;
    if (title) td.title = title; else td.removeAttribute('title');
  });
  updateStats();
  renderIssues();
  renderTabs();
  fitTable();
}

function updateStats() {
  const counts = { error: 0, warn: 0, info: 0 };
  Object.entries(current.ruleCounts).forEach(([id, n]) => { counts[RULES[id].sev] += n; });

  $('rowCount').textContent  = viewRows.length === current.rows.length
    ? `${current.rows.length} rows`
    : `${viewRows.length} of ${current.rows.length} rows`;
  $('errCount').textContent  = highlightOn && counts.error ? `✗ ${counts.error} errors`   : '';
  $('warnCount').textContent = highlightOn && counts.warn  ? `⚠ ${counts.warn} warnings`  : '';
  $('infoCount').textContent = highlightOn && counts.info  ? `ℹ ${counts.info} outliers` : '';
}

/* Let the table fill the window down to the status bar. */
function fitTable() {
  const wrapper = $('tableWrapper');
  const top     = wrapper.getBoundingClientRect().top;
  const status  = $('status').offsetHeight;
  wrapper.style.maxHeight = Math.max(200, window.innerHeight - top - status) + 'px';
}
window.addEventListener('resize', () => { if (current) fitTable(); });

/* Header clicks sort */
$('tableWrapper').addEventListener('click', e => {
  const th = e.target.closest('th[data-c]');
  if (th) sortBy(+th.dataset.c);
});

/* ── Editing ───────────────────────────────────────────────────── */
const wrapperEl = $('tableWrapper');

// On focus, show the raw value (like Excel's formula bar) and explain the cell
wrapperEl.addEventListener('focusin', e => {
  const td = e.target.closest('td[data-c]');
  if (!td) return;
  const row  = current.rows[+td.dataset.i];
  const c    = +td.dataset.c;
  const cell = row.cells[c];

  td.textContent   = editValue(cell);
  td.dataset.orig  = td.textContent;
  selectContents(td);

  const addr    = XLSX.utils.encode_cell({ r: row.r, c: current.cols[c] });
  const parts   = [addr];
  if (cell && cell.f) parts.push('=' + cell.f);
  row.issues[c].forEach(([, reason]) => parts.push(reason));
  setStatus(parts.join('   '));
});

wrapperEl.addEventListener('focusout', e => {
  const td = e.target.closest('td[data-c]');
  if (td) commitEdit(td);
});

wrapperEl.addEventListener('keydown', e => {
  const td = e.target.closest('td[data-c]');
  if (!td) return;
  if (e.key === 'Enter' && !e.shiftKey) {
    e.preventDefault();
    const below = td.parentElement.nextElementSibling;
    td.blur();
    if (below) below.children[td.cellIndex].focus();
  } else if (e.key === 'Escape') {
    td.textContent = td.dataset.orig;
    td.blur();
  }
});

function commitEdit(td) {
  const row  = current.rows[+td.dataset.i];
  const c    = +td.dataset.c;
  const text = td.innerText.replace(/\n+$/, '');

  if (text === td.dataset.orig) {           // nothing changed: keep type and formula
    td.textContent = displayValue(row.cells[c]);
    return;
  }

  const addr   = XLSX.utils.encode_cell({ r: row.r, c: current.cols[c] });
  const parsed = parseLoose(text);

  if (!parsed) {
    delete current.ws[addr];
    row.cells[c] = null;
  } else {
    const cell = row.cells[c] || (current.ws[addr] = {});
    const oldT = cell.t;
    ['f', 'F', 'w', 'h', 'r', 'R'].forEach(k => delete cell[k]);   // a typed value replaces any formula
    cell.t = parsed.t;
    cell.v = parsed.v;

    // Numbers and dates take the column's format (e.g. "#,##0.00 €"),
    // which also replaces a text format "@" that caused number-as-text.
    if (parsed.t === 'n' || parsed.t === 'd') {
      if (oldT !== parsed.t || !cell.z || cell.z === '@') {
        const z = columnFormat(c, parsed.t, cell) || (parsed.t === 'd' ? 'yyyy-mm-dd' : null);
        if (z) cell.z = z; else delete cell.z;
      }
    } else {
      delete cell.z;
    }
    if (parsed.t === 'n' && cell.z) {
      try { cell.w = XLSX.SSF.format(cell.z, cell.v); } catch { /* fall back to plain number */ }
    }
    row.cells[c] = cell;
  }

  td.textContent = displayValue(row.cells[c]);
  dirty = true;
  updateFileBadge();
  analyze(current);
  refreshHighlights();
  setStatus(`${addr} changed`);
}

function columnFormat(c, type, except) {
  const donor = current.rows
    .map(r => r.cells[c])
    .find(x => x && x !== except && x.t === type && x.z && x.z !== 'General' && x.z !== '@');
  return donor ? donor.z : null;
}

function selectContents(el) {
  const range = document.createRange();
  range.selectNodeContents(el);
  const sel = window.getSelection();
  sel.removeAllRanges();
  sel.addRange(range);
}

/* ── Export ────────────────────────────────────────────────────── */
function baseName() {
  return currentFile.replace(/\.[^.]+$/, '') || 'workbook';
}

function exportXLSX() {
  XLSX.writeFile(workbook, 'edited_' + baseName() + '.xlsx', { bookType: 'xlsx' });
  dirty = false;
  updateFileBadge();
  setStatus('Exported edited_' + baseName() + '.xlsx');
}

function exportSheetCSV() {
  const csv  = '﻿' + XLSX.utils.sheet_to_csv(current.ws);   // BOM so Excel reads UTF-8
  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  const url  = URL.createObjectURL(blob);
  const a    = Object.assign(document.createElement('a'), {
    href:     url,
    download: `${baseName()}_${current.name}.csv`,
  });
  a.click();
  URL.revokeObjectURL(url);
  setStatus('Exported sheet "' + current.name + '" as CSV');
}

/* ── Reset to start screen ─────────────────────────────────────── */
function resetAll() {
  if (dirty && !confirm('You have unsaved edits. Discard them?')) return;

  workbook = null; sheets = []; current = null; viewRows = [];
  sortCol = -1; sortDir = 1; highlightOn = true; activeRule = null;
  currentFile = ''; dirty = false;

  $('searchInput').value       = '';
  $('hlBtn').textContent       = '🔍 Inconsistencies';
  $('tableWrapper').innerHTML  = '';
  $('sheetTabs').innerHTML     = '';
  $('issuesBar').innerHTML     = '';
  $('fileInput').value         = '';
  updateFileBadge();

  ['tableWrapper', 'toolbar', 'sheetTabs', 'issuesBar', 'reloadBtn'].forEach(hide);
  show('dropzone');
  setStatus('Ready – load an Excel file to get started');
}

window.addEventListener('beforeunload', e => {
  if (dirty) { e.preventDefault(); e.returnValue = ''; }
});

/* ── Values: display, edit, parse ──────────────────────────────── */
function isBlank(cell) {
  if (!cell || cell.t === 'z' || cell.v === undefined || cell.v === null) return true;
  return cell.t === 's' && String(cell.v).trim() === '';
}

function displayValue(cell) {
  if (isBlank(cell)) return cell && cell.t === 's' ? String(cell.v) : '';
  if (cell.t === 'd') return formatDate(cell.v);   // ignore w: SheetJS renders dates US-style
  if (cell.t === 'e') return errorText(cell);
  if (cell.t === 'b') return cell.v ? 'TRUE' : 'FALSE';
  if (cell.w !== undefined) return cell.w;
  return String(cell.v);
}

function editValue(cell) {
  if (isBlank(cell)) return cell && cell.t === 's' ? String(cell.v) : '';
  if (cell.t === 'n') return String(cell.v);
  return displayValue(cell);
}

function errorText(cell) {
  return cell.w || EXCEL_ERRORS[cell.v] || '#ERROR';
}

/* Read a typed string as number, date, boolean or text.
 * Accepts 1234.5 / 1.234,50 (German) / 1,234.50 (English) and an optional currency sign,
 * plus dates as 2026-03-15 or 15.03.2026. Returns null for an empty string. */
function parseLoose(raw) {
  const s = raw.trim();
  if (s === '') return null;

  const n = s.replace(/^[€$£]\s*|\s*[€$£]$/g, '');
  if (/^[-+]?\d+(\.\d+)?$/.test(n))                                  return { t: 'n', v: parseFloat(n) };
  if (/^[-+]?\d{1,3}(\.\d{3})*(,\d+)?$/.test(n) || /^[-+]?\d+,\d+$/.test(n))
    return { t: 'n', v: parseFloat(n.replace(/\./g, '').replace(',', '.')) };
  if (/^[-+]?\d{1,3}(,\d{3})+(\.\d+)?$/.test(n))                     return { t: 'n', v: parseFloat(n.replace(/,/g, '')) };

  let m;
  if ((m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/)))      { const d = makeDate(+m[1], +m[2], +m[3]); if (d) return { t: 'd', v: d }; }
  if ((m = s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4}|\d{2})$/))) {
    const year = m[3].length === 2 ? 2000 + +m[3] : +m[3];
    const d = makeDate(year, +m[2], +m[1]);
    if (d) return { t: 'd', v: d };
  }

  if (/^(true|wahr)$/i.test(s))   return { t: 'b', v: true };
  if (/^(false|falsch)$/i.test(s)) return { t: 'b', v: false };

  return { t: 's', v: raw.replace(/\n+$/, '') };
}

function makeDate(y, mo, d) {
  if (mo < 1 || mo > 12 || d < 1 || d > 31) return null;
  const date = new Date(y, mo - 1, d);
  return date.getMonth() === mo - 1 ? date : null;
}

function formatDate(d) {
  const p = x => String(x).padStart(2, '0');
  const date = `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}`;
  return (d.getHours() || d.getMinutes()) ? `${date} ${p(d.getHours())}:${p(d.getMinutes())}` : date;
}

/* ── Helpers ───────────────────────────────────────────────────── */
function $(id)    { return document.getElementById(id); }
function show(id) { $(id).classList.remove('hidden'); }
function hide(id) { $(id).classList.add('hidden'); }
function setStatus(msg) { $('status').textContent = msg; }

function updateFileBadge() {
  $('filename').textContent = currentFile ? currentFile + (dirty ? ' (edited)' : '') : 'no file loaded';
}

/* Escape text for use inside HTML elements and quoted attributes */
function esc(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;')
                  .replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

function median(sorted) {
  const mid = Math.floor(sorted.length / 2);
  return sorted.length % 2 ? sorted[mid] : (sorted[mid - 1] + sorted[mid]) / 2;
}

function pct(x)    { return Math.round(x * 100) + ' %'; }
function fmtNum(x) { return Number(x.toFixed(2)).toLocaleString(); }

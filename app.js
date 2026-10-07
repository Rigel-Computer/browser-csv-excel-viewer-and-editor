/* CSV viewer
 *
 * Security notes
 * - Everything taken from the file (headers and values) is HTML-escaped
 *   before it goes into innerHTML, so a crafted CSV cannot inject markup
 *   or scripts.
 * - No inline event handlers: all events are bound here, which lets the
 *   page run under a strict Content-Security-Policy (see index.html).
 * - The file is read locally; this script makes no network requests.
 */
'use strict';

/* ── State ── */
let allData      = [];
let headers      = [];
let filteredRows = [];
let sortCol      = -1;
let sortDir      = 1;
let highlightOn  = true;
let currentFile  = '';

/* ── Theme ─────────────────────────────────────────────────────── */
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
$('exportBtn').addEventListener('click', exportCSV);

/* ── Drag & Drop / File input ───────────────────────────────────── */
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

if (typeof Papa === 'undefined') {
  dz.querySelector('p').textContent = 'PapaParse is missing. Put papaparse.min.js into the vendor folder (see vendor/README-vendor.md).';
}

function handleFile(file) {
  if (!file) return;
  currentFile = file.name;
  $('filename').textContent = file.name;
  setStatus('Loading ' + file.name + ' …');

  Papa.parse(file, {
    header: true,
    skipEmptyLines: true,
    complete(res) {
      headers      = res.meta.fields || [];
      allData      = res.data;
      filteredRows = [...allData];

      show('toolbar');
      show('tableWrapper');
      show('reloadBtn');
      hide('dropzone');

      renderTable();
      setStatus('Loaded: ' + allData.length + ' rows, ' + headers.length + ' columns');
    },
    error(err) { setStatus('Error: ' + err.message); }
  });
}

/* ── Search / Filter ────────────────────────────────────────────── */
$('searchInput').addEventListener('input', function () {
  const q = this.value.toLowerCase();
  filteredRows = q
    ? allData.filter(row => headers.some(h => String(row[h] ?? '').toLowerCase().includes(q)))
    : [...allData];
  if (sortCol >= 0) sortBy(sortCol, false);
  else renderTable();
});

/* ── Highlight toggle ───────────────────────────────────────────── */
function toggleHighlight() {
  highlightOn = !highlightOn;
  $('hlBtn').textContent = highlightOn ? '🔍 Inconsistencies' : '○ Inconsistencies';
  renderTable();
}

/* ── Sort ───────────────────────────────────────────────────────── */
function resetSort() {
  sortCol = -1; sortDir = 1;
  const q = $('searchInput').value.toLowerCase();
  filteredRows = q
    ? allData.filter(row => headers.some(h => String(row[h] ?? '').toLowerCase().includes(q)))
    : [...allData];
  renderTable();
}

function sortBy(colIdx, toggle = true) {
  if (toggle) {
    sortDir = (sortCol === colIdx) ? sortDir * -1 : 1;
    sortCol = colIdx;
  }
  const h = headers[colIdx];
  filteredRows.sort((a, b) => {
    const an = parseFloat(a[h]), bn = parseFloat(b[h]);
    return (!isNaN(an) && !isNaN(bn))
      ? (an - bn) * sortDir
      : String(a[h] ?? '').localeCompare(String(b[h] ?? '')) * sortDir;
  });
  renderTable();
}

/* ── Inconsistency classifier ───────────────────────────────────── */
function classify(val, header, colVals) {
  if (val === '' || val === null || val === undefined) return 'empty';
  const s = String(val).trim();
  if (!s || s === '-' || s.toLowerCase() === 'null' || s.toLowerCase() === 'n/a') return 'empty';

  // Numeric column heuristic (>70 % of non-empty values are numbers)
  const nonEmpty     = colVals.filter(v => v !== '');
  const numericCount = nonEmpty.filter(v => !isNaN(parseFloat(v))).length;
  if (nonEmpty.length > 0 && numericCount / nonEmpty.length > 0.7 && isNaN(parseFloat(s)))
    return 'suspect';

  const lh = header.toLowerCase();

  // Negative value in an amount-like column
  const isAmount = ['amount','total','price','sum','revenue','cost'].some(k => lh.includes(k));
  if (isAmount && parseFloat(s) < 0) return 'suspect';

  // Invalid date in a date-like column
  const isDate = ['date','time','created','updated'].some(k => lh.includes(k));
  if (isDate && isNaN(Date.parse(s))) return 'suspect';

  return '';
}

/* ── Render table ───────────────────────────────────────────────── */
function renderTable() {
  const wrapper = $('tableWrapper');

  // Pre-collect column values for the classifier
  const colVals = {};
  headers.forEach(h => { colVals[h] = filteredRows.map(r => String(r[h] ?? '')); });

  let warns = 0, errs = 0;
  let out = '<table><thead><tr><th class="rn">#</th>';

  headers.forEach((h, i) => {
    const sorted = sortCol === i;
    const icon   = sorted ? (sortDir === 1 ? '▲' : '▼') : '⇅';
    out += `<th class="${sorted ? 'sorted' : ''}" data-c="${i}">${esc(h)}<span class="sort-icon">${icon}</span></th>`;
  });
  out += '</tr></thead><tbody>';

  filteredRows.forEach((row, ri) => {
    out += `<tr><td class="rn">${ri + 1}</td>`;
    headers.forEach((h, c) => {
      const val = row[h] ?? '';
      let cls   = '';
      if (highlightOn) {
        cls = classify(val, h, colVals[h]);
        if (cls === 'empty')   errs++;
        if (cls === 'suspect') warns++;
      }
      out += `<td class="${cls}" contenteditable="true" data-row="${ri}" data-c="${c}">${esc(val)}</td>`;
    });
    out += '</tr>';
  });

  out += '</tbody></table>';
  wrapper.innerHTML = out;

  $('rowCount').textContent  = filteredRows.length + ' rows';
  $('warnCount').textContent = (highlightOn && warns) ? `⚠ ${warns} suspect`   : '';
  $('errCount').textContent  = (highlightOn && errs)  ? `✗ ${errs} empty/null` : '';
}

/* Header clicks sort, leaving a cell saves it */
$('tableWrapper').addEventListener('click', e => {
  const th = e.target.closest('th[data-c]');
  if (th) sortBy(+th.dataset.c);
});
$('tableWrapper').addEventListener('focusout', e => {
  const td = e.target.closest('td[data-c]');
  if (td) cellEdit(td);
});

/* ── Inline cell edit ───────────────────────────────────────────── */
function cellEdit(td) {
  const row = filteredRows[+td.dataset.row];
  const h   = headers[+td.dataset.c];
  const val = td.innerText.trim();

  row[h] = val;   // rows are shared objects, so this also updates allData

  if (highlightOn) {
    const colVals = filteredRows.map(r => String(r[h] ?? ''));
    td.className  = classify(val, h, colVals) || '';
  }
}

/* ── Export ─────────────────────────────────────────────────────── */
function exportCSV() {
  const csv  = Papa.unparse({ fields: headers, data: allData });
  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  const url  = URL.createObjectURL(blob);
  const a    = Object.assign(document.createElement('a'), {
    href:     url,
    download: currentFile ? 'edited_' + currentFile : 'export.csv'
  });
  a.click();
  URL.revokeObjectURL(url);
}

/* ── Reset to start screen ──────────────────────────────────────── */
function resetAll() {
  allData = []; headers = []; filteredRows = [];
  sortCol = -1; sortDir = 1; highlightOn = true; currentFile = '';

  $('filename').textContent    = 'no file loaded';
  $('searchInput').value       = '';
  $('hlBtn').textContent       = '🔍 Inconsistencies';
  $('tableWrapper').innerHTML  = '';
  $('fileInput').value         = '';

  hide('tableWrapper');
  hide('toolbar');
  hide('reloadBtn');
  show('dropzone');
  setStatus('Ready – load a CSV file to get started');
}

/* ── Helpers ────────────────────────────────────────────────────── */
function $(id)    { return document.getElementById(id); }
function show(id) { $(id).classList.remove('hidden'); }
function hide(id) { $(id).classList.add('hidden'); }
function setStatus(msg) { $('status').textContent = msg; }

/* Escape text for use inside HTML elements and quoted attributes */
function esc(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;')
                  .replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

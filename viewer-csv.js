/* CSV viewer
 *
 * Parses CSV/TSV with PapaParse. All values stay text, exactly as in the
 * file, so the export changes nothing but the cells you edited.
 * Uses the helpers from app.js ($, esc, setStatus, fitTable, download …).
 */
'use strict';

const CsvViewer = (() => {

  /* ── State ── */
  let allData      = [];
  let headers      = [];
  let filteredRows = [];
  let sortCol      = -1;
  let sortDir      = 1;
  let highlightOn  = true;
  let currentFile  = '';
  let dirty        = false;

  /* ── Toolbar ───────────────────────────────────────────────────── */
  $('csvHlBtn').addEventListener('click', toggleHighlight);
  $('csvResetSortBtn').addEventListener('click', resetSort);
  $('csvExportBtn').addEventListener('click', exportCSV);
  $('csvSearch').addEventListener('input', () => {
    applyFilter();
    if (sortCol >= 0) sortBy(sortCol, false);
    else renderTable();
  });

  /* ── Open ──────────────────────────────────────────────────────── */
  function open(file) {
    return new Promise((resolve, reject) => {
      Papa.parse(file, {
        header: true,
        skipEmptyLines: true,
        complete(res) {
          headers      = res.meta.fields || [];
          allData      = res.data;
          filteredRows = [...allData];
          currentFile  = file.name;
          dirty        = false;
          showView('csvView');
          renderTable();
          resolve(`Loaded: ${allData.length} rows, ${headers.length} columns`);
        },
        error(err) { reject(err); }
      });
    });
  }

  /* ── Search / filter ───────────────────────────────────────────── */
  function applyFilter() {
    const q = $('csvSearch').value.toLowerCase();
    filteredRows = q
      ? allData.filter(row => headers.some(h => String(row[h] ?? '').toLowerCase().includes(q)))
      : [...allData];
  }

  /* ── Highlight toggle ──────────────────────────────────────────── */
  function toggleHighlight() {
    highlightOn = !highlightOn;
    $('csvHlBtn').textContent = highlightOn ? '🔍 Inconsistencies' : '○ Inconsistencies';
    renderTable();
  }

  /* ── Sort ──────────────────────────────────────────────────────── */
  function resetSort() {
    sortCol = -1; sortDir = 1;
    applyFilter();
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

  /* ── Inconsistency classifier ──────────────────────────────────── */
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
    const isAmount = ['amount', 'total', 'price', 'sum', 'revenue', 'cost'].some(k => lh.includes(k));
    if (isAmount && parseFloat(s) < 0) return 'suspect';

    // Invalid date in a date-like column
    const isDate = ['date', 'time', 'created', 'updated'].some(k => lh.includes(k));
    if (isDate && isNaN(Date.parse(s))) return 'suspect';

    return '';
  }

  /* ── Render ────────────────────────────────────────────────────── */
  function renderTable() {
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
    $('csvTable').innerHTML = out;

    $('csvRowCount').textContent  = filteredRows.length + ' rows';
    $('csvWarnCount').textContent = (highlightOn && warns) ? `⚠ ${warns} suspect`   : '';
    $('csvErrCount').textContent  = (highlightOn && errs)  ? `✗ ${errs} empty/null` : '';
    fit();
  }

  function fit() { fitTable($('csvTable')); }

  /* Header clicks sort, leaving a cell saves it */
  $('csvTable').addEventListener('click', e => {
    const th = e.target.closest('th[data-c]');
    if (th) sortBy(+th.dataset.c);
  });
  $('csvTable').addEventListener('focusout', e => {
    const td = e.target.closest('td[data-c]');
    if (td) cellEdit(td);
  });

  /* ── Inline cell edit ──────────────────────────────────────────── */
  function cellEdit(td) {
    const row = filteredRows[+td.dataset.row];
    const h   = headers[+td.dataset.c];
    const val = td.innerText.trim();
    if (val === String(row[h] ?? '')) return;

    row[h] = val;   // rows are shared objects, so this also updates allData
    dirty  = true;
    setFileBadge(currentFile, true);

    if (highlightOn) {
      const colVals = filteredRows.map(r => String(r[h] ?? ''));
      td.className  = classify(val, h, colVals) || '';
    }
  }

  /* ── Export ────────────────────────────────────────────────────── */
  function exportCSV() {
    const csv = Papa.unparse({ fields: headers, data: allData });
    download(new Blob([csv], { type: 'text/csv;charset=utf-8;' }),
             currentFile ? 'edited_' + currentFile : 'export.csv');
    dirty = false;
    setFileBadge(currentFile, false);
  }

  /* ── Back to the start screen ──────────────────────────────────── */
  function reset() {
    allData = []; headers = []; filteredRows = [];
    sortCol = -1; sortDir = 1; highlightOn = true; currentFile = ''; dirty = false;
    $('csvSearch').value      = '';
    $('csvHlBtn').textContent = '🔍 Inconsistencies';
    $('csvTable').innerHTML   = '';
    hide('csvView');
  }

  return {
    open, reset, fit,
    isDirty:     () => dirty,
    ready:       () => typeof Papa !== 'undefined',
    missingText: () => 'PapaParse is missing. Put papaparse.min.js into the vendor folder (see vendor/README-vendor.md).',
  };
})();

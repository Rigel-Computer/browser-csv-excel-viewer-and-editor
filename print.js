/* Printing
 *
 * Prints only the table that is on screen, with the current search, filter
 * and sort, under a short heading (file, sheet, filter, date). The layout
 * lives in print.css. Works with the Print buttons and with Ctrl+P alike,
 * because the heading is built in the browser's "beforeprint" event.
 *
 * Reads everything it needs from the page itself, so the viewers don't need
 * to know about printing.
 */
'use strict';

(() => {
  const header    = document.getElementById('printHeader');
  const pageTitle = document.title;

  document.querySelectorAll('.print-btn').forEach(btn =>
    btn.addEventListener('click', () => window.print()));

  window.addEventListener('beforeprint', prepare);
  window.addEventListener('afterprint',  cleanUp);

  function prepare() {
    header.textContent = '';
    const view  = ['xlView', 'csvView'].map(id => document.getElementById(id))
                                       .find(el => !el.classList.contains('hidden'));
    const table = view && view.querySelector('table');

    if (!table) {
      line('title', 'Nothing to print: load a file and open a sheet with data first.');
      return;
    }

    // Heading: file and sheet
    const shown = document.getElementById('filename').textContent;
    const sheet = textWithout(view.querySelector('.sheet-tab.active'), '.count');
    line('title', shown + (sheet ? ' – ' + sheet : ''));

    // Browsers offer the page title as file name for "Save as PDF":
    // "Buchungen 2025.xlsx", sheet "Abschluss" -> "Buchungen 2025 - Abschluss"
    const base = shown.replace(/ \(edited\)$/, '').replace(/\.[^.]+$/, '');
    document.title = safeFileName(sheet ? `${base} - ${sheet}` : base) || pageTitle;

    // What is shown: rows, search, issue filter, sort
    const meta = [];
    const rows = table.tBodies[0] ? table.tBodies[0].rows.length : 0;
    meta.push(`${rows} ${rows === 1 ? 'row' : 'rows'}`);

    const search = view.querySelector('input[type="text"]').value.trim();
    if (search) meta.push(`search "${search}"`);

    const chip = textWithout(view.querySelector('.chip.active'), '.n');
    if (chip) meta.push(`only rows with: ${chip}`);

    const sorted = view.querySelector('th.sorted');
    if (sorted) {
      const dir = sorted.querySelector('.sort-icon').textContent.includes('▲') ? 'ascending' : 'descending';
      meta.push(`sorted by ${textWithout(sorted, '.sort-icon, .col-letter')} ${dir}`);
    }

    meta.push('printed ' + timestamp());
    line('meta', meta.join(', '));

    // Legend, only if highlighted cells are in the printout
    const kinds = [];
    if (table.querySelector('td.empty'))   kinds.push(view.id === 'xlView' ? 'red = error' : 'red = empty');
    if (table.querySelector('td.suspect')) kinds.push(view.id === 'xlView' ? 'yellow = warning' : 'yellow = suspect');
    if (table.querySelector('td.info'))    kinds.push('blue = outlier');
    if (kinds.length) line('meta', 'Shaded cells: ' + kinds.join(', '));

    // Wide tables go on landscape pages (see "@page wide" in print.css).
    // The row-number column doesn't count.
    const columns = table.rows[0].cells.length - 1;
    document.documentElement.classList.toggle('print-wide',  columns > 6);
    document.documentElement.classList.toggle('print-xwide', columns > 10);
  }

  function cleanUp() {
    document.documentElement.classList.remove('print-wide', 'print-xwide');
    document.title = pageTitle;
  }

  /* ── Helpers ── */
  function line(cls, text) {
    const div = document.createElement('div');
    div.className   = cls;
    div.textContent = text;          // text only, never HTML
    header.appendChild(div);
  }

  // Text of an element without some of its parts (counters, icons)
  function textWithout(el, selector) {
    if (!el) return '';
    const copy = el.cloneNode(true);
    copy.querySelectorAll(selector).forEach(n => n.remove());
    return copy.textContent.trim();
  }

  // Characters that Windows or macOS don't allow in file names become "-"
  function safeFileName(name) {
    return name.replace(/[\\/:*?"<>|\u0000-\u001f]+/g, '-')
               .replace(/\s+/g, ' ')
               .replace(/^[\s.-]+|[\s.-]+$/g, '');   // no leading/trailing dots, dashes, spaces
  }

  function timestamp() {
    const d = new Date(), p = x => String(x).padStart(2, '0');
    return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}`;
  }
})();

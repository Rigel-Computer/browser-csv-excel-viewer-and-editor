/* CSV + Excel viewer: start screen and shared parts
 *
 * The start screen shows two drop fields, CSV on the left and Excel on the
 * right. Both accept any file: the viewer is chosen from the file's first
 * bytes and its extension, so a file dropped on the "wrong" field still
 * opens in the right viewer. The two viewers live in viewer-csv.js and
 * viewer-excel.js.
 *
 * Security notes
 * - Everything taken from a file is HTML-escaped (esc) or set via
 *   textContent, so a crafted file cannot inject markup or scripts.
 * - No inline event handlers, so the page runs under the strict
 *   Content-Security-Policy set in index.html.
 * - Files are read locally; none of the scripts makes a network request.
 */
'use strict';

let active = null;          // viewer currently shown (CsvViewer / ExcelViewer)

/* ── Helpers shared by both viewers ────────────────────────────── */
function $(id)    { return document.getElementById(id); }
function show(id) { $(id).classList.remove('hidden'); }
function hide(id) { $(id).classList.add('hidden'); }
function setStatus(msg) {
  $('status').textContent = msg;
  if (active) active.fit();   // a longer message can wrap and make the status bar taller
}

/* Escape text for use inside HTML elements and quoted attributes */
function esc(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;')
                  .replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

function setFileBadge(name, edited) {
  $('filename').textContent = name ? name + (edited ? ' (edited)' : '') : 'no file loaded';
}

/* Let a table fill the window down to the status bar */
function fitTable(wrapper) {
  const top    = wrapper.getBoundingClientRect().top;
  const status = $('status').offsetHeight;
  wrapper.style.maxHeight = Math.max(200, window.innerHeight - top - status) + 'px';
}

/* Hand a file to the browser's download */
function download(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a   = Object.assign(document.createElement('a'), { href: url, download: filename });
  a.click();
  URL.revokeObjectURL(url);
}

/* Switch from the start screen to a viewer section */
function showView(id) {
  hide('start');
  show(id);
  show('reloadBtn');
}

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

$('themeBtn').addEventListener('click', () => {
  setTheme(htmlEl.getAttribute('data-theme') === 'dark' ? 'light' : 'dark');
});

/* ── File type detection ───────────────────────────────────────── */
const SIG_OLE2 = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];   // .xls (old binary Office format)
const SIG_ZIP  = [0x50, 0x4B];                                       // "PK": .xlsx and .ods are ZIP files
const SIG_PDF  = [0x25, 0x50, 0x44, 0x46];                           // "%PDF"

/* Returns 'csv', 'excel', or null for files neither viewer can open.
 * Content wins over the extension, so a renamed file still opens correctly. */
async function detectKind(file) {
  const head = new Uint8Array(await file.slice(0, 1024).arrayBuffer());
  const ext  = (/\.([^.]+)$/.exec(file.name) || ['', ''])[1].toLowerCase();
  const startsWith = sig => sig.every((byte, i) => head[i] === byte);

  if (startsWith(SIG_OLE2) || startsWith(SIG_ZIP)) return 'excel';
  if (startsWith(SIG_PDF) || head.includes(0))     return null;      // binary, but no spreadsheet

  // Plain text from here on
  if (['xls', 'xlsx', 'ods'].includes(ext)) {
    // Many programs export HTML or XML tables as ".xls"; SheetJS reads those.
    // Anything else with an Excel extension is a CSV that was saved as .xls.
    const text = new TextDecoder().decode(head).replace(/^﻿/, '').trimStart();
    return text.startsWith('<') ? 'excel' : 'csv';
  }
  return 'csv';
}

/* ── Start screen ──────────────────────────────────────────────── */
const READY_TEXT = 'Ready – drop a CSV or Excel file to get started';
const fileInput  = $('fileInput');
let pickedFrom   = null;    // field that opened the file dialog

document.querySelectorAll('.dropzone').forEach(zone => {
  const field = zone.dataset.kind;
  const pick  = () => { pickedFrom = field; fileInput.click(); };

  zone.addEventListener('click', pick);
  zone.addEventListener('keydown', e => {
    if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); pick(); }
  });
  zone.addEventListener('dragover', e => {
    e.preventDefault();
    e.dataTransfer.dropEffect = 'copy';
    zone.classList.add('over');
  });
  zone.addEventListener('dragleave', e => {
    if (!zone.contains(e.relatedTarget)) zone.classList.remove('over');
  });
  zone.addEventListener('drop', e => {
    e.preventDefault();
    zone.classList.remove('over');
    openFile(e.dataTransfer.files[0], field);
  });
});

fileInput.addEventListener('change', () => {
  openFile(fileInput.files[0], pickedFrom);
  fileInput.value = '';
});

// A drop next to the fields would make the browser open the file itself and
// leave the page (losing any edits). Catch it and show a "not allowed" cursor.
document.addEventListener('dragover', e => {
  e.preventDefault();
  if (!(e.target instanceof Element && e.target.closest('.dropzone'))) e.dataTransfer.dropEffect = 'none';
});
document.addEventListener('drop', e => e.preventDefault());

async function openFile(file, field) {
  if (!file) return;
  if (file.size === 0) { setStatus(`"${file.name}" is empty.`); return; }

  let kind;
  try {
    kind = await detectKind(file);
  } catch (err) {
    setStatus(`Could not read "${file.name}": ${err.message}`);
    return;
  }
  if (!kind) {
    setStatus(`"${file.name}" is neither a CSV nor an Excel file.`);
    return;
  }

  const viewer = kind === 'excel' ? ExcelViewer : CsvViewer;
  if (!viewer.ready()) { setStatus(viewer.missingText()); return; }

  setStatus('Loading ' + file.name + ' …');
  try {
    const summary = await viewer.open(file);
    active = viewer;           // set before setStatus, which refits the table
    setFileBadge(file.name, false);
    const note = field && field !== kind ? ` (recognised as ${kind === 'excel' ? 'Excel' : 'CSV'} file)` : '';
    setStatus(summary + note);
  } catch (err) {
    setStatus(`Could not read "${file.name}": ${err.message}`);
  }
}

/* ── New file / leaving the page ───────────────────────────────── */
$('reloadBtn').addEventListener('click', () => {
  if (active && active.isDirty() && !confirm('You have unsaved edits. Discard them?')) return;
  if (active) active.reset();
  active = null;
  hide('reloadBtn');
  show('start');
  setFileBadge('', false);
  setStatus(READY_TEXT);
});

window.addEventListener('beforeunload', e => {
  if (active && active.isDirty()) { e.preventDefault(); e.returnValue = ''; }
});

window.addEventListener('resize', () => { if (active) active.fit(); });

/* ── Library checks, shown in the matching field ───────────────── */
function versionBelow(v, min) {
  const a = String(v).split('.').map(Number), b = min.split('.').map(Number);
  for (let i = 0; i < 3; i++) if ((a[i] || 0) !== b[i]) return (a[i] || 0) < b[i];
  return false;
}

function warnInField(kind, text) {
  const zone = document.querySelector(`.dropzone[data-kind="${kind}"]`);
  zone.querySelector('.hint').textContent = text;
  zone.classList.add('warning');
}

// app.js loads first (the viewers use its helpers), so wait until all scripts are in
document.addEventListener('DOMContentLoaded', () => {
  if (!CsvViewer.ready()) warnInField('csv', CsvViewer.missingText());
  if (!ExcelViewer.ready()) {
    warnInField('excel', ExcelViewer.missingText());
  } else if (versionBelow(XLSX.version, '0.20.2')) {
    // 0.18.5 from npm/cdnjs still circulates and has known security issues
    warnInField('excel', `SheetJS ${XLSX.version} has known security issues with crafted files. ` +
                         'Replace vendor/xlsx.full.min.js with 0.20.3 (see vendor/README-vendor.md).');
  }
  setStatus(READY_TEXT);
});

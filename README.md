# browser-csv-excel-viewer-and-editor

A lightweight CSV **and Excel** viewer and editor that runs entirely in your browser. No server, no installation, no Excel required. Your files never leave your computer.

![License](https://img.shields.io/badge/license-MIT-blue.svg)
![HTML](https://img.shields.io/badge/no%20build%20step-just%20open-orange.svg)

---

**English** · [Deutsch](README.de.md)

## One page, two viewers

`index.html` opens with two drop fields side by side (stacked on narrow screens): CSV on the left, Excel on the right.

| Viewer | Opens | Export |
|---|---|---|
| CSV | `.csv`, `.tsv`, `.txt` | CSV |
| Excel | `.xlsx`, `.xls`, `.ods`, HTML/XML tables saved as `.xls` | `.xlsx` (whole workbook) or the current sheet as CSV |

**Both fields take any file.** The viewer is chosen by what the file contains, so a file dropped on the "wrong" field still opens correctly; the status bar then says *recognised as Excel file* (or CSV):

| File starts with | Opens in |
|---|---|
| `D0 CF 11 E0 A1 B1 1A E1` (old binary `.xls`) | Excel viewer |
| `PK` (ZIP: `.xlsx`, `.ods`) | Excel viewer |
| `<` and has an Excel extension (HTML/XML export named `.xls`) | Excel viewer |
| other text | CSV viewer, also when it is named `.xls` |
| other binary data, e.g. PDF or images | rejected with a message |

Files dropped next to the fields are caught, so the browser does not leave the page to show the file itself.

## Features

- **Drag & drop** or file picker, from either field
- **Column sorting** — click any header to sort ascending/descending
- **Live search / filter** across all columns simultaneously
- **Inconsistency highlighting** (see below)
- **Inline cell editing** — click any cell to edit it directly
- **Export** the modified data
- **Dark / Light theme** — toggle in the header, preference is remembered
- **New file** button — return to the start screen without reloading the page; asks first if there are unsaved edits

Excel viewer only:

- **Sheet tabs** with a count of flagged cells per sheet
- **Issue chips** — one chip per problem type with its count; click a chip to show only the rows that have that problem
- **Excel row numbers and column letters**, so every finding can be located in the original file
- **Cell details in the status bar** — address, formula and the reason a cell is flagged
- **Edits keep the workbook intact** — formulas, number formats and all other sheets survive the export; an edited number takes the column's number format
- **Totals rows** (`SUM`/`SUBTOTAL` at the end of a sheet) are recognised, skipped by the checks and kept at the bottom when sorting
- Keyboard: <kbd>Enter</kbd> saves and moves down, <kbd>Esc</kbd> discards the edit

## Usage

1. Download or clone the repo
2. Open `index.html` in any modern browser
3. Drop your file onto one of the two fields — that's it

No build step, no npm, no backend, no internet connection needed.

If the Excel field says *SheetJS is missing*, download `xlsx.full.min.js` once into `vendor/` as described in [vendor/README-vendor.md](vendor/README-vendor.md).

## Files

| File | Description |
|---|---|
| `index.html` | The page: start screen and both viewers, markup only |
| `app.js` | Start screen, file type detection, theme and helpers shared by both viewers |
| `viewer-csv.js` | CSV viewer |
| `viewer-excel.js` | Excel viewer, including the inconsistency rules |
| `style.css` | Styles, CSS custom-property theming (dark + light) |
| `vendor/` | Local copies of PapaParse and SheetJS, with versions and checksums in its README |

Keep the folder structure as it is.

## Inconsistency Detection

### CSV viewer

CSV has no types, so the CSV viewer works with simple heuristics:

| Highlight | Meaning |
|---|---|
| 🔴 Red cell | Empty, `null`, `N/A`, or `-` value |
| 🟡 Yellow cell | Text in a column where >70 % of values are numeric |
| 🟡 Yellow cell | Negative number in a column named *amount*, *price*, *total*, *cost*, *revenue*, or *sum* |
| 🟡 Yellow cell | Unparseable value in a column named *date*, *time*, *created*, or *updated* |

### Excel viewer

Excel cells know their type (number, text, date, error) and whether they contain a formula. The Excel viewer first works out what is normal for each column, then flags the cells that break the pattern. Hover a cell to see why it was flagged.

| | Rule | Flags a cell when |
|---|---|---|
| 🔴 | Excel error | it shows `#DIV/0!`, `#REF!`, `#N/A`, `#VALUE!` … |
| 🔴 | Gap in filled column | it is empty while at least 80 % of the column is filled |
| 🟡 | Number stored as text | it is text that reads as a number (`1.250,00`, `€ 99`) in a number column — Excel does not add these up |
| 🟡 | Non-number in number column | it holds other text in a number column |
| 🟡 | Date stored as text | it is text that reads as a date (`15.03.2026`) in a date column |
| 🟡 | Non-date in date column | it holds anything else in a date column |
| 🟡 | Typed value in formula column | it holds a typed-in value where at least 70 % of the column are formulas — a classic overwritten formula |
| 🟡 | Duplicate ID | its value appears more than once in an ID-like column (*ID*, *Nr*, *No*, *…nummer*, *…number*) that is otherwise unique |
| 🟡 | Negative amount | it is negative in an amount column (*amount*, *total*, *price*, *betrag*, *summe*, *netto*, *brutto* …) |
| 🟡 | Leading / trailing spaces | text starts or ends with a space — breaks lookups and grouping |
| 🔵 | Statistical outlier | its size is far off the rest of the column (robust z-score on the order of magnitude, amount and decimal columns only) |

A column counts as a number or date column when more than 70 % of its filled cells are numbers or dates.

## Limitations of the Excel export

The free SheetJS build does not write cell styles, so fill colours, fonts and borders of the original are not kept. Formulas are kept, but their stored results are not recalculated in the browser; if you changed a value a formula depends on, recalculate in Excel or LibreOffice (<kbd>F9</kbd>). A typed value always replaces a formula in that cell. `.xls` and `.ods` files are exported as `.xlsx`.

## Security

The viewers are built for files you may not fully trust, such as exports from other systems.

- **Your data stays in the browser.** Files are read locally. The page makes no network request: the libraries are local copies in `vendor/`, no CDN, no web fonts, no tracking.
- **Network access is switched off.** The page sets a Content-Security-Policy that blocks every connection (`connect-src 'none'`), external images and form posts, and only allows scripts from the viewer's own folder. Even a compromised library could not send your data away through these channels. A CSP cannot stop a script from navigating the whole page to another address, which is why the libraries are pinned, local and checked (see below).
- **Crafted files cannot run code.** Everything taken from a file (headers, cell values, sheet names, formulas) is escaped before it is displayed. Excel macros are never loaded or executed; links in cells are ignored.
- **Exports contain your data as it is.** If a file from someone else contains cells that start with `=`, Excel will treat them as formulas when you open the export, just as it would with the original. The viewers neither add nor remove such content.

## Dependencies

Both libraries live in `vendor/`, so the code that runs is part of the repository and can be reviewed. Versions, licenses, sources and checksums are listed in [vendor/README-vendor.md](vendor/README-vendor.md).

- [PapaParse 5.4.1](https://www.papaparse.com/) (MIT) — CSV parsing and export. No known vulnerabilities in this version.
- [SheetJS CE 0.20.3](https://sheetjs.com/) (Apache-2.0) — Excel reading and writing. Contains the fixes for [CVE-2023-30533](https://github.com/advisories/GHSA-4r6h-8v6p-xvw6) and [CVE-2024-22363](https://github.com/advisories/GHSA-5pgg-2g8v-p4x9). Only available from cdn.sheetjs.com; the `xlsx` package on npm, cdnjs, jsDelivr and unpkg is the vulnerable 0.18.5. Security scanners still flag 0.20.3 because the fix never reached npm, which is a false positive.

## Browser Support

Works in all modern browsers (Chrome, Firefox, Safari, Edge). No Internet Explorer support.

## Disclaimer

This code was written in whole or in part with the help of generative AI (Claude by Anthropic). It was tested with automated browser tests, including deliberately crafted files, but it has not been independently audited.

The software is provided "as is", without warranty of any kind (see [LICENSE](LICENSE)). **Use it at your own risk.** Keep a backup of every file before you edit and export it. The inconsistency checks are heuristics: they help find problems, but they can miss some and flag correct data. They do not replace a proper review of your accounts.

## License

[MIT](LICENSE): you may use, copy, modify and distribute this code, also commercially, as long as the copyright notice and the license text stay with every copy. There is no warranty.

The libraries in `vendor/` keep their own licenses: PapaParse is MIT, SheetJS is Apache-2.0. Their license texts are in `vendor/` and must be passed on with the files.

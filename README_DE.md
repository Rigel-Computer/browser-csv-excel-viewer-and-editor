# browser-csv-excel-viewer-and-editor

Ein schlanker Viewer und Editor für CSV- **und Excel**-Dateien, der vollständig im Browser läuft. Kein Server, keine Installation, kein Excel erforderlich. Die Dateien verlassen den eigenen Rechner nicht.

![Lizenz](https://img.shields.io/badge/lizenz-MIT-blue.svg)
![HTML](https://img.shields.io/badge/kein%20Build-einfach%20öffnen-orange.svg)

---

## Zwei Viewer, ein Look

| Seite | Öffnet | Export |
|---|---|---|
| `index.html` | CSV / TSV | CSV |
| `excel-viewer.html` | `.xlsx`, `.xls`, `.ods` | `.xlsx` (ganze Arbeitsmappe) oder das aktuelle Blatt als CSV |

Beide teilen sich Stylesheet, Theme-Einstellung und Bedienung und verlinken im Header aufeinander.

## Funktionen

- **Drag & Drop** oder Dateiauswahl
- **Spalten sortieren** — Klick auf einen Spaltenkopf sortiert auf- oder absteigend
- **Live-Suche / Filter** über alle Spalten gleichzeitig
- **Inkonsistenz-Highlighting** (siehe unten)
- **Zellen direkt bearbeiten** — einfach in eine Zelle klicken
- **Export** der bearbeiteten Daten
- **Dark- / Light-Theme** — Schalter im Header, Einstellung wird gespeichert
- **Neue Datei**-Button — zurück zur Startseite ohne Seite neu laden

Nur im Excel-Viewer:

- **Blatt-Reiter** mit der Anzahl markierter Zellen pro Blatt
- **Fehler-Chips** — ein Chip pro Fehlerart mit Anzahl; Klick auf einen Chip zeigt nur die Zeilen mit genau diesem Problem
- **Excel-Zeilennummern und Spaltenbuchstaben**, damit sich jeder Fund in der Originaldatei wiederfinden lässt
- **Zelldetails in der Statusleiste** — Adresse, Formel und Grund der Markierung
- **Bearbeiten lässt die Arbeitsmappe intakt** — Formeln, Zahlenformate und alle anderen Blätter bleiben beim Export erhalten; eine bearbeitete Zahl übernimmt das Zahlenformat der Spalte
- **Summenzeilen** (`SUMME`/`TEILERGEBNIS` am Blattende) werden erkannt, von der Prüfung ausgenommen und bleiben beim Sortieren unten
- Tastatur: <kbd>Enter</kbd> speichert und springt eine Zeile tiefer, <kbd>Esc</kbd> verwirft die Änderung

## Verwendung

1. Repo herunterladen oder klonen
2. `index.html` (CSV) oder `excel-viewer.html` (Excel) in einem modernen Browser öffnen
3. Datei per Drag & Drop auf die Seite ziehen — fertig

Kein Build-Schritt, kein npm, kein Backend, keine Internetverbindung nötig.

Meldet der Excel-Viewer *SheetJS is missing*, einmalig `xlsx.full.min.js` nach `vendor/` herunterladen, wie in [vendor/README-vendor.md](vendor/README-vendor.md) beschrieben.

## Dateien

| Datei | Beschreibung |
|---|---|
| `index.html` | CSV-Viewer, nur Markup |
| `app.js` | Logik des CSV-Viewers |
| `excel-viewer.html` | Excel-Viewer, nur Markup |
| `app-excel.js` | Logik des Excel-Viewers, inklusive der Prüfregeln |
| `style.css` | Gemeinsame Styles, CSS-Custom-Properties für Dark- und Light-Theme |
| `vendor/` | Lokale Kopien von PapaParse und SheetJS, Versionen und Prüfsummen in der dortigen README |

Die Ordnerstruktur bitte so belassen.

## Inkonsistenz-Erkennung

### CSV-Viewer

CSV kennt keine Datentypen, deshalb arbeitet der CSV-Viewer mit einfachen Heuristiken:

| Markierung | Bedeutung |
|---|---|
| 🔴 Rote Zelle | Leerer Wert, `null`, `N/A` oder `-` |
| 🟡 Gelbe Zelle | Text in einer Spalte, in der >70 % der Werte numerisch sind |
| 🟡 Gelbe Zelle | Negativer Betrag in einer Spalte namens *amount*, *price*, *total*, *cost*, *revenue* oder *sum* |
| 🟡 Gelbe Zelle | Ungültiger Wert in einer Spalte namens *date*, *time*, *created* oder *updated* |

### Excel-Viewer

Excel-Zellen kennen ihren Typ (Zahl, Text, Datum, Fehler) und wissen, ob sie eine Formel enthalten. Der Excel-Viewer ermittelt zuerst, was in jeder Spalte normal ist, und markiert dann die Zellen, die aus dem Muster fallen. Mit der Maus über eine Zelle fahren zeigt den Grund.

| | Regel | Markiert eine Zelle, wenn |
|---|---|---|
| 🔴 | Excel-Fehler | sie `#DIV/0!`, `#BEZUG!`, `#NV`, `#WERT!` … anzeigt |
| 🔴 | Lücke in gefüllter Spalte | sie leer ist, obwohl mindestens 80 % der Spalte gefüllt sind |
| 🟡 | Zahl als Text gespeichert | sie Text enthält, der wie eine Zahl aussieht (`1.250,00`, `€ 99`), in einer Zahlenspalte — Excel rechnet damit nicht |
| 🟡 | Keine Zahl in Zahlenspalte | sie sonstigen Text in einer Zahlenspalte enthält |
| 🟡 | Datum als Text gespeichert | sie Text enthält, der wie ein Datum aussieht (`15.03.2026`), in einer Datumsspalte |
| 🟡 | Kein Datum in Datumsspalte | sie in einer Datumsspalte etwas anderes enthält |
| 🟡 | Festwert in Formelspalte | sie einen eingetippten Wert enthält, obwohl mindestens 70 % der Spalte Formeln sind — die klassische überschriebene Formel |
| 🟡 | Doppelte ID | ihr Wert in einer ID-artigen Spalte (*ID*, *Nr*, *No*, *…nummer*, *…number*) mehrfach vorkommt, die sonst eindeutig ist |
| 🟡 | Negativer Betrag | sie in einer Betragsspalte negativ ist (*betrag*, *summe*, *preis*, *netto*, *brutto*, *amount*, *total* …) |
| 🟡 | Leerzeichen am Anfang / Ende | Text mit einem Leerzeichen beginnt oder endet — stört SVERWEIS und Gruppierungen |
| 🔵 | Statistischer Ausreißer | ihre Größenordnung weit vom Rest der Spalte abweicht (robuster z-Wert auf die Größenordnung, nur Betrags- und Dezimalspalten) |

Eine Spalte gilt als Zahlen- bzw. Datumsspalte, wenn mehr als 70 % ihrer gefüllten Zellen Zahlen bzw. Datumswerte sind.

Die Oberfläche ist englisch; die Spaltenerkennung versteht deutsche und englische Überschriften.

## Grenzen des Excel-Exports

Die freie SheetJS-Version schreibt keine Zellformatierung, Füllfarben, Schriften und Rahmen des Originals gehen also verloren. Formeln bleiben erhalten, ihre gespeicherten Ergebnisse werden im Browser aber nicht neu berechnet; wer einen Wert ändert, von dem eine Formel abhängt, lässt Excel oder LibreOffice neu rechnen (<kbd>F9</kbd>). Ein eingetippter Wert ersetzt immer die Formel in dieser Zelle. `.xls`- und `.ods`-Dateien werden als `.xlsx` exportiert.

## Sicherheit

Die Viewer sind für Dateien gebaut, denen man nicht blind vertrauen muss, etwa Exporte aus fremden Systemen.

- **Die Daten bleiben im Browser.** Dateien werden lokal gelesen. Keiner der Viewer stellt eine Netzwerkanfrage: Die Bibliotheken liegen als lokale Kopie in `vendor/`, es gibt kein CDN, keine Webfonts, kein Tracking.
- **Netzwerkzugriff ist abgeschaltet.** Beide Seiten setzen eine Content-Security-Policy, die jede Verbindung (`connect-src 'none'`), externe Bilder und Formular-Versand sperrt und nur Skripte aus dem eigenen Ordner zulässt. Selbst eine kompromittierte Bibliothek könnte die Daten über diese Wege nicht verschicken. Eine CSP kann allerdings nicht verhindern, dass ein Skript die ganze Seite auf eine andere Adresse umleitet; deshalb sind die Bibliotheken fest versioniert, lokal und geprüft (siehe unten).
- **Präparierte Dateien können keinen Code ausführen.** Alles, was aus einer Datei kommt (Überschriften, Zellwerte, Blattnamen, Formeln), wird vor der Anzeige maskiert. Excel-Makros werden nie geladen oder ausgeführt, Links in Zellen werden ignoriert.
- **Exporte enthalten die Daten, wie sie sind.** Enthält eine fremde Datei Zellen, die mit `=` beginnen, behandelt Excel sie beim Öffnen des Exports als Formel, genau wie beim Original. Die Viewer fügen solche Inhalte weder hinzu noch entfernen sie sie.

## Abhängigkeiten

Beide Bibliotheken liegen in `vendor/`, der ausgeführte Code ist also Teil des Repos und lässt sich prüfen. Versionen, Lizenzen, Quellen und Prüfsummen stehen in [vendor/README-vendor.md](vendor/README-vendor.md).

- [PapaParse 5.4.1](https://www.papaparse.com/) (MIT) — Parsen und Export von CSV. Keine bekannten Sicherheitslücken in dieser Version.
- [SheetJS CE 0.20.3](https://sheetjs.com/) (Apache-2.0) — Lesen und Schreiben von Excel-Dateien. Enthält die Korrekturen für [CVE-2023-30533](https://github.com/advisories/GHSA-4r6h-8v6p-xvw6) und [CVE-2024-22363](https://github.com/advisories/GHSA-5pgg-2g8v-p4x9). Nur über cdn.sheetjs.com erhältlich; das Paket `xlsx` auf npm, cdnjs, jsDelivr und unpkg ist das verwundbare 0.18.5. Sicherheitsscanner melden 0.20.3 trotzdem, weil die Korrektur nie auf npm veröffentlicht wurde; das ist ein Fehlalarm.

## Browser-Unterstützung

Funktioniert in allen modernen Browsern (Chrome, Firefox, Safari, Edge). Internet Explorer wird nicht unterstützt.

## Lizenz

MIT — mach damit, was du willst.

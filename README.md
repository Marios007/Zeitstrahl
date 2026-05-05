# Zeitstrahl /timeline

Python-Tool zur automatischen Erstellung von Zeitstrahlen aus einer Excel-Datei.  
Die Ausgabe erfolgt als PNG und/oder PDF — horizontal oder vertikal.

Python tool for automatically generating timelines from an Excel file.  
Output is PNG and/or PDF — horizontal or vertical.

---

## Voraussetzungen / Prerequisites

```
pip install openpyxl matplotlib
```

Eigenständige `.exe` (kein Python nötig) / Standalone `.exe` (no Python required):

```
pyinstaller --onefile --hidden-import matplotlib.backends.backend_agg --hidden-import matplotlib.backends.backend_pdf zeitstrahl.py
```

---

## Verwendung / Usage

1. `Zeitstrahl.xlsx` befüllen (siehe Struktur unten) / Fill in `Zeitstrahl.xlsx` (see structure below)
2. `zeitstrahl.py` oder / or `zeitstrahl.exe` ausführen / run
3. Ausgabe im Ordner / Output in folder `output/`

Dateiname / Filename: `YYYYMMDD_HHMMSS_Zeitstrahl_<Intervall>.png`

---

## Excel-Struktur / Excel Structure

### Events — Spalten / Columns A–F (ab Zeile / from row 2)

| Spalte / Column | Inhalt / Content |
|--------|--------|
| A | Datum / Date |
| B | Uhrzeit / Time |
| C | Quelle / Source |
| D | Kategorie / Category (must be defined in column J) |
| E | Text |
| F | Wichtig / Important (`Ja` / `Nein`) |

### Kategorien / Categories — Spalten / Columns J–M (ab Zeile / from row 2)

| Spalte / Column | Inhalt / Content |
|--------|--------|
| J | Kategoriename / Category name |
| K | Farbe / Color (german: `rot`, `blau`, … or hex code) |
| L | Sichtbar / Visible (`Ja` / `Nein`) |
| M | Stem-Länge / Stem length (number, default: `2.5`) |

### Einstellungen / Settings

| Zelle / Cell | Inhalt / Content |
|-------|--------|
| I2 | Tick-Intervall / Tick interval (`1h`, `30min`, `1d`, …) |
| I19 | Projektname / Project name |
| J19 | Ausrichtung / Orientation (`horizontal`, `vertikal`, `beides` / `both`) |
| K19 | Dateiformat / File format (`png`, `pdf`, `beides` / `both`) |
| N2 / O2 | Startdatum / Startzeit — Start date / time (empty = all data) |
| P2 / Q2 | Enddatum / Endzeit — End date / time (empty = all data) |
| R2 | Uhrzeit im Label anzeigen / Show time in label (`Ja` / `Nein`) |
| S2 | Schriftgröße / Font size (`klein` / `small`, `normal`, `groß` / `large`) |

---

## Hervorhebung wichtiger Events / Highlighting Important Events

Events mit `Ja` in Spalte F / Events with `Ja` in column F:
- Stern-Symbol am Ende des Stems / Star symbol at the stem tip
- Fetterer, größerer Label-Text / Bold, larger label text
- Rahmen um den Label-Text / Box around the label text

---

## Unterstützte Farbnamen / Supported Color Names

`schwarz`, `rot`, `grün`, `blau`, `gelb`, `orange`, `lila`, `violett`, `pink`, `rosa`, `grau`, `braun`, `cyan`, `magenta`, `weiß`

Alternativ / Alternatively: beliebiger Hex-Code / any hex code (e.g. `#FF5733`)

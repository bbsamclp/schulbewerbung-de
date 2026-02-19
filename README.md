# Schulbewerbung CSV-Export

Konvertiert CSV-Exporte aus dem Schulverwaltungssystem in aufbereitete XLSX- und PDF-Dateien, gruppiert nach Bildungsgang.

## Voraussetzungen

- Python 3
- `openpyxl` (`pip install openpyxl`)
- `pdflatex` mit den Paketen: `sourcesanspro`, `xltabular`, `fancyhdr`, `geometry`, `babel`

## Verwendung

CSV-Dateien (Semikolon-getrennt) im Script-Verzeichnis ablegen, dann:

```bash
python3 csv_to_xlsx.py
```

### Optionen

| Flag | Beschreibung |
|------|-------------|
| `-v`, `--verbose` | LaTeX-Logfiles (`.log`) und Quelldateien (`.tex`) behalten |

## Ausgabe

Pro Bildungsgang werden erzeugt:

- **`<Bildungsgang>.xlsx`** -- Tabelle mit Stammdaten und Noten
- **`<Bildungsgang>.pdf`** -- Adressliste im A4-Querformat (sortiert nach Rang, dann Name)
- **`uebersicht.pdf`** -- Zusammenfassung aller Bildungsgange mit Anzahl Datensatze

## Lizenz

Dieses Projekt steht unter der [GNU General Public License v3.0](LICENSE).

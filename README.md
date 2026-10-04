# ExcelPdfExporter

Windows-Tool: Gruppiert ein Excel-Tabellenblatt nach den Werten einer Spalte (Standard: `N`)
und erzeugt für jede Gruppe ein eigenes PDF (AutoFilter + PDF-Export).

**Voraussetzungen:** Windows, Microsoft Excel, Python 3.9+

## Start

```powershell
pip install -r requirements.txt
python app.pyw
```

## EXE bauen (ohne Konsolenfenster)

```powershell
.\build.ps1
```

Ergebnis: `dist\ExcelPdfExporter.exe`

## Ausgabestruktur

```
<Ausgabeordner>/<Prefix>/<Prefix><Term>/<Gruppe>.pdf
```

`Prefix` = erstes Wort des Gruppenwerts, `Term` = Semesterkürzel wie `23W`/`24S`
(ohne Term: `<Prefix>/NO_TERM/`). Standard-Ausgabeordner: `pdf_out` neben der Excel-Datei.

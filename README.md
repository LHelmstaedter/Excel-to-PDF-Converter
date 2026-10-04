ExcelPdfExporter

Kleines Windows-Tool, das aus einer Excel-Datei für jeden Wert einer Spalte ein eigenes PDF erzeugt.

⚠️ Wichtig: Das Tool arbeitet derzeit ausschließlich mit Spalte N. Es durchläuft alle unterschiedlichen Werte in Spalte N und exportiert für jeden Wert die Druckansicht des Blatts als PDF. Eine andere Spalte ist aktuell nicht auswählbar.

Was passiert genau?
Excel-Datei laden (Drag & Drop oder Dateiauswahl).
Das Tool liest die eindeutigen Werte aus Spalte N (erste Zeile = Kopfzeile).
Für jeden Wert wird die Tabelle per AutoFilter gefiltert und in der Druckansicht (Druckbereich/Seiteneinrichtung des Blatts) als PDF gespeichert.

Verwendet wird das erste Blatt mit festgelegtem Druckbereich, sonst das aktive Blatt. Die Originaldatei bleibt unverändert.


Voraussetzungen:
Windows, Microsoft Excel, Python 3.9+

Start
pip install -r requirements.txt
python app.pyw

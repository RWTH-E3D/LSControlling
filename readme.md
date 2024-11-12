# Skript zur Verarbeitung von SAP Berichten für den Lehrstuhl-Controllingeinsatz

## Einsatzbereiche und Anforderungen

Das Skript wurde in Python 3.11 geschrieben und kann zur Datenverarbeitung, -filterung und -aggregation 
von RWTH SAP Berichtsdaten benutzt werden und soll eine Grundlage für die Zahlenermittlung
für das Lehrstuhl-Controlling der Fakultät 3 liefern. Anforderung ist, dass
Zugang zu drei SAP Berichten als .csv exportiert vorliegt.

## Vorgehens- und Berechnungsweise

Für das Controlling soll möglichst viel Logik auf SAP-Auswertungen verlagert werden, damit die Zahlen bestmöglich nachvollzogen werden können. Gleichzeitig soll der Aufwand zur Erstellung des Berichts so gering wie möglich gehalten werden. Dies wird durch eine Kombination aus einem Skript und Exportdateien aus SAP ermöglicht.

Das Programm benötigt vier Eingangsdaten aus SAP für die gesamte Finanzstelle: Stammdaten, Budgets, Drittmittelkontostand und Obligos. Aus den Stammdaten können die Laufzeit der PSP-Elemente sowie die Art der Geldgeber extrahiert werden. Das Budget gibt das zur Verfügung stehende Budget an, wobei der Budgetrest aus dem Vorjahr als End-Budget des Vorjahres betrachtet wird. Eventuelle Festlegungen müssen hiervon noch abgezogen werden, da diese dort nicht berücksichtigt sind. Für einige Projektarten ist das Budget jedoch nicht aussagekräftig in Bezug auf die tatsächlichen Ein- und Ausgaben. Hierfür wird der Drittmittelkontostand herangezogen, welcher pro Projekt kumulativ über die Jahre aufaddiert wird, um den Kontostand am Ende des jeweiligen Jahres zu ermitteln. Existiert dieser für ein PSP-Element, wird der Drittmittelkontostand übernommen; andernfalls wird der aus dem Budget ermittelte Kontostand für weitere Auswertungen verwendet.

Anschließend erstellt das Skript standardmäßig einen PDF-Bericht, einen detaillierteren Textbericht sowie eine CSV-Datei mit einer Projektübersicht zum Ende des letzten Jahres. Diese Dateien umfassen dann die entsprechend den oben genannten Vorgaben erzeugten Kontostände, entweder nach Projektart aggregiert und nach Jahren ausgewertet oder nach Jahren aggregiert, um den aktuellen Zustand der Projekte am letzten Tag des entsprechenden Jahres widerzugeben.

Es ist vorgesehen, dass der PDF-Bericht der Leitung der Einrichtung zur Unterschrift vorgelegt werden kann. Der detailliertere Textbericht oder die CSV-Datei sind eher für buchhalterische Überprüfungen und Nachvollziehbarkeit gedacht. 


## Vorbereitung der Eingangsberichte und -dateien

Folgende 4 Berichte sind notwendig. Diese müssen im Ordner `input` abliegen und
die Namen müssen übereinstimmen!

### `input/WPS_PSP_STAMMDATEN_V1.csv`
Dieser Bericht umfasst die Stammdaten. Er wird folgendermaßen erzeugt:
- SAP Berichtsportal öffnen
- in Finanzberichte wechseln
- Finanzcockpit mit aktuellem Jahr und der gewünschten Kostenstelle öffnen
- Rechtsklick → Springen → "Finanzstelle | Stammdaten HHP"
- im neuen Bericht Rechtsklick → Verteilen und exportieren → Nach CSV exportieren
- Datei mit Namen `WPS_PSP_STAMMDATEN_V1.csv` wird erzeugt
- in den Ordner `input` kopieren

### `input/WFI_001_FC_BUDGET_V1.csv`
Dieser Bericht umfasst alle Budgetdaten. Er wird folgendermaßen erzeugt:
- SAP Berichtsportal öffnen
- in Finanzberichte wechseln
- Finanzcockpit mit aktuellem Jahr und der gewünschten Kostenstelle öffnen
- Rechtsklick → Springen → "Finanzstelle | Budget"
- im neuen Bericht bei "Ergebnismenge einschränken" bei Geschäftsjahr "Alle Werte anzeigen" auswählen
- anschließend Rechtsklick → Verteilen und exportieren → Nach CSV exportieren
- Datei mit Namen `WFI_001_FC_BUDGET_V1.csv` wird erzeugt
- in den Ordner `input` kopieren

### `input/WFI_001_FC_OBLIGOS_V1.csv`
Dieser Bericht umfasst alle bestehenden Festlegungen. Er wird folgendermaßen erzeugt:
- SAP Berichtsportal öffnen
- in Finanzberichte wechseln
- Finanzcockpit mit aktuellem Jahr und der gewünschten Kostenstelle öffnen
- Rechtsklick → Springen → "Finanzstelle | Obligo"
- im neuen Bericht bei "Jahr" bitte "Alle Werte anzeigen" auswählen
- anschließend Rechtsklick → Verteilen und exportieren → Nach CSV exportieren
- Datei mit Namen `WFI_001_FC_OBLIGOS_V1.csv` wird erzeugt
- in den Ordner `input` kopieren


### `input/WPSM_004_KSD.csv`
Dieser Bericht umfasst alle Drittmittelkontostände. Er wird folgendermaßen erzeugt:
- SAP Berichtsportal öffnen
- in Finanzberichte wechseln
- Finanzcockpit mit aktuellem Jahr und der gewünschten Kostenstelle öffnen
- Rechtsklick → Springen → "Finanzstelle | Drittm. Kontostand"
- im neuen Bericht Rechtsklick → Verteilen und exportieren → Nach CSV exportieren
- Datei mit Namen `WPSM_004_KSD.csv` wird erzeugt
- in den Ordner `input` kopieren

### `input/PLOT_PSP.csv` (nicht zwingend erforderlich)
Wird diese Datei abgelegt, muss sie folgende Struktur besitzen: In der ersten Zeile steht `PSP`. 
Anschließend wird die Datei zeilenweise um die PSP-Elemente ergänzt, deren Detailplots erzeugt werden sollen. 
Hier wird noch mal geprüft, ob die angegebenen PSP-Elemente in dem Datensatz existieren, ansonsten werden 
sie ignoriert.


Nun stehen alle Daten soweit bereit um das Skript auszuführen. Vier oder fünf CSV Dateien sollten
im Ordner `input` abliegen. Die Namen dürfen nicht verändert oder angepasst werden.  

## Ausführen des Skriptes

Entweder muss eine Python Version auf dem Rechner installiert sein oder auf einem anderen
Rechner muss aus dem Skript nach der unten angegebenen Anleitung eine ausführbare Datei
erzeugt werden.

### Direktes Ausführen des Python-Skriptes

Erst müssen einige Python Bibliotheken installiert werden, damit das Skript lauffähig ist. Dazu gehören unter anderem
Pandas, Matplotlib, Reportlab usw. Zur kompletten Installation der benötigten Pakete kann die mitgelieferte 
`requirements.txt` Datei genutzt werden:

`pip install -r requirements.txt`

Im Anschluss kann das Skript durch den Aufruf von `python lscontrolling.py` gestartet werden. Der Ordner `input` muss
dabei auf der gleichen Ebene liegen, wie die `lscontrolling.py`.

### Installation ohne Python-Kenntnisse

Das Skript kann ebenfalls auf einem externen Rechner als `.exe` kompiliert werden. 
Hierzu bitte folgendermaßen vorgehen:

Python-Terminal öffnen und folgendes ausführen:
- `pip install pyinstaller`
- `pyinstaller.exe --onedir --clean .\lscontrolling.py`

Anschließend wird ein Ordner `lscontrolling` im Verzeichnis `dist` erstellt. In diesem Ordner liegt eine Datei 
`lscontrolling.exe` und ein weiterer Ordner `_internal` der ignoriert werden kann (aber nicht gelöscht werden darf,
da er alle benötigten, internen Bibliotheken enthält). Dieser Ordner 
kann verteilt und ohne Python genutzt werden. Voraussetzung ist, dass in dem gleichen
Ordner, in dem die Datei liegt, auch das Verzeichnis `input` mit den, wie oben
angegebenen Input-Dateien, liegt.

## Berichte und Dateien

Vom Skript her werden standardmäßig 3 Dateien erzeugt:
- ein PDF-Bericht (als Zusammenfassung und Übersicht für Entscheidungsträger)
- ein Text-Bericht (mit weiteren Details für Personen aus der Buchhaltung und zum Nachvollziehen von einzelnen Kontoständen)
- eine CSV-Datei (Detailkontostände am letzten Tag des Vorjahres, um automatisiert weitere Auswertungen zu ermöglichen)

Der PDF-Bericht bietet dabei eine Zusammenfassung über die gesamte IKZ, gruppiert nach Projektarten und ggf. unterteilt
in Sammelkonten und Einzelkonten die bis zum 30.06. des Vorjahres abgeschlossen sind oder derzeit noch laufen (basierend 
auf dem in SAP angegebenen Projektende).

In der PDF-Zusammenfassungstabelle auf der letzten Berichtsseite sind die Summen eingefärbt, um eine bessere 
Übersichtlichkeit zu ermöglichen. Dabei gilt folgende Vereinbarung:
- in Grün werden Zahlen dargestellt, die einen summierte positiven Kontostand aufweisen
- in Orange werden Zahlen dargestellt, die einen negativen Kontostand bei den noch aktuell laufenden Projekten (also Projekten
mit einem SAP-Projektende, das noch nicht erreicht ist). Diese Zahlen deuten an, dass es ggf. zu Problemen kommen kann,
aufgrund eines negativen Kontostandes. Dies kann aber durch noch ausstehende Zahlungen von Fördermittelgebern noch ausgeglichen werden.
- in Rot werden Zahlen dargestellt, die einen negativen Kontostand bei Sammelkonten oder bei Konten, deren SAP Projektende
bereits mehr als ein halbes Jahr vergangen ist. Es kann sein, dass noch Ausgleichszahlungen anstehen, aber das sollte detailliert
untersucht und im Auge behalten werden.

## Anpassungsmöglichkeiten

Um das Skript möglichst flexibel einsetzen zu können, und den Code nicht jedes Mal manuell anpassen zu müssen, gibt es 
die Möglichkeit das Skript per Konfigurationsdatei anzupassen. Hierzu muss im gleichen Ordner eine Datei mit dem Namen 
`config.ini` abgelegt werden. Innerhalb dieser Datei können Anpassungen an der Vorgehensweise des Skriptes vorgenommen 
werden. Liegt diese Datei nicht vor, werden die Standardeinstellungen im Skript genutzt. Die Datei `template_config.ini`
enthält alle derzeit möglichen Einstellungsschlüssel, die manuell eingestellt werden können. Sollten Schlüsselwörter 
nicht vorkommen, werden die entsprechenden Standardwerte genutzt.


Bei Fragen oder Anmerkungen bitte an [J. Frisch](mailto:frisch@e3d.rwth-aachen.de) schreiben.

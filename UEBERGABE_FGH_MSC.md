# Uebergabe FGH MSC

Hinweis:
Die offizielle Bezeichnung lautet jetzt `FGH Modbus Sollwert Manager`.
Bitte kuenftig `UEBERGABE_MODBUS.md` und den Release-Ordner `dist\Release_FGH_Modbus_Sollwert_Manager` verwenden.

## Projekt
Der **FGH MSC** ist eine Windows-Desktop-Anwendung zur Planung und Ausfuehrung von Modbus-TCP-Testablaeufen. Die Anwendung verbindet sich mit einem Modbus-Geraet, schreibt Sollwerte auf bis zu vier konfigurierbare Kanaele und fuehrt einen Testplan stufenweise aus.

Made by Yunus Sevgi.

Version laut Projektstand:
- Name: `FGH MSC`
- Version: `1.0.0`
- Firma: `FGH`

## Hauptfunktionen
- Verbindung zu einem Modbus-TCP-Geraet ueber IP-Adresse, Port und Slave-ID
- Vier frei konfigurierbare Kanaele `S1` bis `S4`
- Wahl des Datentyps pro Kanal
- Globales Registerformat fuer die Schreibreihenfolge
- Testplan mit bis zu 40 Stufen
- Speichern und Laden von Testplaenen als Excel-Datei (`.xlsx`)
- Ereignisprotokoll fuer Verbindung, Schreibvorgaenge und Fehler
- Manuelles Stoppen eines laufenden Testablaufs

## Start der Anwendung
### Entwicklung / Python-Start
Start ueber die Projektumgebung:

```powershell
.\.venv\Scripts\python.exe start_fgh_msc.py
```

Einstiegspunkt:
- `start_fgh_msc.py`

## EXE / Release
Die uebergabefaehige Windows-Version liegt im Release-Ordner:

```text
dist\Release_FGH_MSC
```

Wichtige Datei:

```text
dist\Release_FGH_MSC\FGH_MSC.exe
```

Hinweis:
- Fuer die Nutzung der EXE ist keine separate Python-Installation notwendig.
- Der komplette Release-Ordner muss zusammenbleiben.
- Es sollten keine Dateien aus dem Release-Ordner einzeln verschoben oder geloescht werden.

## EXE bauen
Build-Skript:

```powershell
.\build_exe.ps1
```

Das Skript:
- baut die EXE per PyInstaller
- legt die eigentliche Build-Ausgabe unter `dist\pyinstaller_dist` ab
- erstellt danach den Release-Ordner `dist\Release_FGH_MSC`
- kopiert `README.md`, die Betriebsanleitung und eine Beispiel-Datei in den Release

## Bedienprinzip
### Verbindung
Fuer die Verbindung werden gesetzt:
- IP-Adresse
- Port
- Slave-ID
- Registerformat

### Kanaele
Die Kanaele `S1` bis `S4` besitzen jeweils:
- eine frei waehlbare Bezeichnung fuer den gesteuerten Sollwert
- ein Startregister
- einen Datentyp

### Testplan
Jede aktive Stufe kann enthalten:
- bis zu vier Sollwerte
- eine Laufzeit in Sekunden

Regeln:
- Leeres Sollwertfeld: Kanal wird in dieser Stufe nicht beschrieben
- Wert `0`: der Wert `0` wird aktiv geschrieben
- Eine Stufe ohne Zeit oder ohne Sollwerte wird nicht ausgefuehrt

## Verhalten beim Stoppen
Der Button `Stopp` beendet den weiteren Ablauf der Sequenz.

Wichtig:
- Bereits erfolgreich geschriebene Werte bleiben im Geraet bestehen.
- Es erfolgt kein automatisches Ruecksetzen auf alte oder sichere Werte.
- Der Stopp beendet die weiteren Stufen des Plans.
- Ein bereits gestarteter einzelner Schreibvorgang wird nicht mitten im Telegramm abgebrochen, sondern in der Regel sauber beendet.

Praktische Bedeutung:
- Nach `Stopp` bleibt normalerweise der zuletzt erfolgreich geschriebene Sollwert im Feld bzw. Geraet aktiv.

## Code-Struktur
### Einstieg
- `start_fgh_msc.py`
  Startet die GUI.

### GUI
- `modbus_gui/main_window.py`
  Enthält das Hauptfenster, die Tabellenlogik, die Bedienung, den Verbindungsablauf, Start/Stopp, Excel Laden/Speichern und die Statusanzeige.

### Ablaufsteuerung
- `modbus_gui/sequence_controller.py`
  Fuehrt den Testplan stufenweise aus, verwaltet Zeiten, Wechsel zwischen Stufen, Fehlerbehandlung und Stoppen.

### Modbus-Kommunikation
- `modbus_gui/modbus_service.py`
  Baut die Modbus-TCP-Verbindung auf, fuehrt Schreibbefehle aus und stellt Keepalive-Funktionen bereit.

### Datenmodelle
- `modbus_gui/models.py`
  Definiert Registerformat, Datentypen, Verbindungseinstellungen, Kanaldefinitionen und die Struktur einer Teststufe.

### Werteumwandlung
- `modbus_gui/value_encoder.py`
  Wandelt eingegebene Werte passend zum Datentyp und Registerformat in Modbus-Registerwerte um.

### App-Metadaten
- `modbus_gui/app_info.py`
  Enthält Name, Version, Beschreibung, Autorhinweis und Logo-Bezug.

## Abhaengigkeiten
Projektabhaengigkeiten laut `requirements.txt`:
- `aiohttp==3.13.3`
- `openpyxl==3.1.5`
- `PyQt5==5.15.11`
- `pymodbus==3.12.1`

## Dokumentation
Vorhandene Projektdokumentation:
- `README.md`
- `docs\BETRIEBSANLEITUNG.md`

Die Betriebsanleitung beschreibt:
- Zweck der Anwendung
- Grundprinzip
- Bedienung
- Excel-Dateiaufbau
- wichtige technische Hinweise

## Auslieferungsinhalt
Im Release-Ordner befinden sich:
- `FGH_MSC.exe`
- `README.md`
- `UEBERGABE_FGH_MSC.md`
- `Dokumentation\BETRIEBSANLEITUNG.md`
- `Beispieldateien\Beispiel_Testplan.xlsx`
- `Start_FGH_MSC.bat`
- benoetigte Laufzeitdateien im Unterordner `_internal`

## Hinweise fuer die Abgabe
Empfohlen zur Weitergabe:
1. Den kompletten Ordner `dist\Release_FGH_MSC`
2. Diese Uebergabedatei

Optional fuer Entwickler oder Nachpflege:
- kompletter Quellcodeordner
- virtuelle Umgebung nur bei Bedarf

## Kurzfazit
Das Projekt ist als lauffaehige Windows-GUI fuer Modbus-TCP-Testablaeufe vorbereitet. Fuer Endanwender reicht der Release-Ordner mit der EXE. Fuer technische Nacharbeit stehen Quellcode, Build-Skript und Dokumentation strukturiert im Projekt zur Verfuegung.

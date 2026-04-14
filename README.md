# FGH Modbus Sollwert Manager

Desktop-Anwendung zur Planung und Ausfuehrung von Modbus-TCP-Testablaeufen mit vier konfigurierbaren Kanaelen und Excel-basiertem Testplanformat.

Made by Yunus Sevgi.

## Anwenderfunktionen
- Verbindung zu einem Modbus-TCP-Geraet ueber IP-Adresse, Port und Slave-ID
- Vier Kanaele mit individuellem Startregister und Datentyp
- Globales Registerformat fuer alle Kanaele
- Testplan mit bis zu 40 Stufen
- Speichern und Laden als Excel-Datei (`.xlsx`)
- Ereignisprotokoll fuer Verbindung, Schreibvorgaenge und Fehler

## Start waehrend der Entwicklung
```powershell
.\.venv\Scripts\python.exe start_modbus_sollwert_manager.py
```

## EXE bauen
1. `pyinstaller` in der Projektumgebung installieren
2. Build-Skript ausfuehren:

```powershell
.\build_exe.ps1
```

Danach liegt der uebergabefaehige Release-Ordner unter:

```text
dist\Release_FGH_Modbus_Sollwert_Manager
```

## Dokumentation
- Anwenderdokumentation: [docs/BETRIEBSANLEITUNG.md](/abs/path/c:/Sevgi/projects/Modbus/docs/BETRIEBSANLEITUNG.md)

## Hinweise zur langfristigen Nutzung
- Die EXE kann ohne Python-Installation verwendet werden.
- Testplaene werden als `.xlsx` gespeichert und bleiben dadurch leicht archiviert und austauschbar.
- Ein Update ist nur bei neuen Funktionen, Fehlerkorrekturen oder geaenderten Anforderungen notwendig.
- Fuer den Betrieb sollte immer der komplette Release-Ordner weitergegeben und zusammengehalten werden.

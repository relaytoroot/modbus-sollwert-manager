# FGH Modbus Sollwert Manager

Made by Yunus Sevgi.

## Zweck
Der Modbus Sollwert Manager dient zum Planen, Speichern, Laden und Ausfuehren von Testablaeufen fuer Modbus-TCP-Geraete. Pro Teststufe koennen bis zu vier Sollwerte sowie eine Laufzeit vorgegeben werden.

## Funktionsumfang
- Verbindung zu einem Modbus-TCP-Geraet ueber IP-Adresse, Port und Slave-ID
- Konfiguration von vier Kanaelen mit Startregister und Datentyp
- Globales Registerformat fuer alle Kanaele
- Testplan mit bis zu 40 Stufen
- Speichern und Laden von Testplaenen als Excel-Datei (`.xlsx`)
- Ereignisprotokoll fuer Verbindungsstatus, Schreibvorgaenge und Fehler

## Grundprinzip
- `S1` bis `S4` sind vier Kanaele auf einem Geraet.
- Jeder Kanal besitzt ein eigenes Startregister und einen Datentyp.
- Pro Stufe werden die Felder `Sollwert 1` bis `Sollwert 4` den Kanaelen `S1` bis `S4` zugeordnet.
- Ein leeres Sollwertfeld bedeutet: dieser Kanal wird in der Stufe nicht beschrieben.
- Ein eingetragener Wert `0` bedeutet: der Wert `0` wird aktiv geschrieben.

## Bedienung
1. Verbindung einstellen:
   IP-Adresse, Port, Slave-ID und Registerformat passend zum Zielgeraet eintragen.
2. Kanaele zuordnen:
   Fuer `S1` bis `S4` Startregister und Datentyp gemaess Geraetedokumentation festlegen.
3. Testplan erfassen:
   Pro Stufe bei Bedarf die Zeile aktivieren, Sollwerte eintragen und eine Zeit in Sekunden angeben.
4. Speichern:
   Ueber `Speichern unter` wird der Testplan als Excel-Datei abgelegt.
5. Laden:
   Ueber `Laden` kann eine vorhandene Excel-Datei wieder eingelesen werden.
6. Ausfuehren:
   Erst `Verbinden`, danach `Start` verwenden.

## Excel-Datei
Die Anwendung verwendet drei Arbeitsblaetter:
- `Verbindung`
- `Kanaele`
- `Testplan`

Die Datei sollte nur mit diesen Blaettern weiterverwendet werden, damit das Laden stabil bleibt.

## Wichtige Hinweise
- Ueberlappende Registerbereiche zwischen Kanaelen vermeiden.
- Alle Kanaele verwenden dasselbe Registerformat.
- Mehrregister-Datentypen belegen mehrere aufeinanderfolgende Register:
  `Int16/UInt16 = 1`, `Int32/UInt32/Float32 = 2`, `Int64/UInt64/Float64 = 4`
- Die Sollwerte einer Stufe werden technisch schnell hintereinander, aber nicht exakt gleichzeitig geschrieben.

## Langfristige Nutzung
- Fuer den normalen Einsatz reicht der bereitgestellte EXE-Ordner.
- Testplaene bleiben als `.xlsx` unabhaengig von Python benutzbar.
- Ein Update der Anwendung ist nur noetig, wenn neue Funktionen, Fehlerkorrekturen oder geaenderte Geraeteanforderungen vorliegen.
- Ohne Aenderungsbedarf kann die EXE langfristig unveraendert weiter genutzt werden.
- Es wird empfohlen, den Release-Ordner vollstaendig zusammenzuhalten und nicht einzelne Dateien zu verschieben.

## Auslieferung
Im Release-Ordner befinden sich:
- `FGH_Modbus_Sollwert_Manager.exe`
- diese Anleitung
- eine Beispiel-Excel-Datei

## Support-Hinweise
Bei Problemen zuerst pruefen:
- korrekte IP-Adresse, Port und Slave-ID
- passendes Registerformat
- richtige Startregister und Datentypen
- keine leeren Pflichtfelder in aktiven Stufen

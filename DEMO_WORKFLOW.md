# Demo-Workflow fuer den Modbus Sollwert Manager

Diese Datei ist nur eine interne Git-Hilfe. In der App selbst wird keine Demo- oder Vorzeigeversion genannt.

## Demo-Staende

| Zweck | Branch | Tag |
| --- | --- | --- |
| Startpunkt zeigen | `demo/01-startpunkt` | `demo-01-startpunkt` |
| PT1-Register-Erweiterung zeigen | `demo/02-pt1-register` | `demo-02-pt1-register` |
| Fertiger Stand mit Design, Remote und Haken | `demo/03-fertig-mit-remote` | `demo-03-fertig-mit-remote` |
| Stabiler aktueller Stand | `main` | `v1.0-demo` |

## Schnell wechseln

Auf einen Demo-Stand wechseln:

```powershell
git switch demo/02-pt1-register
```

Zurueck zum aktuellen Stand:

```powershell
git switch main
```

Aktuellen Stand pruefen:

```powershell
git status
```

Historie anzeigen:

```powershell
git log --oneline --decorate --graph --all --max-count=10
```

## GUI aus dem aktuell ausgewaehlten Stand starten

```powershell
.\.venv\Scripts\python.exe start_modbus_sollwert_manager.py
```

Wichtig: Die EXE bleibt immer der zuletzt gebaute Stand. Fuer eine Demo eines alten Git-Standes am besten aus Source starten oder nach dem Wechsel neu bauen.

## EXE aus dem aktuell ausgewaehlten Stand bauen

```powershell
powershell -ExecutionPolicy Bypass -File .\build_exe.ps1
```

## Typischer Ablauf fuer eine Praesentation

1. `git switch demo/01-startpunkt`
2. GUI starten und den frueheren Stand zeigen.
3. `git switch demo/02-pt1-register`
4. GUI starten und erklaeren, was dazugekommen ist.
5. `git switch demo/03-fertig-mit-remote`
6. GUI starten und den fertigen Stand zeigen.
7. Am Ende immer wieder `git switch main`.
